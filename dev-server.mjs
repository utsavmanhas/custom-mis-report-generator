// Local dev server: Express + Vite middleware + the two MIS API routes.
// Run with `npm run dev`. No Vercel CLI needed.

import express from "express";
import { createServer as createViteServer } from "vite";
import dotenv from "dotenv";
import Anthropic from "@anthropic-ai/sdk";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { downloadShortcutResult, getShortcutStatus, startShortcutJob, verifyShortcut, XLSX_MIME } from "./api/shortcutCore.mjs";

dotenv.config({ path: ".env.local" });
dotenv.config();

const app = express();
app.use(express.json({ limit: "20mb" }));

const designLanguage = readFileSync(
  join(process.cwd(), "prompts", "design-language.md"),
  "utf-8",
);

const ANALYZE_SYSTEM = `You are an experienced Indian Chartered Accountant building
a client-facing Management Information System (MIS) workbook from a Trial Balance.

This is PHASE 1 of a two-phase pipeline. In this phase you DO NOT build the
workbook. You analyse the data and tell the orchestrator what to ask the user.

Your task:
1. Detect the business type from account names.
2. Detect operating segments (faculties, departments, projects, SKUs, plants,
   product lines) - whatever the data implies.
3. Propose a sheet roster for the workbook (just names + 1-line purposes).
4. Propose 1-3 allocation logic OPTIONS for shared costs - the user picks.
5. List 6-12 clarifying questions you genuinely need answered. Be specific.
   Don't ask things you can already infer. Ask only what changes the output.
6. Build the plan as a complete MIS pack, not only a P&L.

Rules:
- Do not ask for company name or reporting period if profile already contains
  businessName, periodStart and periodEnd from the trial balance.
- Include Balance Sheet, Tax & Statutory, Fund Flow with next-period projection,
  ChartData, QC Tie-Outs, source schedules and Notes.
- Treat debtor balances as receivables/advances only; credit-balance debtors are
  customer advances/liabilities, not client revenue.
- Use bank statement / bank ledger sources for fund flow when supplied.

Return ONE JSON object inside <output>...</output> tags. No prose outside.

Schema:
interface AnalyzeOutput {
  businessType: string;
  businessTypeReasoning: string;
  detectedSegments: Array<{
    id: string; name: string; kind: string; rationale: string;
    revenueAccountIds: string[]; directCostAccountIds: string[];
  }>;
  allocationLogicOptions: Array<{
    id: string; label: string; description: string;
    appliesTo: string[]; recommended: boolean;
  }>;
  sheetPlan: Array<{ name: string; purpose: string; tabColor?: string; isDrillDown: boolean }>;
  clarifyingQuestions: Array<{
    id: string; section: string; prompt: string; reason: string;
    answerKind: "text"|"number"|"choice"; choices?: string[];
    priority: "high"|"medium"|"low";
  }>;
  unclassifiedAccounts: string[];
}

DESIGN LANGUAGE:
${designLanguage}

The first sheet is always Dashboard. The second is always the consolidated
Income & Expenditure / P&L. After that, one sheet per segment, then Balance
Sheet, Tax & Statutory, Fund Flow, ChartData, schedules, QC Tie-Outs, BRS,
then Notes.`;

const BUILD_SYSTEM = `You are an experienced Indian CA building a SINGLE sheet
of an MIS workbook. The orchestrator has already done Phase 1.
You now build one sheet only.

Return ONE JSON object inside <output>...</output> tags. No prose outside.

Schema:
interface SheetOutput {
  sheet: {
    name: string; tabColor?: string; columnWidths?: number[];
    merges?: string[]; freezePanes?: string;
    rows: Array<Array<{
      value?: string|number|null; formula?: string;
      format?: {
        bold?: boolean; italic?: boolean; fontSize?: number;
        fill?: string; color?: string; align?: "left"|"center"|"right";
        indent?: number;
        numberFormat?: "currency"|"percent"|"number"|"integer"|"text"|"date";
        border?: "thin"|"thick"|"double"|"none";
      };
    }>>;
  };
  notes: string[];
}

DESIGN LANGUAGE:
${designLanguage}

EXTRA RULES per-sheet:
A. If isDrillDown=true, include a BACK cell top-right (red fill C00000, white text, bold).
B. Use cross-sheet formulas with EXACT sheet names. Wrap multi-word names in single quotes.
C. Header rows 1-2: business name (navy fill, white bold) + period.
D. Indian number format for money, 0.0% for percentages.
E. For I&E sheet: segment columns in pairs (amount + %).
F. Use Excel formulas, not pre-computed values, wherever possible.
G. Fund Flow sheets must use bankSources when supplied, otherwise clearly mark
the result provisional and anchor only to TB bank/FD closing balances.
H. Never use "Sample Business" when source metadata has a real company name.`;

const MODEL = process.env.ANTHROPIC_MODEL || "claude-opus-4-7";

function callClaude(systemPrompt, userMessage, maxTokens = 12000) {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) throw new Error("ANTHROPIC_API_KEY missing");
  const client = new Anthropic({ apiKey });
  return client.messages.create({
    model: MODEL,
    max_tokens: maxTokens,
    system: systemPrompt,
    messages: [{ role: "user", content: userMessage }],
  });
}

function extractOutput(response) {
  const text = response.content
    .filter((b) => b.type === "text")
    .map((b) => b.text)
    .join("\n");
  const m = text.match(/<output>([\s\S]*?)<\/output>/);
  if (!m) return { error: "no <output> block", rawPreview: text.slice(0, 500) };
  try {
    return { result: JSON.parse(m[1].trim()) };
  } catch (e) {
    return { error: "invalid JSON", detail: String(e), rawPreview: m[1].slice(0, 500) };
  }
}

app.post("/api/analyze", async (req, res) => {
  try {
    const { profile, trialBalance, bankSources } = req.body || {};
    if (!Array.isArray(trialBalance) || !trialBalance.length) {
      return res.status(400).json({ error: "trialBalance required" });
    }
    const userMsg = [
      "BUSINESS PROFILE:", JSON.stringify(profile, null, 2), "",
      `TRIAL BALANCE (${trialBalance.length} rows):`, JSON.stringify(trialBalance, null, 2),
      "", `BANK / FUND FLOW SOURCES (${bankSources?.length || 0} files):`, JSON.stringify(bankSources || [], null, 2),
      "", "Now produce the Phase 1 JSON described in the system prompt.",
    ].join("\n");
    console.log(`[analyze] calling ${MODEL}, ${trialBalance.length} TB rows`);
    const response = await callClaude(ANALYZE_SYSTEM, userMsg, 8000);
    const out = extractOutput(response);
    if (out.error) return res.status(502).json(out);
    console.log(`[analyze] OK - ${response.usage.input_tokens} in / ${response.usage.output_tokens} out`);
    res.json({ ok: true, model: MODEL, usage: response.usage, result: out.result });
  } catch (err) {
    console.error("[analyze]", err);
    res.status(500).json({ error: err.message, status: err.status });
  }
});

app.post("/api/build-sheet", async (req, res) => {
  try {
    const { profile, trialBalance, answers, segments, allocationChoice, bankSources, sheet, allSheetNames } = req.body || {};
    if (!sheet?.name) return res.status(400).json({ error: "sheet.name required" });
    if (!Array.isArray(trialBalance) || !trialBalance.length) {
      return res.status(400).json({ error: "trialBalance required" });
    }
    const userMsg = [
      "BUSINESS PROFILE:", JSON.stringify(profile, null, 2), "",
      `TRIAL BALANCE (${trialBalance.length} rows):`, JSON.stringify(trialBalance, null, 2), "",
      "DETECTED SEGMENTS:", JSON.stringify(segments, null, 2), "",
      "ALLOCATION CHOICE:", JSON.stringify(allocationChoice, null, 2), "",
      "USER ANSWERS:", JSON.stringify(answers, null, 2), "",
      "BANK / FUND FLOW SOURCES:", JSON.stringify(bankSources || [], null, 2), "",
      "ALL SHEET NAMES (for cross-refs):", JSON.stringify(allSheetNames, null, 2), "",
      `BUILD THIS SHEET: ${JSON.stringify(sheet, null, 2)}`, "",
      "Output the JSON described in the system prompt.",
    ].join("\n");
    console.log(`[build-sheet] '${sheet.name}'`);
    const response = await callClaude(BUILD_SYSTEM, userMsg, 12000);
    const out = extractOutput(response);
    if (out.error) return res.status(502).json(out);
    console.log(`[build-sheet] '${sheet.name}' OK - ${response.usage.input_tokens}/${response.usage.output_tokens}`);
    res.json({ ok: true, model: MODEL, usage: response.usage, result: out.result });
  } catch (err) {
    console.error("[build-sheet]", err);
    res.status(500).json({ error: err.message, status: err.status });
  }
});


// ============================================================
// Shortcut.ai integration (alternative workbook builder)
// ============================================================
// User flow: Phase 1 (Claude) gathers clarifying questions and a sheet plan.
// Phase 2 can be either /api/build-sheet (Claude per-sheet) OR this endpoint,
// which submits one big prompt to Shortcut.ai and gets back the .xlsx.

const SHORTCUT_API_BASE = "https://api.shortcut.ai/api/spreadsheets";
const SHORTCUT_POLL_MS = 4000;
const SHORTCUT_MAX_WAIT_MS = 25 * 60 * 1000;

function shortcutPrompt(payload) {
  const tbLen = Array.isArray(payload.trialBalance) ? payload.trialBalance.length : 0;
  return [
    "Build a client-facing Management Information System (MIS) workbook in Excel from the data below.",
    "",
    "## Business profile", "```json", JSON.stringify(payload.profile, null, 2), "```", "",
    `## Trial balance (${tbLen} rows)`, "```json", JSON.stringify(payload.trialBalance, null, 2), "```", "",
    "## Detected segments", "```json", JSON.stringify(payload.segments, null, 2), "```", "",
    "## Cost-allocation logic chosen", "```json", JSON.stringify(payload.allocationChoice, null, 2), "```", "",
    "## Answers to clarifying questions", "```json", JSON.stringify(payload.answers, null, 2), "```", "",
    "## Required sheet roster (build in this order)", "```json", JSON.stringify(payload.sheetPlan, null, 2), "```", "",
    "## Design requirements",
    "- Sheet 1: Dashboard with KPI tiles + headline narrative",
    "- Sheet 2: Income & Expenditure account, segment-wise, with paired (amount, %) columns per segment",
    "- One drill-down sheet per detected segment",
    "- Supporting schedules where applicable: salary breakdown, sundry creditors, BRS, other receipts",
    "- Notes & Adjustments sheet at the end",
    "- Indian number formatting (#,##,##0). Currency from profile.",
    "- Live Excel formulas, never pre-computed values",
    "- Tab colours: red C00000 for I&E, orange FFC000 for cash/bank, green 00B050 for salary, navy 1F3864 for dashboard",
    "- Cross-sheet formulas using exact sheet names; multi-word names in single quotes",
    "",
    "Produce the workbook now.",
  ].join("\n");
}

app.post("/api/shortcut-build", async (req, res) => {
  const apiKey = process.env.SHORTCUT_API_KEY;
  if (!apiKey) return res.status(500).json({
    error: "SHORTCUT_API_KEY missing.",
    hint: "Add SHORTCUT_API_KEY=sc-... to .env.local and restart this server."
  });
  const body = req.body || {};
  if (!Array.isArray(body.trialBalance) || !body.trialBalance.length) {
    return res.status(400).json({ error: "trialBalance required" });
  }
  const prompt = shortcutPrompt(body);
  try {
    console.log(`[shortcut] submitting job, ${body.trialBalance.length} TB rows`);
    const submit = await fetch(SHORTCUT_API_BASE, {
      method: "POST",
      headers: { Authorization: `Bearer ${apiKey}`, "Content-Type": "application/json" },
      body: JSON.stringify({ prompt }),
    });
    if (!submit.ok) {
      const text = await submit.text();
      return res.status(502).json({ error: `Shortcut submit HTTP ${submit.status}`, detail: text.slice(0, 500) });
    }
    const sj = await submit.json();
    const runId = sj.runId || sj.id;
    if (!runId) return res.status(502).json({ error: "Shortcut response missing runId", detail: JSON.stringify(sj).slice(0, 500) });
    console.log(`[shortcut] runId=${runId}, polling...`);

    const start = Date.now();
    let lastStatus = "submitted";
    while (Date.now() - start < SHORTCUT_MAX_WAIT_MS) {
      await new Promise((r) => setTimeout(r, SHORTCUT_POLL_MS));
      const sr = await fetch(`${SHORTCUT_API_BASE}/${runId}`, { headers: { Authorization: `Bearer ${apiKey}` } });
      if (!sr.ok) {
        const text = await sr.text();
        return res.status(502).json({ error: `Shortcut status HTTP ${sr.status}`, detail: text.slice(0, 500) });
      }
      const stj = await sr.json();
      const status = (stj.status || stj.state || "unknown").toLowerCase();
      lastStatus = status;
      console.log(`[shortcut] status=${status}`);
      if (["done", "complete", "completed", "success", "succeeded"].includes(status)) break;
      if (["failed", "error", "errored", "cancelled", "canceled"].includes(status)) {
        return res.status(502).json({ error: `Shortcut job ${status}`, detail: JSON.stringify(stj).slice(0, 500) });
      }
    }
    if (!["done", "complete", "completed", "success", "succeeded"].includes(lastStatus)) {
      return res.status(504).json({ error: "Shortcut job timed out", lastStatus });
    }

    console.log(`[shortcut] downloading ${runId}`);
    const dl = await fetch(`${SHORTCUT_API_BASE}/${runId}/download`, { headers: { Authorization: `Bearer ${apiKey}` } });
    if (!dl.ok) {
      const text = await dl.text();
      return res.status(502).json({ error: `Shortcut download HTTP ${dl.status}`, detail: text.slice(0, 500) });
    }
    const buf = Buffer.from(await dl.arrayBuffer());
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="mis-shortcut.xlsx"');
    res.setHeader("X-Shortcut-Run-Id", runId);
    res.status(200).send(buf);
  } catch (err) {
    console.error("[shortcut]", err);
    res.status(500).json({ error: err.message });
  }
});

app.get("/api/workbook-build/verify", async (_req, res) => {
  try {
    res.json(await verifyShortcut(process.env));
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message || "Workbook build verify failed" });
  }
});

app.post("/api/workbook-build/start", async (req, res) => {
  try {
    res.json(await startShortcutJob(req.body || {}, process.env));
  } catch (err) {
    const status = /missing/i.test(err.message || "") ? 500 : 400;
    res.status(status).json({ ok: false, error: err.message || "Workbook build start failed" });
  }
});

app.get("/api/workbook-build/status", async (req, res) => {
  try {
    res.json(await getShortcutStatus(String(req.query.runId || ""), process.env));
  } catch (err) {
    res.status(400).json({ ok: false, error: err.message || "Workbook build status failed" });
  }
});

app.get("/api/workbook-build/download", async (req, res) => {
  try {
    const buffer = await downloadShortcutResult(String(req.query.runId || ""), process.env);
    res.setHeader("Content-Type", XLSX_MIME);
    res.setHeader("Content-Disposition", 'attachment; filename="mis-shortcut.xlsx"');
    res.send(buffer);
  } catch (err) {
    res.status(400).json({ ok: false, error: err.message || "Workbook build download failed" });
  }
});

// Vite middleware for the React frontend.
const vite = await createViteServer({
  server: { middlewareMode: true },
  appType: "spa",
});
app.use(vite.middlewares);

const port = process.env.PORT || 3000;
app.listen(port, "127.0.0.1", () => {
  console.log("");
  console.log("==================================================");
  console.log(`  MIS Generator dev server`);
  console.log(`  Open: http://localhost:${port}`);
  console.log(`  Model: ${MODEL}`);
  console.log(`  Anthropic key: ${process.env.ANTHROPIC_API_KEY ? "configured" : "MISSING (Claude flow disabled)"}`);
  console.log(`  Shortcut key:  ${process.env.SHORTCUT_API_KEY  ? "configured" : "missing (Shortcut flow disabled)"}`);
  console.log("==================================================");
  console.log("");
});
