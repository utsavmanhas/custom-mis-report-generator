// Vercel serverless function: /api/shortcut-build
// Submits a single workbook-build job to Shortcut.ai's spreadsheet API
// using the analysis context the wizard already gathered, polls the job
// to completion, and streams the resulting .xlsx back to the browser.

import type { VercelRequest, VercelResponse } from "@vercel/node";

const SHORTCUT_API_BASE = "https://api.shortcut.ai/api/spreadsheets";
const POLL_INTERVAL_MS = 4000;
const MAX_WAIT_MS = 25 * 60 * 1000;

type AnyObj = Record<string, unknown>;

function buildPrompt(payload: AnyObj): string {
  return [
    "Build a client-facing Management Information System (MIS) workbook in Excel from the data below.",
    "",
    "## Business profile", "```json", JSON.stringify(payload.profile, null, 2), "```", "",
    `## Trial balance (${(payload.trialBalance as unknown[])?.length ?? 0} rows)`,
    "```json", JSON.stringify(payload.trialBalance, null, 2), "```", "",
    "## Detected segments", "```json", JSON.stringify(payload.segments, null, 2), "```", "",
    "## Cost-allocation logic chosen", "```json", JSON.stringify(payload.allocationChoice, null, 2), "```", "",
    "## Answers to clarifying questions", "```json", JSON.stringify(payload.answers, null, 2), "```", "",
    "## Bank / fund flow sources", "```json", JSON.stringify(payload.bankSources || [], null, 2), "```", "",
    "## Required sheet roster (build in this order)", "```json", JSON.stringify(payload.sheetPlan, null, 2), "```", "",
    "## Design requirements",
    "- Sheet 1: Dashboard with KPI tiles + headline narrative",
    "- Sheet 2: Income & Expenditure account, segment-wise, with paired (amount, %) columns per segment",
    "- One drill-down sheet per detected segment",
    "- Supporting schedules where applicable: balance sheet, tax & statutory, fund flow with next-period projection, salary breakdown, sundry creditors, BRS, other receipts, ChartData, QC tie-outs",
    "- Use bank statement / bank ledger data for fund flow when supplied; otherwise mark fund flow provisional and only anchor to TB-derived bank/FD closing balances",
    "- Do not use Sample Business if the source metadata provides the real company name and period",
    "- Treat debtor balances as receivables/advances; debtor credit balances are customer advances/liabilities, not client revenue",
    "- Notes & Adjustments sheet at the end",
    "- Indian number formatting (#,##,##0). Currency from profile.",
    "- Live Excel formulas, never pre-computed values",
    "- Tab colours: red C00000 for I&E, orange FFC000 for cash/bank, green 00B050 for salary, navy 1F3864 for dashboard",
    "- Cross-sheet formulas using exact sheet names; multi-word names in single quotes",
    "",
    "Produce the workbook now.",
  ].join("\n");
}

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });
  const apiKey = process.env.SHORTCUT_API_KEY;
  if (!apiKey) return res.status(500).json({
    error: "SHORTCUT_API_KEY is not configured.",
    hint: "Add SHORTCUT_API_KEY=sc-... to .env.local and restart the dev server.",
  });

  let body: AnyObj;
  try { body = typeof req.body === "string" ? JSON.parse(req.body) : (req.body as AnyObj); }
  catch { return res.status(400).json({ error: "Invalid JSON body" }); }

  if (!Array.isArray(body.trialBalance) || !body.trialBalance.length) {
    return res.status(400).json({ error: "trialBalance is required" });
  }

  const prompt = buildPrompt(body);

  try {
    const submit = await fetch(SHORTCUT_API_BASE, {
      method: "POST",
      headers: { Authorization: `Bearer ${apiKey}`, "Content-Type": "application/json" },
      body: JSON.stringify({ prompt }),
    });
    if (!submit.ok) {
      const text = await submit.text();
      return res.status(502).json({ error: `Shortcut submit failed: HTTP ${submit.status}`, detail: text.slice(0, 500) });
    }
    const sj = (await submit.json()) as { runId?: string; id?: string };
    const runId = sj.runId || sj.id;
    if (!runId) return res.status(502).json({ error: "Shortcut response missing runId", detail: JSON.stringify(sj).slice(0, 500) });

    const start = Date.now();
    let lastStatus = "submitted";
    while (Date.now() - start < MAX_WAIT_MS) {
      await new Promise((r) => setTimeout(r, POLL_INTERVAL_MS));
      const sr = await fetch(`${SHORTCUT_API_BASE}/${runId}`, { headers: { Authorization: `Bearer ${apiKey}` } });
      if (!sr.ok) {
        const text = await sr.text();
        return res.status(502).json({ error: `Shortcut status HTTP ${sr.status}`, detail: text.slice(0, 500) });
      }
      const stj = (await sr.json()) as { status?: string; state?: string };
      const status = (stj.status || stj.state || "unknown").toLowerCase();
      lastStatus = status;
      if (["done", "complete", "completed", "success", "succeeded"].includes(status)) break;
      if (["failed", "error", "errored", "cancelled", "canceled"].includes(status)) {
        return res.status(502).json({ error: `Shortcut job ${status}`, detail: JSON.stringify(stj).slice(0, 500) });
      }
    }
    if (!["done", "complete", "completed", "success", "succeeded"].includes(lastStatus)) {
      return res.status(504).json({ error: "Shortcut job timed out", lastStatus });
    }

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
    const e = err as Error;
    res.status(500).json({ error: "Shortcut call failed", detail: e.message });
  }
}
