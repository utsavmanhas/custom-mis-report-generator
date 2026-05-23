// PHASE 1: Analyse a Trial Balance, return clarifying questions + sheet plan.
// No workbook is built here - this is the cheap, fast pass that lets the user
// confirm Claude's interpretation before we spend tokens generating sheets.

import Anthropic from "@anthropic-ai/sdk";
import type { VercelRequest, VercelResponse } from "@vercel/node";
import { readFileSync } from "node:fs";
import { join } from "node:path";

const designLanguage = readFileSync(
  join(process.cwd(), "prompts", "design-language.md"),
  "utf-8",
);

const SYSTEM_PROMPT = `You are an experienced Indian Chartered Accountant building
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

Important source rules:
- If the profile contains businessName, periodStart, and periodEnd inferred from
  the trial balance, DO NOT ask for company name or reporting period again.
- Treat debtor balances as receivables/advances only. Do not allocate client
  revenue from debtors unless an invoice-wise/client-wise revenue register is
  supplied.
- Credit balances in debtor ledgers must be presented as customer advances or
  liabilities, not as ordinary debtor revenue.
- Fund flow should be bank-statement or bank-ledger based when those files are
  supplied. If not supplied, mark fund flow provisional and ask for the bank
  source plus projection assumptions.
- The sheet roster must include Dashboard, P&L / Income & Expenditure, Balance
  Sheet, Tax & Statutory, Fund Flow with next-period projection, ChartData,
  source schedules, QC Tie-Outs, and Notes / Questions.

Return ONE JSON object inside <output>...</output> tags. No prose outside.

Schema:

interface AnalyzeOutput {
  businessType: string;
  businessTypeReasoning: string;
  detectedSegments: Array<{
    id: string;          // slug
    name: string;
    kind: string;        // department | project | product | channel | etc.
    rationale: string;
    revenueAccountIds: string[];   // matching TB row ids
    directCostAccountIds: string[];
  }>;
  allocationLogicOptions: Array<{
    id: string;
    label: string;       // "by student count", "by FTE", "by revenue"
    description: string;
    appliesTo: string[]; // ["shared opex", "salary", "rent"]
    recommended: boolean;
  }>;
  sheetPlan: Array<{
    name: string;
    purpose: string;
    tabColor?: string;     // hex w/o #
    isDrillDown: boolean;
  }>;
  clarifyingQuestions: Array<{
    id: string;
    section: string;
    prompt: string;
    reason: string;
    answerKind: "text" | "number" | "choice";
    choices?: string[];   // when answerKind=="choice"
    priority: "high" | "medium" | "low";
  }>;
  clientPreferenceQuestions: Array<{
    id: string;
    question: string;       // presentation preference, not accounting data
    answerKind: "text" | "choice";
    choices?: string[];
    defaultAnswer?: string;
  }>;
  unclassifiedAccounts: string[];  // TB rowIds you cannot place
}

clientPreferenceQuestions must focus on HOW the client wants to see the output, not WHAT the data is.
Good examples: column layout (monthly vs consolidated), number scale (lakhs/crores), segment orientation
(side-by-side vs separate sheets), KPIs on dashboard, narrative commentary yes/no.
Generate 3-5 of these. If a client brief or reference MIS was provided that already answers a
preference question, omit it.

If a clientBrief is supplied, honour it — it overrides design-language defaults where they conflict.
If a referenceStructure is supplied, mirror its sheet order and column structure as closely as possible.

DESIGN LANGUAGE you must follow when proposing the sheet plan:

${designLanguage}

The first sheet is always Dashboard. The second is always the consolidated
Income & Expenditure / P&L. After that, one sheet per segment, then
Balance Sheet, Tax & Statutory, Fund Flow, ChartData, schedules, QC Tie-Outs,
BRS, then Notes.
`;

// Cost optimisation: Sonnet is 5× cheaper than Opus and handles Phase 1 (classification)
// perfectly well. Override via ANTHROPIC_MODEL_ANALYZE env var if needed.
const MODEL_ANALYZE = process.env.ANTHROPIC_MODEL_ANALYZE?.trim()
  || process.env.ANTHROPIC_MODEL?.trim()
  || "claude-sonnet-4-6";
const ANTHROPIC_VERSION = "2023-06-01";

interface IncomingRequest {
  profile: Record<string, unknown>;
  trialBalance: Array<{
    id: string;
    accountName: string;
    accountGroup: string;
    debit: number;
    credit: number;
  }>;
  bankSources?: unknown[];
  clientBrief?: string;
  referenceStructure?: unknown;
}

export default async function handler(req: VercelRequest, res: VercelResponse) {
  const apiKey = process.env.ANTHROPIC_API_KEY?.trim();

  if (req.method === "GET") {
    if (!apiKey) {
      res.status(200).json({ ok: true, configured: false, model: MODEL_ANALYZE });
      return;
    }
    try {
      const response = await fetch("https://api.anthropic.com/v1/models?limit=20", {
        headers: {
          "x-api-key": apiKey,
          "anthropic-version": ANTHROPIC_VERSION,
        },
      });
      const text = await response.text();
      let data: any = {};
      try {
        data = text ? JSON.parse(text) : {};
      } catch {
        data = { message: text };
      }
      if (!response.ok) {
        res.status(response.status).json({
          ok: false,
          configured: true,
          valid: false,
          model: MODEL_ANALYZE,
          status: response.status,
          error: data.error?.message || data.message || `Anthropic verify HTTP ${response.status}`,
        });
        return;
      }
      const models = Array.isArray(data.data) ? data.data.map((item: any) => item.id).filter(Boolean) : [];
      res.status(200).json({
        ok: true,
        configured: true,
        valid: true,
        model: MODEL_ANALYZE,
        modelAvailable: models.includes(MODEL_ANALYZE),
        models,
      });
      return;
    } catch (err: any) {
      res.status(500).json({
        ok: false,
        configured: true,
        valid: null,
        model: MODEL_ANALYZE,
        error: "Anthropic verification connection failed",
        detail: err.message,
        type: err.name,
      });
      return;
    }
  }

  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  if (!apiKey) {
    res.status(500).json({ error: "ANTHROPIC_API_KEY missing" });
    return;
  }

  let body: IncomingRequest;
  try {
    body = typeof req.body === "string" ? JSON.parse(req.body) : (req.body as IncomingRequest);
  } catch {
    res.status(400).json({ error: "Invalid JSON body" });
    return;
  }
  if (!body?.trialBalance?.length) {
    res.status(400).json({ error: "trialBalance is required" });
    return;
  }

  // Minified JSON: removes ~25% of tokens vs pretty-print with no loss of information
  const userMessage = [
    "BUSINESS PROFILE:", JSON.stringify(body.profile), "",
    `TRIAL BALANCE (${body.trialBalance.length} rows):`,
    JSON.stringify(body.trialBalance),
    "",
    `BANK / FUND FLOW SOURCES (${body.bankSources?.length || 0} files):`,
    JSON.stringify(body.bankSources || []),
    "",
    ...(body.clientBrief?.trim()
      ? ["## Client brief (from accountant)", body.clientBrief.trim(), ""]
      : []),
    ...(body.referenceStructure
      ? ["## Reference MIS structure (client's preferred format)", JSON.stringify(body.referenceStructure), ""]
      : []),
    "Now produce the Phase 1 JSON described in the system prompt.",
  ].join("\n");

  const client = new Anthropic({ apiKey });

  try {
    const response = await client.messages.create({
      model: MODEL_ANALYZE,
      max_tokens: 8000,
      // Prompt caching: system prompt + design language are static across all calls.
      // After the first call, these tokens cost 0.1× (90% cheaper on re-reads).
      system: [
        { type: "text", text: SYSTEM_PROMPT, cache_control: { type: "ephemeral" } },
      ] as any,
      messages: [{ role: "user", content: userMessage }],
    } as any);
    const text = response.content
      .filter((b) => b.type === "text")
      .map((b) => (b as any).text)
      .join("\n");
    const m = text.match(/<output>([\s\S]*?)<\/output>/);
    if (!m) {
      res.status(502).json({ error: "no <output> block", rawPreview: text.slice(0, 500) });
      return;
    }
    let parsed;
    try {
      parsed = JSON.parse(m[1].trim());
    } catch (e) {
      res.status(502).json({ error: "invalid JSON", detail: String(e), rawPreview: m[1].slice(0, 500) });
      return;
    }
    res.status(200).json({ ok: true, model: MODEL_ANALYZE, usage: response.usage, result: parsed });
  } catch (err: any) {
    res.status(err.status || 500).json({
      error: "anthropic call failed",
      detail: err.message,
      status: err.status,
      type: err.name,
      model: MODEL_ANALYZE,
    });
  }
}
