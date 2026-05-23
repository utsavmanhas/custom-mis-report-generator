// PHASE 2: Build ONE sheet of the workbook.
// Called once per sheet in the plan, can be parallelised on the client.
// Returns a ClaudeSheetSpec ready for the renderer.

import Anthropic from "@anthropic-ai/sdk";
import type { VercelRequest, VercelResponse } from "@vercel/node";
import { readFileSync } from "node:fs";
import { join } from "node:path";

const designLanguage = readFileSync(
  join(process.cwd(), "prompts", "design-language.md"),
  "utf-8",
);

const SYSTEM_PROMPT = `You are an experienced Indian CA building a SINGLE sheet
of an MIS workbook. The orchestrator has already done Phase 1 (analysis +
clarifying questions). You now build one sheet only.

Inputs you will receive:
  - the full Trial Balance
  - business profile
  - user's answers to clarifying questions
  - optional bank statement / bank ledger summaries and transactions
  - segments and allocation choices already locked in
  - the SHEET you are building right now (name + purpose + isDrillDown)
  - the LIST of other sheet names so you can write cross-references correctly

Return ONE JSON object inside <output>...</output> tags. No prose outside.

Schema:

interface SheetOutput {
  sheet: {
    name: string;            // exact name from input
    tabColor?: string;       // hex w/o #
    columnWidths?: number[];
    merges?: string[];       // ["A1:H1"]
    freezePanes?: string;    // "B5" etc
    rows: Array<Array<{
      value?: string|number|null;
      formula?: string;       // without leading =
      format?: {
        bold?: boolean;
        italic?: boolean;
        fontSize?: number;
        fill?: string;        // hex
        color?: string;       // hex
        align?: "left"|"center"|"right";
        indent?: number;      // 0..4
        numberFormat?: "currency"|"percent"|"number"|"integer"|"text"|"date";
        border?: "thin"|"thick"|"double"|"none";
      };
    }>>;
  };
  notes: string[];           // anything you want recorded in the Notes sheet
}

DESIGN LANGUAGE you must follow:

${designLanguage}

EXTRA RULES specific to per-sheet generation:

A. If isDrillDown is true, include a BACK cell at the top-right:
   { value: "BACK", format: { bold: true, color: "FFFFFF", fill: "C00000",
     align: "center", border: "thin" } }

B. Use cross-sheet formulas. Reference other sheets by their EXACT name from
   the list provided. Wrap multi-word sheet names in single quotes.

C. Always include a header block at the top:
   row 1: business name (bold, navy fill, white text, merged across columns)
   row 2: report period
   row 3: blank

D. Below the header, the actual data block.

E. Use Indian number format for money. Use 0.0% for percentages.

F. For the I&E sheet specifically: profit-center columns must come in pairs
   (amount + %), and the formulas in the % column must be amount/segment_revenue.

G. If you don't have the data, write a placeholder cell with format.italic = true
   and value = "(Insert from supporting schedule)" rather than a wrong number.

H. 30-80 rows is a good target depth for most sheets. Detail-heavy schedules
   (salary, creditors) can go to 200+.

I. Prefer complete, useful workbook structure over excessive row volume. For
   wide sheets like I&E/P&L, keep the row count compact enough to finish inside
   one API response: major revenue lines, major cost lines, subtotal lines,
   margin lines, variance/commentary rows, and placeholders for deep schedules.

J. Balance Sheet sheets must show liabilities first, then assets, with nested
   schedules for capital, GST/TDS/PF/ESIC payable, creditors, provisions,
   reimbursements, fixed assets, debtors, bank/FD, ITC, advance tax, prepaids
   and suspense/control ledgers.

K. Fund Flow sheets must use bank statement / bank ledger data when supplied.
   If no bank source is supplied, clearly mark the sheet provisional and use
   only TB-derived closing bank/FD as the anchor. Include next-period projection
   rows based on the selected historical basis plus placeholders for bonus,
   tax, capex, loans, advances and collection assumptions.

L. Do not ask for or restate unknown company name / period when profile already
   has businessName, periodStart and periodEnd. Do not use "Sample Business" in
   client output.
`;

// Cost optimisation: Sonnet is 5× cheaper than Opus and produces identical
// structured JSON output for sheet specs. Override via ANTHROPIC_MODEL_BUILD.
const MODEL_BUILD = process.env.ANTHROPIC_MODEL_BUILD?.trim()
  || process.env.ANTHROPIC_MODEL?.trim()
  || "claude-sonnet-4-6";

/**
 * Cost optimisation: filter the Trial Balance to only rows relevant for this
 * specific sheet. Dashboard and I&E need the full TB; segment and schedule
 * sheets only need their own accounts plus a per-group summary of everything
 * else. Saves ~60-80% of TB input tokens for most sheets.
 */
function getRelevantTBRows(
  trialBalance: any[],
  segments: any[],
  sheet: { name: string; purpose?: string },
): any[] {
  const name = sheet.name.toLowerCase();

  // Global sheets need the full TB
  if (
    name.includes("dashboard") ||
    name.includes("p&l") ||
    name.includes("i&e") ||
    name.includes("income") ||
    name.includes("expenditure") ||
    name.includes("consolidated") ||
    name.includes("brs") ||
    name.includes("balance") ||
    name.includes("tax") ||
    name.includes("statutory") ||
    name.includes("fund") ||
    name.includes("cash") ||
    name.includes("chart") ||
    name.includes("qc") ||
    name.includes("notes")
  ) {
    return trialBalance;
  }

  // Collect account IDs that belong to any segment whose name/id appears in the sheet name
  const relevantIds = new Set<string>();
  for (const seg of segments ?? []) {
    const segName = (seg.name ?? "").toLowerCase();
    const segId = (seg.id ?? "").toLowerCase();
    if (name.includes(segName) || (segId.length > 2 && name.includes(segId))) {
      (seg.revenueAccountIds ?? []).forEach((id: string) => relevantIds.add(id));
      (seg.directCostAccountIds ?? []).forEach((id: string) => relevantIds.add(id));
    }
  }

  // No segment matched → return full TB as safe fallback
  if (relevantIds.size === 0) return trialBalance;

  const matched = trialBalance.filter((row) => relevantIds.has(row.id));
  const unmatched = trialBalance.filter((row) => !relevantIds.has(row.id));

  // Summarise unmatched rows by accountGroup so the model retains context
  // without paying full token cost for every individual row
  const groupTotals: Record<string, { debit: number; credit: number; count: number }> = {};
  for (const row of unmatched) {
    const g = row.accountGroup ?? "Other";
    if (!groupTotals[g]) groupTotals[g] = { debit: 0, credit: 0, count: 0 };
    groupTotals[g].debit += row.debit ?? 0;
    groupTotals[g].credit += row.credit ?? 0;
    groupTotals[g].count++;
  }
  const summaryRows = Object.entries(groupTotals).map(([group, t]) => ({
    id: `_summary_${group.replace(/\s+/g, "_")}`,
    accountName: `[${t.count} accounts – ${group} group total]`,
    accountGroup: group,
    debit: t.debit,
    credit: t.credit,
  }));

  return [...matched, ...summaryRows];
}

function extractJsonFromOutput(text: string) {
  const tagged = text.match(/<output>([\s\S]*?)<\/output>/);
  if (tagged) return tagged[1].trim();

  const afterOpenTag = text.match(/<output>\s*([\s\S]*)/);
  const source = (afterOpenTag ? afterOpenTag[1] : text).replace(/<\/output>\s*$/i, "").trim();
  const start = source.indexOf("{");
  if (start < 0) return "";

  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let i = start; i < source.length; i++) {
    const char = source[i];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (char === "\\") {
      escaped = true;
      continue;
    }
    if (char === "\"") {
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (char === "{") depth++;
    if (char === "}") {
      depth--;
      if (depth === 0) return source.slice(start, i + 1).trim();
    }
  }

  return source.slice(start).trim();
}

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== "POST") { res.status(405).json({ error: "Method not allowed" }); return; }
  const apiKey = process.env.ANTHROPIC_API_KEY?.trim();
  if (!apiKey) { res.status(500).json({ error: "ANTHROPIC_API_KEY missing" }); return; }

  let body: any;
  try {
    body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
  } catch {
    res.status(400).json({ error: "Invalid JSON body" }); return;
  }

  const { profile, trialBalance, answers, segments, allocationChoice, bankSources, sheet, allSheetNames, clientInstructions, clientBrief, referenceStructure } = body;
  if (!sheet?.name) { res.status(400).json({ error: "sheet.name is required" }); return; }
  if (!Array.isArray(trialBalance) || !trialBalance.length) { res.status(400).json({ error: "trialBalance required" }); return; }

  // Filter TB to only rows relevant for this sheet — saves 60-80% of TB input tokens
  // on segment/schedule sheets. Dashboard, I&E, BRS always get the full TB.
  const relevantTB = getRelevantTBRows(trialBalance, segments, sheet);

  // Split the user message into two blocks:
  //   Block 1 (cacheable): everything that is IDENTICAL across all sheet builds
  //            → profile, TB, segments, allocation choice, answers
  //   Block 2 (not cached): only what changes per call → sheet target + sheet list
  //
  // Anthropic charges 0.1× input price for cache-hit tokens after the first call,
  // so sheets 2-N are ~90% cheaper on the static portion.
  // Minified JSON saves another ~25% vs pretty-print.
  const cachedBlock = [
    "BUSINESS PROFILE:", JSON.stringify(profile), "",
    `TRIAL BALANCE (${relevantTB.length} rows, ${trialBalance.length} total):`,
    JSON.stringify(relevantTB), "",
    "DETECTED SEGMENTS:", JSON.stringify(segments), "",
    "ALLOCATION CHOICE:", JSON.stringify(allocationChoice), "",
    "USER ANSWERS TO CLARIFYING QUESTIONS:", JSON.stringify(answers),
    "",
    "BANK / FUND FLOW SOURCES:", JSON.stringify(bankSources || []),
    ...(clientBrief?.trim()
      ? ["", "## Client brief (from accountant)", clientBrief.trim()]
      : []),
    ...(referenceStructure
      ? ["", "## Reference MIS structure (client's preferred format)", JSON.stringify(referenceStructure)]
      : []),
  ].join("\n");

  const sheetBlock = [
    "",
    "ALL SHEET NAMES IN WORKBOOK (for cross-references):", JSON.stringify(allSheetNames), "",
    `BUILD THIS SHEET NOW: ${JSON.stringify(sheet)}`,
    "",
    ...(clientInstructions?.trim()
      ? ["## Client instructions (apply to this sheet)", clientInstructions.trim(), ""]
      : []),
    "Output the JSON described in the system prompt.",
  ].join("\n");

  const client = new Anthropic({ apiKey });
  try {
    const response = await client.messages.create({
      model: MODEL_BUILD,
      // 16 000 tokens is sufficient for the largest sheets (salary schedules, deep P&L).
      // 50 000 was wasteful: Opus at max output is ~$3.75/sheet on its own.
      max_tokens: 16000,
      // Cache the system prompt — it is identical for every sheet in this workbook build.
      system: [
        { type: "text", text: SYSTEM_PROMPT, cache_control: { type: "ephemeral" } },
      ] as any,
      messages: [{
        role: "user",
        content: [
          // Cache the static context block — same TB/segments/answers for every sheet
          { type: "text", text: cachedBlock, cache_control: { type: "ephemeral" } } as any,
          // Per-sheet instruction — not cached, changes every call
          { type: "text", text: sheetBlock },
        ],
      }],
    } as any);
    const text = response.content.filter(b => b.type === "text").map(b => (b as any).text).join("\n");
    const outputJson = extractJsonFromOutput(text);
    if (!outputJson) {
      res.status(502).json({ error: "no output JSON found", rawPreview: text.slice(0, 500) });
      return;
    }
    let parsed;
    try { parsed = JSON.parse(outputJson); }
    catch (e) {
      res.status(502).json({
        error: "invalid or incomplete JSON",
        detail: String(e),
        hint: "The model started a sheet response but did not return parseable JSON. Try the build again, or reduce the sheet depth.",
        rawPreview: outputJson.slice(0, 500),
        rawTail: outputJson.slice(-500),
      });
      return;
    }
    res.status(200).json({ ok: true, model: MODEL_BUILD, usage: response.usage, result: parsed });
  } catch (err: any) {
    res.status(err.status || 500).json({
      error: "anthropic call failed",
      detail: err.message,
      status: err.status,
      type: err.name,
      model: MODEL_BUILD,
    });
  }
}
