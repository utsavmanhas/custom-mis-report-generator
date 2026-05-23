// PHASE 1.5: Revise an existing AnalyzeResult based on user's plain-English prompt.
// Called when the user types changes in the Plan Review step and clicks "Revise Plan".
// Returns the same AnalyzeResult schema as Phase 1 plus a short acknowledgement message.

import Anthropic from "@anthropic-ai/sdk";
import type { VercelRequest, VercelResponse } from "@vercel/node";
import { readFileSync } from "node:fs";
import { join } from "node:path";

const designLanguage = readFileSync(
  join(process.cwd(), "prompts", "design-language.md"),
  "utf-8",
);

const SYSTEM_PROMPT = `You are an experienced Indian Chartered Accountant who has
already analysed a Trial Balance and proposed an MIS workbook plan (Phase 1).

The user has now reviewed the plan and typed changes or answers in plain English.
Your job is to revise the plan to honour their instructions exactly.

Rules:
- Only change what the user explicitly asked to change.
- If the user removes a sheet, remove it from sheetPlan.
- If the user adds a sheet, add it with a sensible purpose.
- If the user answers a clarifying question, remove that question from clarifyingQuestions.
- If the user changes reporting preferences (number scale, prior year, etc.), note them
  in a revised clarifyingQuestions entry or in the businessTypeReasoning field.
- Keep all other fields unchanged from the previous plan.

Return two things inside <output>...</output> tags — a JSON object with this schema:

{
  "acknowledgement": "One sentence confirming what you changed.",
  "plan": { ...revised AnalyzeResult here... }
}

The plan field must match the same AnalyzeResult schema:

{
  businessType: string,
  businessTypeReasoning: string,
  detectedSegments: Array<{ id, name, kind, rationale, revenueAccountIds, directCostAccountIds }>,
  allocationLogicOptions: Array<{ id, label, description, appliesTo, recommended }>,
  sheetPlan: Array<{ name, purpose, tabColor?, isDrillDown }>,
  clarifyingQuestions: Array<{ id, section, prompt, reason, answerKind, choices?, priority }>,
  unclassifiedAccounts: string[]
}

${designLanguage}
`;

const MODEL = process.env.ANTHROPIC_MODEL_ANALYZE?.trim()
  || process.env.ANTHROPIC_MODEL?.trim()
  || "claude-sonnet-4-6";

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== "POST") { res.status(405).json({ error: "Method not allowed" }); return; }
  const apiKey = process.env.ANTHROPIC_API_KEY?.trim();
  if (!apiKey) { res.status(500).json({ error: "ANTHROPIC_API_KEY missing" }); return; }

  let body: { currentPlan: unknown; userPrompt: string };
  try {
    body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
  } catch {
    res.status(400).json({ error: "Invalid JSON body" }); return;
  }
  if (!body?.currentPlan || !body?.userPrompt?.trim()) {
    res.status(400).json({ error: "currentPlan and userPrompt are required" }); return;
  }

  const userMessage = [
    "CURRENT PLAN:", JSON.stringify(body.currentPlan), "",
    "USER'S REQUESTED CHANGES:", body.userPrompt.trim(), "",
    "Revise the plan as instructed and return the JSON object.",
  ].join("\n");

  const client = new Anthropic({ apiKey });
  try {
    const response = await client.messages.create({
      model: MODEL,
      max_tokens: 6000,
      system: [{ type: "text", text: SYSTEM_PROMPT, cache_control: { type: "ephemeral" } }] as any,
      messages: [{ role: "user", content: userMessage }],
    } as any);

    const text = response.content.filter((b) => b.type === "text").map((b) => (b as any).text).join("\n");
    const match = text.match(/<output>([\s\S]*?)<\/output>/);
    if (!match) { res.status(502).json({ error: "No <output> block returned", rawPreview: text.slice(0, 400) }); return; }

    let parsed: { acknowledgement: string; plan: unknown };
    try { parsed = JSON.parse(match[1].trim()); }
    catch (e) { res.status(502).json({ error: "Invalid JSON", detail: String(e), rawPreview: match[1].slice(0, 400) }); return; }

    res.status(200).json({ ok: true, model: MODEL, usage: response.usage, result: parsed });
  } catch (err: any) {
    res.status(err.status || 500).json({ error: "Anthropic call failed", detail: err.message, model: MODEL });
  }
}
