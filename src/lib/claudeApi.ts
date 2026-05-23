// Two-phase Claude pipeline client.
// Phase 1: /api/analyze       -> returns clarifying questions + sheet plan
// Phase 2: /api/build-sheet    -> called once per sheet, in parallel
// All API key handling stays on the server.

import type { BankSourceFile, BusinessProfile, ReferenceStructure, TrialBalanceRow } from "../types";

export interface AnalyzeResult {
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
  sheetPlan: Array<{
    name: string; purpose: string; tabColor?: string; isDrillDown: boolean;
  }>;
  clarifyingQuestions: Array<{
    id: string; section: string; prompt: string; reason: string;
    answerKind: "text" | "number" | "choice"; choices?: string[];
    priority: "high" | "medium" | "low";
  }>;
  clientPreferenceQuestions: Array<{
    id: string; question: string;
    answerKind: "text" | "choice"; choices?: string[];
    defaultAnswer?: string;
  }>;
  unclassifiedAccounts: string[];
}

export interface SheetSpec {
  name: string;
  tabColor?: string;
  columnWidths?: number[];
  merges?: string[];
  freezePanes?: string;
  rows: Array<Array<{
    value?: string | number | null;
    formula?: string;
    format?: {
      bold?: boolean; italic?: boolean; fontSize?: number;
      fill?: string; color?: string;
      align?: "left" | "center" | "right";
      indent?: number;
      numberFormat?: "currency" | "percent" | "number" | "integer" | "text" | "date";
      border?: "thin" | "thick" | "double" | "none";
    };
  }>>;
}

export interface BuildSheetResult {
  sheet: SheetSpec;
  notes: string[];
}

interface ApiResponse<T> {
  ok: true;
  model: string;
  usage?: { input_tokens: number; output_tokens: number };
  result: T;
}

interface ApiError {
  error: string;
  detail?: string;
  hint?: string;
  rawPreview?: string;
  status?: number;
}

async function postJson<T>(url: string, body: unknown): Promise<ApiResponse<T>> {
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const e = (await res.json().catch(() => ({}))) as ApiError;
    const parts = [e.error || `HTTP ${res.status}`];
    if (e.hint) parts.push(`Hint: ${e.hint}`);
    if (e.detail) parts.push(e.detail);
    if (e.rawPreview) parts.push(`Preview: ${e.rawPreview.slice(0, 200)}`);
    throw new Error(parts.join(" - "));
  }
  return (await res.json()) as ApiResponse<T>;
}

function trimTb(rows: TrialBalanceRow[]) {
  return rows.map((r) => ({
    id: r.id,
    accountName: r.accountName,
    accountGroup: r.accountGroup,
    accountPath: r.accountPath,
    debit: r.debit,
    credit: r.credit,
    category: r.category,
    misGroup: r.misGroup,
    riskFlags: r.riskFlags,
  }));
}

function trimBankSources(bankSources: BankSourceFile[] = []) {
  return bankSources.map((source) => ({
    id: source.id,
    sourceType: source.sourceType,
    fileName: source.fileName,
    rowsImported: source.rowsImported,
    summary: source.summary,
    transactions: source.transactions.slice(0, 500).map((transaction) => ({
      id: transaction.id,
      date: transaction.date,
      description: transaction.description,
      debit: transaction.debit,
      credit: transaction.credit,
      balance: transaction.balance,
      sourceSheet: transaction.sourceSheet,
      sourceRowNumber: transaction.sourceRowNumber,
    })),
  }));
}

export async function analyzeWithClaude(input: {
  profile: BusinessProfile;
  rows: TrialBalanceRow[];
  bankSources?: BankSourceFile[];
  clientBrief?: string;
  referenceStructure?: ReferenceStructure | null;
}) {
  return postJson<AnalyzeResult>("/api/analyze", {
    profile: input.profile,
    trialBalance: trimTb(input.rows),
    bankSources: trimBankSources(input.bankSources),
    clientBrief: input.clientBrief || undefined,
    referenceStructure: input.referenceStructure || undefined,
  });
}

export async function buildSheetWithClaude(input: {
  profile: BusinessProfile;
  rows: TrialBalanceRow[];
  segments: AnalyzeResult["detectedSegments"];
  allocationChoice: AnalyzeResult["allocationLogicOptions"][number] | null;
  answers: Array<{ id: string; question: string; answer: string }>;
  bankSources?: BankSourceFile[];
  sheet: AnalyzeResult["sheetPlan"][number];
  allSheetNames: string[];
  clientInstructions?: string;
  clientBrief?: string;
  referenceStructure?: ReferenceStructure | null;
}) {
  return postJson<BuildSheetResult>("/api/build-sheet", {
    profile: input.profile,
    trialBalance: trimTb(input.rows),
    segments: input.segments,
    allocationChoice: input.allocationChoice,
    answers: input.answers,
    bankSources: trimBankSources(input.bankSources),
    sheet: input.sheet,
    allSheetNames: input.allSheetNames,
    clientInstructions: input.clientInstructions,
    clientBrief: input.clientBrief || undefined,
    referenceStructure: input.referenceStructure || undefined,
  });
}

export interface RevisePlanResult {
  acknowledgement: string;
  plan: AnalyzeResult;
}

export async function revisePlanWithClaude(input: {
  currentPlan: AnalyzeResult;
  userPrompt: string;
}) {
  return postJson<RevisePlanResult>("/api/revise-plan", {
    currentPlan: input.currentPlan,
    userPrompt: input.userPrompt,
  });
}



// ============================================================
// Shortcut.ai builder - alternative to per-sheet build with Claude
// ============================================================
//
// Sends one job to Shortcut with the same Phase-1 context Claude produced,
// and gets back the .xlsx as a Blob.  The server polls the job to completion;
// the browser sees a single fetch that returns binary on success.

export async function buildWorkbookWithShortcut(input: {
  profile: BusinessProfile;
  rows: TrialBalanceRow[];
  bankSources?: BankSourceFile[];
  segments: AnalyzeResult["detectedSegments"];
  allocationChoice: AnalyzeResult["allocationLogicOptions"][number] | null;
  answers: Array<{ id: string; question: string; answer: string }>;
  sheetPlan: AnalyzeResult["sheetPlan"];
}): Promise<Blob> {
  const trialBalance = input.rows.map((r) => ({
    id: r.id,
    accountName: r.accountName,
    accountGroup: r.accountGroup,
    debit: r.debit,
    credit: r.credit,
  }));
  const res = await fetch("/api/shortcut-build", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      profile: input.profile,
      trialBalance,
      segments: input.segments,
      allocationChoice: input.allocationChoice,
      answers: input.answers,
      sheetPlan: input.sheetPlan,
      bankSources: trimBankSources(input.bankSources),
    }),
  });
  if (!res.ok) {
    const e = (await res.json().catch(() => ({}))) as ApiError;
    const parts = [e.error || `HTTP ${res.status}`];
    if (e.hint) parts.push(`Hint: ${e.hint}`);
    if (e.detail) parts.push(e.detail);
    throw new Error(parts.join(" - "));
  }
  return await res.blob();
}

export function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.rel = "noopener";
  link.style.display = "none";
  document.body.appendChild(link);
  link.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, view: window }));
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}
