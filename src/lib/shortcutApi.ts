import type { BankSourceFile, BusinessProfile, ProfitCenter, QuestionAnswer, StaffMember, TrialBalanceRow } from "../types";
import type { AnalyzeResult } from "./claudeApi";

export interface ShortcutVerifyResult {
  ok: boolean;
  configured: boolean;
  valid?: boolean;
  message?: string;
  error?: string;
}

export interface ShortcutStartResult {
  ok: true;
  runId: string;
  status: "queued" | "running";
  filename: string;
  sourceFileId?: string;
}

export type ShortcutStatus =
  | { runId: string; status: "queued" | "running" }
  | { runId: string; status: "completed"; downloadUrl?: string; summary?: string; artifacts?: string[] }
  | { runId: string; status: "failed" | "error"; error: string };

interface ApiError {
  error?: string;
  message?: string;
}

async function readJson<T>(response: Response): Promise<T> {
  const data = (await response.json().catch(() => ({}))) as ApiError;
  if (!response.ok) {
    throw new Error(data.error || data.message || `HTTP ${response.status}`);
  }
  return data as T;
}

export async function verifyShortcut() {
  return readJson<ShortcutVerifyResult>(await fetch("/api/workbook-build/verify"));
}

export async function startShortcut(input: {
  profile: BusinessProfile;
  rows: TrialBalanceRow[];
  bankSources: BankSourceFile[];
  profitCenters: ProfitCenter[];
  staff: StaffMember[];
  answers: QuestionAnswer[];
  reportAnswers: Array<{ id: string; question: string; answer: string }>;
  analysis: AnalyzeResult | null;
}) {
  const response = await fetch("/api/workbook-build/start", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  });
  return readJson<ShortcutStartResult>(response);
}

export async function getShortcutStatus(runId: string) {
  return readJson<ShortcutStatus>(await fetch(`/api/workbook-build/status?runId=${encodeURIComponent(runId)}`));
}

export async function downloadShortcut(runId: string) {
  const response = await fetch(`/api/workbook-build/download?runId=${encodeURIComponent(runId)}`);
  if (!response.ok) {
    const data = (await response.json().catch(() => ({}))) as ApiError;
    throw new Error(data.error || data.message || `HTTP ${response.status}`);
  }
  return response.blob();
}
