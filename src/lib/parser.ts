import * as XLSX from "xlsx";
import { classifyAccount } from "./classification";
import type { TrialBalanceRow } from "../types";

const accountHeaders = ["account", "ledger", "particular", "name", "description"];
const groupHeaders = ["group", "parent", "account group", "ledger group"];
const debitHeaders = ["debit", "dr", "debit amount", "debits"];
const creditHeaders = ["credit", "cr", "credit amount", "credits"];
const balanceHeaders = ["balance", "closing", "closing balance", "net", "amount"];

function normalizeHeader(value: unknown) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[().:_-]/g, " ")
    .trim();
}

function headerMatches(value: string, candidates: string[]) {
  return candidates.some((candidate) => value === candidate || value.includes(candidate));
}

function findColumn(headers: string[], candidates: string[]) {
  return headers.findIndex((header) => headerMatches(header, candidates));
}

function headerScore(row: unknown[]) {
  const headers = row.map(normalizeHeader);
  let score = 0;
  if (findColumn(headers, accountHeaders) >= 0) score += 4;
  if (findColumn(headers, debitHeaders) >= 0) score += 2;
  if (findColumn(headers, creditHeaders) >= 0) score += 2;
  if (findColumn(headers, balanceHeaders) >= 0) score += 1;
  if (findColumn(headers, groupHeaders) >= 0) score += 1;
  return score;
}

function parseNumber(value: unknown) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (value == null) return 0;

  const raw = String(value).trim();
  if (!raw) return 0;

  const isCredit = /\bcr\b/i.test(raw);
  const isDebit = /\bdr\b/i.test(raw);
  const hasParens = /^\(.*\)$/.test(raw);
  const cleaned = raw
    .replace(/\b(cr|dr)\b/gi, "")
    .replace(/[₹$€£,\s]/g, "")
    .replace(/[()]/g, "");

  const parsed = Number.parseFloat(cleaned);
  if (!Number.isFinite(parsed)) return 0;

  if (isCredit) return -Math.abs(parsed);
  if (isDebit) return Math.abs(parsed);
  if (hasParens) return -Math.abs(parsed);
  return parsed;
}

function toRawRecord(headers: string[], row: unknown[]) {
  return headers.reduce<Record<string, string | number | null>>((acc, header, index) => {
    acc[header || `Column ${index + 1}`] = row[index] as string | number | null;
    return acc;
  }, {});
}

export async function parseTrialBalanceFile(file: File): Promise<TrialBalanceRow[]> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const rows: TrialBalanceRow[] = [];

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
    const headerIndex = aoa.slice(0, 25).reduce(
      (best, row, index) => {
        const score = headerScore(row);
        return score > best.score ? { index, score } : best;
      },
      { index: -1, score: 0 },
    ).index;

    if (headerIndex < 0) return;

    const originalHeaders = aoa[headerIndex].map((header, index) => String(header || `Column ${index + 1}`).trim());
    const normalizedHeaders = originalHeaders.map(normalizeHeader);
    const accountIndex = findColumn(normalizedHeaders, accountHeaders);
    const groupIndex = findColumn(normalizedHeaders, groupHeaders);
    const debitIndex = findColumn(normalizedHeaders, debitHeaders);
    const creditIndex = findColumn(normalizedHeaders, creditHeaders);
    const balanceIndex = findColumn(normalizedHeaders, balanceHeaders);

    if (accountIndex < 0) return;

    aoa.slice(headerIndex + 1).forEach((row, rowIndex) => {
      const accountName = String(row[accountIndex] ?? "").trim();
      if (!accountName || /^total\b/i.test(accountName)) return;

      const accountGroup = groupIndex >= 0 ? String(row[groupIndex] ?? "").trim() : "";
      let debit = debitIndex >= 0 ? Math.abs(parseNumber(row[debitIndex])) : 0;
      let credit = creditIndex >= 0 ? Math.abs(parseNumber(row[creditIndex])) : 0;

      if (!debit && !credit && balanceIndex >= 0) {
        const balance = parseNumber(row[balanceIndex]);
        if (balance >= 0) debit = Math.abs(balance);
        if (balance < 0) credit = Math.abs(balance);
      }

      if (!debit && !credit) return;

      const classification = classifyAccount(accountName, accountGroup);

      rows.push({
        id: `${sheetName}-${rowIndex}-${accountName}`.replace(/\s+/g, "-").toLowerCase(),
        sourceSheet: sheetName,
        accountName,
        accountGroup,
        debit,
        credit,
        balance: debit - credit,
        category: classification.category,
        misGroup: classification.misGroup,
        profitCenterId: "",
        allocationBase: "revenue",
        confidence: classification.confidence,
        raw: toRawRecord(originalHeaders, row),
      });
    });
  });

  return rows;
}
