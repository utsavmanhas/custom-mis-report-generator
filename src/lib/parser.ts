import * as XLSX from "xlsx";
import { classifyAccount } from "./classification";
import type {
  BankMonthlySummary,
  BankSourceFile,
  BankTransaction,
  ParsedTrialBalance,
  ReferenceStructure,
  TrialBalanceGroupRow,
  TrialBalanceMetadata,
  TrialBalanceRow,
} from "../types";

export async function parseReferenceStructure(file: File): Promise<ReferenceStructure> {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: "array" });
  return {
    sheets: wb.SheetNames.map((name) => {
      const ws = wb.Sheets[name];
      const aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: "" });
      const headers = (aoa[0] ?? []).map((c) => String(c ?? "").trim()).filter(Boolean);
      const sampleRows = aoa.slice(1, 4).map((row) =>
        (row as unknown[]).map((c) => String(c ?? "").trim())
      );
      return { name, headers, sampleRows, rowCount: Math.max(0, aoa.length - 1) };
    }),
  };
}

const accountHeaders = ["account", "ledger", "particular", "particulars", "name", "description"];
const groupHeaders = ["group", "parent", "account group", "ledger group"];
const debitHeaders = ["debit", "dr", "withdrawal", "withdrawals", "payment", "payments", "paid"];
const creditHeaders = ["credit", "cr", "deposit", "deposits", "receipt", "receipts", "received"];
const balanceHeaders = ["balance", "closing", "closing balance", "net", "amount"];

const HEADER_SCAN_DEPTH = 30;
const HEADER_LOOKBACK = 6;

const monthIndex: Record<string, number> = {
  jan: 0,
  january: 0,
  feb: 1,
  february: 1,
  mar: 2,
  march: 2,
  apr: 3,
  april: 3,
  may: 4,
  jun: 5,
  june: 5,
  jul: 6,
  july: 6,
  aug: 7,
  august: 7,
  sep: 8,
  sept: 8,
  september: 8,
  oct: 9,
  october: 9,
  nov: 10,
  november: 10,
  dec: 11,
  december: 11,
};

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

function isActualHeaderCell(value: string) {
  return (
    ["particular", "particulars", "ledger", "ledger name", "account name", "description", "debit", "dr", "credit", "cr", "closing balance", "balance", "amount", "account group", "group"].includes(value) ||
    /^closing\s+balance$/.test(value) ||
    /^debit\s+amount$/.test(value) ||
    /^credit\s+amount$/.test(value)
  );
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

function locateColumns(aoa: unknown[][]) {
  const top = aoa.slice(0, HEADER_SCAN_DEPTH);
  const ncols = top.reduce((max, row) => Math.max(max, row.length), 0);
  let accountIdx = -1;
  let debitIdx = -1;
  let creditIdx = -1;
  let balanceIdx = -1;
  let groupIdx = -1;
  let lastHeaderRow = -1;

  for (let column = 0; column < ncols; column += 1) {
    const colText = top.map((row) => normalizeHeader(row[column])).filter(Boolean).join(" | ");
    if (!colText) continue;

    if (accountIdx < 0 && headerMatches(colText, accountHeaders)) accountIdx = column;
    if (debitIdx < 0 && /\bdebit\b|\bdr\b|\bwithdrawal|\bpayment/.test(colText)) debitIdx = column;
    if (creditIdx < 0 && /\bcredit\b|\bcr\b|\bdeposit|\breceipt/.test(colText)) creditIdx = column;
    if (balanceIdx < 0 && /\bbalance\b|\bclosing\b|\bnet\b|\bamount\b/.test(colText)) balanceIdx = column;
    if (groupIdx < 0 && /\bgroup\b|\bparent\b/.test(colText)) groupIdx = column;

    for (let row = 0; row < top.length; row += 1) {
      const value = normalizeHeader(top[row][column]);
      if (!value) continue;
      if (isActualHeaderCell(value)) {
        if (row > lastHeaderRow) lastHeaderRow = row;
      }
    }
  }

  if (accountIdx < 0) return null;
  if (debitIdx < 0 && creditIdx < 0 && balanceIdx < 0) return null;
  if (lastHeaderRow < 0) return null;

  return { accountIdx, debitIdx, creditIdx, balanceIdx, groupIdx, lastHeaderRow };
}

function parseDatePart(value: string) {
  const match = value.trim().match(/^(\d{1,2})[-/\s]([A-Za-z]{3,9})[-/\s](\d{2,4})$/);
  if (!match) return "";
  const day = Number.parseInt(match[1], 10);
  const month = monthIndex[match[2].toLowerCase()];
  let year = Number.parseInt(match[3], 10);
  if (year < 100) year += 2000;
  if (!Number.isFinite(day) || month == null || !Number.isFinite(year)) return "";
  const date = new Date(Date.UTC(year, month, day));
  return date.toISOString().slice(0, 10);
}

function parsePeriodText(value: string) {
  const match = value.match(/(\d{1,2}[-/\s][A-Za-z]{3,9}[-/\s]\d{2,4})\s*(?:to|-|through)\s*(\d{1,2}[-/\s][A-Za-z]{3,9}[-/\s]\d{2,4})/i);
  if (!match) return { periodText: "", periodStart: "", periodEnd: "" };
  return {
    periodText: match[0],
    periodStart: parseDatePart(match[1]),
    periodEnd: parseDatePart(match[2]),
  };
}

function extractMetadata(aoa: unknown[][], fileName: string): TrialBalanceMetadata {
  const topRows = aoa.slice(0, 12).map((row) => row.map((cell) => String(cell ?? "").trim()).filter(Boolean));
  const flattened = topRows.flat();
  const period = flattened.map(parsePeriodText).find((item) => item.periodStart && item.periodEnd) || { periodText: "", periodStart: "", periodEnd: "" };
  const companyName =
    flattened.find((value) => {
      const normalized = value.toLowerCase();
      return Boolean(
        value &&
          !/trial balance|particulars|closing balance|debit|credit|^\d{1,2}[-/\s]/i.test(value) &&
          !normalized.includes("to ") &&
          value.length > 2,
      );
    }) || "";

  return {
    companyName,
    periodText: period.periodText,
    periodStart: period.periodStart,
    periodEnd: period.periodEnd,
    totalDebit: 0,
    totalCredit: 0,
    sourceFileName: fileName,
  };
}

function groupLevel(accountName: string) {
  const name = accountName.trim().toLowerCase();
  if (/^(capital account|current liabilities|fixed assets|current assets|indirect incomes|indirect expenses|direct incomes|direct expenses|sales accounts|purchase accounts|loans \(liability\)|loan liabilities)$/.test(name)) {
    return 1;
  }
  if (/^(duties\s*&\s*taxes|provisions?|sundry creditors|reimbursement account|deposits? \(?asset\)?|sundry debtors|bank accounts|cash-in-hand|loans\s*&\s*advances|stock-in-hand|secured loans|unsecured loans)$/.test(name)) {
    return 2;
  }
  return 0;
}

function shouldLeaveSubgroup(subgroup: string | undefined, accountName: string) {
  const group = String(subgroup || "").toLowerCase();
  const name = accountName.toLowerCase();
  if (group === "bank accounts") return !/\bbank\b|\bfd\b|fixed deposit|cash/.test(name);
  if (group === "sundry debtors") return /\bbank\b|\bfd\b|input tax|\bitc\b|advance tax|prepaid|preliminary|suspense/.test(name);
  if (group === "deposits (asset)" || group === "deposit (asset)") return !/deposit/.test(name);
  return false;
}

function makeRowId(sheetName: string, rowNumber: number, accountName: string) {
  return `${sheetName}-${rowNumber}-${accountName}`.replace(/\s+/g, "-").replace(/[^a-z0-9-_]/gi, "").toLowerCase();
}

function riskFlagsFor(row: Pick<TrialBalanceRow, "accountName" | "debit" | "credit">, allRows: Array<Pick<TrialBalanceRow, "accountName">>) {
  const flags: string[] = [];
  const normalized = row.accountName.toLowerCase().trim();
  const duplicateCount = allRows.filter((item) => item.accountName.toLowerCase().trim() === normalized).length;
  if (duplicateCount > 1) flags.push("Duplicate ledger name");
  if (/suspense/.test(normalized)) flags.push("Suspense/control ledger");
  if (row.debit > 0 && row.credit > 0) flags.push("Ledger has both debit and credit balance columns populated");
  return flags;
}

export async function parseTrialBalanceWorkbook(file: File): Promise<ParsedTrialBalance> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const rows: TrialBalanceRow[] = [];
  const groupRows: TrialBalanceGroupRow[] = [];
  const warnings: string[] = [];
  let metadata: TrialBalanceMetadata = {
    companyName: "",
    periodText: "",
    periodStart: "",
    periodEnd: "",
    totalDebit: 0,
    totalCredit: 0,
    sourceFileName: file.name,
  };

  workbook.SheetNames.forEach((sheetName, sheetIndex) => {
    const sheet = workbook.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
    if (sheetIndex === 0) metadata = extractMetadata(aoa, file.name);

    const located = locateColumns(aoa);
    if (!located) return;
    const { accountIdx, debitIdx, creditIdx, balanceIdx, groupIdx, lastHeaderRow } = located;

    const headerRow = aoa[lastHeaderRow] ?? [];
    const originalHeaders: string[] = [];
    const ncols = aoa.reduce((max, row) => Math.max(max, row.length), 0);
    for (let column = 0; column < ncols; column += 1) {
      originalHeaders.push(String(headerRow[column] ?? "").trim() || `Column ${column + 1}`);
    }

    const pathStack: string[] = [];
    aoa.slice(lastHeaderRow + 1).forEach((row, rowOffset) => {
      const sourceRowNumber = lastHeaderRow + rowOffset + 2;
      const accountName = String(row[accountIdx] ?? "").trim();
      if (!accountName) return;

      let debit = debitIdx >= 0 ? Math.abs(parseNumber(row[debitIdx])) : 0;
      let credit = creditIdx >= 0 ? Math.abs(parseNumber(row[creditIdx])) : 0;
      if (!debit && !credit && balanceIdx >= 0) {
        const balance = parseNumber(row[balanceIdx]);
        if (balance >= 0) debit = Math.abs(balance);
        if (balance < 0) credit = Math.abs(balance);
      }

      if (/^grand total$/i.test(accountName)) {
        metadata.totalDebit = debit;
        metadata.totalCredit = credit;
        groupRows.push({
          sourceSheet: sheetName,
          sourceRowNumber,
          accountName,
          accountPath: ["Grand Total"],
          debit,
          credit,
          rowType: "total",
        });
        return;
      }

      const level = groupLevel(accountName);
      if (level > 0) {
        pathStack.splice(level - 1);
        pathStack[level - 1] = accountName;
        groupRows.push({
          sourceSheet: sheetName,
          sourceRowNumber,
          accountName,
          accountPath: pathStack.slice(0, level),
          debit,
          credit,
          rowType: "group",
        });
        return;
      }

      if (!debit && !credit) return;

      if (pathStack.length > 1 && shouldLeaveSubgroup(pathStack[1], accountName)) {
        pathStack.splice(1);
      }

      const explicitGroup = groupIdx >= 0 ? String(row[groupIdx] ?? "").trim() : "";
      const accountPath = [...pathStack, accountName].filter(Boolean);
      const accountGroup = explicitGroup || pathStack.join(" > ");
      const classification = classifyAccount(accountName, accountGroup);

      rows.push({
        id: makeRowId(sheetName, sourceRowNumber, accountName),
        sourceSheet: sheetName,
        sourceRowNumber,
        accountName,
        accountGroup,
        accountPath,
        hierarchyLevel: Math.max(1, accountPath.length),
        debit,
        credit,
        balance: debit - credit,
        category: classification.category,
        misGroup: classification.misGroup,
        profitCenterId: "",
        allocationBase: "revenue",
        confidence: classification.confidence,
        riskFlags: [],
        raw: toRawRecord(originalHeaders, row),
      });
    });
  });

  rows.forEach((row) => {
    row.riskFlags = riskFlagsFor(row, rows);
  });

  if (!metadata.companyName) warnings.push("Company name was not detected in the trial balance header.");
  if (!metadata.periodStart || !metadata.periodEnd) warnings.push("Reporting period was not detected in the trial balance header.");
  if (!metadata.totalDebit && !metadata.totalCredit) {
    metadata.totalDebit = rows.reduce((sum, row) => sum + row.debit, 0);
    metadata.totalCredit = rows.reduce((sum, row) => sum + row.credit, 0);
  }

  return { rows, metadata, groupRows, warnings };
}

export async function parseTrialBalanceFile(file: File): Promise<TrialBalanceRow[]> {
  return (await parseTrialBalanceWorkbook(file)).rows;
}

function locateBankColumns(aoa: unknown[][]) {
  const top = aoa.slice(0, HEADER_SCAN_DEPTH);
  const ncols = top.reduce((max, row) => Math.max(max, row.length), 0);
  let dateIdx = -1;
  let descriptionIdx = -1;
  let debitIdx = -1;
  let creditIdx = -1;
  let amountIdx = -1;
  let balanceIdx = -1;
  let lastHeaderRow = -1;

  for (let column = 0; column < ncols; column += 1) {
    const colText = top.map((row) => normalizeHeader(row[column])).filter(Boolean).join(" | ");
    if (!colText) continue;
    if (dateIdx < 0 && /\b(date|txn date|transaction date|value date|voucher date|vch)\b/.test(colText)) dateIdx = column;
    if (descriptionIdx < 0 && /\b(narration|description|particulars|remarks|ledger|account|details)\b/.test(colText)) descriptionIdx = column;
    if (debitIdx < 0 && /\b(debit|withdrawal|payment|paid|dr)\b/.test(colText)) debitIdx = column;
    if (creditIdx < 0 && /\b(credit|deposit|receipt|received|cr)\b/.test(colText)) creditIdx = column;
    if (amountIdx < 0 && /\bamount\b/.test(colText)) amountIdx = column;
    if (balanceIdx < 0 && /\bbalance\b|\bclosing\b/.test(colText)) balanceIdx = column;

    for (let row = 0; row < top.length; row += 1) {
      const value = normalizeHeader(top[row][column]);
      if (value && /\b(date|narration|description|particulars|debit|credit|amount|balance|withdrawal|deposit|receipt|payment)\b/.test(value)) {
        lastHeaderRow = Math.max(lastHeaderRow, row);
      }
    }
  }

  if (dateIdx < 0 && descriptionIdx < 0) return null;
  if (debitIdx < 0 && creditIdx < 0 && amountIdx < 0) return null;
  return { dateIdx, descriptionIdx, debitIdx, creditIdx, amountIdx, balanceIdx, lastHeaderRow: Math.max(lastHeaderRow, 0) };
}

function parseDate(value: unknown) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d)).toISOString().slice(0, 10);
  }
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const periodDate = parseDatePart(raw);
  if (periodDate) return periodDate;
  // DD-Mon-YYYY used by Tally exports (e.g. "1-Apr-2024", "01-Apr-2024")
  const monName = raw.match(/^(\d{1,2})[-\s]([A-Za-z]{3,})[-\s](\d{4})$/);
  if (monName) {
    const months: Record<string, number> = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };
    const m = months[monName[2].toLowerCase().slice(0, 3)];
    if (m !== undefined) {
      const d = new Date(Date.UTC(Number.parseInt(monName[3], 10), m, Number.parseInt(monName[1], 10)));
      if (!Number.isNaN(d.getTime())) return d.toISOString().slice(0, 10);
    }
  }
  const normal = raw.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/);
  if (normal) {
    let year = Number.parseInt(normal[3], 10);
    if (year < 100) year += 2000;
    const day = Number.parseInt(normal[1], 10);
    const month = Number.parseInt(normal[2], 10) - 1;
    const date = new Date(Date.UTC(year, month, day));
    if (!Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  }
  const fallback = new Date(raw);
  return Number.isNaN(fallback.getTime()) ? "" : fallback.toISOString().slice(0, 10);
}

function monthKey(date: string) {
  return date ? date.slice(0, 7) : "Un dated";
}

function summariseTransactions(transactions: BankTransaction[]): BankSourceFile["summary"] {
  const sorted = [...transactions].sort((a, b) => (a.date || "").localeCompare(b.date || "") || a.sourceRowNumber - b.sourceRowNumber);
  const monthly = new Map<string, BankMonthlySummary>();
  sorted.forEach((transaction) => {
    const key = monthKey(transaction.date);
    const current = monthly.get(key) || {
      month: key,
      receipts: 0,
      payments: 0,
      openingBalance: null,
      closingBalance: null,
      transactionCount: 0,
    };
    current.receipts += transaction.credit;
    current.payments += transaction.debit;
    current.transactionCount += 1;
    if (current.openingBalance == null && transaction.balance != null) {
      current.openingBalance = transaction.balance - transaction.credit + transaction.debit;
    }
    if (transaction.balance != null) current.closingBalance = transaction.balance;
    monthly.set(key, current);
  });

  return {
    totalReceipts: transactions.reduce((sum, transaction) => sum + transaction.credit, 0),
    totalPayments: transactions.reduce((sum, transaction) => sum + transaction.debit, 0),
    openingBalance: sorted.find((transaction) => transaction.balance != null)?.balance ?? null,
    closingBalance: [...sorted].reverse().find((transaction) => transaction.balance != null)?.balance ?? null,
    monthly: Array.from(monthly.values()).sort((a, b) => a.month.localeCompare(b.month)),
    warnings: [
      ...(transactions.length ? [] : ["No bank transactions were detected from this file."]),
      ...(transactions.length && transactions.every((t) => t.balance == null)
        ? ["Balance column was not detected — opening/closing balance will not be available."]
        : []),
    ],
  };
}

export async function parseBankSourceFile(file: File, sourceType: BankSourceFile["sourceType"]): Promise<BankSourceFile> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sourceFileId = makeRowId(sourceType, Date.now(), file.name);
  const transactions: BankTransaction[] = [];
  const warnings: string[] = [];

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
    const located = locateBankColumns(aoa);
    if (!located) return;

    const headerRow = aoa[located.lastHeaderRow] ?? [];
    const ncols = aoa.reduce((max, row) => Math.max(max, row.length), 0);
    const originalHeaders = Array.from({ length: ncols }, (_, index) => String(headerRow[index] ?? "").trim() || `Column ${index + 1}`);

    aoa.slice(located.lastHeaderRow + 1).forEach((row, offset) => {
      const sourceRowNumber = located.lastHeaderRow + offset + 2;
      const date = located.dateIdx >= 0 ? parseDate(row[located.dateIdx]) : "";
      const description = located.descriptionIdx >= 0 ? String(row[located.descriptionIdx] ?? "").trim() : "";
      let debit = located.debitIdx >= 0 ? Math.abs(parseNumber(row[located.debitIdx])) : 0;
      let credit = located.creditIdx >= 0 ? Math.abs(parseNumber(row[located.creditIdx])) : 0;

      if (!debit && !credit && located.amountIdx >= 0) {
        const amount = parseNumber(row[located.amountIdx]);
        if (amount < 0) debit = Math.abs(amount);
        if (amount > 0) credit = Math.abs(amount);
      }

      const balance = located.balanceIdx >= 0 ? parseNumber(row[located.balanceIdx]) : 0;
      if (!date && !description && !debit && !credit) return;
      if (!debit && !credit) return;

      transactions.push({
        id: makeRowId(sheetName, sourceRowNumber, description || "bank-row"),
        sourceFileId,
        sourceType,
        sourceSheet: sheetName,
        sourceRowNumber,
        date,
        description,
        debit,
        credit,
        balance: Number.isFinite(balance) && balance !== 0 ? balance : null,
        raw: toRawRecord(originalHeaders, row),
      });
    });
  });

  if (!transactions.length) warnings.push("No dated bank rows with debit/credit amounts were detected.");
  const summary = summariseTransactions(transactions);
  summary.warnings.push(...warnings);

  return {
    id: sourceFileId,
    sourceType,
    fileName: file.name,
    rowsImported: transactions.length,
    transactions,
    summary,
  };
}

export const _internal = {
  accountHeaders,
  groupHeaders,
  debitHeaders,
  creditHeaders,
  balanceHeaders,
  HEADER_LOOKBACK,
  parsePeriodText,
  groupLevel,
};
