import * as XLSX from "xlsx";
import { classifyAccount } from "./classification";
import type { AccountCategory, BankSourceType, BankTransaction, ParsedBankStatement, ParsedTrialBalance, TrialBalanceMetadata, TrialBalanceRow } from "../types";

const accountHeaders = ["account", "ledger", "particular", "name", "description"];
const groupHeaders = ["group", "parent", "account group", "ledger group"];
const debitHeaders = ["debit", "dr", "debit amount", "debits", "withdrawal", "withdrawals", "payment", "payments", "paid", "debit amt"];
const creditHeaders = ["credit", "cr", "credit amount", "credits", "deposit", "deposits", "receipt", "receipts", "received", "credit amt"];
const balanceHeaders = ["balance", "closing", "closing balance", "net", "amount", "running balance"];
const dateHeaders = ["date", "transaction date", "value date", "posting date", "voucher date", "vch date"];
const narrationHeaders = ["narration", "description", "particular", "particulars", "remarks", "counterparty", "name", "ledger", "account"];
const amountHeaders = ["amount", "transaction amount", "net amount"];
const tallyGroupHeadings = new Set([
  "capital account",
  "current liabilities",
  "duties & taxes",
  "duties and taxes",
  "provisions",
  "sundry creditors",
  "reimbursement account",
  "fixed assets",
  "current assets",
  "deposits asset",
  "deposits (asset)",
  "sundry debtors",
  "bank accounts",
  "indirect incomes",
  "indirect expenses",
]);

const tallyGroupOverrides: Array<{
  groups: RegExp[];
  category: AccountCategory;
  misGroup: string;
  confidence: number;
}> = [
  {
    groups: [/capital account/],
    category: "equity",
    misGroup: "Equity",
    confidence: 0.95,
  },
  {
    groups: [/current liabilities/, /duties (?:&|and) taxes/, /provisions/, /sundry creditors/],
    category: "current-liability",
    misGroup: "Current Liabilities",
    confidence: 0.92,
  },
  {
    groups: [/fixed assets/],
    category: "fixed-asset",
    misGroup: "Fixed Assets",
    confidence: 0.92,
  },
  {
    groups: [/current assets/, /deposits asset/, /sundry debtors/, /bank accounts/],
    category: "current-asset",
    misGroup: "Current Assets",
    confidence: 0.9,
  },
  {
    groups: [/indirect incomes/],
    category: "revenue",
    misGroup: "Operating Revenue",
    confidence: 0.74,
  },
];

function normalizeHeader(value: unknown) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[().:_-]/g, " ")
    .trim();
}

function normalizeName(value: unknown) {
  return normalizeHeader(value).replace(/\s+/g, " ");
}

function titleCase(value: string) {
  return value.trim();
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

function parseDateValue(value: unknown) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number" && Number.isFinite(value)) {
    const date = XLSX.SSF.parse_date_code(value);
    if (date) return `${date.y.toString().padStart(4, "0")}-${date.m.toString().padStart(2, "0")}-${date.d.toString().padStart(2, "0")}`;
  }

  const raw = String(value ?? "").trim();
  const match = raw.match(/(\d{1,2})[-/\s]([A-Za-z]{3,}|\d{1,2})[-/\s](\d{2,4})/);
  if (!match) return "";

  const day = Number.parseInt(match[1], 10);
  const monthToken = match[2].toLowerCase();
  const months: Record<string, number> = {
    jan: 1,
    january: 1,
    feb: 2,
    february: 2,
    mar: 3,
    march: 3,
    apr: 4,
    april: 4,
    may: 5,
    jun: 6,
    june: 6,
    jul: 7,
    july: 7,
    aug: 8,
    august: 8,
    sep: 9,
    sept: 9,
    september: 9,
    oct: 10,
    october: 10,
    nov: 11,
    november: 11,
    dec: 12,
    december: 12,
  };
  const month = months[monthToken] || Number.parseInt(monthToken, 10);
  let year = Number.parseInt(match[3], 10);
  if (year < 100) year += year >= 70 ? 1900 : 2000;
  if (!day || !month || !year) return "";
  return `${year.toString().padStart(4, "0")}-${month.toString().padStart(2, "0")}-${day.toString().padStart(2, "0")}`;
}

function toRawRecord(headers: string[], row: unknown[]) {
  return headers.reduce<Record<string, string | number | null>>((acc, header, index) => {
    acc[header || `Column ${index + 1}`] = row[index] as string | number | null;
    return acc;
  }, {});
}

function isTallyGroupHeading(accountName: string) {
  return tallyGroupHeadings.has(normalizeName(accountName));
}

function isGrandTotal(accountName: string) {
  return /^grand total$/i.test(accountName.trim()) || /^total\b/i.test(accountName.trim());
}

function classifyTallyAccount(accountName: string, accountGroup: string) {
  const classification = classifyAccount(accountName, accountGroup);
  const group = normalizeName(accountGroup);
  const balanceSheetOverride = tallyGroupOverrides.find(
    (item) => item.category !== "revenue" && item.groups.some((pattern) => pattern.test(group)),
  );

  if (balanceSheetOverride) {
    if (classification.category === balanceSheetOverride.category && classification.misGroup !== balanceSheetOverride.misGroup) {
      return {
        category: classification.category,
        misGroup: classification.misGroup,
        confidence: Math.max(classification.confidence, balanceSheetOverride.confidence),
      };
    }

    return {
      category: balanceSheetOverride.category,
      misGroup: balanceSheetOverride.misGroup,
      confidence: Math.max(classification.confidence, balanceSheetOverride.confidence),
    };
  }

  if (classification.category !== "unknown" && !["tax", "people-cost"].includes(classification.category)) {
    return classification;
  }

  const override = tallyGroupOverrides.find((item) => item.groups.some((pattern) => pattern.test(group)));
  if (!override) return classification;

  if (override.category === "revenue" && classification.category === "other-income") {
    return classification;
  }

  return {
    category: override.category,
    misGroup: override.misGroup,
    confidence: Math.max(classification.confidence, override.confidence),
  };
}

function findTallyHeader(aoa: unknown[][]) {
  for (let rowIndex = 0; rowIndex < Math.min(25, aoa.length); rowIndex += 1) {
    const normalizedRow = aoa[rowIndex].map(normalizeHeader);
    const accountIndex = findColumn(normalizedRow, accountHeaders);
    if (accountIndex < 0) continue;

    for (let debitCreditRowIndex = rowIndex + 1; debitCreditRowIndex < Math.min(rowIndex + 5, aoa.length); debitCreditRowIndex += 1) {
      const debitCreditRow = aoa[debitCreditRowIndex].map(normalizeHeader);
      const debitIndex = findColumn(debitCreditRow, debitHeaders);
      const creditIndex = findColumn(debitCreditRow, creditHeaders);

      if (debitIndex >= 0 && creditIndex >= 0) {
        return {
          accountIndex,
          debitIndex,
          creditIndex,
          headerIndex: rowIndex,
          debitCreditRowIndex,
        };
      }
    }
  }

  return null;
}

function inferTrialBalanceMetadata(workbook: XLSX.WorkBook, fileName = ""): TrialBalanceMetadata {
  let businessName = "";
  let periodStart = "";
  let periodEnd = "";
  let sourceSheetName = workbook.SheetNames[0] || "";

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
    for (let index = 0; index < Math.min(10, aoa.length); index += 1) {
      const text = aoa[index].map((cell) => String(cell ?? "").trim()).filter(Boolean).join(" ");
      if (!businessName && text && !/trial balance|particulars|closing balance|debit|credit/i.test(text) && /[A-Za-z]/.test(text)) {
        businessName = titleCase(text);
        sourceSheetName = sheetName;
      }

      const period = text.match(/(\d{1,2}[-/\s][A-Za-z]{3,}[-/\s]\d{2,4})\s*(?:to|-)\s*(\d{1,2}[-/\s][A-Za-z]{3,}[-/\s]\d{2,4})/i);
      if (period) {
        periodStart = parseDateValue(period[1]);
        periodEnd = parseDateValue(period[2]);
      }
    }
    if (businessName || periodStart || periodEnd) break;
  }

  return {
    businessName,
    periodStart,
    periodEnd,
    sourceFileName: fileName,
    sourceSheetName,
    totalDebit: 0,
    totalCredit: 0,
  };
}

function parseTallyTrialBalanceSheet(sheetName: string, aoa: unknown[][]): TrialBalanceRow[] {
  const header = findTallyHeader(aoa);
  if (!header) return [];

  const rows: TrialBalanceRow[] = [];
  const accountPath: string[] = [];

  aoa.slice(header.debitCreditRowIndex + 1).forEach((row, index) => {
    const absoluteRowNumber = header.debitCreditRowIndex + index + 2;
    const accountName = String(row[header.accountIndex] ?? "").trim();
    if (!accountName || isGrandTotal(accountName)) return;

    const debit = Math.abs(parseNumber(row[header.debitIndex]));
    const credit = Math.abs(parseNumber(row[header.creditIndex]));
    if (!debit && !credit) return;

    if (isTallyGroupHeading(accountName)) {
      accountPath.push(accountName);
      if (accountPath.length > 2) accountPath.shift();
      return;
    }

    const accountGroup = accountPath.at(-1) || "";
    const classification = classifyTallyAccount(accountName, accountGroup);
    if (/sundry debtors/i.test(accountGroup) && credit > debit) {
      classification.category = "current-liability";
      classification.misGroup = "Customer Advances / Credit Debtors";
      classification.confidence = 0.96;
    }
    if (/sundry creditors/i.test(accountGroup) && debit > credit) {
      classification.category = "current-asset";
      classification.misGroup = "Supplier Advances / Debit Balance Creditors";
      classification.confidence = 0.94;
    }
    if (/reimbursement account/i.test(accountGroup)) {
      classification.category = debit >= credit ? "current-asset" : "current-liability";
      classification.misGroup = debit >= credit ? "Employee Advances / Reimbursements" : "Employee Payables / Reimbursements";
      classification.confidence = 0.9;
    }

    rows.push({
      id: `${sheetName}-${absoluteRowNumber}-${accountName}`.replace(/\s+/g, "-").toLowerCase(),
      sourceSheet: sheetName,
      accountName,
      accountGroup,
      accountPath: [...accountPath, accountName],
      debit,
      credit,
      balance: debit - credit,
      category: classification.category,
      misGroup: classification.misGroup,
      profitCenterId: "",
      allocationBase: "revenue",
      confidence: classification.confidence,
      raw: {
        Particulars: accountName,
        Debit: debit,
        Credit: credit,
        "Account Group": accountGroup,
        "Account Path": [...accountPath, accountName].join(" > "),
        "Source Row": absoluteRowNumber,
      },
    });
  });

  return rows;
}

export function parseTrialBalanceWorkbook(buffer: ArrayBuffer, fileName = ""): ParsedTrialBalance {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const rows: TrialBalanceRow[] = [];
  const metadata = inferTrialBalanceMetadata(workbook, fileName);
  const warnings: string[] = [];

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
    const tallyRows = parseTallyTrialBalanceSheet(sheetName, aoa);
    if (tallyRows.length) {
      rows.push(...tallyRows);
      return;
    }

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
        accountPath: accountGroup ? [accountGroup, accountName] : [accountName],
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

  const seen = new Map<string, number>();
  rows.forEach((row) => {
    const key = `${normalizeName(row.accountName)}|${row.debit}|${row.credit}`;
    seen.set(key, (seen.get(key) || 0) + 1);
  });
  const duplicateSuspense = rows.some((row) => /suspense/i.test(row.accountName) && (seen.get(`${normalizeName(row.accountName)}|${row.debit}|${row.credit}`) || 0) > 1);
  if (duplicateSuspense) warnings.push("Duplicate Suspense A/c balance detected. Confirm whether this is a genuine duplicate or two separate balances.");

  const emitted = new Set<string>();
  const calculationRows = rows.filter((row) => {
    const key = `${normalizeName(row.accountName)}|${row.debit}|${row.credit}`;
    if (/suspense/i.test(row.accountName) && emitted.has(key)) return false;
    emitted.add(key);
    return true;
  });

  metadata.totalDebit = calculationRows.reduce((sum, row) => sum + row.debit, 0);
  metadata.totalCredit = calculationRows.reduce((sum, row) => sum + row.credit, 0);

  return { rows: calculationRows, metadata, warnings };
}

export async function parseTrialBalanceFile(file: File): Promise<TrialBalanceRow[]> {
  const parsed = parseTrialBalanceWorkbook(await file.arrayBuffer(), file.name);
  return parsed.rows;
}

function classifyBankTransaction(narration: string, debit: number, credit: number) {
  const text = narration.toLowerCase();
  const isReceipt = credit > debit;
  const isInterAccountTransfer = /self|own account|internal transfer|inter bank|interbank|contra|fund transfer|transfer to|transfer from|upi.*self|neft.*self|rtgs.*self/.test(text);

  if (isInterAccountTransfer) {
    return {
      category: "Inter-bank transfer",
      fundFlowGroup: "Inter-bank transfers",
      isInterAccountTransfer: true,
    };
  }
  if (/gst|tds|tax|pf|epf|esi|esic|challan|income tax|professional tax/.test(text)) {
    return {
      category: "Tax / statutory",
      fundFlowGroup: "GST / TDS / statutory",
      isInterAccountTransfer: false,
    };
  }
  if (/salary|payroll|bonus|wages|stipend|employee|staff/.test(text)) {
    return {
      category: "People payments",
      fundFlowGroup: "Salary / bonus / payroll",
      isInterAccountTransfer: false,
    };
  }
  if (/fd|fixed deposit|sweep|liquid fund|treasury|interest/.test(text)) {
    return {
      category: "Treasury / FD",
      fundFlowGroup: isReceipt ? "FD maturity / interest receipts" : "FD placement / treasury outflow",
      isInterAccountTransfer: false,
    };
  }
  if (/loan|emi|nbfc|bank finance|principal repayment|borrowing/.test(text)) {
    return {
      category: "Loans / financing",
      fundFlowGroup: isReceipt ? "Loan receipts" : "Loan repayments",
      isInterAccountTransfer: false,
    };
  }
  if (/capital|partner|director|shareholder|owner|drawing|drawings|dividend/.test(text)) {
    return {
      category: "Owner / capital movement",
      fundFlowGroup: isReceipt ? "Owner / capital receipts" : "Owner / capital payments",
      isInterAccountTransfer: false,
    };
  }
  if (/laptop|computer|asset|equipment|furniture|camera|server|capex|vehicle/.test(text)) {
    return {
      category: "Capex / fixed assets",
      fundFlowGroup: "Capex / fixed assets",
      isInterAccountTransfer: false,
    };
  }
  if (/advance|reimbursement|reimburse|imprest|deposit|security deposit|prepaid/.test(text)) {
    return {
      category: "Advances / reimbursements",
      fundFlowGroup: isReceipt ? "Advances / reimbursements received" : "Advances / reimbursements paid",
      isInterAccountTransfer: false,
    };
  }
  if (/receipt|received|fee|client|invoice|export|customer|collection|professional|consulting|service/.test(text) || isReceipt) {
    return {
      category: isReceipt ? "Receipts" : "Operating payments",
      fundFlowGroup: isReceipt ? "Client / operating receipts" : "Vendor / operating payments",
      isInterAccountTransfer: false,
    };
  }
  if (/rent|office|software|subscription|audit|legal|travel|hotel|conveyance|vendor|consultant|professional|microsoft|aws|google|workspace|insurance|repair|maintenance|bank charge|charges/.test(text)) {
    return {
      category: "Operating payments",
      fundFlowGroup: "Vendor / operating payments",
      isInterAccountTransfer: false,
    };
  }
  return {
    category: isReceipt ? "Unclassified receipts" : "Unclassified payments",
    fundFlowGroup: isReceipt ? "Unclassified receipts" : "Unclassified payments",
    isInterAccountTransfer: false,
  };
}

export function parseBankStatementWorkbook(buffer: ArrayBuffer, sourceType: BankSourceType = "bank-statement", fileName = ""): ParsedBankStatement {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const transactions: BankTransaction[] = [];
  const warnings: string[] = [];

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
    const headerIndex = aoa.slice(0, 30).reduce(
      (best, row, index) => {
        const headers = row.map(normalizeHeader);
        let score = 0;
        if (findColumn(headers, dateHeaders) >= 0) score += 3;
        if (findColumn(headers, narrationHeaders) >= 0) score += 3;
        if (findColumn(headers, debitHeaders) >= 0 || findColumn(headers, creditHeaders) >= 0 || findColumn(headers, amountHeaders) >= 0) score += 3;
        if (findColumn(headers, balanceHeaders) >= 0) score += 1;
        return score > best.score ? { index, score } : best;
      },
      { index: -1, score: 0 },
    ).index;
    if (headerIndex < 0) return;

    const originalHeaders = aoa[headerIndex].map((header, index) => String(header || `Column ${index + 1}`).trim());
    const normalizedHeaders = originalHeaders.map(normalizeHeader);
    const dateIndex = findColumn(normalizedHeaders, dateHeaders);
    const narrationIndex = findColumn(normalizedHeaders, narrationHeaders);
    const debitIndex = findColumn(normalizedHeaders, debitHeaders);
    const creditIndex = findColumn(normalizedHeaders, creditHeaders);
    const amountIndex = findColumn(normalizedHeaders, amountHeaders);
    const balanceIndex = findColumn(normalizedHeaders, balanceHeaders);
    if (dateIndex < 0 || narrationIndex < 0) return;

    aoa.slice(headerIndex + 1).forEach((row, rowIndex) => {
      const date = parseDateValue(row[dateIndex]);
      const narration = String(row[narrationIndex] ?? "").trim();
      if (!date || !narration) return;
      let debit = debitIndex >= 0 ? Math.abs(parseNumber(row[debitIndex])) : 0;
      let credit = creditIndex >= 0 ? Math.abs(parseNumber(row[creditIndex])) : 0;
      if (!debit && !credit && amountIndex >= 0) {
        const signed = parseNumber(row[amountIndex]);
        if (signed < 0) debit = Math.abs(signed);
        else credit = Math.abs(signed);
      }
      if (!debit && !credit) return;
      if (sourceType === "bank-ledger") {
        [debit, credit] = [credit, debit];
      }
      const classification = classifyBankTransaction(narration, debit, credit);
      transactions.push({
        id: `${sourceType}-${fileName || "bank-source"}-${sheetName}-${rowIndex}-${date}-${narration}`.replace(/\s+/g, "-").toLowerCase(),
        sourceType,
        sourceFileName: fileName,
        sourceSheet: sheetName,
        date,
        narration,
        debit,
        credit,
        amount: credit - debit,
        balance: balanceIndex >= 0 ? parseNumber(row[balanceIndex]) : 0,
        bankName: sheetName,
        accountName: sheetName,
        category: classification.category,
        fundFlowGroup: classification.fundFlowGroup,
        isInterAccountTransfer: classification.isInterAccountTransfer,
        raw: toRawRecord(originalHeaders, row),
      });
    });
  });

  if (!transactions.length) warnings.push(`No ${sourceType === "bank-ledger" ? "bank ledger" : "bank statement"} transactions detected. Fund flow will remain provisional until bank source data is uploaded.`);
  return { transactions, warnings };
}
