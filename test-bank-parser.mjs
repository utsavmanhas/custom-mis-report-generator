/**
 * Bank ledger parser test — exercises locateBankColumns + parseBankSourceFile
 * logic directly in Node.js without needing a browser or test framework.
 *
 * Run: node test-bank-parser.mjs
 */

import * as XLSX from "xlsx";

// ── Re-implement the parser logic in JS so we can test it directly ───────────

const HEADER_SCAN_DEPTH = 30;

function normalizeHeader(value) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[().:_-]/g, " ")
    .trim();
}

function parseNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (value == null) return 0;
  const raw = String(value).trim();
  if (!raw) return 0;
  const isCredit = /\bcr\b/i.test(raw);
  const isDebit  = /\bdr\b/i.test(raw);
  const hasParens = /^\(.*\)$/.test(raw);
  const cleaned = raw.replace(/\b(cr|dr)\b/gi, "").replace(/[₹$€£,\s]/g, "").replace(/[()]/g, "");
  const parsed = Number.parseFloat(cleaned);
  if (!Number.isFinite(parsed)) return 0;
  if (isCredit) return -Math.abs(parsed);
  if (isDebit)  return Math.abs(parsed);
  if (hasParens) return -Math.abs(parsed);
  return parsed;
}

// ── FIXED version of locateBankColumns ───────────────────────────────────────

function locateBankColumns(aoa) {
  const top = aoa.slice(0, HEADER_SCAN_DEPTH);
  const ncols = top.reduce((max, row) => Math.max(max, row.length), 0);

  const headerKeywordRe = /^(date|txn date|transaction date|value date|voucher date|vch date|narration|description|particulars|remarks|debit|dr|credit|cr|amount|balance|closing balance|withdrawal|withdrawals|deposit|deposits|receipt|receipts|payment|payments|details|ledger|account|vch|vch type|vch no)$/;
  let lastHeaderRow = -1;
  let bestScore = 0;
  for (let row = 0; row < top.length; row++) {
    let score = 0;
    for (let col = 0; col < ncols; col++) {
      const v = normalizeHeader(top[row][col]);
      if (v && headerKeywordRe.test(v)) score++;
    }
    if (score > bestScore) { bestScore = score; lastHeaderRow = row; }
    else if (score > 0 && score === bestScore && row > lastHeaderRow) lastHeaderRow = row;
  }

  if (lastHeaderRow < 0) {
    for (let row = 0; row < top.length; row++) {
      for (let col = 0; col < ncols; col++) {
        const v = normalizeHeader(top[row][col]);
        if (v && /\b(date|narration|description|particulars|debit|credit|amount|balance|withdrawal|deposit|receipt|payment)\b/.test(v)) {
          lastHeaderRow = row; break;
        }
      }
      if (lastHeaderRow >= 0) break;
    }
  }

  if (lastHeaderRow < 0) return null;

  const headerRow = top[lastHeaderRow] ?? [];
  let dateIdx = -1, descriptionIdx = -1, debitIdx = -1, creditIdx = -1, amountIdx = -1, balanceIdx = -1;

  for (let column = 0; column < headerRow.length; column++) {
    const value = normalizeHeader(headerRow[column]);
    if (!value) continue;
    if (dateIdx < 0        && /\b(date|txn date|transaction date|value date|voucher date|vch)\b/.test(value)) dateIdx = column;
    if (descriptionIdx < 0 && /\b(narration|description|particulars|remarks|ledger|account|details)\b/.test(value)) descriptionIdx = column;
    if (debitIdx < 0       && /\b(debit|withdrawals?|payment|paid|dr)\b/.test(value)) debitIdx = column;
    if (creditIdx < 0      && /\b(credit|deposits?|receipts?|received|cr)\b/.test(value)) creditIdx = column;
    if (amountIdx < 0      && /\bamount\b/.test(value)) amountIdx = column;
    if (balanceIdx < 0     && /\bbalance\b|\bclosing\b/.test(value)) balanceIdx = column;
  }

  if (dateIdx < 0 && descriptionIdx < 0) return null;
  if (debitIdx < 0 && creditIdx < 0 && amountIdx < 0) return null;

  let descriptionEndIdx = descriptionIdx;
  if (descriptionIdx >= 0) {
    const firstAmountCol = [debitIdx, creditIdx, amountIdx]
      .filter((i) => i > descriptionIdx)
      .reduce((min, i) => Math.min(min, i), headerRow.length);
    for (let col = descriptionIdx + 1; col < firstAmountCol; col++) {
      const h = normalizeHeader(headerRow[col]);
      if (h && /\b(vch|voucher|type|no|number|ref|reference|cheque|chq|mode|bank)\b/.test(h)) break;
      if (!h) descriptionEndIdx = col;
    }
  }

  return { dateIdx, descriptionIdx, descriptionEndIdx, debitIdx, creditIdx, amountIdx, balanceIdx, lastHeaderRow };
}

// ── OLD (buggy) version for comparison ───────────────────────────────────────

function locateBankColumnsOLD(aoa) {
  const top = aoa.slice(0, HEADER_SCAN_DEPTH);
  const ncols = top.reduce((max, row) => Math.max(max, row.length), 0);
  let dateIdx = -1, descriptionIdx = -1, debitIdx = -1, creditIdx = -1, amountIdx = -1, balanceIdx = -1, lastHeaderRow = -1;

  for (let column = 0; column < ncols; column++) {
    const colText = top.map((row) => normalizeHeader(row[column])).filter(Boolean).join(" | ");
    if (!colText) continue;
    if (dateIdx < 0        && /\b(date|txn date|transaction date|value date|voucher date|vch)\b/.test(colText)) dateIdx = column;
    if (descriptionIdx < 0 && /\b(narration|description|particulars|remarks|ledger|account|details)\b/.test(colText)) descriptionIdx = column;
    if (debitIdx < 0       && /\b(debit|withdrawal|payment|paid|dr)\b/.test(colText)) debitIdx = column;
    if (creditIdx < 0      && /\b(credit|deposit|receipt|received|cr)\b/.test(colText)) creditIdx = column;
    if (amountIdx < 0      && /\bamount\b/.test(colText)) amountIdx = column;
    if (balanceIdx < 0     && /\bbalance\b|\bclosing\b/.test(colText)) balanceIdx = column;

    for (let row = 0; row < top.length; row++) {
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

// ── Parse transactions (same as parser.ts) ───────────────────────────────────

function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d)).toISOString().slice(0, 10);
  }
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const normal = raw.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/);
  if (normal) {
    let year = Number.parseInt(normal[3], 10);
    if (year < 100) year += 2000;
    const d = new Date(Date.UTC(year, Number.parseInt(normal[2], 10) - 1, Number.parseInt(normal[1], 10)));
    return Number.isNaN(d.getTime()) ? "" : d.toISOString().slice(0, 10);
  }
  const monName = raw.match(/^(\d{1,2})[-\s]([A-Za-z]{3,})[-\s](\d{4})$/);
  if (monName) {
    const months = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
    const m = months[monName[2].toLowerCase().slice(0, 3)];
    if (m !== undefined) {
      const d = new Date(Date.UTC(Number.parseInt(monName[3], 10), m, Number.parseInt(monName[1], 10)));
      if (!Number.isNaN(d.getTime())) return d.toISOString().slice(0, 10);
    }
  }
  return "";
}

function extractTransactions(aoa, located) {
  const transactions = [];
  aoa.slice(located.lastHeaderRow + 1).forEach((row, offset) => {
    const sourceRowNumber = located.lastHeaderRow + offset + 2;
    const date = located.dateIdx >= 0 ? parseDate(row[located.dateIdx]) : "";

    const descParts = [];
    for (let col = located.descriptionIdx; col <= located.descriptionEndIdx; col++) {
      const v = String(row[col] ?? "").trim();
      if (v) descParts.push(v);
    }
    const description = descParts.join(" ");

    let debit  = located.debitIdx >= 0  ? Math.abs(parseNumber(row[located.debitIdx]))  : 0;
    let credit = located.creditIdx >= 0 ? Math.abs(parseNumber(row[located.creditIdx])) : 0;

    if (!debit && !credit && located.amountIdx >= 0) {
      const amount = parseNumber(row[located.amountIdx]);
      if (amount < 0) debit  = Math.abs(amount);
      if (amount > 0) credit = Math.abs(amount);
    }

    if (!date && !description && !debit && !credit) return;
    if (!debit && !credit) return;
    if (/\bopening\s*balance\b/i.test(description)) return;

    transactions.push({ sourceRowNumber, date, description, debit, credit });
  });
  return transactions;
}

// ── Helper to build XLSX buffer from 2-D array ────────────────────────────────

function makeXlsx(aoa) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

function parseAoa(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
}

// ── Test runner ───────────────────────────────────────────────────────────────

let passed = 0;
let failed = 0;

function run(label, aoa, expectedTxns) {
  const buf = makeXlsx(aoa);
  const parsed = parseAoa(buf);

  const newResult = locateBankColumns(parsed);
  const oldResult = locateBankColumnsOLD(parsed);

  const newTxns = newResult ? extractTransactions(parsed, newResult) : [];
  const oldTxns = oldResult ? extractTransactions(parsed, oldResult) : [];

  const ok = newTxns.length === expectedTxns;
  if (ok) {
    passed++;
    console.log(`  ✓  ${label} — ${newTxns.length} txn(s) detected`);
    if (oldTxns.length !== newTxns.length) {
      console.log(`       (OLD code would have found ${oldTxns.length} — bug confirmed & fixed)`);
    }
  } else {
    failed++;
    console.log(`  ✗  ${label}`);
    console.log(`       Expected ${expectedTxns} txn(s), got ${newTxns.length}`);
    console.log(`       locateBankColumns result:`, newResult);
    if (newTxns.length > 0) {
      console.log(`       First txn:`, newTxns[0]);
    }
  }
}

function detail(label, aoa) {
  const buf = makeXlsx(aoa);
  const parsed = parseAoa(buf);
  const result = locateBankColumns(parsed);
  if (!result) { console.log(`  [detail] ${label}: locateBankColumns → null`); return; }
  const txns = extractTransactions(parsed, result);
  console.log(`  [detail] ${label}: headerRow=${result.lastHeaderRow}, debitIdx=${result.debitIdx}, creditIdx=${result.creditIdx}, dateIdx=${result.dateIdx}, descriptionIdx=${result.descriptionIdx}`);
  txns.forEach((t, i) => console.log(`    txn[${i}]: ${t.date} | ${t.description.padEnd(30)} | debit=${t.debit} credit=${t.credit}`));
}

// ── FIXTURE 1: Classic Tally bank ledger (separate Debit/Credit columns) ─────
// This is the main failing case: narration has "payment"/"deposit"/"receipt"
// which caused old code to misassign debitIdx/creditIdx to the narration column.

console.log("\n=== Fixture 1: Tally bank ledger with narration keywords ===");
const tally1 = [
  ["Acme Pvt Ltd"],
  ["Bank Accounts > HDFC Current Account"],
  ["1-Apr-2024 to 31-Mar-2025"],
  [],
  ["Date",       "Particulars",                         "Vch Type", "Vch No.", "Debit",   "Credit", "Balance"],
  ["1-Apr-2024", "Opening Balance",                      "",         "",         "",       "",      "2,00,000 Dr"],
  ["2-Apr-2024", "To Sales Invoice — Receipt from XYZ",  "Receipt",  "001",      "",       50000,   "2,50,000 Dr"],
  ["5-Apr-2024", "By Bank Payment to ABC Suppliers",     "Payment",  "002",      30000,    "",      "2,20,000 Dr"],
  ["10-Apr-2024","By Cash Deposit — deposit to petty",   "Contra",   "003",      "",       10000,   "2,30,000 Dr"],
  ["15-Apr-2024","To Vendor — Cash Payment settlement",  "Payment",  "004",      15000,    "",      "2,15,000 Dr"],
  ["30-Apr-2024","By Salary Payment — staff payroll",    "Payment",  "005",      80000,    "",      "1,35,000 Dr"],
];

run("Tally ledger — narration has payment/deposit/receipt keywords", tally1, 5);
detail("Tally ledger detail", tally1);

// ── FIXTURE 2: Bank statement (simple, no company header) ─────────────────────

console.log("\n=== Fixture 2: Simple bank statement (no company header rows) ===");
const stmt1 = [
  ["Date",       "Description",                   "Debit",  "Credit", "Balance"],
  ["01/04/2024", "NEFT Credit from Customer A",   "",       45000,    145000],
  ["03/04/2024", "UPI Payment to Vendor",         12000,    "",       133000],
  ["07/04/2024", "Cash Withdrawal",               5000,     "",       128000],
  ["15/04/2024", "Interest Credit",               "",       1200,     129200],
  ["30/04/2024", "GST Payment",                   18000,    "",       111200],
];

run("Bank statement — simple format", stmt1, 5);
detail("Bank statement detail", stmt1);

// ── FIXTURE 3: Bank statement with Withdrawal/Deposit column names ────────────

console.log("\n=== Fixture 3: Withdrawal / Deposits column naming ===");
const stmt2 = [
  ["Txn Date",    "Narration",                      "Withdrawals", "Deposits", "Balance"],
  ["01-04-2024",  "Opening balance",                 "",            "",          75000],
  ["02-04-2024",  "Salary credited",                 "",            120000,      195000],
  ["05-04-2024",  "Rent payment",                    25000,         "",          170000],
  ["10-04-2024",  "Client receipt — Project Alpha",  "",            80000,       250000],
];

run("Withdrawal/Deposits column names", stmt2, 3);
detail("Withdrawal/Deposits detail", stmt2);

// ── FIXTURE 4: Single Amount column with Dr/Cr suffix (Tally compact export) ──

console.log("\n=== Fixture 4: Single Amount column with Dr/Cr suffix ===");
const tally2 = [
  ["Company: Beta Solutions"],
  ["Ledger: ICICI Bank OD"],
  [],
  ["Date",       "Particulars",                     "Amount"],
  ["1-Apr-2024", "To Sales Receipt ABC Corp",       "75,000 Cr"],
  ["3-Apr-2024", "By Purchase Payment XYZ Ltd",     "45,000 Dr"],
  ["8-Apr-2024", "By Rent payment office",          "25,000 Dr"],
  ["15-Apr-2024","To Consulting Receipt DEF",       "30,000 Cr"],
];

run("Single Amount column with Dr/Cr suffix", tally2, 4);
detail("Single Amount detail", tally2);

// ── FIXTURE 5: Header is NOT on row 0 — company info precedes it ─────────────

console.log("\n=== Fixture 5: Header preceded by 6 rows of company metadata ===");
const tally3 = [
  ["XYZ Exports Pvt Ltd"],
  ["Trial Balance — Bank Accounts"],
  ["Period: 1-Apr-2024 to 31-Mar-2025"],
  ["Currency: INR"],
  [],
  [],
  ["Date",       "Particulars",                    "Debit",  "Credit",  "Balance"],
  ["01-04-2024", "To Export Receipt USD payment",  "",       250000,    750000],
  ["05-04-2024", "By Bank Payment RTGS",           100000,   "",        650000],
  ["20-04-2024", "To Advance Receipt deposit",     "",       50000,     700000],
];

run("Header at row 6 — company metadata above", tally3, 3);
detail("Header at row 6 detail", tally3);

// ── FIXTURE 6: The exact old bug scenario ─────────────────────────────────────
// Narration col is col 1, Debit is col 2. Old code set debitIdx=1 (narration).
// Verify the old code was broken and new code fixes it.

console.log("\n=== Fixture 6: Bug regression — narration poisons column detection ===");
const bugFix = [
  ["Date",       "Particulars",                        "Debit", "Credit", "Balance"],
  ["01-04-2024", "Being payment to vendor",             25000,  "",       75000],
  ["02-04-2024", "Being receipt from client",           "",     50000,    125000],
  ["03-04-2024", "Bank deposit confirmation receipt",   "",     10000,    135000],
];

const bufBug = makeXlsx(bugFix);
const parsedBug = parseAoa(bufBug);
const oldRes = locateBankColumnsOLD(parsedBug);
const newRes = locateBankColumns(parsedBug);
const oldTxns = oldRes ? extractTransactions(parsedBug, oldRes) : [];
const newTxns = newRes ? extractTransactions(parsedBug, newRes) : [];

console.log(`  OLD debitIdx → col ${oldRes?.debitIdx ?? "null"} (should be 2)`);
console.log(`  NEW debitIdx → col ${newRes?.debitIdx ?? "null"} (should be 2)`);
console.log(`  OLD found ${oldTxns.length} transactions (should be 3 — bug shows 0 or wrong)`);
console.log(`  NEW found ${newTxns.length} transactions (should be 3)`);

if (newRes?.debitIdx === 2 && newTxns.length === 3) {
  passed++;
  console.log("  ✓  Bug regression: new code correctly assigns debitIdx=2 and finds 3 txns");
} else {
  failed++;
  console.log("  ✗  Bug regression: fix did not work as expected");
}
if (oldRes?.debitIdx !== 2 || oldTxns.length !== 3) {
  console.log("  ✓  Confirmed: old code had the bug (debitIdx was wrong or 0 txns found)");
}

// ── FIXTURE 7: Edge — balance-only column with Dr/Cr doesn't steal debitIdx ──

console.log("\n=== Fixture 7: Balance column has Dr/Cr — should not pollute debitIdx ===");
const tally4 = [
  ["Date",       "Particulars",    "Debit",   "Credit",  "Balance"],
  ["01-04-24",   "Opening Bal",    "",         "",        "1,00,000 Dr"],
  ["05-04-24",   "To Receipt",     "",         20000,     "1,20,000 Dr"],
  ["10-04-24",   "By Payment",     15000,      "",        "1,05,000 Dr"],
];

run("Balance column has Dr/Cr suffix — debitIdx stays on Debit column", tally4, 2);
detail("Balance Dr/Cr detail", tally4);

// ── FIXTURE 8: Actual "Bank Ledger.xlsx" from Downloads ──────────────────────

console.log("\n=== Fixture 8: Actual Bank Ledger.xlsx ===");
import { readFileSync } from "fs";

const realBuf = readFileSync("C:\\Users\\Administrator\\Downloads\\Bank Ledger.xlsx");
const realWb  = XLSX.read(realBuf, { type: "buffer", cellDates: true });

console.log(`  Sheets: ${realWb.SheetNames.join(", ")}`);

for (const sheetName of realWb.SheetNames) {
  const ws  = realWb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  console.log(`\n  Sheet: "${sheetName}" (${aoa.length} rows)`);

  // Show top 10 rows so we can see the raw layout
  console.log("  Raw top rows:");
  aoa.slice(0, 10).forEach((row, i) => {
    const cells = row.map(c => String(c ?? "").slice(0, 25)).join(" | ");
    if (cells.trim()) console.log(`    row ${i}: ${cells}`);
  });

  const located = locateBankColumns(aoa);
  const locatedOld = locateBankColumnsOLD(aoa);

  if (!located) {
    console.log("  NEW locateBankColumns → null (sheet skipped)");
  } else {
    console.log(`  NEW → headerRow=${located.lastHeaderRow}, dateIdx=${located.dateIdx}, descriptionIdx=${located.descriptionIdx}, debitIdx=${located.debitIdx}, creditIdx=${located.creditIdx}, amountIdx=${located.amountIdx}, balanceIdx=${located.balanceIdx}`);
    const txns = extractTransactions(aoa, located);
    console.log(`  NEW → ${txns.length} transactions detected`);
    if (txns.length > 0) {
      console.log("  First 5 transactions:");
      txns.slice(0, 5).forEach((t, i) =>
        console.log(`    [${i}] ${t.date} | ${t.description.slice(0, 35).padEnd(35)} | debit=${t.debit} credit=${t.credit}`)
      );
      const totalDebit  = txns.reduce((s, t) => s + t.debit,  0);
      const totalCredit = txns.reduce((s, t) => s + t.credit, 0);
      console.log(`  Totals → debit=${totalDebit.toLocaleString("en-IN")}  credit=${totalCredit.toLocaleString("en-IN")}`);
    }
    if (txns.length > 0) { passed++; } else { failed++; }
  }

  if (!locatedOld) {
    console.log("  OLD locateBankColumns → null");
  } else {
    const oldTxns = extractTransactions(aoa, locatedOld);
    console.log(`  OLD → debitIdx=${locatedOld.debitIdx}, creditIdx=${locatedOld.creditIdx} → ${oldTxns.length} transactions`);
  }
}

// ── Results ───────────────────────────────────────────────────────────────────

console.log(`\n${"─".repeat(55)}`);
console.log(`  ${passed + failed} tests — ${passed} passed, ${failed} failed`);
if (failed === 0) {
  console.log("  All tests passed. Bank ledger detection fix is working.\n");
} else {
  console.log(`  ${failed} test(s) failed. Check output above.\n`);
  process.exitCode = 1;
}
