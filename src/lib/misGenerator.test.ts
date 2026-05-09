import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import * as XLSX from "xlsx";
import { parseBankStatementWorkbook, parseTrialBalanceWorkbook } from "./parser";
import { buildQuestions } from "./questions";
import { generateMisWorkbook } from "./reporting";
import type { BankTransaction, BusinessProfile, ProfitCenter } from "../types";

const simpleTbPath = "C:/Users/Administrator/OneDrive/Desktop/Simple_1TrialBal.xlsx";

function arrayBufferFromFile(path: string) {
  const buffer = readFileSync(path);
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function arrayBufferFromRows(rows: unknown[][]) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "ICICI Bank");
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
}

function profileFromSimple(): BusinessProfile {
  return {
    businessName: "SIMPLE INSIGHTS PRIVATE LIMITED",
    legalEntity: "SIMPLE INSIGHTS PRIVATE LIMITED",
    businessType: "professional-services",
    cadence: "annual",
    periodStart: "2025-04-01",
    periodEnd: "2026-03-31",
    currency: "INR",
    website: "",
    geography: "India",
    publicNotes: "",
    sourceUrls: "",
    allocationBase: "revenue",
    fundFlowBasis: "trial-balance",
  };
}

function defaultCenter(): ProfitCenter {
  return {
    id: "client-all",
    name: "All clients",
    kind: "project",
    owner: "",
    segment: "Research analytics",
    revenueDriver: "Invoice register pending",
    manualRevenue: 0,
    manualDirectCost: 0,
    priorRevenue: 0,
    priorDirectCost: 0,
    studentCount: 0,
    teachingStaffCount: 0,
    nonTeachingStaffCount: 0,
    averageRevenueRate: 0,
    variableCostRate: 0,
    utilizationPercent: 0,
    allocationWeight: 1,
    notes: "TB revenue is not split by invoice register yet.",
  };
}

describe("Simple trial balance parser", () => {
  it("infers source metadata and excludes group headings from TB rows", () => {
    const parsed = parseTrialBalanceWorkbook(arrayBufferFromFile(simpleTbPath), "Simple_1TrialBal.xlsx");

    expect(parsed.metadata.businessName).toBe("SIMPLE INSIGHTS PRIVATE LIMITED");
    expect(parsed.metadata.periodStart).toBe("2025-04-01");
    expect(parsed.metadata.periodEnd).toBe("2026-03-31");
    expect(parsed.metadata.totalDebit).toBeCloseTo(98644108.52, 2);
    expect(parsed.metadata.totalCredit).toBeCloseTo(98644108.52, 2);
    expect(parsed.rows.some((row) => row.accountName === "Grand Total")).toBe(false);
    expect(parsed.rows.some((row) => row.accountName === "Current Liabilities")).toBe(false);
    expect(parsed.rows.filter((row) => row.category === "unknown")).toHaveLength(0);
    expect(parsed.warnings.some((warning) => /Suspense/i.test(warning))).toBe(true);

    const reliant = parsed.rows.find((row) => /Reliant AI/i.test(row.accountName));
    expect(reliant?.category).toBe("current-liability");
    expect(reliant?.misGroup).toMatch(/Customer Advances/);
    expect(parsed.rows.find((row) => /TDS \(Salary\)/i.test(row.accountName))?.misGroup).toBe("TDS Payable");
    expect(parsed.rows.find((row) => /ICICI FD ACCOUNT/i.test(row.accountName))?.misGroup).toBe("Fixed Deposits / Treasury");
    expect(parsed.rows.find((row) => /Bonus Account/i.test(row.accountName))?.misGroup).toBe("Bonus / Variable Compensation");
    expect(parsed.rows.find((row) => /Consultancy Fees/i.test(row.accountName))?.misGroup).toBe("Professional & Consultancy");
  });
});

describe("bank source parser and source-led questions", () => {
  it("keeps bank statements and bank ledgers separate and normalizes ledger direction", () => {
    const statement = parseBankStatementWorkbook(
      arrayBufferFromRows([
        ["Date", "Narration", "Debit", "Credit", "Balance"],
        ["01-Apr-2025", "Client receipt from customer", "", 1000, 11000],
        ["02-Apr-2025", "Software subscription payment", 250, "", 10750],
      ]),
      "bank-statement",
      "statement.xlsx",
    );
    const ledger = parseBankStatementWorkbook(
      arrayBufferFromRows([
        ["Date", "Particulars", "Debit", "Credit", "Balance"],
        ["01-Apr-2025", "Client receipt from customer", 1000, "", 11000],
        ["02-Apr-2025", "Software subscription payment", "", 250, 10750],
      ]),
      "bank-ledger",
      "ledger.xlsx",
    );

    expect(statement.transactions[0].sourceType).toBe("bank-statement");
    expect(statement.transactions[0].credit).toBe(1000);
    expect(statement.transactions[1].debit).toBe(250);
    expect(ledger.transactions[0].sourceType).toBe("bank-ledger");
    expect(ledger.transactions[0].credit).toBe(1000);
    expect(ledger.transactions[1].debit).toBe(250);
  });

  it("does not ask redundant profile or TB-category questions when the TB already provides them", () => {
    const parsed = parseTrialBalanceWorkbook(arrayBufferFromFile(simpleTbPath), "Simple_1TrialBal.xlsx");
    const questions = buildQuestions(profileFromSimple(), parsed.rows, [defaultCenter()], [], [], []);
    const questionText = questions.map((question) => question.prompt).join("\n");

    expect(questionText).not.toMatch(/time period|reporting period|company name/i);
    expect(questionText).not.toMatch(/confirm the category|classify.*ledger|unclassified ledger/i);
    expect(questions.some((question) => question.id === "upload-bank-statements-ledgers")).toBe(true);
  });
});

describe("client-ready workbook output", () => {
  it("exports AKEV-style schedules and QC values for Simple", () => {
    const parsed = parseTrialBalanceWorkbook(arrayBufferFromFile(simpleTbPath), "Simple_1TrialBal.xlsx");
    const output = generateMisWorkbook({
      profile: profileFromSimple(),
      rows: parsed.rows,
      centers: [defaultCenter()],
      staff: [],
      questions: [],
      answers: [],
      bankTransactions: [],
    });
    const wb = output.workbook;

    expect(wb.SheetNames).toContain("Balance Sheet");
    expect(wb.SheetNames).toContain("Tax & Statutory");
    expect(wb.SheetNames).toContain("Fund Flow");
    expect(wb.SheetNames).toContain("Fund Flow Projection");
    expect(wb.SheetNames).toContain("Bank Statement Audit");
    expect(wb.SheetNames).toContain("Bank Ledger Audit");
    expect(wb.SheetNames).toContain("Bank Reconciliation");
    expect(wb.SheetNames).toContain("Fund Flow Assumptions");
    expect(wb.SheetNames).toContain("ChartData");
    expect(wb.SheetNames).toContain("QC Tie-Outs");
    expect(wb.SheetNames).toContain("Account Allocation");

    const control = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets["00 Control Panel"], { header: 1, defval: "" });
    expect(JSON.stringify(control)).not.toMatch(/Sample Business/i);
    expect(JSON.stringify(control)).toMatch(/SIMPLE INSIGHTS PRIVATE LIMITED/);

    const qc = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets["QC Tie-Outs"], { header: 1, defval: "" });
    const qcText = JSON.stringify(qc);
    expect(qcText).toMatch(/79,502,679\.52|79502679\.52/);
    expect(qcText).toMatch(/15,360,815\.97|15360815\.97/);
  });

  it("exports bank-backed fund flow sheets without duplicate statement/ledger actuals", () => {
    const parsed = parseTrialBalanceWorkbook(arrayBufferFromFile(simpleTbPath), "Simple_1TrialBal.xlsx");
    const statement = parseBankStatementWorkbook(
      arrayBufferFromRows([
        ["Date", "Narration", "Debit", "Credit", "Balance"],
        ["01-Apr-2025", "Client receipt from customer", "", 1000, 11000],
        ["02-Apr-2025", "Salary payment", 250, "", 10750],
      ]),
      "bank-statement",
      "statement.xlsx",
    ).transactions;
    const ledger = parseBankStatementWorkbook(
      arrayBufferFromRows([
        ["Date", "Particulars", "Debit", "Credit", "Balance"],
        ["01-Apr-2025", "Client receipt from customer", 1000, "", 11000],
        ["02-Apr-2025", "Salary payment", "", 250, 10750],
      ]),
      "bank-ledger",
      "ledger.xlsx",
    ).transactions;
    const bankTransactions: BankTransaction[] = [...statement, ...ledger];
    const output = generateMisWorkbook({
      profile: { ...profileFromSimple(), fundFlowBasis: "bank-statement-and-ledger" },
      rows: parsed.rows,
      centers: [defaultCenter()],
      staff: [],
      questions: [],
      answers: [],
      bankTransactions,
    });
    const fundFlow = XLSX.utils.sheet_to_json<unknown[]>(output.workbook.Sheets["Fund Flow"], { header: 1, defval: "" });
    const fundFlowText = JSON.stringify(fundFlow);

    expect(fundFlowText).toMatch(/Bank statement actuals/);
    expect(fundFlowText).toMatch(/Client \/ operating receipts/);
    expect(fundFlowText).toMatch(/Salary \/ bonus \/ payroll/);
    expect(fundFlowText).toMatch(/1000/);
    expect(fundFlowText).toMatch(/250/);
    expect(fundFlowText).not.toMatch(/2000/);
  });
});
