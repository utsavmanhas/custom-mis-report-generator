import { categorySignedAmount } from "./classification";
import { generateCustomMisWorkbook } from "./customReport";
import { generateUniversityMisWorkbook } from "./universityReport";
import type { AccountCategory, BankTransaction, BusinessProfile, GeneratedQuestion, ProfitCenter, QuestionAnswer, StaffMember, TrialBalanceRow, WorkbookIssue } from "../types";

function amountForCategory(row: TrialBalanceRow, category: AccountCategory) {
  return row.category === category ? Math.abs(categorySignedAmount(row)) : 0;
}

function totalByCategory(rows: TrialBalanceRow[], category: AccountCategory) {
  return rows.reduce((sum, row) => sum + amountForCategory(row, category), 0);
}

export function buildWorkbookIssues(
  profile: BusinessProfile,
  rows: TrialBalanceRow[],
  centers: ProfitCenter[],
  staff: StaffMember[],
  questions: GeneratedQuestion[],
  bankTransactions: BankTransaction[] = [],
): WorkbookIssue[] {
  const issues: WorkbookIssue[] = [];
  const totalDebit = rows.reduce((sum, row) => sum + row.debit, 0);
  const totalCredit = rows.reduce((sum, row) => sum + row.credit, 0);
  const imbalance = Math.abs(totalDebit - totalCredit);

  if (!rows.length) {
    issues.push({ label: "Missing accounting source", detail: "No trial balance or ledger file has been imported.", severity: "high" });
  }

  if (imbalance > 1) {
    issues.push({ label: "Trial balance mismatch", detail: `Debit and credit differ by ${profile.currency} ${imbalance.toLocaleString()}.`, severity: "high" });
  }

  const unknownCount = rows.filter((row) => row.category === "unknown").length;
  if (unknownCount) {
    issues.push({ label: "Unclassified ledgers", detail: `${unknownCount} ledger rows need category confirmation.`, severity: "medium" });
  }

  if (!centers.length) {
    issues.push({ label: "No operating units", detail: "At least one project, department, product, channel, plant, or custom unit is needed for granular MIS output.", severity: "high" });
  }

  if (totalByCategory(rows, "people-cost") > 0 && !staff.length) {
    issues.push({ label: "People allocation missing", detail: "Payroll exists in the trial balance but no staff assignment data has been entered.", severity: "medium" });
  }

  const hasBankStatement = bankTransactions.some((txn) => txn.sourceType === "bank-statement");
  const hasBankLedger = bankTransactions.some((txn) => txn.sourceType === "bank-ledger");
  if (rows.length && !bankTransactions.length) {
    issues.push({
      label: "Fund flow source missing",
      detail: "Upload bank statements and bank ledgers before treating fund flow as client-ready.",
      severity: "high",
    });
  }
  if (rows.length && bankTransactions.length && !hasBankStatement) {
    issues.push({
      label: "Bank statement missing",
      detail: "Bank ledger is useful for classification, but actual fund flow needs bank statement movement.",
      severity: "high",
    });
  }
  if (rows.length && bankTransactions.length && !hasBankLedger) {
    issues.push({
      label: "Bank ledger missing",
      detail: "Upload bank ledgers to support classification and reconcile book movement to statement movement.",
      severity: "medium",
    });
  }

  if (questions.some((question) => question.priority === "high")) {
    issues.push({ label: "Open critical questions", detail: `${questions.filter((question) => question.priority === "high").length} high-priority MIS questions remain open.`, severity: "medium" });
  }

  return issues;
}

export function generateMisWorkbook(params: {
  profile: BusinessProfile;
  rows: TrialBalanceRow[];
  centers: ProfitCenter[];
  staff: StaffMember[];
  questions: GeneratedQuestion[];
  answers: QuestionAnswer[];
  bankTransactions?: BankTransaction[];
}) {
  const issues = buildWorkbookIssues(params.profile, params.rows, params.centers, params.staff, params.questions, params.bankTransactions || []);

  if (params.profile.businessType === "university") {
    return generateUniversityMisWorkbook({ ...params, issues });
  }

  return generateCustomMisWorkbook({ ...params, issues });
}
