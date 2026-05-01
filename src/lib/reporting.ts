import { categorySignedAmount } from "./classification";
import { generateCustomMisWorkbook } from "./customReport";
import { generateUniversityMisWorkbook } from "./universityReport";
import type { AccountCategory, BusinessProfile, GeneratedQuestion, ProfitCenter, QuestionAnswer, StaffMember, TrialBalanceRow, WorkbookIssue } from "../types";

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
}) {
  const issues = buildWorkbookIssues(params.profile, params.rows, params.centers, params.staff, params.questions);

  if (params.profile.businessType === "university") {
    return generateUniversityMisWorkbook({ ...params, issues });
  }

  return generateCustomMisWorkbook({ ...params, issues });
}
