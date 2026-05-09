import * as XLSX from "xlsx";
import { allocationLabels, businessTypeLabels, categoryLabels, categorySignedAmount } from "./classification";
import { getBusinessTemplate } from "./businessTemplates";
import type { BankTransaction, BusinessProfile, GeneratedQuestion, ProfitCenter, QuestionAnswer, StaffMember, TrialBalanceRow, WorkbookIssue } from "../types";

const moneyFormat = '#,##0.00;[Red](#,##0.00);-';
const percentFormat = "0.0%";

interface CustomWorkbookParams {
  profile: BusinessProfile;
  rows: TrialBalanceRow[];
  centers: ProfitCenter[];
  staff: StaffMember[];
  questions: GeneratedQuestion[];
  answers: QuestionAnswer[];
  issues: WorkbookIssue[];
  bankTransactions?: BankTransaction[];
}

interface UnitFinancials {
  center: ProfitCenter;
  driverRow: number;
  ledgerRevenue: number;
  driverRevenue: number;
  otherIncome: number;
  directCost: number;
  driverDirectCost: number;
  peopleCost: number;
  financeTax: number;
  sharedWeight: number;
  sharedOpex: number;
  priorRevenue: number;
  priorCost: number;
}

interface FundFlowModel {
  months: string[];
  projectionMonths: string[];
  openingBalances: number[];
  closingBalances: number[];
  closingCash: number;
  receiptRows: Array<[string, number[], string]>;
  paymentRows: Array<[string, number[], string]>;
  internalTransferValues: number[];
}

function makeSheet(data: unknown[][]) {
  return XLSX.utils.aoa_to_sheet(data);
}

function setFormula(sheet: XLSX.WorkSheet, address: string, formula: string, format = moneyFormat, cachedValue = 0) {
  sheet[address] = { t: "n", v: cachedValue, f: formula, z: format };
}

function col(index: number) {
  return XLSX.utils.encode_col(index);
}

function sanitizeSheetName(name: string) {
  return name.replace(/[\\/?*[\]:]/g, " ").slice(0, 31).trim() || "Sheet";
}

function sheetCell(sheetName: string, address: string) {
  return `'${sheetName.replace(/'/g, "''")}'!${address}`;
}

function periodMonths(profile: BusinessProfile) {
  if (profile.periodStart && profile.periodEnd) {
    const start = new Date(`${profile.periodStart}T00:00:00`);
    const end = new Date(`${profile.periodEnd}T00:00:00`);
    if (!Number.isNaN(start.getTime()) && !Number.isNaN(end.getTime())) {
      return Math.max(1, (end.getFullYear() - start.getFullYear()) * 12 + end.getMonth() - start.getMonth() + 1);
    }
  }

  if (profile.cadence === "annual") return 12;
  if (profile.cadence === "quarterly") return 3;
  return 1;
}

function dateRange(profile: BusinessProfile) {
  if (profile.periodStart && profile.periodEnd) return `${profile.periodStart} to ${profile.periodEnd}`;
  return `${profile.cadence} reporting period`;
}

function monthLabels(profile: BusinessProfile) {
  const months = periodMonths(profile);
  const start = profile.periodStart ? new Date(`${profile.periodStart}T00:00:00`) : null;
  return Array.from({ length: months }, (_, index) => {
    if (start && !Number.isNaN(start.getTime())) {
      const month = new Date(start.getFullYear(), start.getMonth() + index, 1);
      return month.toLocaleString("en-US", { month: "short", year: "2-digit" });
    }
    return `Month ${index + 1}`;
  });
}

function projectionMonthLabels(profile: BusinessProfile, count = 12) {
  const periodEnd = profile.periodEnd ? new Date(`${profile.periodEnd}T00:00:00`) : null;
  const start =
    periodEnd && !Number.isNaN(periodEnd.getTime())
      ? new Date(periodEnd.getFullYear(), periodEnd.getMonth() + 1, 1)
      : new Date(new Date().getFullYear(), new Date().getMonth() + 1, 1);

  return Array.from({ length: count }, (_, index) => {
    const month = new Date(start.getFullYear(), start.getMonth() + index, 1);
    return month.toLocaleString("en-US", { month: "short", year: "2-digit" });
  });
}

function monthKey(date: string) {
  const parsed = new Date(`${date}T00:00:00`);
  if (Number.isNaN(parsed.getTime())) return "";
  return parsed.toLocaleString("en-US", { month: "short", year: "2-digit" });
}

function amount(row: TrialBalanceRow) {
  return Math.abs(categorySignedAmount(row));
}

function numeric(value: number | undefined) {
  return Number.isFinite(value) ? Number(value) : 0;
}

function driverRevenue(center: ProfitCenter) {
  return numeric(center.studentCount) * numeric(center.averageRevenueRate) * (numeric(center.utilizationPercent) ? numeric(center.utilizationPercent) / 100 : 1);
}

function driverDirectCost(center: ProfitCenter) {
  return numeric(center.studentCount) * numeric(center.variableCostRate);
}

function assignedAmount(rows: TrialBalanceRow[], centerId: string, categories: string[]) {
  return rows.filter((row) => row.profitCenterId === centerId && categories.includes(row.category)).reduce((sum, row) => sum + amount(row), 0);
}

function unassignedAmount(rows: TrialBalanceRow[], categories: string[]) {
  return rows.filter((row) => !row.profitCenterId && categories.includes(row.category)).reduce((sum, row) => sum + amount(row), 0);
}

function fteForCenter(staff: StaffMember[], centerId: string) {
  return staff.reduce(
    (sum, person) => sum + person.assignments.filter((assignment) => assignment.profitCenterId === centerId).reduce((inner, assignment) => inner + assignment.fte, 0),
    0,
  );
}

function salaryForCenter(staff: StaffMember[], centerId: string, months: number) {
  return staff.reduce((sum, person) => {
    const assignedFte = person.assignments.reduce((inner, assignment) => inner + assignment.fte, 0);
    if (!assignedFte) return sum;
    const centerFte = person.assignments.filter((assignment) => assignment.profitCenterId === centerId).reduce((inner, assignment) => inner + assignment.fte, 0);
    return sum + person.monthlyCost * months * (centerFte / assignedFte);
  }, 0);
}

function allocationWeights(profile: BusinessProfile, rows: TrialBalanceRow[], centers: ProfitCenter[], staff: StaffMember[]) {
  const revenueByCenter = centers.map((center) => assignedAmount(rows, center.id, ["revenue"]) + numeric(center.manualRevenue) + numeric(center.studentCount) * numeric(center.averageRevenueRate));
  const totalRevenue = revenueByCenter.reduce((sum, value) => sum + value, 0);
  const totalFte = centers.reduce((sum, center) => sum + fteForCenter(staff, center.id), 0);
  const totalManual = centers.reduce((sum, center) => sum + Math.max(numeric(center.allocationWeight), 0), 0);

  return centers.map((center, index) => {
    if (profile.allocationBase === "revenue" && totalRevenue) return revenueByCenter[index] / totalRevenue;
    if (profile.allocationBase === "fte" && totalFte) return fteForCenter(staff, center.id) / totalFte;
    if (profile.allocationBase === "headcount" && staff.length) {
      const headcount = staff.filter((person) => person.assignments.some((assignment) => assignment.profitCenterId === center.id && assignment.fte > 0)).length;
      return headcount / staff.length;
    }
    if (profile.allocationBase === "manual" && totalManual) return Math.max(numeric(center.allocationWeight), 0) / totalManual;
    return centers.length ? 1 / centers.length : 0;
  });
}

function buildUnitFinancials(profile: BusinessProfile, rows: TrialBalanceRow[], centers: ProfitCenter[], staff: StaffMember[]) {
  const months = periodMonths(profile);
  const weights = allocationWeights(profile, rows, centers, staff);
  const revenuePool = unassignedAmount(rows, ["revenue"]);
  const otherIncomePool = unassignedAmount(rows, ["other-income"]);
  const directPool = unassignedAmount(rows, ["direct-cost"]);
  const peoplePool = unassignedAmount(rows, ["people-cost"]);
  const sharedOpexPool = unassignedAmount(rows, ["operating-expense"]);
  const sharedFinanceTaxPool = unassignedAmount(rows, ["finance-cost", "tax"]);
  const useManualRevenue = centers.some((center) => numeric(center.manualRevenue) > 0);
  const useManualDirectCost = centers.some((center) => numeric(center.manualDirectCost) > 0);

  return centers.map<UnitFinancials>((center, index) => ({
    center,
    driverRow: index + 2,
    ledgerRevenue: useManualRevenue ? numeric(center.manualRevenue) : assignedAmount(rows, center.id, ["revenue"]) + revenuePool * weights[index],
    driverRevenue: useManualRevenue ? 0 : driverRevenue(center),
    otherIncome: assignedAmount(rows, center.id, ["other-income"]) + otherIncomePool * weights[index],
    directCost: useManualDirectCost ? numeric(center.manualDirectCost) : assignedAmount(rows, center.id, ["direct-cost"]) + directPool * weights[index],
    driverDirectCost: useManualDirectCost ? 0 : driverDirectCost(center),
    peopleCost: salaryForCenter(staff, center.id, months) || assignedAmount(rows, center.id, ["people-cost"]) + peoplePool * weights[index],
    financeTax: assignedAmount(rows, center.id, ["finance-cost", "tax"]) + sharedFinanceTaxPool * weights[index],
    sharedWeight: weights[index],
    sharedOpex: sharedOpexPool * weights[index],
    priorRevenue: numeric(center.priorRevenue),
    priorCost: numeric(center.priorDirectCost),
  }));
}

function totalRevenue(unit: UnitFinancials) {
  return unit.ledgerRevenue + unit.driverRevenue + unit.otherIncome;
}

function totalCost(unit: UnitFinancials) {
  return unit.directCost + unit.driverDirectCost + unit.peopleCost + unit.sharedOpex + unit.financeTax;
}

function ebitda(unit: UnitFinancials) {
  return totalRevenue(unit) - totalCost(unit);
}

function margin(unit: UnitFinancials) {
  const revenue = totalRevenue(unit);
  return revenue ? ebitda(unit) / revenue : 0;
}

function centerNameMatch(accountName: string, centerName: string) {
  const account = accountName.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
  const center = centerName.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
  if (!account || !center) return false;
  if (account.includes(center) || center.includes(account)) return true;

  const centerTokens = center.split(" ").filter((token) => token.length > 2);
  if (!centerTokens.length) return false;
  return centerTokens.every((token) => account.includes(token));
}

function receivableForCenter(rows: TrialBalanceRow[], centerName: string) {
  return rows
    .filter((row) => row.category === "current-asset" && /sundry debtors/i.test(row.accountGroup) && centerNameMatch(row.accountName, centerName))
    .reduce((sum, row) => sum + amount(row), 0);
}

function payableTotal(rows: TrialBalanceRow[]) {
  return rows.filter((row) => row.category === "current-liability").reduce((sum, row) => sum + amount(row), 0);
}

function cashAndBankTotal(rows: TrialBalanceRow[]) {
  return rows.filter((row) => row.category === "current-asset" && /bank|fd|cash/i.test(row.accountName)).reduce((sum, row) => sum + amount(row), 0);
}

function setAutoFilter(sheet: XLSX.WorkSheet, rows: number, cols: number) {
  if (rows > 1 && cols > 1) sheet["!autofilter"] = { ref: `A1:${col(cols - 1)}${rows}` };
}

function appendCoverSheet(workbook: XLSX.WorkBook, profile: BusinessProfile, financials: UnitFinancials[], rows: TrialBalanceRow[]) {
  const revenue = financials.reduce((sum, unit) => sum + totalRevenue(unit), 0);
  const cost = financials.reduce((sum, unit) => sum + totalCost(unit), 0);
  const receivables = rows.filter((row) => row.category === "current-asset" && /sundry debtors/i.test(row.accountGroup)).reduce((sum, row) => sum + amount(row), 0);
  const sheet = makeSheet([
    ["Client MIS Pack"],
    [],
    ["Business", profile.businessName || profile.legalEntity],
    ["Legal entity", profile.legalEntity],
    ["Period", dateRange(profile)],
    ["Currency", profile.currency],
    ["Prepared from", "Trial balance + engagement demo allocation"],
    [],
    ["Headline"],
    ["Revenue", revenue],
    ["Total cost", cost],
    ["Operating surplus", revenue - cost],
    ["Operating margin", revenue ? (revenue - cost) / revenue : 0],
    ["Client receivables", receivables],
    ["Cash and bank", cashAndBankTotal(rows)],
    ["Current liabilities", payableTotal(rows)],
    [],
    ["Pack contents"],
    ["1", "Management Summary"],
    ["2", "Engagement P&L"],
    ["3", "Client Scorecard"],
    ["4", "Cost Pool Allocation"],
    ["5", "Working Capital"],
    ["6", "Assumptions & Data Quality"],
  ]);
  sheet["B13"] = { ...(sheet["B13"] || {}), z: percentFormat };
  sheet["!cols"] = [{ wch: 26 }, { wch: 42 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Cover");
}

function appendManagementSummary(workbook: XLSX.WorkBook, profile: BusinessProfile, financials: UnitFinancials[], rows: TrialBalanceRow[], issues: WorkbookIssue[]) {
  const revenue = financials.reduce((sum, unit) => sum + totalRevenue(unit), 0);
  const direct = financials.reduce((sum, unit) => sum + unit.directCost + unit.driverDirectCost, 0);
  const people = financials.reduce((sum, unit) => sum + unit.peopleCost, 0);
  const overhead = financials.reduce((sum, unit) => sum + unit.sharedOpex, 0);
  const financeTax = financials.reduce((sum, unit) => sum + unit.financeTax, 0);
  const cost = direct + people + overhead + financeTax;
  const surplus = revenue - cost;
  const topRevenue = [...financials].sort((a, b) => totalRevenue(b) - totalRevenue(a))[0];
  const bestMargin = [...financials].sort((a, b) => margin(b) - margin(a))[0];
  const weakestMargin = [...financials].sort((a, b) => margin(a) - margin(b))[0];
  const receivables = rows.filter((row) => row.category === "current-asset" && /sundry debtors/i.test(row.accountGroup)).reduce((sum, row) => sum + amount(row), 0);

  const sheet = makeSheet([
    [`Management Summary (${dateRange(profile)})`],
    [],
    ["KPI", "Amount / value", "% of revenue", "Commentary"],
    ["Revenue", revenue, 1, "Engagement revenue loaded from demo allocation. Replace with invoice-level revenue for final client pack."],
    ["Direct delivery cost", direct, revenue ? direct / revenue : 0, "Consultancy/vendor delivery costs traced to engagements."],
    ["People cost", people, revenue ? people / revenue : 0, "Allocated from staffing matrix using FTE weights."],
    ["Shared overhead", overhead, revenue ? overhead / revenue : 0, "Allocated using selected shared-cost basis."],
    ["Finance and tax", financeTax, revenue ? financeTax / revenue : 0, "Finance and statutory cost allocation."],
    ["Total cost", cost, revenue ? cost / revenue : 0, "Total operating cost after allocation."],
    ["Operating surplus", surplus, revenue ? surplus / revenue : 0, "Engagement-level surplus after allocated costs."],
    ["Client receivables", receivables, revenue ? receivables / revenue : 0, "Debtor balance from trial balance, matched to engagements where possible."],
    ["Cash and bank", cashAndBankTotal(rows), "", "Corporate cash/bank/FD balance."],
    ["Current liabilities", payableTotal(rows), "", "Statutory dues, provisions, creditors, and employee payables."],
    [],
    ["Management observations", "", "", ""],
    ["Largest revenue engagement", topRevenue?.center.name || "", topRevenue ? totalRevenue(topRevenue) : 0, ""],
    ["Best margin engagement", bestMargin?.center.name || "", bestMargin ? margin(bestMargin) : 0, ""],
    ["Lowest margin engagement", weakestMargin?.center.name || "", weakestMargin ? margin(weakestMargin) : 0, ""],
    ["Open issues", issues.length, "", issues.length ? "Review Assumptions & Data Quality." : "No blocking issues generated by the app."],
  ]);
  [4, 5, 6, 7, 8, 9, 10, 11, 17, 18].forEach((row) => {
    if (sheet[`C${row}`]) sheet[`C${row}`] = { ...(sheet[`C${row}`] || {}), z: percentFormat };
  });
  sheet["!cols"] = [{ wch: 28 }, { wch: 18 }, { wch: 14 }, { wch: 88 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Management Summary");
}

function appendClientReadyEngagementPnl(workbook: XLSX.WorkBook, profile: BusinessProfile, financials: UnitFinancials[], staff: StaffMember[], rows: TrialBalanceRow[]) {
  const body: Array<Array<string | number>> = financials.map((unit) => {
    const revenue = totalRevenue(unit);
    const direct = unit.directCost + unit.driverDirectCost;
    const grossProfit = revenue - direct;
    const cost = totalCost(unit);
    const fte = fteForCenter(staff, unit.center.id);
    const receivable = receivableForCenter(rows, unit.center.name);
    return [
      unit.center.name,
      unit.center.segment,
      unit.center.owner,
      unit.center.revenueDriver,
      revenue,
      direct,
      grossProfit,
      revenue ? grossProfit / revenue : 0,
      unit.peopleCost,
      unit.sharedOpex,
      unit.financeTax,
      cost,
      revenue - cost,
      revenue ? (revenue - cost) / revenue : 0,
      numeric(unit.center.studentCount),
      fte,
      revenue ? receivable / revenue : 0,
      unit.center.notes,
    ];
  });
  const totals: number[] = [];
  body.forEach((row) => {
    [4, 5, 6, 8, 9, 10, 11, 12, 14, 15].forEach((index) => {
      totals[index] = Number(totals[index] || 0) + Number(row[index] || 0);
    });
  });
  const sheet = makeSheet([
    [`Engagement P&L (${dateRange(profile)})`],
    [],
    [
      "Engagement",
      "Service line",
      "Owner",
      "Fee model",
      "Revenue",
      "Direct cost",
      "Gross profit",
      "Gross margin",
      "People cost",
      "Shared overhead",
      "Finance / tax",
      "Total cost",
      "Operating surplus",
      "Operating margin",
      "Billable hours",
      "Assigned FTE",
      "Receivable days proxy",
      "Notes",
    ],
    ...body,
    [
      "Total",
      "",
      "",
      "",
      totals[4] || 0,
      totals[5] || 0,
      totals[6] || 0,
      totals[4] ? (totals[6] || 0) / totals[4] : 0,
      totals[8] || 0,
      totals[9] || 0,
      totals[10] || 0,
      totals[11] || 0,
      totals[12] || 0,
      totals[4] ? (totals[12] || 0) / totals[4] : 0,
      totals[14] || 0,
      totals[15] || 0,
      "",
      "",
    ],
  ]);
  const rowCount = body.length + 4;
  setAutoFilter(sheet, rowCount, 18);
  sheet["!cols"] = [
    { wch: 24 },
    { wch: 30 },
    { wch: 18 },
    { wch: 30 },
    ...Array.from({ length: 10 }, () => ({ wch: 16 })),
    { wch: 16 },
    { wch: 14 },
    { wch: 18 },
    { wch: 70 },
  ];
  XLSX.utils.book_append_sheet(workbook, sheet, "Engagement P&L");
}

function appendClientScorecard(workbook: XLSX.WorkBook, financials: UnitFinancials[], staff: StaffMember[], rows: TrialBalanceRow[]) {
  const totalRev = financials.reduce((sum, unit) => sum + totalRevenue(unit), 0);
  const sheet = makeSheet([
    ["Client Scorecard"],
    [],
    ["Client", "Revenue share", "Margin", "Revenue", "Operating surplus", "Assigned FTE", "Revenue / FTE", "Receivables", "AR / revenue", "Status"],
    ...financials.map((unit) => {
      const revenue = totalRevenue(unit);
      const fte = fteForCenter(staff, unit.center.id);
      const ar = receivableForCenter(rows, unit.center.name);
      const unitMargin = margin(unit);
      return [
        unit.center.name,
        totalRev ? revenue / totalRev : 0,
        unitMargin,
        revenue,
        ebitda(unit),
        fte,
        fte ? revenue / fte : 0,
        ar,
        revenue ? ar / revenue : 0,
        unitMargin < 0.25 ? "Needs review" : unitMargin < 0.35 ? "Watch" : "Healthy",
      ];
    }),
  ]);
  setAutoFilter(sheet, financials.length + 3, 10);
  sheet["!cols"] = [{ wch: 24 }, { wch: 14 }, { wch: 12 }, { wch: 16 }, { wch: 18 }, { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Client Scorecard");
}

function appendCostPoolAllocation(workbook: XLSX.WorkBook, rows: TrialBalanceRow[], financials: UnitFinancials[]) {
  const pools = [
    ["Direct costs not tagged to client", unassignedAmount(rows, ["direct-cost"])],
    ["People cost from staffing matrix", financials.reduce((sum, unit) => sum + unit.peopleCost, 0)],
    ["Shared operating expense", financials.reduce((sum, unit) => sum + unit.sharedOpex, 0)],
    ["Finance and tax", financials.reduce((sum, unit) => sum + unit.financeTax, 0)],
  ];
  const sheet = makeSheet([
    ["Cost Pool Allocation"],
    [],
    ["Cost pool", "Pool amount", ...financials.map((unit) => unit.center.name), "Total allocated"],
    ...pools.map(([label, pool]) => [
      label,
      pool,
      ...financials.map((unit) => {
        if (label === "People cost from staffing matrix") return unit.peopleCost;
        if (label === "Shared operating expense") return unit.sharedOpex;
        if (label === "Finance and tax") return unit.financeTax;
        return unit.directCost;
      }),
      financials.reduce((sum, unit) => {
        if (label === "People cost from staffing matrix") return sum + unit.peopleCost;
        if (label === "Shared operating expense") return sum + unit.sharedOpex;
        if (label === "Finance and tax") return sum + unit.financeTax;
        return sum + unit.directCost;
      }, 0),
    ]),
  ]);
  sheet["!cols"] = [{ wch: 32 }, { wch: 16 }, ...financials.map(() => ({ wch: 18 })), { wch: 18 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Cost Pool Allocation");
}

function appendAssumptionsQuality(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], staff: StaffMember[], issues: WorkbookIssue[]) {
  const unknownRows = rows.filter((row) => row.category === "unknown");
  const sheet = makeSheet([
    ["Assumptions & Data Quality"],
    [],
    ["Area", "Status", "Detail", "Client-ready action"],
    ["Accounting source", rows.length ? "Loaded" : "Missing", `${rows.length} ledger rows imported for ${dateRange(profile)}.`, "Use final TB exported after books close."],
    ["Revenue split", "Demo assumption", "Manual engagement revenue is loaded from debtor-balance proxy.", "Replace with invoice-wise revenue by client/project."],
    ["Direct cost split", "Demo assumption", "Manual engagement direct cost is allocated using debtor-balance proxy.", "Replace with vendor/project tagging."],
    ["People allocation", staff.length ? "Loaded" : "Missing", `${staff.length} staffing rows loaded. Client columns are FTE/time weights.`, "Replace pods with named employees and monthly CTC."],
    ["Receivables", "Partial", "Debtor balances are matched to engagement names where possible.", "Add ageing buckets, invoice dates, collection status, and WIP."],
    ["Unknown ledgers", unknownRows.length ? "Needs review" : "Clear", `${unknownRows.length} ledgers are unclassified.`, "Classify every unknown ledger before client circulation."],
    ...issues.map((issue) => ["Generated issue", issue.severity, issue.detail, issue.label]),
  ]);
  sheet["!cols"] = [{ wch: 24 }, { wch: 18 }, { wch: 80 }, { wch: 60 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Assumptions & Data Quality");
}

function applyTableFormatting(sheet: XLSX.WorkSheet, rows: number, cols: number) {
  setAutoFilter(sheet, rows, cols);
  sheet["!freeze"] = { xSplit: 0, ySplit: 1 };
}

function appendControlPanel(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], centers: ProfitCenter[], staff: StaffMember[], questions: GeneratedQuestion[], issues: WorkbookIssue[]) {
  const totalDebit = rows.reduce((sum, row) => sum + row.debit, 0);
  const totalCredit = rows.reduce((sum, row) => sum + row.credit, 0);
  const revenueRows = rows.filter((row) => row.category === "revenue");
  const debtorRows = rows.filter((row) => row.category === "current-asset" && /sundry debtors/i.test(row.accountGroup));
  const payrollRows = rows.filter((row) => row.category === "people-cost");
  const directRows = rows.filter((row) => row.category === "direct-cost");
  const missing: string[] = [];

  if (!rows.length) missing.push("Trial balance / ledger source");
  if (!revenueRows.length) missing.push("Ledger revenue mapping");
  if (!debtorRows.length) missing.push("Client receivables / debtor schedule");
  if (!payrollRows.length) missing.push("Payroll ledgers");
  if (!staff.length) missing.push("Staffing matrix / employee roster");
  if (!directRows.length && !centers.some((center) => numeric(center.manualDirectCost) > 0)) missing.push("Vendor or engagement direct-cost tagging");
  if (questions.some((question) => question.priority === "high")) missing.push("Open high-priority question answers");

  const sheet = makeSheet([
    ["00 Control Panel"],
    [],
    ["Field", "Value", "Client-ready interpretation"],
    ["Business", profile.businessName || profile.legalEntity, ""],
    ["Legal entity", profile.legalEntity, ""],
    ["Period", dateRange(profile), ""],
    ["Currency", profile.currency, ""],
    ["Export timestamp", new Date().toISOString(), ""],
    ["Workbook mode", "AI-assisted MIS production pack", "Use the app as source of truth; use Claude/Shortcut only for Excel-native finishing."],
    ["TB status", rows.length ? "Loaded" : "Missing", rows.length ? `${rows.length} ledger rows imported.` : "Demo-only export. Upload TB before circulating."],
    ["TB debit", totalDebit, ""],
    ["TB credit", totalCredit, ""],
    ["TB imbalance", Math.abs(totalDebit - totalCredit), Math.abs(totalDebit - totalCredit) > 1 ? "Review source TB." : "Balanced within tolerance."],
    ["Profit centers", centers.length, ""],
    ["Staff rows", staff.length, ""],
    ["Open checks", issues.length, issues.length ? "Review 10 Assumptions & Data Quality." : "No generated issues."],
    [],
    ["Data completeness", "Status", "Detail"],
    ["Trial balance", rows.length ? "Loaded" : "Missing", rows.length ? "Raw ledger schedules are populated." : "Raw ledger and working-capital schedules are intentionally blank/zero."],
    ["Revenue bridge", revenueRows.length || centers.some((center) => numeric(center.manualRevenue) > 0) ? "Available" : "Missing", "Manual engagement revenue is marked as proxy where invoice register is unavailable."],
    ["People allocation", staff.length ? "Available" : "Missing", staff.length ? "Staff/pod rows allocated by FTE weights." : "Add employee roster or staffing pods."],
    ["Working capital", debtorRows.length ? "Partial" : "Missing", debtorRows.length ? "Debtors are matched to clients by ledger name where possible." : "Upload debtor/AR ageing for client-ready pack."],
    ["Direct costs", directRows.length || centers.some((center) => numeric(center.manualDirectCost) > 0) ? "Available" : "Missing", "Direct costs should be vendor/project-tagged for final MIS."],
    [],
    ["Missing schedules", missing.length ? missing.join("; ") : "None flagged by current rules", ""],
    [],
    ["Third-party AI warning", "Using Claude for Excel or Shortcut can transmit workbook financial data, salary/pod costs, ledgers, receivables, and assumptions to that provider.", "Use only with client/firm approval and the right data protection settings."],
  ]);
  setFormula(sheet, "B13", "ABS(B11-B12)", moneyFormat, Math.abs(totalDebit - totalCredit));
  sheet["!cols"] = [{ wch: 28 }, { wch: 40 }, { wch: 90 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "00 Control Panel");
}

function appendProductionManagementSummary(workbook: XLSX.WorkBook, profile: BusinessProfile, financials: UnitFinancials[], rows: TrialBalanceRow[], issues: WorkbookIssue[]) {
  const revenue = financials.reduce((sum, unit) => sum + totalRevenue(unit), 0);
  const direct = financials.reduce((sum, unit) => sum + unit.directCost + unit.driverDirectCost, 0);
  const people = financials.reduce((sum, unit) => sum + unit.peopleCost, 0);
  const overhead = financials.reduce((sum, unit) => sum + unit.sharedOpex, 0);
  const financeTax = financials.reduce((sum, unit) => sum + unit.financeTax, 0);
  const cost = direct + people + overhead + financeTax;
  const receivables = rows.filter((row) => row.category === "current-asset" && /sundry debtors/i.test(row.accountGroup)).reduce((sum, row) => sum + amount(row), 0);
  const currentLiabilities = payableTotal(rows);
  const cash = cashAndBankTotal(rows);
  const topRevenue = [...financials].sort((a, b) => totalRevenue(b) - totalRevenue(a))[0];
  const bestMargin = [...financials].sort((a, b) => margin(b) - margin(a))[0];
  const lowMargin = [...financials].sort((a, b) => margin(a) - margin(b))[0];

  const sheet = makeSheet([
    [`01 Management Summary (${dateRange(profile)})`],
    [],
    ["Metric", "Amount / value", "% of revenue", "Source schedule", "Commentary"],
    ["Revenue", revenue, revenue ? 1 : 0, "03 Revenue Register", rows.length ? "Ties to ledger revenue and/or explicit engagement revenue bridge." : "Demo-only; no TB loaded."],
    ["Direct delivery cost", direct, revenue ? direct / revenue : 0, "05 Vendor Direct Cost Register", "Vendor/project direct cost and manual engagement direct cost."],
    ["People cost", people, revenue ? people / revenue : 0, "04 People Cost Register", "Allocated from staffing matrix."],
    ["Shared overhead", overhead, revenue ? overhead / revenue : 0, "06 Shared Cost Pools", "Allocated using selected shared cost basis."],
    ["Finance and tax", financeTax, revenue ? financeTax / revenue : 0, "06 Shared Cost Pools", "Finance and tax cost pools."],
    ["Total cost", cost, revenue ? cost / revenue : 0, "02 Engagement P&L", "Fully loaded engagement cost."],
    ["Operating surplus", revenue - cost, revenue ? (revenue - cost) / revenue : 0, "02 Engagement P&L", "Revenue less total allocated cost."],
    ["Cash and bank", cash, "", "07 Working Capital", "Corporate liquidity from TB bank/FD/cash ledgers."],
    ["Client receivables", receivables, revenue ? receivables / revenue : 0, "07 Working Capital", "Client debtor balances matched to engagements where possible."],
    ["Current liabilities", currentLiabilities, "", "07 Working Capital", "Statutory dues, provisions, creditors and employee payables."],
    ["Net working capital proxy", receivables + cash - currentLiabilities, "", "07 Working Capital", "Proxy because ageing/WIP schedules are not uploaded."],
    [],
    ["Management observation", "Value", "Interpretation", "", ""],
    ["Largest revenue engagement", topRevenue?.center.name || "", topRevenue ? totalRevenue(topRevenue) : 0, "", ""],
    ["Lowest margin engagement", lowMargin?.center.name || "", lowMargin ? margin(lowMargin) : 0, "", ""],
    ["Open data-quality checks", issues.length, "", "10 Assumptions & Data Quality", "Resolve before client circulation."],
  ]);

  [4, 5, 6, 7, 8, 9, 10, 12, 18].forEach((row) => {
    if (sheet[`C${row}`]) sheet[`C${row}`] = { ...(sheet[`C${row}`] || {}), z: percentFormat };
  });
  const pnlLastRow = financials.length + 3;
  setFormula(sheet, "B4", `SUM('02 Engagement P&L'!E4:E${pnlLastRow})`, moneyFormat, revenue);
  setFormula(sheet, "C4", "IF($B$4=0,0,B4/$B$4)", percentFormat, revenue ? 1 : 0);
  setFormula(sheet, "B5", `SUM('02 Engagement P&L'!H4:H${pnlLastRow})`, moneyFormat, direct);
  setFormula(sheet, "C5", "IF($B$4=0,0,B5/$B$4)", percentFormat, revenue ? direct / revenue : 0);
  setFormula(sheet, "B6", `SUM('02 Engagement P&L'!K4:K${pnlLastRow})`, moneyFormat, people);
  setFormula(sheet, "C6", "IF($B$4=0,0,B6/$B$4)", percentFormat, revenue ? people / revenue : 0);
  setFormula(sheet, "B7", `SUM('02 Engagement P&L'!L4:L${pnlLastRow})`, moneyFormat, overhead);
  setFormula(sheet, "C7", "IF($B$4=0,0,B7/$B$4)", percentFormat, revenue ? overhead / revenue : 0);
  setFormula(sheet, "B8", `SUM('02 Engagement P&L'!M4:M${pnlLastRow})`, moneyFormat, financeTax);
  setFormula(sheet, "C8", "IF($B$4=0,0,B8/$B$4)", percentFormat, revenue ? financeTax / revenue : 0);
  setFormula(sheet, "B9", "SUM(B5:B8)", moneyFormat, cost);
  setFormula(sheet, "C9", "IF($B$4=0,0,B9/$B$4)", percentFormat, revenue ? cost / revenue : 0);
  setFormula(sheet, "B10", "B4-B9", moneyFormat, revenue - cost);
  setFormula(sheet, "C10", "IF($B$4=0,0,B10/$B$4)", percentFormat, revenue ? (revenue - cost) / revenue : 0);
  setFormula(sheet, "B11", `SUM('02 Engagement P&L'!T4:T${pnlLastRow})`, moneyFormat, receivables);
  setFormula(sheet, "C11", "IF($B$4=0,0,B11/$B$4)", percentFormat, revenue ? receivables / revenue : 0);
  setFormula(sheet, "B12", `SUMIFS('09 Ledger Mapping Audit'!G:G,'09 Ledger Mapping Audit'!H:H,"Current asset",'09 Ledger Mapping Audit'!C:C,"*Bank*")+SUMIFS('09 Ledger Mapping Audit'!G:G,'09 Ledger Mapping Audit'!H:H,"Current asset",'09 Ledger Mapping Audit'!C:C,"*FD*")`, moneyFormat, cash);
  setFormula(sheet, "B13", `SUMIF('09 Ledger Mapping Audit'!H:H,"Current liability",'09 Ledger Mapping Audit'!G:G)`, moneyFormat, currentLiabilities);
  setFormula(sheet, "B14", "B11+B12-B13", moneyFormat, receivables + cash - currentLiabilities);
  sheet["B16"] = { t: "s", v: topRevenue?.center.name || "", f: `INDEX('02 Engagement P&L'!A4:A${pnlLastRow},MATCH(MAX('02 Engagement P&L'!E4:E${pnlLastRow}),'02 Engagement P&L'!E4:E${pnlLastRow},0))` };
  setFormula(sheet, "C16", `MAX('02 Engagement P&L'!E4:E${pnlLastRow})`, moneyFormat, topRevenue ? totalRevenue(topRevenue) : 0);
  sheet["B17"] = { t: "s", v: bestMargin?.center.name || "", f: `INDEX('02 Engagement P&L'!A4:A${pnlLastRow},MATCH(MAX('02 Engagement P&L'!P4:P${pnlLastRow}),'02 Engagement P&L'!P4:P${pnlLastRow},0))` };
  setFormula(sheet, "C17", `MAX('02 Engagement P&L'!P4:P${pnlLastRow})`, percentFormat, bestMargin ? margin(bestMargin) : 0);
  sheet["B18"] = { t: "s", v: lowMargin?.center.name || "", f: `INDEX('02 Engagement P&L'!A4:A${pnlLastRow},MATCH(MIN('02 Engagement P&L'!P4:P${pnlLastRow}),'02 Engagement P&L'!P4:P${pnlLastRow},0))` };
  setFormula(sheet, "C18", `MIN('02 Engagement P&L'!P4:P${pnlLastRow})`, percentFormat, lowMargin ? margin(lowMargin) : 0);
  sheet["!cols"] = [{ wch: 30 }, { wch: 18 }, { wch: 14 }, { wch: 28 }, { wch: 86 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "01 Management Summary");
}

function appendProductionEngagementPnl(workbook: XLSX.WorkBook, profile: BusinessProfile, financials: UnitFinancials[], staff: StaffMember[], rows: TrialBalanceRow[]) {
  const body = financials.map((unit) => {
    const revenue = totalRevenue(unit);
    const direct = unit.directCost + unit.driverDirectCost;
    const cost = totalCost(unit);
    const receivable = receivableForCenter(rows, unit.center.name);
    const fte = fteForCenter(staff, unit.center.id);
    return [
      unit.center.name,
      unit.center.segment,
      unit.center.owner,
      unit.center.revenueDriver,
      revenue,
      unit.ledgerRevenue,
      unit.driverRevenue,
      direct,
      revenue - direct,
      revenue ? (revenue - direct) / revenue : 0,
      unit.peopleCost,
      unit.sharedOpex,
      unit.financeTax,
      cost,
      revenue - cost,
      revenue ? (revenue - cost) / revenue : 0,
      numeric(unit.center.studentCount),
      fte,
      fte ? revenue / fte : 0,
      receivable,
      revenue ? receivable / revenue : 0,
      unit.priorRevenue,
      unit.priorCost,
      unit.center.notes,
    ];
  });
  const sheet = makeSheet([
    [`02 Engagement P&L (${dateRange(profile)})`],
    [],
    [
      "Engagement",
      "Service line",
      "Owner",
      "Fee model",
      "Revenue",
      "Manual / ledger revenue",
      "Driver revenue",
      "Direct cost",
      "Gross profit",
      "Gross margin",
      "People cost",
      "Shared overhead",
      "Finance / tax",
      "Total cost",
      "Operating surplus",
      "Operating margin",
      "Billable hours",
      "Assigned FTE",
      "Revenue / FTE",
      "Receivables",
      "AR / revenue",
      "Prior revenue",
      "Prior direct cost",
      "Notes",
    ],
    ...body,
    [
      "Total",
      "",
      "",
      "",
      ...Array.from({ length: 19 }, (_, index) => (index === 5 || index === 11 || index === 17 ? "" : 0)),
      "",
    ],
  ]);
  const totalRow = body.length + 4;
  [5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 17, 18, 19, 20, 22, 23].forEach((oneBasedCol) => {
    const letter = col(oneBasedCol - 1);
    setFormula(sheet, `${letter}${totalRow}`, `SUM(${letter}4:${letter}${totalRow - 1})`);
  });
  setFormula(sheet, `J${totalRow}`, `IF(E${totalRow}=0,0,I${totalRow}/E${totalRow})`, percentFormat);
  setFormula(sheet, `P${totalRow}`, `IF(E${totalRow}=0,0,O${totalRow}/E${totalRow})`, percentFormat);
  setFormula(sheet, `U${totalRow}`, `IF(E${totalRow}=0,0,T${totalRow}/E${totalRow})`, percentFormat);
  applyTableFormatting(sheet, totalRow, 24);
  sheet["!cols"] = [{ wch: 24 }, { wch: 30 }, { wch: 20 }, { wch: 32 }, ...Array.from({ length: 19 }, () => ({ wch: 16 })), { wch: 70 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "02 Engagement P&L");
}

function appendRevenueRegister(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], financials: UnitFinancials[]) {
  const ledgerRevenueRows = rows.filter((row) => row.category === "revenue" || row.category === "other-income");
  const bridgeRows = financials.flatMap((unit) => [
    [
      "Engagement bridge",
      unit.center.name,
      unit.center.segment,
      unit.center.owner,
      unit.ledgerRevenue,
      "Manual/proxy allocation",
      unit.center.revenueDriver,
      unit.center.notes,
    ],
    ...(unit.driverRevenue
      ? [["Driver build-up", unit.center.name, unit.center.segment, unit.center.owner, unit.driverRevenue, "Billable hours x billing rate x recovery", unit.center.revenueDriver, "Used only where manual revenue is not entered."]]
      : []),
  ]);
  const sheet = makeSheet([
    [`03 Revenue Register (${dateRange(profile)})`],
    [],
    ["Source", "Engagement / ledger", "Service line / group", "Owner / sheet", "Amount", "Basis", "Fee arrangement", "Assumption / trace"],
    ...bridgeRows,
    ...ledgerRevenueRows.map((row) => ["Trial balance ledger", row.accountName, row.accountGroup, row.sourceSheet, amount(row), row.category === "revenue" ? "Operating revenue ledger" : "Other income ledger", "", `Source row ${row.raw["Source Row"] ?? ""}`]),
    ...(!ledgerRevenueRows.length ? [["Missing source", "", "", "", 0, "No revenue ledger loaded", "", "Upload TB or invoice register."]] : []),
  ]);
  applyTableFormatting(sheet, bridgeRows.length + ledgerRevenueRows.length + 4, 8);
  sheet["!cols"] = [{ wch: 22 }, { wch: 36 }, { wch: 30 }, { wch: 22 }, { wch: 16 }, { wch: 34 }, { wch: 34 }, { wch: 80 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "03 Revenue Register");
}

function appendPeopleCostRegister(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], financials: UnitFinancials[], staff: StaffMember[]) {
  const months = periodMonths(profile);
  const register = staff.flatMap((person) => {
    const totalFte = person.assignments.reduce((sum, assignment) => sum + assignment.fte, 0);
    return financials.map((unit) => {
      const fte = person.assignments.find((assignment) => assignment.profitCenterId === unit.center.id)?.fte || 0;
      const cost = totalFte ? person.monthlyCost * months * (fte / totalFte) : 0;
      return [person.name, person.role, person.department, person.monthlyCost, months, person.monthlyCost * months, unit.center.name, fte, cost, totalFte ? fte / totalFte : 0, "Staffing matrix FTE weight"];
    });
  });
  const payrollRows = rows.filter((row) => row.category === "people-cost");
  const sheet = makeSheet([
    [`04 People Cost Register (${dateRange(profile)})`],
    [],
    ["Name / ledger", "Role / group", "Department / sheet", "Monthly cost", "Months", "Period cost", "Engagement", "Assigned FTE", "Allocated cost", "% of person", "Source / note"],
    ...register,
    ...payrollRows.map((row) => [row.accountName, row.accountGroup, row.sourceSheet, "", "", amount(row), "Unassigned TB payroll pool", "", amount(row), "", `TB source row ${row.raw["Source Row"] ?? ""}`]),
    ...(!staff.length ? [["Missing roster", "", "", "", "", 0, "", "", 0, "", "Add employee roster or staffing pods."]] : []),
  ]);
  applyTableFormatting(sheet, register.length + payrollRows.length + 4, 11);
  sheet["!cols"] = [{ wch: 32 }, { wch: 24 }, { wch: 24 }, { wch: 16 }, { wch: 10 }, { wch: 16 }, { wch: 28 }, { wch: 14 }, { wch: 16 }, { wch: 14 }, { wch: 42 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "04 People Cost Register");
}

function appendVendorDirectCostRegister(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], financials: UnitFinancials[]) {
  const directCandidates = rows.filter((row) => ["direct-cost", "operating-expense"].includes(row.category) && /consult|professional|travel|lodging|boarding|conveyance|car hire|software|subscription|training|workshop|laptop|research|vendor|guidepoint|microsoft/i.test(`${row.accountName} ${row.accountGroup}`));
  const manualDirect = financials.map((unit) => ["Manual engagement direct cost", unit.center.name, unit.center.segment, unit.center.owner, unit.directCost + unit.driverDirectCost, "Engagement direct-cost bridge", unit.center.notes]);
  const sheet = makeSheet([
    [`05 Vendor / Direct Cost Register (${dateRange(profile)})`],
    [],
    ["Source", "Vendor / ledger / engagement", "Group / service line", "Owner / sheet", "Amount", "Basis", "Trace / assumption"],
    ...manualDirect,
    ...directCandidates.map((row) => ["Trial balance candidate", row.accountName, row.accountGroup, row.sourceSheet, amount(row), categoryLabels[row.category], `Source row ${row.raw["Source Row"] ?? ""}`]),
    ...(!directCandidates.length && !manualDirect.length ? [["Missing source", "", "", "", 0, "No direct-cost records loaded", "Upload vendor/project tagging."]] : []),
  ]);
  applyTableFormatting(sheet, manualDirect.length + directCandidates.length + 4, 7);
  sheet["!cols"] = [{ wch: 24 }, { wch: 40 }, { wch: 28 }, { wch: 22 }, { wch: 16 }, { wch: 30 }, { wch: 80 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "05 Vendor Direct Cost Register");
}

function appendSharedCostPoolsProduction(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], financials: UnitFinancials[]) {
  const poolRows = rows.filter((row) => !row.profitCenterId && ["operating-expense", "finance-cost", "tax", "direct-cost", "people-cost"].includes(row.category));
  const sheet = makeSheet([
    [`06 Shared Cost Pools (${allocationLabels[profile.allocationBase]})`],
    [],
    ["Ledger", "Category", "MIS group", "Pool amount", "Allocation basis", ...financials.map((unit) => unit.center.name), "Total allocated", "Trace"],
    ...poolRows.map((row) => {
      const pool = amount(row);
      return [
        row.accountName,
        categoryLabels[row.category],
        row.misGroup,
        pool,
        allocationLabels[profile.allocationBase],
        ...financials.map((unit) => pool * unit.sharedWeight),
        pool,
        `${row.sourceSheet} row ${row.raw["Source Row"] ?? ""}`,
      ];
    }),
    ...(!poolRows.length ? [["No shared cost ledgers loaded", "", "", 0, allocationLabels[profile.allocationBase], ...financials.map(() => 0), 0, rows.length ? "All costs directly mapped/manual." : "Trial balance missing."]] : []),
  ]);
  applyTableFormatting(sheet, poolRows.length + 4, 7 + financials.length);
  sheet["!cols"] = [{ wch: 38 }, { wch: 20 }, { wch: 24 }, { wch: 16 }, { wch: 22 }, ...financials.map(() => ({ wch: 16 })), { wch: 16 }, { wch: 26 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "06 Shared Cost Pools");
}

function appendWorkingCapitalProduction(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], financials: UnitFinancials[]) {
  const bsRows = rows.filter((row) => ["current-asset", "current-liability", "fixed-asset", "equity"].includes(row.category));
  const sheet = makeSheet([
    [`07 Working Capital (${dateRange(profile)})`],
    [],
    ["Account", "Category", "Group", "Debit", "Credit", "Signed balance", "Matched engagement", ...financials.map((unit) => unit.center.name), "Shared / corporate", "Trace"],
    ...bsRows.map((row) => {
      const matched = financials.find((unit) => centerNameMatch(row.accountName, unit.center.name));
      return [
        row.accountName,
        categoryLabels[row.category],
        row.accountGroup,
        row.debit,
        row.credit,
        categorySignedAmount(row),
        matched?.center.name || "",
        ...financials.map((unit) => (matched?.center.id === unit.center.id ? amount(row) : 0)),
        matched ? 0 : amount(row),
        `${row.sourceSheet} row ${row.raw["Source Row"] ?? ""}`,
      ];
    }),
    ...(!bsRows.length ? [["Trial balance missing", "", "", 0, 0, 0, "", ...financials.map(() => 0), 0, "Upload TB to populate working capital."]] : []),
  ]);
  applyTableFormatting(sheet, bsRows.length + 4, 9 + financials.length);
  sheet["!cols"] = [{ wch: 38 }, { wch: 20 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 24 }, ...financials.map(() => ({ wch: 16 })), { wch: 18 }, { wch: 26 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "07 Working Capital");
}

function appendCashFlowProduction(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], financials: UnitFinancials[]) {
  const surplus = financials.reduce((sum, unit) => sum + ebitda(unit), 0);
  const receivables = rows.filter((row) => row.category === "current-asset" && /sundry debtors/i.test(row.accountGroup)).reduce((sum, row) => sum + amount(row), 0);
  const depositsPrepaids = rows.filter((row) => row.category === "current-asset" && /deposit|prepaid|advance|itc|suspense|preliminary/i.test(`${row.accountName} ${row.accountGroup}`)).reduce((sum, row) => sum + amount(row), 0);
  const liabilities = payableTotal(rows);
  const fixedAssets = rows.filter((row) => row.category === "fixed-asset").reduce((sum, row) => sum + amount(row), 0);
  const cash = cashAndBankTotal(rows);
  const sheet = makeSheet([
    [`08 Cash Flow (${dateRange(profile)})`],
    [],
    ["Particulars", "Amount", "Formula / source", "Client-ready caveat"],
    ["Opening cash / bank", 0, "Prior-period bank schedule not uploaded", "Replace with opening cash/bank from prior MIS."],
    ["Operating surplus", surplus, "02 Engagement P&L", ""],
    ["Less: closing receivables / WIP proxy", -receivables, "07 Working Capital client debtors", "Needs invoice ageing for exact cash bridge."],
    ["Less: deposits, prepaid and advances", -depositsPrepaids, "07 Working Capital", ""],
    ["Add: current liabilities / payables", liabilities, "07 Working Capital", ""],
    ["Operating cash flow proxy", null, "SUM(B5:B8)", ""],
    ["Less: capex / fixed assets", -fixedAssets, "07 Working Capital fixed assets", ""],
    ["Financing / owner movement", 0, "Equity movement not available from single TB", "Requires prior period balances."],
    ["Closing cash / bank", cash, "07 Working Capital cash and bank", ""],
    ["Cash bridge check", null, "B4+B9+B10+B11-B12", "Should be zero only after true opening cash and movements are available."],
  ]);
  setFormula(sheet, "B9", "SUM(B5:B8)");
  setFormula(sheet, "B13", "B4+B9+B10+B11-B12");
  sheet["!cols"] = [{ wch: 34 }, { wch: 18 }, { wch: 36 }, { wch: 68 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "08 Cash Flow");
}

function rowsMatching(rows: TrialBalanceRow[], category: string, pattern?: RegExp) {
  return rows.filter((row) => row.category === category && (!pattern || pattern.test(`${row.accountName} ${row.accountGroup}`)));
}

function pbtFromTrialBalance(rows: TrialBalanceRow[]) {
  const revenue = rowsMatching(rows, "revenue").reduce((sum, row) => sum + amount(row), 0);
  const otherIncome = rowsMatching(rows, "other-income").reduce((sum, row) => sum + amount(row), 0);
  const expenses = rows
    .filter((row) => ["direct-cost", "people-cost", "operating-expense", "finance-cost", "tax"].includes(row.category))
    .reduce((sum, row) => sum + amount(row), 0);
  return revenue + otherIncome - expenses;
}

function sumValues(values: number[]) {
  return values.reduce((sum, value) => sum + value, 0);
}

function averagePositive(values: number[]) {
  const positive = values.filter((value) => value > 0);
  return positive.length ? sumValues(positive) / positive.length : 0;
}

function flowText(txn: BankTransaction) {
  return `${txn.category} ${txn.fundFlowGroup} ${txn.narration}`.toLowerCase();
}

function actualFundFlowTransactions(bankTransactions: BankTransaction[]) {
  const statementRows = bankTransactions.filter((txn) => txn.sourceType === "bank-statement");
  return statementRows.length ? statementRows : bankTransactions;
}

function monthlyBankValues(
  months: string[],
  transactions: BankTransaction[],
  predicate: (txn: BankTransaction) => boolean,
  selector: (txn: BankTransaction) => number,
) {
  const values = new Map(months.map((month) => [month, 0]));
  transactions.forEach((txn) => {
    const key = monthKey(txn.date);
    if (!key || !values.has(key) || !predicate(txn)) return;
    values.set(key, (values.get(key) || 0) + selector(txn));
  });
  return months.map((month) => values.get(month) || 0);
}

function latestBankBalance(transactions: BankTransaction[]) {
  const preferred = transactions.some((txn) => txn.sourceType === "bank-statement" && txn.balance)
    ? transactions.filter((txn) => txn.sourceType === "bank-statement")
    : transactions;
  const byAccount = new Map<string, BankTransaction>();

  preferred
    .filter((txn) => txn.balance)
    .forEach((txn) => {
      const key = `${txn.sourceType}-${txn.sourceFileName || ""}-${txn.accountName || txn.bankName}`;
      const existing = byAccount.get(key);
      if (!existing || txn.date > existing.date) byAccount.set(key, txn);
    });

  return Array.from(byAccount.values()).reduce((sum, txn) => sum + txn.balance, 0);
}

function buildFundFlowModel(profile: BusinessProfile, rows: TrialBalanceRow[], bankTransactions: BankTransaction[]): FundFlowModel {
  const months = monthLabels(profile);
  const projectionMonths = projectionMonthLabels(profile);
  const actualTransactions = actualFundFlowTransactions(bankTransactions).filter((txn) => !txn.isInterAccountTransfer);
  const internalTransfers = actualFundFlowTransactions(bankTransactions).filter((txn) => txn.isInterAccountTransfer);
  const matches = (pattern: RegExp) => (txn: BankTransaction) => pattern.test(flowText(txn));
  const payment = (txn: BankTransaction) => txn.debit;
  const receipt = (txn: BankTransaction) => txn.credit;
  const receiptRows: FundFlowModel["receiptRows"] = [
    [
      "Client / operating receipts",
      monthlyBankValues(
        months,
        actualTransactions,
        (txn) => txn.credit > 0 && matches(/client|operating receipt|receipt|fee|invoice|collection/)(txn) && !matches(/unclassified|gst|tds|tax|statutory|fd|treasury|interest|loan|owner|capital|advance|reimbursement/)(txn),
        receipt,
      ),
      "Bank receipts tagged from statement/ledger narrations.",
    ],
    ["GST / tax refunds / statutory receipts", monthlyBankValues(months, actualTransactions, (txn) => txn.credit > 0 && matches(/gst|tds|tax|statutory/)(txn), receipt), "Credits linked to GST/TDS/tax/statutory narrations."],
    ["FD maturity / interest receipts", monthlyBankValues(months, actualTransactions, (txn) => txn.credit > 0 && matches(/fd|treasury|interest/)(txn), receipt), "FD maturity, sweep-back, treasury and interest credits."],
    ["Loan / owner receipts", monthlyBankValues(months, actualTransactions, (txn) => txn.credit > 0 && matches(/loan|owner|capital|financing/)(txn), receipt), "Borrowing, capital, director or owner receipts."],
    ["Advances / reimbursements received", monthlyBankValues(months, actualTransactions, (txn) => txn.credit > 0 && matches(/advance|reimbursement/)(txn), receipt), "Recoveries, advances, deposits or reimbursement receipts."],
    ["Unclassified receipts", monthlyBankValues(months, actualTransactions, (txn) => txn.credit > 0 && matches(/unclassified receipt/)(txn), receipt), "Review these before client circulation."],
  ];
  const paymentRows: FundFlowModel["paymentRows"] = [
    ["GST / TDS / PF / ESIC payments", monthlyBankValues(months, actualTransactions, matches(/gst|tds|tax|statutory|pf|esic/), payment), "Statutory cash outflow identified from bank narrations."],
    ["Salary / bonus / payroll", monthlyBankValues(months, actualTransactions, matches(/salary|bonus|payroll|people/), payment), "Payroll and bonus cash outflow."],
    ["Vendor / operating payments", monthlyBankValues(months, actualTransactions, matches(/vendor|operating payment|rent|office|software|audit|legal|travel|charges/), payment), "Operating vendor and overhead payments."],
    ["Capex / fixed assets", monthlyBankValues(months, actualTransactions, matches(/capex|fixed assets|asset|equipment|laptop|computer/), payment), "Asset purchases and capital expenditure."],
    ["FD placement / treasury outflow", monthlyBankValues(months, actualTransactions, matches(/fd placement|treasury outflow|fixed deposit|fd|treasury/), payment), "FD placements, sweep-outs and treasury movement."],
    ["Loan / owner payments", monthlyBankValues(months, actualTransactions, matches(/loan|owner|capital|financing|emi|drawing/), payment), "Loan repayments, owner payouts and financing outflow."],
    ["Advances / reimbursements paid", monthlyBankValues(months, actualTransactions, matches(/advance|reimbursement|deposit|prepaid/), payment), "Advances, deposits, prepaids and reimbursements paid."],
    ["Unclassified payments", monthlyBankValues(months, actualTransactions, (txn) => txn.debit > 0 && matches(/unclassified payment/)(txn), payment), "Review these before client circulation."],
  ];
  const internalTransferValues = monthlyBankValues(months, internalTransfers, () => true, (txn) => txn.debit + txn.credit);
  const receiptTotals = months.map((_, index) => receiptRows.reduce((sum, row) => sum + row[1][index], 0));
  const paymentTotals = months.map((_, index) => paymentRows.reduce((sum, row) => sum + row[1][index], 0));
  const closingCash = cashAndBankTotal(rows) || latestBankBalance(actualFundFlowTransactions(bankTransactions));
  const openingFirst = actualTransactions.length ? closingCash - sumValues(receiptTotals) + sumValues(paymentTotals) : 0;
  const openingBalances: number[] = [];
  const closingBalances: number[] = [];
  let running = openingFirst;

  months.forEach((_, index) => {
    openingBalances.push(running);
    running += receiptTotals[index] - paymentTotals[index];
    closingBalances.push(running);
  });

  if (!actualTransactions.length && closingCash) {
    closingBalances[closingBalances.length - 1] = closingCash;
  }

  return { months, projectionMonths, openingBalances, closingBalances, closingCash, receiptRows, paymentRows, internalTransferValues };
}

function balanceLineRows(rows: TrialBalanceRow[], category: string, pattern?: RegExp) {
  return rowsMatching(rows, category, pattern).map((row) => ["", `  - ${row.accountName}`, "", "", amount(row), row.accountGroup, row.raw["Source Row"] ?? ""]);
}

function appendBalanceSheetProduction(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[]) {
  const pbt = pbtFromTrialBalance(rows);
  const liabilities = [
    ["1", "Capital Account", "", "", rowsMatching(rows, "equity").reduce((sum, row) => sum + amount(row), 0), "", ""],
    ...balanceLineRows(rows, "equity"),
    ["2", "Taxes / Statutory Payable", "", "", rowsMatching(rows, "current-liability", /gst|tds|pf|esi|esic|tax|duties/i).reduce((sum, row) => sum + amount(row), 0), "", ""],
    ...balanceLineRows(rows, "current-liability", /gst|tds|pf|esi|esic|tax|duties/i),
    ["3", "Accounts Payable / Creditors", "", "", rowsMatching(rows, "current-liability", /creditor|payable|audit|microsoft|guidepoint|manhas|yellow|profectus|abcom/i).reduce((sum, row) => sum + amount(row), 0), "", ""],
    ...balanceLineRows(rows, "current-liability", /creditor|payable|audit|microsoft|guidepoint|manhas|yellow|profectus|abcom/i),
    ["4", "Customer Advances / Credit Debtors", "", "", rowsMatching(rows, "current-liability", /debtor|advance|reliant/i).reduce((sum, row) => sum + amount(row), 0), "", ""],
    ...balanceLineRows(rows, "current-liability", /debtor|advance|reliant/i),
    ["5", "Other Current Liabilities", "", "", rowsMatching(rows, "current-liability").filter((row) => !/gst|tds|pf|esi|esic|tax|duties|creditor|payable|audit|microsoft|guidepoint|manhas|yellow|profectus|abcom|debtor|advance|reliant/i.test(`${row.accountName} ${row.accountGroup}`)).reduce((sum, row) => sum + amount(row), 0), "", ""],
    ...rowsMatching(rows, "current-liability").filter((row) => !/gst|tds|pf|esi|esic|tax|duties|creditor|payable|audit|microsoft|guidepoint|manhas|yellow|profectus|abcom|debtor|advance|reliant/i.test(`${row.accountName} ${row.accountGroup}`)).map((row) => ["", `  - ${row.accountName}`, "", "", amount(row), row.accountGroup, row.raw["Source Row"] ?? ""]),
    ["6", "Profit & Loss A/c - current period PBT", "", "", pbt, "", "Derived from P&L rows"],
  ];
  const assets = [
    ["1", "Fixed Assets", "", "", rowsMatching(rows, "fixed-asset").reduce((sum, row) => sum + amount(row), 0), "", ""],
    ...balanceLineRows(rows, "fixed-asset"),
    ["2", "Sundry Debtors / Receivables", "", "", rowsMatching(rows, "current-asset", /sundry debtors|debtor|receivable|alnylam|jntl|kenvue|replimune|vifor|xoma/i).reduce((sum, row) => sum + amount(row), 0), "", ""],
    ...balanceLineRows(rows, "current-asset", /sundry debtors|debtor|receivable|alnylam|jntl|kenvue|replimune|vifor|xoma/i),
    ["3", "Cash, Bank and Fixed Deposits", "", "", cashAndBankTotal(rows), "", ""],
    ...balanceLineRows(rows, "current-asset", /bank|cash|fd/i),
    ["4", "Advances, ITC, Prepaids and Deposits", "", "", rowsMatching(rows, "current-asset", /advance|input tax|itc|prepaid|deposit|suspense|preliminary|reimbursement/i).reduce((sum, row) => sum + amount(row), 0), "", ""],
    ...balanceLineRows(rows, "current-asset", /advance|input tax|itc|prepaid|deposit|suspense|preliminary|reimbursement/i),
    ["5", "Other Current Assets", "", "", rowsMatching(rows, "current-asset").filter((row) => !/sundry debtors|debtor|receivable|alnylam|jntl|kenvue|replimune|vifor|xoma|bank|cash|fd|advance|input tax|itc|prepaid|deposit|suspense|preliminary|reimbursement/i.test(`${row.accountName} ${row.accountGroup}`)).reduce((sum, row) => sum + amount(row), 0), "", ""],
    ...rowsMatching(rows, "current-asset").filter((row) => !/sundry debtors|debtor|receivable|alnylam|jntl|kenvue|replimune|vifor|xoma|bank|cash|fd|advance|input tax|itc|prepaid|deposit|suspense|preliminary|reimbursement/i.test(`${row.accountName} ${row.accountGroup}`)).map((row) => ["", `  - ${row.accountName}`, "", "", amount(row), row.accountGroup, row.raw["Source Row"] ?? ""]),
  ];
  const liabilityTotal = liabilities.reduce((sum, row) => sum + (String(row[0]) ? Number(row[4]) || 0 : 0), 0);
  const assetTotal = assets.reduce((sum, row) => sum + (String(row[0]) ? Number(row[4]) || 0 : 0), 0);
  const sheet = makeSheet([
    [`Balance Sheet - ${profile.businessName || profile.legalEntity || "Untitled MIS report"}`],
    [`As on ${profile.periodEnd || dateRange(profile)}`],
    [],
    ["S. No.", "Particulars", "", "", "Amount", "Group / trace", "Source row"],
    ["", "LIABILITIES & EQUITY", "", "", "", "", ""],
    ...liabilities,
    ["", "Total Liabilities & Equity", "", "", liabilityTotal, "", ""],
    [],
    ["", "ASSETS", "", "", "", "", ""],
    ...assets,
    ["", "Total Assets", "", "", assetTotal, "", ""],
    ["", "Balance Check", "", "", assetTotal - liabilityTotal, "", "Should be near zero after P&L and balance sheet mapping."],
  ]);
  sheet["!cols"] = [{ wch: 10 }, { wch: 46 }, { wch: 4 }, { wch: 4 }, { wch: 18 }, { wch: 30 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Balance Sheet");
}

function appendTaxStatutorySchedule(workbook: XLSX.WorkBook, rows: TrialBalanceRow[]) {
  const taxRows = rows.filter((row) => /gst|tds|pf|esi|esic|professional tax|advance tax|input tax|itc|tax/i.test(`${row.accountName} ${row.accountGroup}`));
  const sheet = makeSheet([
    ["Tax & Statutory Schedule"],
    [],
    ["Account", "Category", "Debit", "Credit", "Signed balance", "Treatment", "Client-ready action"],
    ...taxRows.map((row) => [
      row.accountName,
      categoryLabels[row.category],
      row.debit,
      row.credit,
      categorySignedAmount(row),
      row.category === "current-liability" ? "Payable / statutory liability" : row.category === "current-asset" ? "Recoverable / advance tax / ITC" : "P&L tax or interest cost",
      /advance tax/i.test(row.accountName) ? "Ask whether to compute current tax provision." : /itc|input tax/i.test(row.accountName) ? "Reconcile against GST payable and returns." : "Tie to challans/returns before client circulation.",
    ]),
    [],
    ["Tax provision prompt", "", "", "", "", "Advance tax exists but income-tax provision may not be booked.", "Confirm whether MIS should show PBT only or compute current tax provision."],
  ]);
  sheet["!cols"] = [{ wch: 42 }, { wch: 20 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 34 }, { wch: 72 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Tax & Statutory");
}

function appendFundFlowStatement(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], bankTransactions: BankTransaction[]) {
  const model = buildFundFlowModel(profile, rows, bankTransactions);
  const receiptTotals = model.months.map((_, index) => model.receiptRows.reduce((sum, row) => sum + row[1][index], 0));
  const paymentTotals = model.months.map((_, index) => model.paymentRows.reduce((sum, row) => sum + row[1][index], 0));
  const netMovement = model.months.map((_, index) => receiptTotals[index] - paymentTotals[index]);
  const hasBank = bankTransactions.length > 0;
  const sheet = makeSheet([
    ["COMPANY NAME", profile.businessName || profile.legalEntity || "Untitled MIS report"],
    ["REPORT TYPE", "FUND FLOW STATEMENTS"],
    ["PERIOD", dateRange(profile)],
    ["BASIS", hasBank ? (bankTransactions.some((txn) => txn.sourceType === "bank-statement") ? "Bank statement actuals; bank ledger supports classification/reconciliation" : "Bank ledger provisional actuals") : profile.fundFlowBasis === "manual" ? "Manual schedule pending" : "Trial balance proxy only"],
    ["PARTICULARS / MONTHS", ...model.months, "Total / Closing"],
    ["BEGINNING CASH / BANK / FD", ...model.openingBalances, model.openingBalances[0] || 0],
    ["CASH / BANK / FD PER TB", ...model.months.map((_, index) => (index === model.months.length - 1 ? model.closingCash : 0)), model.closingCash],
    [],
    ["( + ) CASH / BANK RECEIPTS", ...model.months.map(() => ""), ""],
    ...model.receiptRows.map(([label, values, note]) => [label, ...values, sumValues(values), note]),
    ["TOTAL RECEIPTS", ...receiptTotals, sumValues(receiptTotals)],
    [],
    ["( - ) TAX PAYMENTS / STATUTORY", ...model.months.map(() => ""), ""],
    ...model.paymentRows.slice(0, 1).map(([label, values, note]) => [label, ...values, sumValues(values), note]),
    ["Advance tax / ITC per TB", ...model.months.map(() => 0), rowsMatching(rows, "current-asset", /advance tax|itc|input tax/i).reduce((sum, row) => sum + amount(row), 0), "Balance-sheet tax assets; cash timing comes from bank source."],
    [],
    ["( - ) OPERATING PAYMENTS", ...model.months.map(() => ""), ""],
    ...model.paymentRows.slice(1).map(([label, values, note]) => [label, ...values, sumValues(values), note]),
    ["TOTAL PAYMENTS", ...paymentTotals, sumValues(paymentTotals)],
    ["NET MOVEMENT", ...netMovement, sumValues(netMovement)],
    ["CLOSING CASH / BANK / FD", ...model.closingBalances, model.closingBalances.at(-1) || model.closingCash],
    ["Internal transfers excluded", ...model.internalTransferValues, sumValues(model.internalTransferValues), "Shown for audit only; excluded from receipts/payments."],
    [],
    ["Caveat", hasBank ? "Derived from uploaded bank sources. Review unclassified bank rows and reconciliation before client circulation." : "Bank statement/ledger not uploaded. Upload bank sources before treating this as a true fund flow.", ...model.months.map(() => ""), ""],
  ]);
  sheet["!cols"] = [{ wch: 36 }, ...model.months.map(() => ({ wch: 14 })), { wch: 18 }, { wch: 70 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Fund Flow");
}

function appendFundFlowProjection(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], bankTransactions: BankTransaction[]) {
  const model = buildFundFlowModel(profile, rows, bankTransactions);
  const actualReceiptTotals = model.months.map((_, index) => model.receiptRows.reduce((sum, row) => sum + row[1][index], 0));
  const actualSalary = model.paymentRows.find(([label]) => /salary|bonus/i.test(label))?.[1] || [];
  const actualOpex = model.paymentRows.find(([label]) => /vendor|operating/i.test(label))?.[1] || [];
  const actualTax = model.paymentRows.find(([label]) => /gst|tds|pf|esic/i.test(label))?.[1] || [];
  const actualCapex = model.paymentRows.find(([label]) => /capex/i.test(label))?.[1] || [];
  const monthsInPeriod = periodMonths(profile);
  const revenueBase = averagePositive(actualReceiptTotals) || rowsMatching(rows, "revenue").reduce((sum, row) => sum + amount(row), 0) / monthsInPeriod;
  const salaryBase = averagePositive(actualSalary) || rowsMatching(rows, "people-cost").reduce((sum, row) => sum + amount(row), 0) / monthsInPeriod;
  const opexBase =
    averagePositive(actualOpex) ||
    rows.filter((row) => ["direct-cost", "operating-expense", "finance-cost"].includes(row.category)).reduce((sum, row) => sum + amount(row), 0) / monthsInPeriod;
  const taxBase = averagePositive(actualTax);
  const capexBase = averagePositive(actualCapex);
  const projectedReceipts = model.projectionMonths.map((_, index) => revenueBase * 1.1 * (1 + Math.floor(index / 12) * 0.05));
  const projectedSalary = model.projectionMonths.map(() => salaryBase * 1.1);
  const projectedOpex = model.projectionMonths.map(() => opexBase * 1.08);
  const projectedTax = model.projectionMonths.map(() => taxBase);
  const projectedCapex = model.projectionMonths.map(() => capexBase);
  const projectedPayments = model.projectionMonths.map((_, index) => projectedSalary[index] + projectedOpex[index] + projectedTax[index] + projectedCapex[index]);
  const projectedNet = model.projectionMonths.map((_, index) => projectedReceipts[index] - projectedPayments[index]);
  const openingBalances: number[] = [];
  const closingBalances: number[] = [];
  let running = model.closingBalances.at(-1) || model.closingCash;

  model.projectionMonths.forEach((_, index) => {
    openingBalances.push(running);
    running += projectedNet[index];
    closingBalances.push(running);
  });

  const sheet = makeSheet([
    ["FUND FLOW PROJECTION"],
    ["Company", profile.businessName || profile.legalEntity || "Untitled MIS report"],
    ["Projection basis", "Default next 12 months based on historical monthly bank movement where available; otherwise annual TB run-rate."],
    [],
    ["Particulars / Months", ...model.projectionMonths, "Total"],
    ["Opening cash / bank / FD", ...openingBalances, openingBalances[0] || 0],
    ["Projected client receipts", ...projectedReceipts, sumValues(projectedReceipts)],
    ["Projected salary / bonus / payroll", ...projectedSalary.map((value) => -value), -sumValues(projectedSalary)],
    ["Projected vendor / operating payments", ...projectedOpex.map((value) => -value), -sumValues(projectedOpex)],
    ["Projected tax / statutory payments", ...projectedTax.map((value) => -value), -sumValues(projectedTax)],
    ["Projected capex / FD / treasury outflow", ...projectedCapex.map((value) => -value), -sumValues(projectedCapex)],
    ["Net projected movement", ...projectedNet, sumValues(projectedNet)],
    ["Projected closing cash / bank / FD", ...closingBalances, closingBalances.at(-1) || 0],
    [],
    ["Default assumptions", "Revenue +10%; payroll +10%; operating cost +8%; tax/capex use historical average where visible. Override through the question register before client circulation."],
  ]);
  sheet["!cols"] = [{ wch: 38 }, ...model.projectionMonths.map(() => ({ wch: 14 })), { wch: 18 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Fund Flow Projection");
}

function appendBankSourceAudit(workbook: XLSX.WorkBook, profile: BusinessProfile, bankTransactions: BankTransaction[], sourceType: "bank-statement" | "bank-ledger", sheetName: string) {
  const sourceRows = bankTransactions.filter((txn) => txn.sourceType === sourceType);
  const sheet = makeSheet([
    [sourceType === "bank-statement" ? "Bank Statement Audit" : "Bank Ledger Audit"],
    [],
    ["Fund flow basis selected", profile.fundFlowBasis],
    ["Transactions loaded", sourceRows.length],
    ["Status", sourceRows.length ? "Bank data loaded; review categories before client circulation." : `Missing ${sourceType === "bank-statement" ? "bank statement" : "bank ledger"}. Fund flow remains provisional.`],
    [],
    ["Date", "Source file", "Bank / account", "Narration", "Payment", "Receipt", "Net amount", "Balance", "Category", "Fund-flow group", "Internal transfer", "Source sheet"],
    ...sourceRows.map((txn) => [
      txn.date,
      txn.sourceFileName,
      txn.accountName || txn.bankName,
      txn.narration,
      txn.debit,
      txn.credit,
      txn.amount,
      txn.balance,
      txn.category,
      txn.fundFlowGroup,
      txn.isInterAccountTransfer ? "Yes" : "No",
      txn.sourceSheet,
    ]),
    ...(!sourceRows.length
      ? [["", "", "", `Upload ${sourceType === "bank-statement" ? "bank statements" : "bank ledgers"} for every bank/cash/FD account used in the period.`, 0, 0, 0, 0, "Missing source", "", "", ""]]
      : []),
  ]);
  sheet["!cols"] = [{ wch: 14 }, { wch: 28 }, { wch: 24 }, { wch: 64 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 24 }, { wch: 30 }, { wch: 16 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
}

function appendBankReconciliation(workbook: XLSX.WorkBook, rows: TrialBalanceRow[], bankTransactions: BankTransaction[]) {
  const statementRows = bankTransactions.filter((txn) => txn.sourceType === "bank-statement");
  const ledgerRows = bankTransactions.filter((txn) => txn.sourceType === "bank-ledger");
  const statementClosing = latestBankBalance(statementRows);
  const ledgerClosing = latestBankBalance(ledgerRows);
  const tbCash = cashAndBankTotal(rows);
  const statementReceipts = statementRows.reduce((sum, txn) => sum + (txn.isInterAccountTransfer ? 0 : txn.credit), 0);
  const statementPayments = statementRows.reduce((sum, txn) => sum + (txn.isInterAccountTransfer ? 0 : txn.debit), 0);
  const ledgerReceipts = ledgerRows.reduce((sum, txn) => sum + (txn.isInterAccountTransfer ? 0 : txn.credit), 0);
  const ledgerPayments = ledgerRows.reduce((sum, txn) => sum + (txn.isInterAccountTransfer ? 0 : txn.debit), 0);
  const sheet = makeSheet([
    ["Bank Reconciliation"],
    [],
    ["Check", "Bank statement", "Bank ledger", "Trial balance", "Status / action"],
    ["Rows loaded", statementRows.length, ledgerRows.length, rows.length, bankTransactions.length ? "Loaded" : "Upload bank statement and ledger."],
    ["Receipts", statementReceipts, ledgerReceipts, "", "Statement is used for actual cash receipts when present."],
    ["Payments", statementPayments, ledgerPayments, "", "Statement is used for actual cash payments when present."],
    ["Net movement", statementReceipts - statementPayments, ledgerReceipts - ledgerPayments, "", "Differences indicate timing, missing account, or unreconciled entries."],
    ["Closing balance from source", statementClosing, ledgerClosing, tbCash, "Tie statement/ledger closing to TB bank + FD."],
    ["Statement vs TB difference", statementClosing ? statementClosing - tbCash : "", "", "", statementClosing ? "Review if not near zero." : "Statement balance unavailable."],
    ["Ledger vs TB difference", "", ledgerClosing ? ledgerClosing - tbCash : "", "", ledgerClosing ? "Review if not near zero." : "Ledger balance unavailable."],
    [],
    ["TB bank / FD accounts", "Category", "Debit", "Credit", "Amount"],
    ...rows.filter((row) => row.category === "current-asset" && /bank|cash|fd|fixed deposit/i.test(`${row.accountName} ${row.accountGroup}`)).map((row) => [row.accountName, row.misGroup, row.debit, row.credit, amount(row)]),
  ]);
  sheet["!cols"] = [{ wch: 36 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 68 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Bank Reconciliation");
}

function appendFundFlowAssumptions(workbook: XLSX.WorkBook, profile: BusinessProfile, bankTransactions: BankTransaction[]) {
  const hasStatement = bankTransactions.some((txn) => txn.sourceType === "bank-statement");
  const hasLedger = bankTransactions.some((txn) => txn.sourceType === "bank-ledger");
  const unclassified = bankTransactions.filter((txn) => /unclassified/i.test(`${txn.category} ${txn.fundFlowGroup}`)).length;
  const sheet = makeSheet([
    ["Fund Flow Assumptions"],
    [],
    ["Area", "Default used by generator", "When user input is still needed"],
    ["Actual period", profile.periodStart && profile.periodEnd ? dateRange(profile) : "Missing from source", "Only ask if the TB does not contain a readable period."],
    ["Cash basis", hasStatement ? "Bank statement actuals" : hasLedger ? "Bank ledger provisional actuals" : "TB proxy only", "Upload statement and ledger for client-ready fund flow."],
    ["Bank ledger use", hasLedger ? "Used for classification/reconciliation" : "Missing", "Needed where statement narration is weak."],
    ["Projection horizon", "Next 12 months after report period", "Override only if management wants a different horizon."],
    ["Projection revenue", "Historical bank receipt average or annual TB revenue / months, plus 10%", "Override with management budget/growth rate."],
    ["Projection costs", "Historical bank payment average or TB cost run-rate; payroll +10%, opex +8%", "Override for hiring, bonus, vendor changes, capex, EMI, tax calendar."],
    ["Internal transfers", "Excluded from receipts/payments, shown as audit row", "Confirm only if bank tags look wrong."],
    ["Unclassified bank movements", unclassified, unclassified ? "Review unclassified rows before client circulation." : "No action."],
  ]);
  sheet["!cols"] = [{ wch: 28 }, { wch: 56 }, { wch: 72 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Fund Flow Assumptions");
}

function appendChartDataProduction(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], financials: UnitFinancials[], bankTransactions: BankTransaction[]) {
  const revenue = rowsMatching(rows, "revenue").reduce((sum, row) => sum + amount(row), 0);
  const otherIncome = rowsMatching(rows, "other-income").reduce((sum, row) => sum + amount(row), 0);
  const people = rowsMatching(rows, "people-cost").reduce((sum, row) => sum + amount(row), 0);
  const direct = rowsMatching(rows, "direct-cost").reduce((sum, row) => sum + amount(row), 0);
  const opex = rowsMatching(rows, "operating-expense").reduce((sum, row) => sum + amount(row), 0);
  const financeTax = rows.filter((row) => ["finance-cost", "tax"].includes(row.category)).reduce((sum, row) => sum + amount(row), 0);
  const actualBankRows = actualFundFlowTransactions(bankTransactions).filter((txn) => !txn.isInterAccountTransfer);
  const receipts = actualBankRows.reduce((sum, txn) => sum + txn.credit, 0);
  const payments = actualBankRows.reduce((sum, txn) => sum + txn.debit, 0);
  const model = buildFundFlowModel(profile, rows, bankTransactions);
  const netActual = model.months.map((_, index) => {
    const monthReceipts = model.receiptRows.reduce((sum, row) => sum + row[1][index], 0);
    const monthPayments = model.paymentRows.reduce((sum, row) => sum + row[1][index], 0);
    return monthReceipts - monthPayments;
  });
  const averageNet = netActual.length ? sumValues(netActual) / netActual.length : 0;
  let projectedClosing = model.closingBalances.at(-1) || model.closingCash;
  const projectedCashTrend = model.projectionMonths.map((month) => {
    projectedClosing += averageNet;
    return [month, projectedClosing];
  });
  const sheet = makeSheet([
    ["Revenue Mix", "Amount", "", "Expense Mix", "Amount", "", "Client / Unit", "Revenue", "PBT / Surplus", "", "Asset Composition", "Amount", "", "Liability Composition", "Amount"],
    ["Operating Revenue", revenue, "", "People Costs", people, "", ...((financials[0] && [financials[0].center.name, totalRevenue(financials[0]), ebitda(financials[0])]) || ["", 0, 0]), "", "Fixed Assets", rowsMatching(rows, "fixed-asset").reduce((sum, row) => sum + amount(row), 0), "", "Capital", rowsMatching(rows, "equity").reduce((sum, row) => sum + amount(row), 0)],
    ["Other Income", otherIncome, "", "Direct Costs", direct, "", ...((financials[1] && [financials[1].center.name, totalRevenue(financials[1]), ebitda(financials[1])]) || ["", 0, 0]), "", "Debtors / Receivables", rowsMatching(rows, "current-asset", /debtor|receivable|alnylam|jntl|kenvue|replimune|vifor|xoma/i).reduce((sum, row) => sum + amount(row), 0), "", "Current Liabilities", rowsMatching(rows, "current-liability").reduce((sum, row) => sum + amount(row), 0)],
    ["", "", "", "Operating Expenses", opex, "", ...((financials[2] && [financials[2].center.name, totalRevenue(financials[2]), ebitda(financials[2])]) || ["", 0, 0]), "", "Bank + FD", cashAndBankTotal(rows), "", "Current Period PBT", pbtFromTrialBalance(rows)],
    ["", "", "", "Finance / Tax / Other", financeTax, "", ...((financials[3] && [financials[3].center.name, totalRevenue(financials[3]), ebitda(financials[3])]) || ["", 0, 0]), "", "Advances / ITC / Prepaids", rowsMatching(rows, "current-asset", /advance|itc|input tax|prepaid|deposit|suspense|preliminary/i).reduce((sum, row) => sum + amount(row), 0), "", "", ""],
    [],
    ["Fund Flow", "Amount", "", "Receipts vs Payments", "Amount"],
    ["Bank receipts", receipts, "", "Receipts", receipts],
    ["Bank payments", payments, "", "Payments", payments],
    ["Net bank movement", receipts - payments, "", "TB cash/bank/FD", cashAndBankTotal(rows)],
    [],
    ["Actual Cash Trend", "Closing cash / bank / FD", "", "Projected Cash Trend", "Projected closing cash / bank / FD"],
    ...model.months.map((month, index) => [month, model.closingBalances[index] || 0, "", projectedCashTrend[index]?.[0] || "", projectedCashTrend[index]?.[1] || 0]),
  ]);
  sheet["!cols"] = Array.from({ length: 15 }, () => ({ wch: 20 }));
  XLSX.utils.book_append_sheet(workbook, sheet, "ChartData");
}

function appendQcTieOuts(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], financials: UnitFinancials[]) {
  const debit = rows.reduce((sum, row) => sum + row.debit, 0);
  const credit = rows.reduce((sum, row) => sum + row.credit, 0);
  const revenue = rowsMatching(rows, "revenue").reduce((sum, row) => sum + amount(row), 0);
  const pbt = pbtFromTrialBalance(rows);
  const clientMargins = financials.map((unit) => Math.round(margin(unit) * 1000000) / 1000000);
  const repeatedMargins = clientMargins.length > 1 && new Set(clientMargins).size === 1;
  const suspenseRows = rows.filter((row) => /suspense/i.test(row.accountName));
  const sheet = makeSheet([
    ["QC Tie-Outs"],
    [],
    ["Check", "Value", "Status", "Action"],
    ["Business name inferred", profile.businessName, profile.businessName && !/sample business/i.test(profile.businessName) ? "OK" : "Needs review", "Use source company name before client circulation."],
    ["Reporting period inferred", dateRange(profile), profile.periodStart && profile.periodEnd ? "OK" : "Needs review", "Use source TB period where available."],
    ["TB debit", debit, Math.abs(debit - credit) <= 1 ? "OK" : "Mismatch", ""],
    ["TB credit", credit, Math.abs(debit - credit) <= 1 ? "OK" : "Mismatch", ""],
    ["TB imbalance", Math.abs(debit - credit), Math.abs(debit - credit) <= 1 ? "OK" : "Mismatch", "Review parser/group heading exclusion."],
    ["Operating revenue", revenue, revenue ? "OK" : "Missing", "For Simple fixture this should be 79,502,679.52."],
    ["Profit before tax margin", revenue ? pbt / revenue : 0, revenue ? "OK" : "Missing", "Formula is PBT / Revenue."],
    ["Cash + bank + FD", cashAndBankTotal(rows), cashAndBankTotal(rows) ? "OK" : "Missing", "For Simple fixture this should be 15,360,815.97."],
    ["Credit debtor / customer advance", rowsMatching(rows, "current-liability", /reliant|debtor/i).reduce((sum, row) => sum + amount(row), 0), "Review", "Credit-balance debtors must not be ordinary receivables or revenue allocation base."],
    ["Repeated client margins", repeatedMargins ? "Yes" : "No", repeatedMargins ? "Needs review" : "OK", "Avoid mechanically identical margins unless allocation is intentionally revenue-share only."],
    ["Suspense rows", suspenseRows.length, suspenseRows.length ? "Needs review" : "OK", "Confirm duplicate Suspense A/c balances."],
  ]);
  sheet["B10"] = { ...(sheet["B10"] || {}), z: percentFormat };
  sheet["!cols"] = [{ wch: 34 }, { wch: 24 }, { wch: 16 }, { wch: 78 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "QC Tie-Outs");
}

function appendLedgerMappingAuditProduction(workbook: XLSX.WorkBook, rows: TrialBalanceRow[], centers: ProfitCenter[]) {
  const sheet = makeSheet([
    ["Source Sheet", "Source Row", "Account", "Group", "Debit", "Credit", "Signed balance", "MIS Category", "MIS Group", "Profit center", "Confidence"],
    ...rows.map((row) => [
      row.sourceSheet,
      row.raw["Source Row"] ?? "",
      row.accountName,
      row.accountGroup,
      row.debit,
      row.credit,
      categorySignedAmount(row),
      categoryLabels[row.category],
      row.misGroup,
      centers.find((center) => center.id === row.profitCenterId)?.name || "Shared / unassigned",
      row.confidence,
    ]),
    ...(!rows.length ? [["Trial balance missing", "", "", "", 0, 0, 0, "", "", "", ""]] : []),
  ]);
  applyTableFormatting(sheet, rows.length + 1, 11);
  sheet["!cols"] = [{ wch: 16 }, { wch: 12 }, { wch: 42 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 20 }, { wch: 24 }, { wch: 28 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "09 Ledger Mapping Audit");
}

function appendAccountAllocationMatrix(workbook: XLSX.WorkBook, rows: TrialBalanceRow[]) {
  const summary = Array.from(new Set(rows.map((row) => row.misGroup))).map((group) => {
    const groupRows = rows.filter((row) => row.misGroup === group);
    return [
      group,
      categoryLabels[groupRows[0]?.category || "unknown"],
      groupRows.length,
      groupRows.reduce((sum, row) => sum + amount(row), 0),
    ];
  });

  const sheet = makeSheet([
    ["Account Allocation"],
    ["Auto-mapped from trial balance ledger names, account groups, debit/credit orientation, and Simple-specific ledger vocabulary. No client input required for this base allocation."],
    [],
    ["MIS Group", "Broad category", "Ledger count", "Mapped amount"],
    ...summary,
    [],
    ["Source row", "Account path", "Account", "Account group", "Debit", "Credit", "Broad category", "Detailed MIS group", "Balance treatment", "Mapping basis"],
    ...rows.map((row) => [
      row.raw["Source Row"] ?? "",
      row.accountPath?.join(" > ") || [row.accountGroup, row.accountName].filter(Boolean).join(" > "),
      row.accountName,
      row.accountGroup,
      row.debit,
      row.credit,
      categoryLabels[row.category],
      row.misGroup,
      row.category === "current-liability" || row.category === "equity" || row.category === "revenue" || row.category === "other-income" ? "Credit-normal" : "Debit-normal",
      row.confidence >= 0.94 ? "Specific ledger rule" : row.accountGroup ? "TB group + ledger keywords" : "Ledger keywords",
    ]),
  ]);
  sheet["!cols"] = [{ wch: 12 }, { wch: 58 }, { wch: 42 }, { wch: 28 }, { wch: 14 }, { wch: 14 }, { wch: 20 }, { wch: 38 }, { wch: 18 }, { wch: 28 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Account Allocation");
}

function appendAssumptionsQualityProduction(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], staff: StaffMember[], issues: WorkbookIssue[], questions: GeneratedQuestion[], answers: QuestionAnswer[]) {
  const answerById = new Map(answers.map((answer) => [answer.id, answer]));
  const assumptionRows = [
    ["Accounting source", rows.length ? "Actual loaded" : "Missing", `${rows.length} ledger rows imported.`, rows.length ? "Use final TB after books close." : "Upload TB before client circulation."],
    ["Revenue split", "Proxy where invoice register is missing", "Manual engagement revenue may be based on debtor balance proxy.", "Replace with invoice-wise revenue by engagement."],
    ["People cost", staff.length ? "Staffing matrix loaded" : "Missing", `${staff.length} staff/pod rows. Client columns are FTE weights.`, "Replace pods with named employees, CTC, and timesheets."],
    ["Direct costs", "Partial", "Vendor/project tagging is inferred/manual unless source register is uploaded.", "Add vendor ledger with project tags."],
    ["Working capital", rows.length ? "TB based" : "Missing", "Debtors/payables/cash are from TB; ageing is not available.", "Add AR/AP ageing and bank statements."],
    ["Third-party AI", "Requires approval", "Claude/Shortcut may transmit workbook data to a third party.", "Use only with appropriate client/firm data consent."],
  ];
  const sheet = makeSheet([
    ["10 Assumptions & Data Quality"],
    [],
    ["Area", "Status", "Detail", "Client-ready action"],
    ...assumptionRows,
    [],
    ["Generated issue", "Severity", "Detail", "Resolution"],
    ...issues.map((issue) => [issue.label, issue.severity, issue.detail, "Resolve before client circulation."]),
    [],
    ["Open question", "Priority", "Status", "Answer"],
    ...questions.map((question) => [question.prompt, question.priority, answerById.get(question.id)?.status || "open", answerById.get(question.id)?.answer || ""]),
  ]);
  sheet["!cols"] = [{ wch: 34 }, { wch: 20 }, { wch: 88 }, { wch: 72 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "10 Assumptions & Data Quality");
}

function appendAiBuildInstructions(workbook: XLSX.WorkBook) {
  const sheet = makeSheet([
    ["AI_BUILD_INSTRUCTIONS"],
    [],
    ["Warning", "This workbook may contain client financials, salaries/pod costs, ledgers, receivables, and assumptions. Using Claude for Excel or Shortcut may transmit that data to the provider. Confirm permission and data protection settings before use."],
    [],
    ["Tool", "Prompt"],
    [
      "Claude for Excel / Shortcut",
      "Review this workbook as a client-ready MIS pack. Do not invent accounting data. Preserve all source schedules. Apply professional Shoolini-style formatting, freeze panes, filters, consistent number formats, and clear section headers across all sheets.",
    ],
    [
      "Claude for Excel / Shortcut",
      "Verify formulas and tie-outs across 01 Management Summary, 02 Engagement P&L, 03 Revenue Register, 04 People Cost Register, 05 Vendor Direct Cost Register, 06 Shared Cost Pools, 07 Working Capital, and 08 Cash Flow. Flag any mismatch in a new QA section; do not overwrite source values.",
    ],
    [
      "Claude for Excel / Shortcut",
      "Create margin heatmaps on 02 Engagement P&L and Client Scorecard views. Highlight operating margin below 25%, receivables above 25% of revenue, negative surplus, and zero/blank source schedules.",
    ],
    [
      "Claude for Excel / Shortcut",
      "Add a navigation strip or hyperlinks from 00 Control Panel to every major sheet and back. Keep source schedules intact and visible.",
    ],
    [
      "Claude for Excel / Shortcut",
      "Create board-style charts from 01 Management Summary and 02 Engagement P&L: revenue by engagement, operating surplus by engagement, margin by engagement, and cost mix waterfall.",
    ],
    [
      "Claude for Excel / Shortcut",
      "Convert dense schedules into Excel tables where appropriate and ensure filters are enabled. Do not change ledger classifications unless the user explicitly confirms the accounting treatment.",
    ],
  ]);
  sheet["!cols"] = [{ wch: 28 }, { wch: 130 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "AI_BUILD_INSTRUCTIONS");
}

function appendAiFormulaBuildSpec(workbook: XLSX.WorkBook) {
  const sheet = makeSheet([
    ["AI_FORMULA_BUILD_SPEC"],
    [],
    ["Purpose", "Use this sheet with Claude for Excel or Shortcut. The app exports source schedules and values; the Excel AI should convert the production pack into a formula-driven client workbook."],
    ["Data warning", "This workbook may include client financials, salary/pod costs, ledgers, receivables and assumptions. Use Claude/Shortcut only after client/firm approval."],
    [],
    ["Step", "Instruction for Claude / Shortcut"],
    [
      "1",
      "Treat 03 Revenue Register, 04 People Cost Register, 05 Vendor Direct Cost Register, 06 Shared Cost Pools, 07 Working Capital, and 09 Ledger Mapping Audit as source/data sheets. Do not overwrite raw values in these sheets.",
    ],
    [
      "2",
      "Convert 01 Management Summary, 02 Engagement P&L, 08 Cash Flow, Client Scorecard, Cost Pool Allocation, and Unit Profitability into formula-driven reports that reference the source/data sheets.",
    ],
    [
      "3",
      "For 02 Engagement P&L: Revenue should SUMIFS 03 Revenue Register by engagement. Direct cost should SUMIFS 05 Vendor Direct Cost Register by engagement. People cost should SUMIFS 04 People Cost Register by engagement. Shared overhead and finance/tax should SUMIFS 06 Shared Cost Pools by engagement. Receivables should SUMIFS 07 Working Capital by matched engagement.",
    ],
    [
      "4",
      "For 01 Management Summary: every amount should reference 02 Engagement P&L or 07 Working Capital, not hardcoded values. Add formulas for revenue, direct cost, people cost, shared overhead, finance/tax, total cost, surplus, margin, cash, receivables, current liabilities, and net working capital.",
    ],
    [
      "5",
      "For 08 Cash Flow: create formulas that bridge operating surplus to cash using receivables, deposits/prepaids/advances, current liabilities, fixed assets/capex, and closing cash. Mark opening cash and prior-period movement as assumptions if not available.",
    ],
    [
      "6",
      "Create a QA/tie-out block that checks: TB debit = TB credit, revenue register total = Engagement P&L revenue, people register total = Engagement P&L people cost, shared cost pools total = allocated shared overhead/finance/tax, working capital totals = ledger mapping audit balance sheet categories.",
    ],
    [
      "7",
      "Apply client-ready formatting: freeze panes, filters, number formats, INR formatting, margin heatmaps, section dividers, hidden gridlines if appropriate, and navigation links back to 00 Control Panel.",
    ],
    [
      "8",
      "Create board charts from 01 Management Summary and 02 Engagement P&L: revenue by engagement, operating surplus by engagement, margin by engagement, cost mix waterfall, and receivables by client.",
    ],
    [],
    ["Copy/paste master prompt"],
    [
      "Prompt",
      "You are an expert Excel financial modeller. Turn this MIS production pack into a client-ready, formula-driven workbook. Keep source sheets intact. Build formulas from source schedules into summary/reporting sheets. Add tie-out checks, heatmaps, navigation, and board charts. Do not invent missing accounting data; clearly label assumptions and data gaps. Use Shoolini-style MIS depth: detailed schedules, formula-driven totals, traceability, and management-ready formatting.",
    ],
  ]);
  sheet["!cols"] = [{ wch: 26 }, { wch: 140 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "AI_FORMULA_BUILD_SPEC");
}

function appendMethodology(workbook: XLSX.WorkBook, profile: BusinessProfile) {
  const template = getBusinessTemplate(profile.businessType);
  const sheet = makeSheet([
    ["Custom MIS Generator"],
    ["Business", profile.businessName],
    ["Legal entity", profile.legalEntity],
    ["Business type", businessTypeLabels[profile.businessType]],
    ["Reporting period", dateRange(profile)],
    ["Currency", profile.currency],
    ["Default allocation base", allocationLabels[profile.allocationBase]],
    ["Primary operating unit", template.unitPlural],
    ["Next-level analysis", template.subUnitLabel],
    ["Core formula", `${template.metricLabels.primary} x ${template.metricLabels.averageRate}`],
    ["Variable-cost formula", `${template.metricLabels.primary} x ${template.metricLabels.variableCostRate}`],
    ["Deep-dive sheet", template.deepDiveSheet],
    [],
    ["Workbook logic"],
    ["1", "Raw accounting data is preserved and mapped into MIS categories."],
    ["2", "The business type chooses the operating unit and the driver fields."],
    ["3", "Revenue is built from ledger/manual revenue plus driver-based revenue."],
    ["4", "Direct costs are traced first; shared costs are allocated only after direct attribution."],
    ["5", "People cost follows named staff assignment equivalent / FTE."],
    ["6", "Question Register and Deep Dive Requirements capture open assumptions and data gaps."],
  ]);
  sheet["!cols"] = [{ wch: 28 }, { wch: 100 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Methodology");
}

function appendExecutiveDashboard(workbook: XLSX.WorkBook, profile: BusinessProfile, financials: UnitFinancials[], issues: WorkbookIssue[]) {
  const template = getBusinessTemplate(profile.businessType);
  const revenue = financials.reduce((sum, unit) => sum + totalRevenue(unit), 0);
  const cost = financials.reduce((sum, unit) => sum + totalCost(unit), 0);
  const surplus = revenue - cost;
  const topRevenue = [...financials].sort((a, b) => totalRevenue(b) - totalRevenue(a))[0];
  const weakestMargin = [...financials].sort((a, b) => margin(a) - margin(b))[0];
  const highIssues = issues.filter((issue) => issue.severity === "high").length;

  const sheet = makeSheet([
    [`Executive Dashboard (${dateRange(profile)})`],
    [],
    ["Metric", "Value", "Interpretation"],
    ["Business", profile.businessName || profile.legalEntity, businessTypeLabels[profile.businessType]],
    ["Revenue", revenue, "Manual/ledger revenue plus driver revenue where entered."],
    ["Total cost", cost, "Direct cost, people cost, shared overhead, finance and tax."],
    ["Operating surplus", surplus, "Revenue less all allocated costs."],
    ["Operating margin", revenue ? surplus / revenue : 0, "Surplus / revenue."],
    [`Top ${template.unitLabel.toLowerCase()} by revenue`, topRevenue?.center.name || "", topRevenue ? totalRevenue(topRevenue) : 0],
    ["Weakest margin unit", weakestMargin?.center.name || "", weakestMargin ? margin(weakestMargin) : 0],
    ["Open high-risk gaps", highIssues, "Review Question Register and Data Gaps before relying on final profitability."],
    [],
    ["Management read-out"],
    ["1", "Profitability is directional where revenue/costs are manually allocated from debtor or effort proxies."],
    ["2", "Final MIS should replace demo assumptions with invoice-level revenue, staff timesheets, and vendor/project tagging."],
    ["3", "Working capital is split by client name where the debtor ledger matches the engagement name; unmatched balances remain corporate."],
  ]);
  sheet["B8"] = { ...(sheet["B8"] || {}), z: percentFormat };
  sheet["C10"] = { ...(sheet["C10"] || {}), z: percentFormat };
  sheet["!cols"] = [{ wch: 34 }, { wch: 22 }, { wch: 90 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Executive Dashboard");
}

function appendDriverBuildUp(workbook: XLSX.WorkBook, profile: BusinessProfile, financials: UnitFinancials[]) {
  const template = getBusinessTemplate(profile.businessType);
  const sheet = makeSheet([
    [
      template.unitLabel,
      template.metricLabels.primary,
      template.metricLabels.averageRate,
      template.metricLabels.utilization,
      "Driver revenue",
      template.metricLabels.variableCostRate,
      "Driver direct cost",
      template.metricLabels.secondary,
      template.metricLabels.tertiary,
      template.segmentLabel,
      template.ownerLabel,
      template.revenueDriverLabel,
    ],
    ...financials.map(({ center }) => [
      center.name,
      numeric(center.studentCount),
      numeric(center.averageRevenueRate),
      numeric(center.utilizationPercent),
      null,
      numeric(center.variableCostRate),
      null,
      numeric(center.teachingStaffCount),
      numeric(center.nonTeachingStaffCount),
      center.segment || "",
      center.owner || "",
      center.revenueDriver || "",
    ]),
  ]);

  financials.forEach((unit) => {
    setFormula(sheet, `E${unit.driverRow}`, `B${unit.driverRow}*C${unit.driverRow}*IF(D${unit.driverRow}=0,1,D${unit.driverRow}/100)`);
    setFormula(sheet, `G${unit.driverRow}`, `B${unit.driverRow}*F${unit.driverRow}`);
  });

  sheet["!cols"] = [
    { wch: 28 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 20 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 24 },
    { wch: 22 },
    { wch: 24 },
  ];
  XLSX.utils.book_append_sheet(workbook, sheet, "Driver Build-Up");
}

function appendEngagementSummary(workbook: XLSX.WorkBook, profile: BusinessProfile, financials: UnitFinancials[], staff: StaffMember[]) {
  const template = getBusinessTemplate(profile.businessType);
  const sheet = makeSheet([
    [`${template.unitLabel} MIS Summary (${dateRange(profile)})`],
    [],
    [
      template.unitLabel,
      template.segmentLabel,
      template.ownerLabel,
      "Fee arrangement",
      "Revenue",
      "Direct cost",
      "People cost",
      "Shared overhead",
      "Finance / tax",
      "Total cost",
      "Operating surplus",
      "Margin %",
      template.metricLabels.primary,
      "Assigned FTE",
      "Revenue / FTE",
      "Cost / FTE",
      "Notes",
    ],
    ...financials.map((unit) => {
      const fte = fteForCenter(staff, unit.center.id);
      return [
        unit.center.name,
        unit.center.segment,
        unit.center.owner,
        unit.center.revenueDriver,
        totalRevenue(unit),
        unit.directCost + unit.driverDirectCost,
        unit.peopleCost,
        unit.sharedOpex,
        unit.financeTax,
        totalCost(unit),
        ebitda(unit),
        margin(unit),
        numeric(unit.center.studentCount),
        fte,
        fte ? totalRevenue(unit) / fte : 0,
        fte ? totalCost(unit) / fte : 0,
        unit.center.notes,
      ];
    }),
  ]);

  sheet["!cols"] = [
    { wch: 28 },
    { wch: 28 },
    { wch: 22 },
    { wch: 28 },
    { wch: 16 },
    { wch: 16 },
    { wch: 16 },
    { wch: 16 },
    { wch: 16 },
    { wch: 16 },
    { wch: 16 },
    { wch: 12 },
    { wch: 18 },
    { wch: 14 },
    { wch: 16 },
    { wch: 16 },
    { wch: 60 },
  ];
  XLSX.utils.book_append_sheet(workbook, sheet, "Engagement Summary");
}

function appendSharedCostAllocation(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], financials: UnitFinancials[]) {
  const totalCol = 3 + financials.length;
  const sheet = makeSheet([
    ["Cost pool", "Pool amount", "Allocation basis", ...financials.map((unit) => unit.center.name), "Total allocated"],
    ["Allocation weight", 1, allocationLabels[profile.allocationBase], ...financials.map((unit) => unit.sharedWeight), null],
    ["Unassigned operating expenses", unassignedAmount(rows, ["operating-expense"]), allocationLabels[profile.allocationBase], ...financials.map(() => null), null],
    ["Unassigned finance and tax", unassignedAmount(rows, ["finance-cost", "tax"]), allocationLabels[profile.allocationBase], ...financials.map(() => null), null],
    ["Unassigned direct cost", unassignedAmount(rows, ["direct-cost"]), allocationLabels[profile.allocationBase], ...financials.map(() => null), null],
    ["Unassigned people cost", unassignedAmount(rows, ["people-cost"]), "Review roster before allocating", ...financials.map(() => null), null],
  ]);

  for (let row = 3; row <= 6; row += 1) {
    financials.forEach((_, index) => {
      const unitCol = col(3 + index);
      setFormula(sheet, `${unitCol}${row}`, `$B${row}*${unitCol}$2`);
    });
    setFormula(sheet, `${col(totalCol)}${row}`, `SUM(${col(3)}${row}:${col(totalCol - 1)}${row})`);
  }
  setFormula(sheet, `${col(totalCol)}2`, `SUM(${col(3)}2:${col(totalCol - 1)}2)`, percentFormat);

  sheet["!cols"] = [{ wch: 30 }, { wch: 16 }, { wch: 26 }, ...financials.map(() => ({ wch: 18 })), { wch: 18 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Shared Cost Allocation");
}

function appendLedgerCategorySummary(workbook: XLSX.WorkBook, rows: TrialBalanceRow[], centers: ProfitCenter[]) {
  const categories = Array.from(new Set(rows.map((row) => row.category)));
  const sheet = makeSheet([
    ["MIS Category", "MIS Group", "Assigned to profit centers", "Shared / unassigned", "Total", "Ledger count", "Review note"],
    ...categories.map((category) => {
      const categoryRows = rows.filter((row) => row.category === category);
      const assigned = categoryRows.filter((row) => row.profitCenterId).reduce((sum, row) => sum + amount(row), 0);
      const shared = categoryRows.filter((row) => !row.profitCenterId).reduce((sum, row) => sum + amount(row), 0);
      return [
        categoryLabels[category],
        categoryRows[0]?.misGroup || "",
        assigned,
        shared,
        assigned + shared,
        categoryRows.length,
        category === "unknown" ? "Needs manual classification" : centers.length && assigned === 0 && ["revenue", "direct-cost", "people-cost"].includes(category) ? "Consider direct center tagging" : "",
      ];
    }),
  ]);
  sheet["!cols"] = [{ wch: 22 }, { wch: 24 }, { wch: 22 }, { wch: 22 }, { wch: 18 }, { wch: 14 }, { wch: 34 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Ledger Category Summary");
}

function appendUnitProfitability(workbook: XLSX.WorkBook, profile: BusinessProfile, financials: UnitFinancials[]) {
  const template = getBusinessTemplate(profile.businessType);
  const months = periodMonths(profile);
  const rows = [
    "Ledger / manual revenue",
    "Driver revenue",
    "Other income",
    "Total revenue",
    "Direct ledger cost",
    "Driver direct cost",
    "People cost",
    "Allocated shared overhead",
    "Finance and tax",
    "Total cost",
    "Gross profit",
    "EBITDA / operating surplus",
    "Profit margin",
    template.metricLabels.primary,
    `Revenue per ${template.metricLabels.primary.toLowerCase()}`,
    `Cost per ${template.metricLabels.primary.toLowerCase()}`,
  ];
  const totalCol = 1 + financials.length * 2;
  const totalPctCol = totalCol + 1;
  const priorCol = totalCol + 2;
  const annualCol = totalCol + 3;
  const sheet = makeSheet([
    [`${template.unitLabel} Profitability (${dateRange(profile)})`],
    ["Line item", ...financials.flatMap((unit) => [unit.center.name, "%"]), "Total", "%", "Prior period", "Annualised figures"],
    ...rows.map((line) => [line, ...financials.flatMap(() => [null, null]), null, null, null, null]),
  ]);

  const row = {
    ledgerRevenue: 3,
    driverRevenue: 4,
    otherIncome: 5,
    totalRevenue: 6,
    directCost: 7,
    driverCost: 8,
    peopleCost: 9,
    shared: 10,
    financeTax: 11,
    totalCost: 12,
    grossProfit: 13,
    ebitda: 14,
    margin: 15,
    primary: 16,
    revenuePerPrimary: 17,
    costPerPrimary: 18,
  };

  financials.forEach((unit, index) => {
    const valueCol = col(1 + index * 2);
    const pctCol = col(2 + index * 2);
    const driverRow = unit.driverRow;

    sheet[`${valueCol}${row.ledgerRevenue}`] = { t: "n", v: unit.ledgerRevenue, z: moneyFormat };
    setFormula(sheet, `${valueCol}${row.driverRevenue}`, sheetCell("Driver Build-Up", `E${driverRow}`));
    sheet[`${valueCol}${row.otherIncome}`] = { t: "n", v: unit.otherIncome, z: moneyFormat };
    setFormula(sheet, `${valueCol}${row.totalRevenue}`, `SUM(${valueCol}${row.ledgerRevenue}:${valueCol}${row.otherIncome})`);
    sheet[`${valueCol}${row.directCost}`] = { t: "n", v: unit.directCost, z: moneyFormat };
    setFormula(sheet, `${valueCol}${row.driverCost}`, sheetCell("Driver Build-Up", `G${driverRow}`));
    sheet[`${valueCol}${row.peopleCost}`] = { t: "n", v: unit.peopleCost, z: moneyFormat };
    sheet[`${valueCol}${row.shared}`] = { t: "n", v: unit.sharedOpex, z: moneyFormat };
    sheet[`${valueCol}${row.financeTax}`] = { t: "n", v: unit.financeTax, z: moneyFormat };
    setFormula(sheet, `${valueCol}${row.totalCost}`, `SUM(${valueCol}${row.directCost}:${valueCol}${row.financeTax})`);
    setFormula(sheet, `${valueCol}${row.grossProfit}`, `${valueCol}${row.totalRevenue}-${valueCol}${row.directCost}-${valueCol}${row.driverCost}`);
    setFormula(sheet, `${valueCol}${row.ebitda}`, `${valueCol}${row.totalRevenue}-${valueCol}${row.totalCost}`);
    setFormula(sheet, `${valueCol}${row.margin}`, `IF(${valueCol}${row.totalRevenue}=0,0,${valueCol}${row.ebitda}/${valueCol}${row.totalRevenue})`, percentFormat);
    setFormula(sheet, `${valueCol}${row.primary}`, sheetCell("Driver Build-Up", `B${driverRow}`));
    setFormula(sheet, `${valueCol}${row.revenuePerPrimary}`, `IF(${valueCol}${row.primary}=0,0,${valueCol}${row.totalRevenue}/${valueCol}${row.primary})`);
    setFormula(sheet, `${valueCol}${row.costPerPrimary}`, `IF(${valueCol}${row.primary}=0,0,${valueCol}${row.totalCost}/${valueCol}${row.primary})`);

    for (let line = row.ledgerRevenue; line <= row.costPerPrimary; line += 1) {
      if (line === row.margin) continue;
      setFormula(sheet, `${pctCol}${line}`, `IF(${valueCol}${row.totalRevenue}=0,0,${valueCol}${line}/${valueCol}${row.totalRevenue})`, percentFormat);
    }
  });

  for (let line = row.ledgerRevenue; line <= row.costPerPrimary; line += 1) {
    const totalLetter = col(totalCol);
    const totalPctLetter = col(totalPctCol);
    const priorLetter = col(priorCol);
    const annualLetter = col(annualCol);
    const parts = financials.map((_, index) => `${col(1 + index * 2)}${line}`).join(",");

    if (line === row.margin) {
      setFormula(sheet, `${totalLetter}${line}`, `IF(${totalLetter}${row.totalRevenue}=0,0,${totalLetter}${row.ebitda}/${totalLetter}${row.totalRevenue})`, percentFormat);
      setFormula(sheet, `${totalPctLetter}${line}`, `${totalLetter}${line}`, percentFormat);
      setFormula(sheet, `${annualLetter}${line}`, `${totalLetter}${line}`, percentFormat);
    } else if (line === row.revenuePerPrimary) {
      setFormula(sheet, `${totalLetter}${line}`, `IF(${totalLetter}${row.primary}=0,0,${totalLetter}${row.totalRevenue}/${totalLetter}${row.primary})`);
      setFormula(sheet, `${annualLetter}${line}`, `${totalLetter}${line}/${months}*12`);
    } else if (line === row.costPerPrimary) {
      setFormula(sheet, `${totalLetter}${line}`, `IF(${totalLetter}${row.primary}=0,0,${totalLetter}${row.totalCost}/${totalLetter}${row.primary})`);
      setFormula(sheet, `${annualLetter}${line}`, `${totalLetter}${line}/${months}*12`);
    } else {
      setFormula(sheet, `${totalLetter}${line}`, `SUM(${parts})`);
      setFormula(sheet, `${totalPctLetter}${line}`, `IF(${totalLetter}${row.totalRevenue}=0,0,${totalLetter}${line}/${totalLetter}${row.totalRevenue})`, percentFormat);
      if (line < row.primary) setFormula(sheet, `${annualLetter}${line}`, `${totalLetter}${line}/${months}*12`);
    }

    if (line === row.ledgerRevenue) sheet[`${priorLetter}${line}`] = { t: "n", v: financials.reduce((sum, unit) => sum + unit.priorRevenue, 0), z: moneyFormat };
    if (line === row.totalCost) sheet[`${priorLetter}${line}`] = { t: "n", v: financials.reduce((sum, unit) => sum + unit.priorCost, 0), z: moneyFormat };
    if (line === row.ebitda) setFormula(sheet, `${priorLetter}${line}`, `${priorLetter}${row.ledgerRevenue}-${priorLetter}${row.totalCost}`);
  }

  sheet["!cols"] = [{ wch: 34 }, ...financials.flatMap(() => [{ wch: 18 }, { wch: 12 }]), { wch: 18 }, { wch: 12 }, { wch: 18 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Unit Profitability");
}

function appendExecutiveMis(workbook: XLSX.WorkBook, profile: BusinessProfile, financials: UnitFinancials[]) {
  const template = getBusinessTemplate(profile.businessType);
  const totalCol = col(1 + financials.length * 2);
  const sheet = makeSheet([
    [`Executive MIS (${dateRange(profile)})`],
    [],
    ["Metric", "Formula / source", "Value"],
    ["Business type", businessTypeLabels[profile.businessType], ""],
    ["Primary operating unit", template.unitPlural, ""],
    ["Total revenue", "Unit Profitability total revenue", null],
    ["Total cost", "Unit Profitability total cost", null],
    ["Operating surplus", "Unit Profitability EBITDA", null],
    ["Operating margin", "Surplus / revenue", null],
    [`Total ${template.metricLabels.primary}`, "Driver Build-Up", null],
    [`Revenue per ${template.metricLabels.primary.toLowerCase()}`, "Revenue / primary metric", null],
    [`Cost per ${template.metricLabels.primary.toLowerCase()}`, "Cost / primary metric", null],
    ["Largest unit by revenue", "Review Unit Profitability", financials[0]?.center.name || ""],
    ["Open high-priority questions", "Question Register", ""],
  ]);
  setFormula(sheet, "C6", sheetCell("Unit Profitability", `${totalCol}6`));
  setFormula(sheet, "C7", sheetCell("Unit Profitability", `${totalCol}12`));
  setFormula(sheet, "C8", sheetCell("Unit Profitability", `${totalCol}14`));
  setFormula(sheet, "C9", `IF(C6=0,0,C8/C6)`, percentFormat);
  setFormula(sheet, "C10", sheetCell("Unit Profitability", `${totalCol}16`));
  setFormula(sheet, "C11", `IF(C10=0,0,C6/C10)`);
  setFormula(sheet, "C12", `IF(C10=0,0,C7/C10)`);
  sheet["!cols"] = [{ wch: 34 }, { wch: 42 }, { wch: 22 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Executive MIS");
}

function appendPeopleAllocation(workbook: XLSX.WorkBook, profile: BusinessProfile, financials: UnitFinancials[], staff: StaffMember[]) {
  const months = periodMonths(profile);
  const sheet = makeSheet([
    ["Name", "Role", "Department", "Monthly cost", "Period cost", ...financials.map((unit) => unit.center.name), "Total assigned FTE"],
    ...staff.map((person) => [
      person.name,
      person.role,
      person.department,
      person.monthlyCost,
      person.monthlyCost * months,
      ...financials.map((unit) => person.assignments.find((assignment) => assignment.profitCenterId === unit.center.id)?.fte || 0),
      person.assignments.reduce((sum, assignment) => sum + assignment.fte, 0),
    ]),
  ]);
  sheet["!cols"] = [{ wch: 26 }, { wch: 22 }, { wch: 22 }, { wch: 16 }, { wch: 16 }, ...financials.map(() => ({ wch: 16 })), { wch: 16 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "People Allocation");
}

function appendDeepDive(workbook: XLSX.WorkBook, profile: BusinessProfile, financials: UnitFinancials[]) {
  const template = getBusinessTemplate(profile.businessType);
  const sheetName = sanitizeSheetName(template.deepDiveSheet);
  const sheet = makeSheet([
    [`${template.deepDiveSheet} Requirements`],
    [],
    ["Section", "Required detail", "Purpose", "Answer / source", template.unitLabel],
    ...template.scheduleRows.flatMap((row) =>
      financials.map((unit) => [row.section, row.requirement, row.purpose, unit.center.notes || "", unit.center.name]),
    ),
  ]);
  sheet["!cols"] = [{ wch: 22 }, { wch: 62 }, { wch: 58 }, { wch: 46 }, { wch: 28 }];
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
}

function appendWorkingCapital(workbook: XLSX.WorkBook, rows: TrialBalanceRow[], financials: UnitFinancials[]) {
  function byCategory(category: string) {
    return rows.filter((row) => row.category === category);
  }

  function byGroup(pattern: RegExp) {
    return rows.filter((row) => pattern.test(row.accountGroup.toLowerCase()));
  }

  function splitRows(sourceRows: TrialBalanceRow[]) {
    const centerAmounts = financials.map((unit) =>
      sourceRows.filter((row) => row.profitCenterId === unit.center.id || centerNameMatch(row.accountName, unit.center.name)).reduce((sum, row) => sum + amount(row), 0),
    );
    const total = sourceRows.reduce((sum, row) => sum + amount(row), 0);
    const centerTotal = centerAmounts.reduce((sum, value) => sum + value, 0);
    return [...centerAmounts, Math.max(total - centerTotal, 0), total];
  }

  const lines: Array<[string, TrialBalanceRow[]]> = [
    ["Client receivables / debtors", byGroup(/sundry debtors/)],
    ["Cash and bank", rows.filter((row) => row.category === "current-asset" && /bank|fd|cash/i.test(row.accountName))],
    ["Deposits, prepaid and advances", rows.filter((row) => row.category === "current-asset" && /deposit|prepaid|advance|itc|suspense|preliminary/i.test(`${row.accountName} ${row.accountGroup}`))],
    ["Other current assets", byCategory("current-asset")],
    ["Current liabilities and statutory dues", byCategory("current-liability")],
    ["Fixed assets / capex", byCategory("fixed-asset")],
    ["Equity / reserves", byCategory("equity")],
  ];

  const sheet = makeSheet([
    ["Balance sheet / working capital lens", ...financials.map((unit) => unit.center.name), "Shared / corporate", "Total"],
    ...lines.map(([label, sourceRows]) => [label, ...splitRows(sourceRows)]),
  ]);
  sheet["!cols"] = [{ wch: 36 }, ...financials.map(() => ({ wch: 18 })), { wch: 20 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Working Capital");
}

function appendCashFlow(workbook: XLSX.WorkBook) {
  const sheet = makeSheet([
    ["Cash Flow Bridge"],
    [],
    ["Particulars", "Amount"],
    ["Opening cash / bank", 0],
    ["Operating surplus", null],
    ["Change in receivables / WIP", 0],
    ["Change in inventory / advances", 0],
    ["Change in payables / deferred revenue", 0],
    ["Operating cash flow", null],
    ["Capex / fixed assets", null],
    ["Financing / owner movement", 0],
    ["Closing cash / bank", null],
  ]);
  setFormula(sheet, "B5", sheetCell("Executive MIS", "C8"));
  setFormula(sheet, "B9", "SUM(B5:B8)");
  setFormula(sheet, "B10", `SUM('Working Capital'!B4:ZZ4)`);
  setFormula(sheet, "B12", "B4+B9-B10+B11");
  sheet["!cols"] = [{ wch: 34 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Cash Flow");
}

function appendRawTrialBalance(workbook: XLSX.WorkBook, rows: TrialBalanceRow[], centers: ProfitCenter[]) {
  const sheet = makeSheet([
    ["Source Sheet", "Account", "Group", "Debit", "Credit", "Balance", "MIS Category", "MIS Group", "Operating Unit", "Confidence"],
    ...rows.map((row) => [
      row.sourceSheet,
      row.accountName,
      row.accountGroup,
      row.debit,
      row.credit,
      row.balance,
      categoryLabels[row.category],
      row.misGroup,
      centers.find((center) => center.id === row.profitCenterId)?.name || "Shared / unassigned",
      row.confidence,
    ]),
  ]);
  sheet["!cols"] = [{ wch: 16 }, { wch: 38 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 20 }, { wch: 22 }, { wch: 26 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Raw Trial Balance");
}

function appendQuestionAndGapSheets(workbook: XLSX.WorkBook, questions: GeneratedQuestion[], answers: QuestionAnswer[], issues: WorkbookIssue[], rows: TrialBalanceRow[]) {
  const questionSheet = makeSheet([
    ["Section", "Priority", "Question", "Reason", "Status", "Answer"],
    ...questions.map((question) => {
      const answer = answers.find((item) => item.id === question.id);
      return [question.section, question.priority, question.prompt, question.reason, answer?.status || "open", answer?.answer || ""];
    }),
    ...answers.filter((answer) => !questions.some((question) => question.id === answer.id)).map((answer) => ["User supplied", "low", answer.question, "", answer.status, answer.answer]),
  ]);
  questionSheet["!cols"] = [{ wch: 24 }, { wch: 12 }, { wch: 78 }, { wch: 72 }, { wch: 16 }, { wch: 72 }];
  XLSX.utils.book_append_sheet(workbook, questionSheet, "Question Register");

  const gapSheet = makeSheet([
    ["Severity", "Issue", "Detail"],
    ...issues.map((issue) => [issue.severity, issue.label, issue.detail]),
    ...rows.filter((row) => row.category === "unknown").map((row) => ["medium", "Unclassified ledger", `${row.accountName} (${row.accountGroup || "No group"})`]),
  ]);
  gapSheet["!cols"] = [{ wch: 12 }, { wch: 30 }, { wch: 92 }];
  XLSX.utils.book_append_sheet(workbook, gapSheet, "Data Gaps");
}

export function generateCustomMisWorkbook(params: CustomWorkbookParams) {
  const { profile, rows, centers, staff, questions, answers, issues, bankTransactions = [] } = params;
  const template = getBusinessTemplate(profile.businessType);
  const safeCenters = centers.length ? centers : [];
  const financials = buildUnitFinancials(profile, rows, safeCenters, staff);
  const workbook = XLSX.utils.book_new();

  appendControlPanel(workbook, profile, rows, safeCenters, staff, questions, issues);
  appendProductionManagementSummary(workbook, profile, financials, rows, issues);
  appendProductionEngagementPnl(workbook, profile, financials, staff, rows);
  appendRevenueRegister(workbook, profile, rows, financials);
  appendPeopleCostRegister(workbook, profile, rows, financials, staff);
  appendVendorDirectCostRegister(workbook, profile, rows, financials);
  appendSharedCostPoolsProduction(workbook, profile, rows, financials);
  appendWorkingCapitalProduction(workbook, profile, rows, financials);
  appendCashFlowProduction(workbook, profile, rows, financials);
  appendBalanceSheetProduction(workbook, profile, rows);
  appendTaxStatutorySchedule(workbook, rows);
  appendFundFlowStatement(workbook, profile, rows, bankTransactions);
  appendFundFlowProjection(workbook, profile, rows, bankTransactions);
  appendBankSourceAudit(workbook, profile, bankTransactions, "bank-statement", "Bank Statement Audit");
  appendBankSourceAudit(workbook, profile, bankTransactions, "bank-ledger", "Bank Ledger Audit");
  appendBankReconciliation(workbook, rows, bankTransactions);
  appendFundFlowAssumptions(workbook, profile, bankTransactions);
  appendChartDataProduction(workbook, profile, rows, financials, bankTransactions);
  appendQcTieOuts(workbook, profile, rows, financials);
  appendLedgerMappingAuditProduction(workbook, rows, centers);
  appendAccountAllocationMatrix(workbook, rows);
  appendAssumptionsQualityProduction(workbook, profile, rows, staff, issues, questions, answers);
  appendAiBuildInstructions(workbook);
  appendAiFormulaBuildSpec(workbook);
  appendCoverSheet(workbook, profile, financials, rows);
  appendManagementSummary(workbook, profile, financials, rows, issues);
  appendClientReadyEngagementPnl(workbook, profile, financials, staff, rows);
  appendClientScorecard(workbook, financials, staff, rows);
  appendCostPoolAllocation(workbook, rows, financials);
  appendAssumptionsQuality(workbook, profile, rows, staff, issues);
  appendExecutiveDashboard(workbook, profile, financials, issues);
  appendMethodology(workbook, profile);
  appendDriverBuildUp(workbook, profile, financials);
  appendEngagementSummary(workbook, profile, financials, staff);
  appendSharedCostAllocation(workbook, profile, rows, financials);
  appendLedgerCategorySummary(workbook, rows, centers);
  appendUnitProfitability(workbook, profile, financials);
  appendExecutiveMis(workbook, profile, financials);
  appendPeopleAllocation(workbook, profile, financials, staff);
  appendDeepDive(workbook, profile, financials);
  appendWorkingCapital(workbook, rows, financials);
  appendCashFlow(workbook);
  appendRawTrialBalance(workbook, rows, centers);
  appendQuestionAndGapSheets(workbook, questions, answers, issues, rows);

  return {
    workbook,
    issues,
    filename: `${sanitizeSheetName(profile.businessName || "Custom MIS")}-${profile.cadence}-${sanitizeSheetName(template.unitLabel)}-MIS.xlsx`.replace(/\s+/g, "-"),
  };
}
