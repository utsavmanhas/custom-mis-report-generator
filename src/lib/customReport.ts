import * as XLSX from "xlsx";
import { allocationLabels, businessTypeLabels, categoryLabels, categorySignedAmount } from "./classification";
import { getBusinessTemplate } from "./businessTemplates";
import { buildBalanceSheetModel, buildChartData, buildFundFlowModel, buildTaxSchedule, isBankOrFd, rowPath } from "./financialModel";
import type { BankSourceFile, BusinessProfile, GeneratedQuestion, ProfitCenter, QuestionAnswer, StaffMember, TrialBalanceRow, WorkbookIssue } from "../types";

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
  bankSources?: BankSourceFile[];
}

interface UnitFinancials {
  center: ProfitCenter;
  driverRow: number;
  ledgerRevenue: number;
  otherIncome: number;
  directCost: number;
  peopleCost: number;
  financeTax: number;
  sharedWeight: number;
  sharedOpex: number;
  priorRevenue: number;
  priorCost: number;
}

function makeSheet(data: unknown[][]) {
  return XLSX.utils.aoa_to_sheet(data);
}

function setFormula(sheet: XLSX.WorkSheet, address: string, formula: string, format = moneyFormat) {
  sheet[address] = { t: "n", f: formula, z: format };
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

function amount(row: TrialBalanceRow) {
  return Math.abs(categorySignedAmount(row));
}

function numeric(value: number | undefined) {
  return Number.isFinite(value) ? Number(value) : 0;
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
  const sharedOpexPool = unassignedAmount(rows, ["operating-expense"]);
  const sharedFinanceTaxPool = unassignedAmount(rows, ["finance-cost", "tax"]);

  return centers.map<UnitFinancials>((center, index) => ({
    center,
    driverRow: index + 2,
    ledgerRevenue: assignedAmount(rows, center.id, ["revenue"]) + numeric(center.manualRevenue),
    otherIncome: assignedAmount(rows, center.id, ["other-income"]),
    directCost: assignedAmount(rows, center.id, ["direct-cost", "operating-expense"]) + numeric(center.manualDirectCost),
    peopleCost: salaryForCenter(staff, center.id, months) || assignedAmount(rows, center.id, ["people-cost"]),
    financeTax: assignedAmount(rows, center.id, ["finance-cost", "tax"]) + sharedFinanceTaxPool * weights[index],
    sharedWeight: weights[index],
    sharedOpex: sharedOpexPool * weights[index],
    priorRevenue: numeric(center.priorRevenue),
    priorCost: numeric(center.priorDirectCost),
  }));
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
  const categories = [
    ["Current assets", "current-asset"],
    ["Current liabilities", "current-liability"],
    ["Fixed assets / capex", "fixed-asset"],
    ["Equity / reserves", "equity"],
  ];
  const totalCol = col(1 + financials.length);
  const sheet = makeSheet([
    ["Balance sheet / working capital lens", ...financials.map((unit) => unit.center.name), "Total"],
    ...categories.map(([label, category]) => [
      label,
      ...financials.map((unit) => assignedAmount(rows, unit.center.id, [category])),
      rows.filter((row) => row.category === category).reduce((sum, row) => sum + amount(row), 0),
    ]),
    [],
    ["Receivables ageing", ...financials.map(() => 0), 0],
    ["Inventory / WIP ageing", ...financials.map(() => 0), 0],
    ["Advances / deferred revenue", ...financials.map(() => 0), 0],
    ["Payables ageing", ...financials.map(() => 0), 0],
  ]);
  [2, 3, 4, 5, 7, 8, 9, 10].forEach((row) => {
    setFormula(sheet, `${totalCol}${row}`, `SUM(B${row}:${col(financials.length)}${row})`);
  });
  sheet["!cols"] = [{ wch: 32 }, ...financials.map(() => ({ wch: 18 })), { wch: 18 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Working Capital");
}

function appendCashFlow(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], bankSources: BankSourceFile[]) {
  const fundFlow = buildFundFlowModel(profile, rows, bankSources);
  const sheet = makeSheet([
    ["Fund Flow / Cash Flow Bridge"],
    [],
    ["Basis", fundFlow.basis],
    ["Status", fundFlow.status],
    ["Opening cash / bank", fundFlow.openingCashBank ?? 0],
    ["Closing cash / bank / FD", fundFlow.closingCashBank],
    [],
    ["Month", "Receipts", "Payments", "Net movement", "Opening", "Closing", "Txns", "Projection?"],
    ...fundFlow.actualMonths.map((month) => [month.month, month.receipts, month.payments, month.netMovement, month.openingBalance ?? "", month.closingBalance ?? "", month.transactionCount, "No"]),
    ...fundFlow.projectionMonths.map((month) => [month.month, month.receipts, month.payments, month.netMovement, month.openingBalance ?? "", month.closingBalance ?? "", month.transactionCount, "Yes"]),
    [],
    ["Notes"],
    ...fundFlow.notes.map((note) => [note]),
  ]);
  sheet["!cols"] = [{ wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 10 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Fund Flow");
}

function appendBalanceSheet(workbook: XLSX.WorkBook, rows: TrialBalanceRow[]) {
  const balanceSheet = buildBalanceSheetModel(rows);
  const data = [
    ["Balance Sheet Source"],
    [],
    ["Liabilities", "", "", "", "", ""],
    ["Section", "Account", "Amount", "Debit", "Credit", "Path / note"],
    ...balanceSheet.liabilities.map((line) => [line.section, line.account, line.amount, line.debit, line.credit, line.note ? `${line.path} - ${line.note}` : line.path]),
    ["Total liabilities", "", balanceSheet.totals.liabilities, "", "", ""],
    [],
    ["Assets", "", "", "", "", ""],
    ["Section", "Account", "Amount", "Debit", "Credit", "Path"],
    ...balanceSheet.assets.map((line) => [line.section, line.account, line.amount, line.debit, line.credit, line.path]),
    ["Total assets", "", balanceSheet.totals.assets, "", "", ""],
    ["Assets minus liabilities", "", balanceSheet.totals.difference, "", "", ""],
  ];
  const sheet = makeSheet(data);
  sheet["!cols"] = [{ wch: 36 }, { wch: 42 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 80 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Balance Sheet");
}

function appendTaxStatutory(workbook: XLSX.WorkBook, rows: TrialBalanceRow[]) {
  const taxRows = buildTaxSchedule(rows);
  const sheet = makeSheet([
    ["Tax & Statutory Schedule"],
    [],
    ["Section", "Account", "Debit", "Credit", "Recoverable / advance", "Payable", "Path"],
    ...taxRows.map((row) => [row.section, row.account, row.debit, row.credit, row.assetAmount, row.liabilityAmount, row.path]),
    [],
    ["Tax provision prompt", "Enter expected income-tax provision, deferred tax, assessments, and payment timing in the Question Register."],
  ]);
  sheet["!cols"] = [{ wch: 30 }, { wch: 38 }, { wch: 16 }, { wch: 16 }, { wch: 22 }, { wch: 18 }, { wch: 72 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Tax & Statutory");
}

function appendChartData(workbook: XLSX.WorkBook, rows: TrialBalanceRow[], bankSources: BankSourceFile[]) {
  const chartData = buildChartData(rows, bankSources);
  const sheet = makeSheet([
    ["Chart", "Label", "Account", "Value"],
    ...chartData.revenueMix.map((item) => ["Revenue mix", item.label, item.label, item.value]),
    ...chartData.expenseMix.map((item) => ["Expense mix", item.label, item.label, item.value]),
    ...chartData.assetComposition.map((item) => ["Asset composition", item.label, item.account, item.value]),
    ...chartData.liabilityComposition.map((item) => ["Liability composition", item.label, item.account, item.value]),
    ...chartData.receiptsVsPayments.map((item) => ["Receipts vs payments", item.label, item.label, item.value]),
  ]);
  sheet["!cols"] = [{ wch: 24 }, { wch: 34 }, { wch: 42 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "ChartData");
}

function appendQcTieOut(workbook: XLSX.WorkBook, profile: BusinessProfile, rows: TrialBalanceRow[], bankSources: BankSourceFile[]) {
  const ledgerDebit = rows.reduce((sum, row) => sum + row.debit, 0);
  const ledgerCredit = rows.reduce((sum, row) => sum + row.credit, 0);
  const totalDebit = profile.tbTotalDebit || ledgerDebit;
  const totalCredit = profile.tbTotalCredit || ledgerCredit;
  const revenue = rows.filter((row) => row.category === "revenue").reduce((sum, row) => sum + Math.abs(categorySignedAmount(row)), 0);
  const otherIncome = rows.filter((row) => row.category === "other-income").reduce((sum, row) => sum + Math.abs(categorySignedAmount(row)), 0);
  const bankAndFd = rows.filter(isBankOrFd).reduce((sum, row) => sum + Math.max(row.debit - row.credit, 0), 0);
  const duplicateRows = rows.filter((row) => row.riskFlags?.some((flag) => /duplicate|suspense/i.test(flag)));
  const sheet = makeSheet([
    ["QC Tie-Outs"],
    [],
    ["Check", "Value", "Expected / note"],
    ["Source TB debit", totalDebit, "Should equal source TB credit"],
    ["Source TB credit", totalCredit, "Should equal source TB debit"],
    ["TB difference", totalDebit - totalCredit, "Should be zero"],
    ["Imported ledger debit", ledgerDebit, "Ledger-row sum after excluding group subtotals"],
    ["Imported ledger credit", ledgerCredit, "Ledger-row sum after excluding group subtotals"],
    ["Operating revenue", revenue, "From revenue ledgers only"],
    ["Other income", otherIncome, "Shown separately from operating revenue"],
    ["Bank + FD closing", bankAndFd, "TB-derived bank and FD balances"],
    ["Bank source files", bankSources.length, "Statement / ledger uploads"],
    [],
    ["Risk flag", "Account", "Path"],
    ...duplicateRows.map((row) => [(row.riskFlags || []).join(", "), row.accountName, rowPath(row)]),
  ]);
  sheet["!cols"] = [{ wch: 28 }, { wch: 24 }, { wch: 80 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "QC Tie-Outs");
}

function appendRawTrialBalance(workbook: XLSX.WorkBook, rows: TrialBalanceRow[], centers: ProfitCenter[]) {
  const sheet = makeSheet([
    ["Source Sheet", "Source Row", "Account", "Group / Path", "Debit", "Credit", "Balance", "MIS Category", "MIS Group", "Operating Unit", "Confidence", "Risk Flags"],
    ...rows.map((row) => [
      row.sourceSheet,
      row.sourceRowNumber || "",
      row.accountName,
      rowPath(row),
      row.debit,
      row.credit,
      row.balance,
      categoryLabels[row.category],
      row.misGroup,
      centers.find((center) => center.id === row.profitCenterId)?.name || "Shared / unassigned",
      row.confidence,
      (row.riskFlags || []).join(", "),
    ]),
  ]);
  sheet["!cols"] = [{ wch: 16 }, { wch: 12 }, { wch: 38 }, { wch: 54 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 20 }, { wch: 22 }, { wch: 26 }, { wch: 12 }, { wch: 32 }];
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
  const { profile, rows, centers, staff, questions, answers, issues } = params;
  const bankSources = params.bankSources || [];
  const template = getBusinessTemplate(profile.businessType);
  const safeCenters = centers.length ? centers : [];
  const financials = buildUnitFinancials(profile, rows, safeCenters, staff);
  const workbook = XLSX.utils.book_new();

  appendMethodology(workbook, profile);
  appendDriverBuildUp(workbook, profile, financials);
  appendSharedCostAllocation(workbook, profile, rows, financials);
  appendUnitProfitability(workbook, profile, financials);
  appendExecutiveMis(workbook, profile, financials);
  appendBalanceSheet(workbook, rows);
  appendTaxStatutory(workbook, rows);
  appendPeopleAllocation(workbook, profile, financials, staff);
  appendDeepDive(workbook, profile, financials);
  appendWorkingCapital(workbook, rows, financials);
  appendCashFlow(workbook, profile, rows, bankSources);
  appendChartData(workbook, rows, bankSources);
  appendQcTieOut(workbook, profile, rows, bankSources);
  appendRawTrialBalance(workbook, rows, centers);
  appendQuestionAndGapSheets(workbook, questions, answers, issues, rows);

  return {
    workbook,
    issues,
    filename: `${sanitizeSheetName(profile.businessName || "Custom MIS")}-${profile.cadence}-${sanitizeSheetName(template.unitLabel)}-MIS.xlsx`.replace(/\s+/g, "-"),
  };
}
