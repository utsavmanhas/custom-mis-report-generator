import * as XLSX from "xlsx";
import { allocationLabels, businessTypeLabels, categoryLabels, categorySignedAmount } from "./classification";
import type { BusinessProfile, GeneratedQuestion, ProfitCenter, QuestionAnswer, StaffMember, TrialBalanceRow, WorkbookIssue } from "../types";

const moneyFormat = '#,##0.00;[Red](#,##0.00);-';
const percentFormat = "0.0%";

interface UniversityWorkbookParams {
  profile: BusinessProfile;
  rows: TrialBalanceRow[];
  centers: ProfitCenter[];
  staff: StaffMember[];
  questions: GeneratedQuestion[];
  answers: QuestionAnswer[];
  issues: WorkbookIssue[];
}

interface EntityFinancials {
  id: string;
  name: string;
  currentFees: number;
  otherRevenue: number;
  teachingSalary: number;
  nonTeachingSalary: number;
  unclassifiedSalary: number;
  directExpense: number;
  assetPurchase: number;
  sharedExpense: number;
  priorRevenue: number;
  priorExpense: number;
  annualRevenue: number;
  annualExpense: number;
  studentCount: number;
  teachingStaffCount: number;
  nonTeachingStaffCount: number;
}

function aoa(data: unknown[][]) {
  return XLSX.utils.aoa_to_sheet(data);
}

function setFormula(sheet: XLSX.WorkSheet, address: string, formula: string, format = moneyFormat) {
  sheet[address] = { t: "n", f: formula, z: format };
}

function col(index: number) {
  return XLSX.utils.encode_col(index);
}

function sheetCell(sheetName: string, address: string) {
  return `'${sheetName.replace(/'/g, "''")}'!${address}`;
}

function sanitizeSheetName(name: string) {
  return name.replace(/[\\/?*[\]:]/g, " ").slice(0, 31).trim() || "Sheet";
}

function dateRange(profile: BusinessProfile) {
  if (profile.periodStart && profile.periodEnd) return `${profile.periodStart} to ${profile.periodEnd}`;
  return `${profile.cadence} reporting period`;
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

function emptyCenter(name: string): ProfitCenter {
  return {
    id: `virtual-${name.toLowerCase().replace(/\s+/g, "-")}`,
    name,
    kind: "custom",
    owner: "",
    segment: "",
    revenueDriver: "",
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
    notes: "",
  };
}

function hasText(center: ProfitCenter, words: string[]) {
  const text = `${center.name} ${center.kind} ${center.revenueDriver}`.toLowerCase();
  return words.some((word) => text.includes(word));
}

function splitUniversityCenters(centers: ProfitCenter[]) {
  const training = centers.filter((center) => hasText(center, ["training", "inspire", "workshop"]));
  const research = centers.filter((center) => hasText(center, ["research", "grant", "serb", "dbt", "drdo", "dae", "project"]));
  const hostel = centers.filter((center) => hasText(center, ["hostel", "housing", "residence"]));
  const transport = centers.filter((center) => hasText(center, ["transport", "bus", "fleet"]));
  const other = centers.filter((center) => hasText(center, ["other", "misc", "ancillary"]));
  const specialIds = new Set([...training, ...research, ...hostel, ...transport, ...other].map((center) => center.id));
  const academic = centers.filter((center) => !specialIds.has(center.id) && (center.kind === "department" || center.kind === "batch" || center.kind === "custom"));

  return {
    academic: academic.length ? academic : [emptyCenter("Faculty Wise Income")],
    training: training.length ? training : [emptyCenter("Training Projects")],
    research: research.length ? research : [emptyCenter("Research Projects")],
    hostel: hostel.length ? hostel : [emptyCenter("Hostel")],
    transport: transport.length ? transport : [emptyCenter("Transport")],
    other: other.length ? other : [emptyCenter("Others")],
  };
}

function signedAmount(row: TrialBalanceRow) {
  return Math.abs(categorySignedAmount(row));
}

function rowsForCenter(rows: TrialBalanceRow[], centerId: string, categories: string[]) {
  return rows.filter((row) => row.profitCenterId === centerId && categories.includes(row.category));
}

function assignedAmount(rows: TrialBalanceRow[], centerId: string, categories: string[]) {
  return rowsForCenter(rows, centerId, categories).reduce((sum, row) => sum + signedAmount(row), 0);
}

function unassignedAmount(rows: TrialBalanceRow[], categories: string[]) {
  return rows.filter((row) => !row.profitCenterId && categories.includes(row.category)).reduce((sum, row) => sum + signedAmount(row), 0);
}

function centerWeight(center: ProfitCenter, centers: ProfitCenter[]) {
  const total = centers.reduce((sum, item) => sum + Math.max(item.allocationWeight, 0), 0);
  if (total) return Math.max(center.allocationWeight, 0) / total;
  return centers.length ? 1 / centers.length : 0;
}

function salarySplit(staff: StaffMember[], centerId: string, months: number) {
  return staff.reduce(
    (acc, person) => {
      const assignedFte = person.assignments.reduce((sum, assignment) => sum + assignment.fte, 0);
      const centerFte = person.assignments.find((assignment) => assignment.profitCenterId === centerId)?.fte || 0;
      if (!assignedFte || !centerFte) return acc;

      const cost = person.monthlyCost * months * (centerFte / assignedFte);
      const role = `${person.role} ${person.department}`.toLowerCase();
      if (/teacher|teaching|faculty|professor|lecturer|dean/.test(role)) acc.teaching += cost;
      else acc.nonTeaching += cost;
      return acc;
    },
    { teaching: 0, nonTeaching: 0 },
  );
}

function createFinancials(params: {
  center: ProfitCenter;
  group: ProfitCenter[];
  rows: TrialBalanceRow[];
  staff: StaffMember[];
  months: number;
  fallbackUnassigned?: boolean;
}) {
  const { center, group, rows, staff, months, fallbackUnassigned = false } = params;
  const fallback = fallbackUnassigned ? centerWeight(center, group) : 0;
  const salary = salarySplit(staff, center.id, months);
  const assignedPayroll = assignedAmount(rows, center.id, ["people-cost"]);
  const unassignedPayroll = unassignedAmount(rows, ["people-cost"]) * fallback;
  const unclassifiedSalary = Math.max(0, assignedPayroll + unassignedPayroll - salary.teaching - salary.nonTeaching);
  const currentFees =
    assignedAmount(rows, center.id, ["revenue"]) +
    center.manualRevenue +
    unassignedAmount(rows, ["revenue"]) * fallback;
  const otherRevenue = assignedAmount(rows, center.id, ["other-income"]) + unassignedAmount(rows, ["other-income"]) * fallback;
  const directExpense =
    assignedAmount(rows, center.id, ["direct-cost", "operating-expense", "finance-cost", "tax"]) +
    center.manualDirectCost +
    unassignedAmount(rows, ["direct-cost"]) * fallback;
  const assetPurchase = assignedAmount(rows, center.id, ["fixed-asset"]) + unassignedAmount(rows, ["fixed-asset"]) * fallback;
  const annualFactor = 12 / months;

  return {
    id: center.id,
    name: center.name,
    currentFees,
    otherRevenue,
    teachingSalary: salary.teaching,
    nonTeachingSalary: salary.nonTeaching,
    unclassifiedSalary,
    directExpense,
    assetPurchase,
    sharedExpense: 0,
    priorRevenue: center.priorRevenue,
    priorExpense: center.priorDirectCost,
    annualRevenue: (currentFees + otherRevenue) * annualFactor,
    annualExpense: (salary.teaching + salary.nonTeaching + unclassifiedSalary + directExpense + assetPurchase) * annualFactor,
    studentCount: center.studentCount,
    teachingStaffCount: center.teachingStaffCount,
    nonTeachingStaffCount: center.nonTeachingStaffCount,
  } satisfies EntityFinancials;
}

function rowNumber(rowMap: Map<string, number>, label: string) {
  const row = rowMap.get(label);
  if (!row) throw new Error(`Missing university report row: ${label}`);
  return row;
}

function appendStatementSheet(workbook: XLSX.WorkBook, sheetName: string, title: string, entities: EntityFinancials[], profile: BusinessProfile, options?: { projectLabel?: string }) {
  const valueColumns = entities.map((_, index) => 1 + index * 2);
  const totalCol = 1 + entities.length * 2;
  const totalPctCol = totalCol + 1;
  const priorCol = totalCol + 2;
  const annualCol = totalCol + 3;
  const rowMap = new Map<string, number>();
  const months = periodMonths(profile);

  const rows: unknown[][] = [
    [title],
    [`Period: ${dateRange(profile)}`],
    [],
    [
      "Particulars",
      ...entities.flatMap((entity) => [entity.name, ""]),
      "Total",
      "%",
      "Prior period",
      "Annualised figures",
    ],
    ["", ...entities.flatMap(() => ["Amount", "% of revenue"]), "Upto period", "% of revenue", "Prior", "Annualised"],
    [],
  ];

  function add(label: string, values: number[], prior = 0, annual = 0) {
    rowMap.set(label, rows.length + 1);
    rows.push([label, ...entities.flatMap((_, index) => [values[index] || 0, null]), null, null, prior, annual]);
  }

  add("TOTAL REVENUE", entities.map(() => 0), entities.reduce((sum, item) => sum + item.priorRevenue, 0), entities.reduce((sum, item) => sum + item.annualRevenue, 0));
  add(options?.projectLabel || "Current Year Fees", entities.map((item) => item.currentFees));
  add("Other Revenue", entities.map((item) => item.otherRevenue));
  rows.push([]);
  add("TOTAL EXPENSES", entities.map(() => 0), entities.reduce((sum, item) => sum + item.priorExpense, 0), entities.reduce((sum, item) => sum + item.annualExpense, 0));
  add("Salary", entities.map(() => 0));
  add("Teaching Staff", entities.map((item) => item.teachingSalary));
  add("Non-Teaching Staff", entities.map((item) => item.nonTeachingSalary));
  add("Unclassified Payroll", entities.map((item) => item.unclassifiedSalary));
  add("Direct / Misc. Expenses", entities.map((item) => item.directExpense));
  add("Lab Equipment / Assets", entities.map((item) => item.assetPurchase));
  add("Allocated Shared Overheads", entities.map((item) => item.sharedExpense));
  rows.push([]);
  add("Excess of Income over Expenditure", entities.map(() => 0));
  add("Surplus Margin", entities.map(() => 0));
  rows.push([]);
  add("Students", entities.map((item) => item.studentCount));
  add("Teaching Staff Count", entities.map((item) => item.teachingStaffCount));
  add("Non-Teaching Staff Count", entities.map((item) => item.nonTeachingStaffCount));
  add("Revenue per Student", entities.map(() => 0));
  add("Cost per Student", entities.map(() => 0));

  const sheet = aoa(rows);
  const revenueRow = rowNumber(rowMap, "TOTAL REVENUE");
  const feeRow = rowNumber(rowMap, options?.projectLabel || "Current Year Fees");
  const otherRevenueRow = rowNumber(rowMap, "Other Revenue");
  const expenseRow = rowNumber(rowMap, "TOTAL EXPENSES");
  const salaryRow = rowNumber(rowMap, "Salary");
  const teachingRow = rowNumber(rowMap, "Teaching Staff");
  const nonTeachingRow = rowNumber(rowMap, "Non-Teaching Staff");
  const unclassifiedPayrollRow = rowNumber(rowMap, "Unclassified Payroll");
  const directRow = rowNumber(rowMap, "Direct / Misc. Expenses");
  const assetRow = rowNumber(rowMap, "Lab Equipment / Assets");
  const sharedRow = rowNumber(rowMap, "Allocated Shared Overheads");
  const surplusRow = rowNumber(rowMap, "Excess of Income over Expenditure");
  const marginRow = rowNumber(rowMap, "Surplus Margin");
  const studentsRow = rowNumber(rowMap, "Students");
  const revenuePerStudentRow = rowNumber(rowMap, "Revenue per Student");
  const costPerStudentRow = rowNumber(rowMap, "Cost per Student");

  valueColumns.forEach((valueCol) => {
    const valueLetter = col(valueCol);
    const pctLetter = col(valueCol + 1);
    setFormula(sheet, `${valueLetter}${revenueRow}`, `SUM(${valueLetter}${feeRow}:${valueLetter}${otherRevenueRow})`);
    setFormula(sheet, `${valueLetter}${expenseRow}`, `SUM(${valueLetter}${salaryRow},${valueLetter}${directRow}:${valueLetter}${sharedRow})`);
    setFormula(sheet, `${valueLetter}${salaryRow}`, `SUM(${valueLetter}${teachingRow}:${valueLetter}${unclassifiedPayrollRow})`);
    setFormula(sheet, `${valueLetter}${surplusRow}`, `${valueLetter}${revenueRow}-${valueLetter}${expenseRow}`);
    setFormula(sheet, `${valueLetter}${marginRow}`, `IF(${valueLetter}${revenueRow}=0,0,${valueLetter}${surplusRow}/${valueLetter}${revenueRow})`, percentFormat);
    setFormula(sheet, `${valueLetter}${revenuePerStudentRow}`, `IF(${valueLetter}${studentsRow}=0,0,${valueLetter}${revenueRow}/${valueLetter}${studentsRow})`);
    setFormula(sheet, `${valueLetter}${costPerStudentRow}`, `IF(${valueLetter}${studentsRow}=0,0,${valueLetter}${expenseRow}/${valueLetter}${studentsRow})`);

    for (let row = feeRow; row <= costPerStudentRow; row += 1) {
      if (row === marginRow) continue;
      setFormula(sheet, `${pctLetter}${row}`, `IF(${valueLetter}${revenueRow}=0,0,${valueLetter}${row}/${valueLetter}${revenueRow})`, percentFormat);
    }
  });

  for (const [label, row] of rowMap.entries()) {
    const totalLetter = col(totalCol);
    const totalPctLetter = col(totalPctCol);
    const priorLetter = col(priorCol);
    const annualLetter = col(annualCol);
    const parts = valueColumns.map((valueCol) => `${col(valueCol)}${row}`).join(",");

    if (label === "Surplus Margin") {
      setFormula(sheet, `${totalLetter}${row}`, `IF(${totalLetter}${revenueRow}=0,0,${totalLetter}${surplusRow}/${totalLetter}${revenueRow})`, percentFormat);
      setFormula(sheet, `${totalPctLetter}${row}`, `${totalLetter}${row}`, percentFormat);
      setFormula(sheet, `${annualLetter}${row}`, `IF(${annualLetter}${revenueRow}=0,0,${annualLetter}${surplusRow}/${annualLetter}${revenueRow})`, percentFormat);
    } else if (label === "TOTAL REVENUE") {
      setFormula(sheet, `${totalLetter}${row}`, `SUM(${parts})`);
      setFormula(sheet, `${totalPctLetter}${row}`, "1", percentFormat);
      setFormula(sheet, `${annualLetter}${row}`, `${totalLetter}${row}/${months}*12`);
    } else if (label === "TOTAL EXPENSES" || label === "Salary" || label === "Excess of Income over Expenditure") {
      setFormula(sheet, `${totalLetter}${row}`, `SUM(${parts})`);
      setFormula(sheet, `${totalPctLetter}${row}`, `IF(${totalLetter}${revenueRow}=0,0,${totalLetter}${row}/${totalLetter}${revenueRow})`, percentFormat);
      setFormula(sheet, `${annualLetter}${row}`, `${totalLetter}${row}/${months}*12`);
    } else if (label === "Revenue per Student") {
      setFormula(sheet, `${totalLetter}${row}`, `IF(${totalLetter}${studentsRow}=0,0,${totalLetter}${revenueRow}/${totalLetter}${studentsRow})`);
      setFormula(sheet, `${annualLetter}${row}`, `IF(${totalLetter}${studentsRow}=0,0,${annualLetter}${revenueRow}/${totalLetter}${studentsRow})`);
    } else if (label === "Cost per Student") {
      setFormula(sheet, `${totalLetter}${row}`, `IF(${totalLetter}${studentsRow}=0,0,${totalLetter}${expenseRow}/${totalLetter}${studentsRow})`);
      setFormula(sheet, `${annualLetter}${row}`, `IF(${totalLetter}${studentsRow}=0,0,${annualLetter}${expenseRow}/${totalLetter}${studentsRow})`);
    } else {
      setFormula(sheet, `${totalLetter}${row}`, `SUM(${parts})`);
      if (!["Students", "Teaching Staff Count", "Non-Teaching Staff Count"].includes(label)) {
        setFormula(sheet, `${totalPctLetter}${row}`, `IF(${totalLetter}${revenueRow}=0,0,${totalLetter}${row}/${totalLetter}${revenueRow})`, percentFormat);
        setFormula(sheet, `${annualLetter}${row}`, `${totalLetter}${row}/${months}*12`);
      }
    }

    if (label === "Excess of Income over Expenditure") {
      setFormula(sheet, `${priorLetter}${row}`, `${priorLetter}${revenueRow}-${priorLetter}${expenseRow}`);
    }
  }

  sheet["!cols"] = [{ wch: 34 }, ...entities.flatMap(() => [{ wch: 18 }, { wch: 12 }]), { wch: 18 }, { wch: 12 }, { wch: 18 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(workbook, sheet, sanitizeSheetName(sheetName));

  return { sheetName: sanitizeSheetName(sheetName), rowMap, totalCol, annualCol };
}

function appendIncomeAndExpenditure(
  workbook: XLSX.WorkBook,
  profile: BusinessProfile,
  refs: Array<{ label: string; sheet: string; totalCol: number; revenueRow: number; expenseRow: number; surplusRow: number; priorRevenue: number; priorExpense: number }>,
) {
  const rows: unknown[][] = [
    [`Income And Expenditure Account (${dateRange(profile)})`],
    [],
    ["Particulars", ...refs.flatMap((ref) => [ref.label, ""]), "Total", "%", "Prior period", "Annualised figures"],
    ["", ...refs.flatMap(() => ["Amount", "% of revenue"]), "Upto period", "% of revenue", "Prior", "Annualised"],
    [],
    ["Total Revenue", ...refs.flatMap(() => [null, null]), null, null, refs.reduce((sum, ref) => sum + ref.priorRevenue, 0), null],
    ["Total Expenses", ...refs.flatMap(() => [null, null]), null, null, refs.reduce((sum, ref) => sum + ref.priorExpense, 0), null],
    ["Salary", ...refs.flatMap(() => [null, null]), null, null, null, null],
    ["Misc. & Utilities Expenses", ...refs.flatMap(() => [null, null]), null, null, null, null],
    ["Excess of Income over Expenditure", ...refs.flatMap(() => [null, null]), null, null, null, null],
    ["Surplus Margin", ...refs.flatMap(() => [null, null]), null, null, null, null],
    [],
    ["Provision for security refund", ...refs.flatMap(() => [0, null]), null, null, null, null],
    ["Transferred to IT/equipment replacement fund", ...refs.flatMap(() => [0, null]), null, null, null, null],
    ["Transferred to building refreshment fund", ...refs.flatMap(() => [0, null]), null, null, null, null],
    ["Net after provisions", ...refs.flatMap(() => [null, null]), null, null, null, null],
  ];

  const sheet = aoa(rows);
  const totalCol = 1 + refs.length * 2;
  const totalPctCol = totalCol + 1;
  const priorCol = totalCol + 2;
  const annualCol = totalCol + 3;
  const months = periodMonths(profile);
  const row = {
    revenue: 6,
    expenses: 7,
    salary: 8,
    misc: 9,
    surplus: 10,
    margin: 11,
    security: 13,
    itFund: 14,
    building: 15,
    net: 16,
  };

  refs.forEach((ref, index) => {
    const valueCol = 1 + index * 2;
    const valueLetter = col(valueCol);
    const pctLetter = col(valueCol + 1);
    const sourceCol = col(ref.totalCol);

    setFormula(sheet, `${valueLetter}${row.revenue}`, sheetCell(ref.sheet, `${sourceCol}${ref.revenueRow}`));
    setFormula(sheet, `${valueLetter}${row.expenses}`, sheetCell(ref.sheet, `${sourceCol}${ref.expenseRow}`));
    setFormula(sheet, `${valueLetter}${row.salary}`, sheetCell(ref.sheet, `${sourceCol}${ref.expenseRow + 1}`));
    setFormula(sheet, `${valueLetter}${row.misc}`, `${valueLetter}${row.expenses}-${valueLetter}${row.salary}`);
    setFormula(sheet, `${valueLetter}${row.surplus}`, sheetCell(ref.sheet, `${sourceCol}${ref.surplusRow}`));
    setFormula(sheet, `${valueLetter}${row.margin}`, `IF(${valueLetter}${row.revenue}=0,0,${valueLetter}${row.surplus}/${valueLetter}${row.revenue})`, percentFormat);
    setFormula(sheet, `${valueLetter}${row.net}`, `${valueLetter}${row.surplus}-${valueLetter}${row.security}-${valueLetter}${row.itFund}-${valueLetter}${row.building}`);

    [row.revenue, row.expenses, row.salary, row.misc, row.surplus, row.security, row.itFund, row.building, row.net].forEach((line) => {
      setFormula(sheet, `${pctLetter}${line}`, `IF(${valueLetter}${row.revenue}=0,0,${valueLetter}${line}/${valueLetter}${row.revenue})`, percentFormat);
    });
  });

  for (let line = row.revenue; line <= row.net; line += 1) {
    if (line === 12) continue;
    const parts = refs.map((_, index) => `${col(1 + index * 2)}${line}`).join(",");
    const totalLetter = col(totalCol);
    const totalPctLetter = col(totalPctCol);
    const annualLetter = col(annualCol);

    if (line === row.margin) {
      setFormula(sheet, `${totalLetter}${line}`, `IF(${totalLetter}${row.revenue}=0,0,${totalLetter}${row.surplus}/${totalLetter}${row.revenue})`, percentFormat);
      setFormula(sheet, `${totalPctLetter}${line}`, `${totalLetter}${line}`, percentFormat);
      setFormula(sheet, `${annualLetter}${line}`, `${totalLetter}${line}`, percentFormat);
    } else {
      setFormula(sheet, `${totalLetter}${line}`, `SUM(${parts})`);
      setFormula(sheet, `${totalPctLetter}${line}`, `IF(${totalLetter}${row.revenue}=0,0,${totalLetter}${line}/${totalLetter}${row.revenue})`, percentFormat);
      setFormula(sheet, `${annualLetter}${line}`, `${totalLetter}${line}/${months}*12`);
    }
  }

  setFormula(sheet, `${col(priorCol)}${row.surplus}`, `${col(priorCol)}${row.revenue}-${col(priorCol)}${row.expenses}`);
  sheet["!cols"] = [{ wch: 38 }, ...refs.flatMap(() => [{ wch: 18 }, { wch: 12 }]), { wch: 18 }, { wch: 12 }, { wch: 18 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Income & Expenditure");
}

function appendSalarySheet(workbook: XLSX.WorkBook, sheetName: string, title: string, centers: ProfitCenter[], staff: StaffMember[], profile: BusinessProfile) {
  const months = periodMonths(profile);
  const rows: unknown[][] = [[title], [], ["S.No.", "Particulars", "Role", "Department", "Assigned FTE", "Monthly cost", "Period cost", "Assigned to"]];
  let serial = 1;

  centers.forEach((center) => {
    const people = staff.filter((person) => person.assignments.some((assignment) => assignment.profitCenterId === center.id && assignment.fte > 0));
    rows.push([people.length, `${center.name} - salary`, "", "", "", "", null, ""]);
    const subtotalRow = rows.length;
    people.forEach((person) => {
      const assignedFte = person.assignments.reduce((sum, assignment) => sum + assignment.fte, 0);
      const centerFte = person.assignments.find((assignment) => assignment.profitCenterId === center.id)?.fte || 0;
      rows.push([serial, person.name, person.role, person.department, centerFte, person.monthlyCost, assignedFte ? person.monthlyCost * months * (centerFte / assignedFte) : 0, center.name]);
      serial += 1;
    });
    const sheetSubtotalRow = subtotalRow + 1;
    const firstDetail = sheetSubtotalRow + 1;
    const lastDetail = rows.length;
    if (lastDetail >= firstDetail) {
      rows[subtotalRow - 1][6] = { t: "n", f: `SUM(G${firstDetail}:G${lastDetail})`, z: moneyFormat };
    }
  });

  const sheet = aoa(rows);
  sheet["!cols"] = [{ wch: 10 }, { wch: 36 }, { wch: 24 }, { wch: 24 }, { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 26 }];
  XLSX.utils.book_append_sheet(workbook, sheet, sanitizeSheetName(sheetName));
}

function appendMonthlySupport(workbook: XLSX.WorkBook, profile: BusinessProfile) {
  const months = monthLabels(profile);
  const totalCol = col(months.length + 1);
  const supportRows = [
    "Tuition Fees",
    "Hostel Fees",
    "Transport Fees",
    "Research Receipts",
    "Training Receipts",
    "Other Receipts",
    "Salary Paid",
    "Sundry Creditors Paid",
    "Assets Purchase",
    "Advance to Employees",
    "Direct Expenses",
    "Finance Costs",
  ];
  const rows: unknown[][] = [["Monthly Ledger Input"], [], ["Particulars", ...months, "TOTAL"], ...supportRows.map((label) => [label, ...months.map(() => 0), null])];
  const sheet = aoa(rows);
  supportRows.forEach((_, index) => {
    const row = index + 4;
    setFormula(sheet, `${totalCol}${row}`, `SUM(B${row}:${col(months.length)}${row})`);
  });
  sheet["!cols"] = [{ wch: 28 }, ...months.map(() => ({ wch: 14 })), { wch: 16 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Monthly Input");
}

function appendCashFlow(workbook: XLSX.WorkBook, profile: BusinessProfile) {
  const months = monthLabels(profile);
  const totalCol = col(months.length + 1);
  const rows: unknown[][] = [
    ["Cash Flow Statement"],
    [],
    ["Particulars", ...months, "Total"],
    ["Opening Balance", ...months.map(() => 0), null],
    [],
    ["Receipts", ...months.map(() => null), null],
    ["Tuition Fees", ...months.map((_, index) => ({ t: "n", f: `'Monthly Input'!${col(index + 1)}4`, z: moneyFormat })), null],
    ["Hostel Fees", ...months.map((_, index) => ({ t: "n", f: `'Monthly Input'!${col(index + 1)}5`, z: moneyFormat })), null],
    ["Transport Fees", ...months.map((_, index) => ({ t: "n", f: `'Monthly Input'!${col(index + 1)}6`, z: moneyFormat })), null],
    ["Research Projects", ...months.map((_, index) => ({ t: "n", f: `'Monthly Input'!${col(index + 1)}7`, z: moneyFormat })), null],
    ["Training Projects", ...months.map((_, index) => ({ t: "n", f: `'Monthly Input'!${col(index + 1)}8`, z: moneyFormat })), null],
    ["Other Receipts", ...months.map((_, index) => ({ t: "n", f: `'Monthly Input'!${col(index + 1)}9`, z: moneyFormat })), null],
    [],
    ["Payments", ...months.map(() => null), null],
    ["Salary Paid", ...months.map((_, index) => ({ t: "n", f: `'Monthly Input'!${col(index + 1)}10`, z: moneyFormat })), null],
    ["Sundry Creditors", ...months.map((_, index) => ({ t: "n", f: `'Monthly Input'!${col(index + 1)}11`, z: moneyFormat })), null],
    ["Assets Purchase", ...months.map((_, index) => ({ t: "n", f: `'Monthly Input'!${col(index + 1)}12`, z: moneyFormat })), null],
    ["Advance to Employees", ...months.map((_, index) => ({ t: "n", f: `'Monthly Input'!${col(index + 1)}13`, z: moneyFormat })), null],
    ["Direct Expenses", ...months.map((_, index) => ({ t: "n", f: `'Monthly Input'!${col(index + 1)}14`, z: moneyFormat })), null],
    ["Finance Costs", ...months.map((_, index) => ({ t: "n", f: `'Monthly Input'!${col(index + 1)}15`, z: moneyFormat })), null],
    [],
    ["Closing Balance", ...months.map(() => null), null],
  ];
  const sheet = aoa(rows);

  for (let index = 0; index < months.length; index += 1) {
    const letter = col(index + 1);
    setFormula(sheet, `${letter}6`, `SUM(${letter}7:${letter}12)`);
    setFormula(sheet, `${letter}14`, `SUM(${letter}15:${letter}20)`);
    setFormula(sheet, `${letter}22`, `${letter}4+${letter}6-${letter}14`);
    if (index < months.length - 1) setFormula(sheet, `${col(index + 2)}4`, `${letter}22`);
  }

  [4, 6, 7, 8, 9, 10, 11, 12, 14, 15, 16, 17, 18, 19, 20, 22].forEach((row) => {
    setFormula(sheet, `${totalCol}${row}`, `SUM(B${row}:${col(months.length)}${row})`);
  });

  sheet["!cols"] = [{ wch: 28 }, ...months.map(() => ({ wch: 14 })), { wch: 16 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Cash Flow Statement");
}

function appendMonthlyLedgerSheet(workbook: XLSX.WorkBook, sheetName: string, title: string, labels: string[], profile: BusinessProfile) {
  const months = monthLabels(profile);
  const totalCol = col(months.length + 1);
  const rows: unknown[][] = [[title], [], ["Particulars", ...months, "TOTAL", "Prior period"], ...labels.map((label) => [label, ...months.map(() => 0), null, 0]), ["Total", ...months.map(() => null), null, null]];
  const sheet = aoa(rows);
  labels.forEach((_, index) => {
    const row = index + 4;
    setFormula(sheet, `${totalCol}${row}`, `SUM(B${row}:${col(months.length)}${row})`);
  });
  const totalRow = labels.length + 4;
  for (let index = 0; index < months.length; index += 1) setFormula(sheet, `${col(index + 1)}${totalRow}`, `SUM(${col(index + 1)}4:${col(index + 1)}${totalRow - 1})`);
  setFormula(sheet, `${totalCol}${totalRow}`, `SUM(${totalCol}4:${totalCol}${totalRow - 1})`);
  sheet["!cols"] = [{ wch: 34 }, ...months.map(() => ({ wch: 14 })), { wch: 16 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(workbook, sheet, sanitizeSheetName(sheetName));
}

function appendBankReconciliation(workbook: XLSX.WorkBook, profile: BusinessProfile) {
  const rows: unknown[][] = [
    ["BANK RECONCILIATION"],
    [],
    ["NAME", "", `: ${profile.businessName}`],
    ["PERIOD", "", `: ${dateRange(profile)}`],
    ["ACCOUNT NO.", "", ": Bank account"],
    ["NAME OF BANK", "", ": Bank"],
    [],
    ["", "", "", `${profile.currency || "INR"} - Currency`],
    [],
    ["Balance per Books of Accounts", "", "", 0],
    [],
    ["Balance per Bank Statement", "", "", 0],
    [],
    ["Ch. No.", "Date", "Vendor Names", "Amount"],
    ["", "", "", 0],
    ["", "", "", 0],
    ["", "", "", 0],
    ["", "", "", 0],
    ["", "", "Total", null],
    ["", "", "Proper Bank Balance", null],
    ["", "", "Unreconciled Difference", null],
  ];
  const sheet = aoa(rows);
  setFormula(sheet, "D19", "SUM(D15:D18)");
  setFormula(sheet, "D20", "D12+D19");
  setFormula(sheet, "D21", "D10-D20");
  sheet["!cols"] = [{ wch: 18 }, { wch: 14 }, { wch: 32 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Bank Reconciliation");
}

function appendRawTrialBalance(workbook: XLSX.WorkBook, rows: TrialBalanceRow[], centers: ProfitCenter[]) {
  const sheet = aoa([
    ["Source Sheet", "Account", "Group", "Debit", "Credit", "Balance", "MIS Category", "MIS Group", "Profit Center", "Confidence"],
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
  sheet["!cols"] = [{ wch: 16 }, { wch: 38 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 20 }, { wch: 26 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Raw Trial Balance");
}

function appendRegisters(workbook: XLSX.WorkBook, questions: GeneratedQuestion[], answers: QuestionAnswer[], issues: WorkbookIssue[], rows: TrialBalanceRow[]) {
  const questionSheet = aoa([
    ["Section", "Priority", "Question", "Reason", "Status", "Answer"],
    ...questions.map((question) => {
      const answer = answers.find((item) => item.id === question.id);
      return [question.section, question.priority, question.prompt, question.reason, answer?.status || "open", answer?.answer || ""];
    }),
    ...answers.filter((answer) => !questions.some((question) => question.id === answer.id)).map((answer) => ["User supplied", "low", answer.question, "", answer.status, answer.answer]),
  ]);
  questionSheet["!cols"] = [{ wch: 22 }, { wch: 12 }, { wch: 78 }, { wch: 72 }, { wch: 16 }, { wch: 72 }];
  XLSX.utils.book_append_sheet(workbook, questionSheet, "Question Register");

  const gapSheet = aoa([
    ["Severity", "Issue", "Detail"],
    ...issues.map((issue) => [issue.severity, issue.label, issue.detail]),
    ...rows.filter((row) => row.category === "unknown").map((row) => ["medium", "Unclassified ledger", `${row.accountName} (${row.accountGroup || "No group"})`]),
  ]);
  gapSheet["!cols"] = [{ wch: 12 }, { wch: 30 }, { wch: 92 }];
  XLSX.utils.book_append_sheet(workbook, gapSheet, "Data Gaps");
}

function appendMethodology(workbook: XLSX.WorkBook, profile: BusinessProfile) {
  const sheet = aoa([
    ["Custom University MIS Generator"],
    ["Business", profile.businessName],
    ["Legal entity", profile.legalEntity],
    ["Business type", businessTypeLabels[profile.businessType]],
    ["Reporting period", dateRange(profile)],
    ["Currency", profile.currency],
    ["Default allocation base", allocationLabels[profile.allocationBase]],
    ["Website", profile.website],
    ["Geography", profile.geography],
    [],
    ["Detailed workbook tabs"],
    ["1", "Income & Expenditure consolidates academic, corporate, project, hostel, transport, and other verticals."],
    ["2", "Faculty Wise Fees breaks academic performance by department or batch."],
    ["3", "Corporate Centre, Training Projects, Research Project, Hostel, Transport, and Others follow standalone income statement schedules."],
    ["4", "Salary sheets preserve person-level assignment equivalents for department and cost-center attribution."],
    ["5", "Monthly Input, Cash Flow, creditors, advances, expenses, assets, bank reconciliation, notes, questions, and data gaps support auditability."],
  ]);
  sheet["!cols"] = [{ wch: 28 }, { wch: 100 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Methodology");
}

export function generateUniversityMisWorkbook(params: UniversityWorkbookParams) {
  const { profile, rows, centers, staff, questions, answers, issues } = params;
  const workbook = XLSX.utils.book_new();
  const months = periodMonths(profile);
  const split = splitUniversityCenters(centers);
  const hasAssignedRows = rows.some((row) => row.profitCenterId);

  appendMethodology(workbook, profile);
  appendMonthlySupport(workbook, profile);
  appendCashFlow(workbook, profile);

  const academicEntities = split.academic.map((center) => createFinancials({ center, group: split.academic, rows, staff, months, fallbackUnassigned: !hasAssignedRows }));
  const trainingEntities = split.training.map((center) => createFinancials({ center, group: split.training, rows, staff, months }));
  const researchEntities = split.research.map((center) => createFinancials({ center, group: split.research, rows, staff, months }));
  const hostelEntities = split.hostel.map((center) => createFinancials({ center, group: split.hostel, rows, staff, months }));
  const transportEntities = split.transport.map((center) => createFinancials({ center, group: split.transport, rows, staff, months }));
  const otherEntities = split.other.map((center) => createFinancials({ center, group: split.other, rows, staff, months, fallbackUnassigned: hasAssignedRows }));

  const facultyRef = appendStatementSheet(workbook, "Faculty Wise Fees", `Faculty Wise Income Statement (${dateRange(profile)})`, academicEntities, profile);
  const corporateEntity: EntityFinancials = {
    ...createFinancials({ center: emptyCenter("Corporate Centre"), group: [emptyCenter("Corporate Centre")], rows, staff, months }),
    currentFees: 0,
    otherRevenue: 0,
    directExpense: unassignedAmount(rows, ["operating-expense", "finance-cost", "tax"]),
    unclassifiedSalary: hasAssignedRows ? unassignedAmount(rows, ["people-cost"]) : 0,
  };
  const corporateRef = appendStatementSheet(workbook, "Corporate Centre", `Corporate Centre Cost (${dateRange(profile)})`, [corporateEntity], profile, { projectLabel: "Internal Recoveries" });
  const trainingRef = appendStatementSheet(workbook, "Training Projects", `Training Projects Income Statement (${dateRange(profile)})`, trainingEntities, profile, { projectLabel: "Project Fees" });
  const researchRef = appendStatementSheet(workbook, "Research Project", `Research Projects Income Statement (${dateRange(profile)})`, researchEntities, profile, { projectLabel: "Project Fees" });
  const hostelRef = appendStatementSheet(workbook, "Hostel", `Hostel Income Statement (${dateRange(profile)})`, hostelEntities, profile);
  const transportRef = appendStatementSheet(workbook, "Transport", `Transport Income Statement (${dateRange(profile)})`, transportEntities, profile);
  const otherRef = appendStatementSheet(workbook, "Others", `Others Income Statement (${dateRange(profile)})`, otherEntities, profile, { projectLabel: "Ancillary Receipts" });

  appendIncomeAndExpenditure(workbook, profile, [
    { label: "Faculty Wise Income", sheet: facultyRef.sheetName, totalCol: facultyRef.totalCol, revenueRow: rowNumber(facultyRef.rowMap, "TOTAL REVENUE"), expenseRow: rowNumber(facultyRef.rowMap, "TOTAL EXPENSES"), surplusRow: rowNumber(facultyRef.rowMap, "Excess of Income over Expenditure"), priorRevenue: academicEntities.reduce((sum, item) => sum + item.priorRevenue, 0), priorExpense: academicEntities.reduce((sum, item) => sum + item.priorExpense, 0) },
    { label: "Corporate Centre", sheet: corporateRef.sheetName, totalCol: corporateRef.totalCol, revenueRow: rowNumber(corporateRef.rowMap, "TOTAL REVENUE"), expenseRow: rowNumber(corporateRef.rowMap, "TOTAL EXPENSES"), surplusRow: rowNumber(corporateRef.rowMap, "Excess of Income over Expenditure"), priorRevenue: 0, priorExpense: corporateEntity.priorExpense },
    { label: "Training Projects", sheet: trainingRef.sheetName, totalCol: trainingRef.totalCol, revenueRow: rowNumber(trainingRef.rowMap, "TOTAL REVENUE"), expenseRow: rowNumber(trainingRef.rowMap, "TOTAL EXPENSES"), surplusRow: rowNumber(trainingRef.rowMap, "Excess of Income over Expenditure"), priorRevenue: trainingEntities.reduce((sum, item) => sum + item.priorRevenue, 0), priorExpense: trainingEntities.reduce((sum, item) => sum + item.priorExpense, 0) },
    { label: "Research Projects", sheet: researchRef.sheetName, totalCol: researchRef.totalCol, revenueRow: rowNumber(researchRef.rowMap, "TOTAL REVENUE"), expenseRow: rowNumber(researchRef.rowMap, "TOTAL EXPENSES"), surplusRow: rowNumber(researchRef.rowMap, "Excess of Income over Expenditure"), priorRevenue: researchEntities.reduce((sum, item) => sum + item.priorRevenue, 0), priorExpense: researchEntities.reduce((sum, item) => sum + item.priorExpense, 0) },
    { label: "Hostel", sheet: hostelRef.sheetName, totalCol: hostelRef.totalCol, revenueRow: rowNumber(hostelRef.rowMap, "TOTAL REVENUE"), expenseRow: rowNumber(hostelRef.rowMap, "TOTAL EXPENSES"), surplusRow: rowNumber(hostelRef.rowMap, "Excess of Income over Expenditure"), priorRevenue: hostelEntities.reduce((sum, item) => sum + item.priorRevenue, 0), priorExpense: hostelEntities.reduce((sum, item) => sum + item.priorExpense, 0) },
    { label: "Transport", sheet: transportRef.sheetName, totalCol: transportRef.totalCol, revenueRow: rowNumber(transportRef.rowMap, "TOTAL REVENUE"), expenseRow: rowNumber(transportRef.rowMap, "TOTAL EXPENSES"), surplusRow: rowNumber(transportRef.rowMap, "Excess of Income over Expenditure"), priorRevenue: transportEntities.reduce((sum, item) => sum + item.priorRevenue, 0), priorExpense: transportEntities.reduce((sum, item) => sum + item.priorExpense, 0) },
    { label: "Others", sheet: otherRef.sheetName, totalCol: otherRef.totalCol, revenueRow: rowNumber(otherRef.rowMap, "TOTAL REVENUE"), expenseRow: rowNumber(otherRef.rowMap, "TOTAL EXPENSES"), surplusRow: rowNumber(otherRef.rowMap, "Excess of Income over Expenditure"), priorRevenue: otherEntities.reduce((sum, item) => sum + item.priorRevenue, 0), priorExpense: otherEntities.reduce((sum, item) => sum + item.priorExpense, 0) },
  ]);

  appendSalarySheet(workbook, "Faculty Wise-Salary", "FACULTY WISE - SALARY", split.academic, staff, profile);
  appendSalarySheet(workbook, "Corporate-Salary", "SALARY Corporate Centre", [emptyCenter("Corporate Centre")], staff, profile);
  appendSalarySheet(workbook, "Transport-Salary", "TRANSPORT WISE SALARY", split.transport, staff, profile);
  appendMonthlyLedgerSheet(workbook, "Other Receipts", "Other Receipts", ["Development charges", "Examination Fees", "Fine / Late Fees", "Interest Income", "Rent Income", "Misc. Income", "Prospectus Fee", "Registration Fee"], profile);
  appendMonthlyLedgerSheet(workbook, "Sundry Creditors", "Sundry Creditors", ["Advertising and Marketing", "Scientific Supplies", "Repairs and Maintenance", "Utilities", "Professional Fees", "Security Services", "Housekeeping", "Other Vendors"], profile);
  appendMonthlyLedgerSheet(workbook, "Advance - Emp", "Loans & Advances", ["Faculty advances", "Staff imprest", "Travel advances", "Project advances", "Other advances"], profile);
  appendMonthlyLedgerSheet(workbook, "Expenses-Cash", "Expenses incurred in cash/cheque", ["Corporate Centre", "Faculty Wise Expenses", "Hostel Wise Expenses", "Research Project Expenses", "Training Project Expenses", "Transport Expenses"], profile);
  appendMonthlyLedgerSheet(workbook, "Assets Purchase", "Purchase of Fixed Assets", ["Furniture & Fixture", "Computer", "Electric Equipments", "Sports Equipments", "Lab Equipment", "Library Books", "Office Equipment"], profile);
  appendBankReconciliation(workbook, profile);
  const notes = aoa([
    ["Adjustments made in preparation of MIS report"],
    [],
    ["S. no.", "Particulars", "Adjustments"],
    ...answers.filter((answer) => answer.answer.trim()).map((answer, index) => [index + 1, answer.question, answer.answer]),
  ]);
  notes["!cols"] = [{ wch: 10 }, { wch: 56 }, { wch: 100 }];
  XLSX.utils.book_append_sheet(workbook, notes, "Notes");
  appendRawTrialBalance(workbook, rows, centers);
  appendRegisters(workbook, questions, answers, issues, rows);

  return {
    workbook,
    issues,
    filename: `${sanitizeSheetName(profile.businessName || "University MIS")}-${profile.cadence}-University-MIS.xlsx`.replace(/\s+/g, "-"),
  };
}
