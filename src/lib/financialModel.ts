import { categorySignedAmount } from "./classification";
import type { BankMonthlySummary, BankSourceFile, BusinessProfile, TrialBalanceRow } from "../types";

export interface ScheduleLine {
  section: string;
  account: string;
  amount: number;
  debit: number;
  credit: number;
  path: string;
  note?: string;
}

export interface BalanceSheetModel {
  liabilities: ScheduleLine[];
  assets: ScheduleLine[];
  totals: {
    liabilities: number;
    assets: number;
    difference: number;
    bankAndFd: number;
    debtors: number;
    customerAdvances: number;
  };
}

export interface TaxScheduleLine {
  section: string;
  account: string;
  debit: number;
  credit: number;
  assetAmount: number;
  liabilityAmount: number;
  path: string;
}

export interface FundFlowMonth {
  month: string;
  receipts: number;
  payments: number;
  netMovement: number;
  openingBalance: number | null;
  closingBalance: number | null;
  transactionCount: number;
  projection: boolean;
}

export interface FundFlowModel {
  basis: BusinessProfile["fundFlowBasis"];
  status: "actual-from-bank-source" | "provisional-tb-proxy" | "manual-assumptions-required";
  openingCashBank: number | null;
  closingCashBank: number;
  actualMonths: FundFlowMonth[];
  projectionMonths: FundFlowMonth[];
  notes: string[];
}

function lower(value: string | undefined) {
  return String(value || "").toLowerCase();
}

export function rowPath(row: TrialBalanceRow) {
  return (row.accountPath?.length ? row.accountPath : [row.accountGroup, row.accountName].filter(Boolean)).join(" > ");
}

export function rowNet(row: TrialBalanceRow) {
  return row.debit - row.credit;
}

export function rowAbs(row: TrialBalanceRow) {
  return Math.abs(rowNet(row));
}

export function balanceSheetAmount(row: TrialBalanceRow) {
  if (row.category === "current-liability" || row.category === "equity") return Math.max(row.credit - row.debit, 0);
  return Math.max(row.debit - row.credit, 0);
}

export function isBankOrFd(row: TrialBalanceRow) {
  const text = lower(`${row.accountName} ${rowPath(row)}`);
  return row.category === "current-asset" && (/\bbank\b|bank accounts|cash/.test(text) || /\bfd\b|fixed deposit/.test(lower(row.accountName)));
}

export function isDebtor(row: TrialBalanceRow) {
  return /sundry debtors|debtor|receivable/.test(lower(rowPath(row)));
}

export function isCreditor(row: TrialBalanceRow) {
  return /sundry creditors|creditor|payable/.test(lower(rowPath(row)));
}

function liabilitySection(row: TrialBalanceRow) {
  const text = lower(`${row.accountName} ${rowPath(row)}`);
  if (/capital|shareholder|reserve|retained/.test(text)) return "Capital & equity";
  if (/gst|tds|pf|epf|esi|esic|edli|duties|tax/.test(text)) return "GST, TDS, PF/ESIC and statutory payables";
  if (/provision|employee payable|bonus payable|audit fee/.test(text)) return "Provisions and employee payables";
  if (/sundry creditors|creditor/.test(text)) return "Sundry creditors";
  if (/reimbursement/.test(text)) return "Reimbursements payable";
  if (isDebtor(row) && row.credit > row.debit) return "Customer advances / credit-balance debtors";
  return "Other current liabilities";
}

function assetSection(row: TrialBalanceRow) {
  const text = lower(`${row.accountName} ${rowPath(row)}`);
  if (/fixed assets|computer|laptop|office equipment|plant|machinery/.test(text)) return "Fixed assets";
  if (/deposit/.test(text)) return "Deposits";
  if (isDebtor(row)) return "Trade receivables / debtors";
  if (isBankOrFd(row)) return "Bank, cash and fixed deposits";
  if (/input tax credit|\bitc\b|gst credit/.test(text)) return "Input tax credit";
  if (/advance tax/.test(text)) return "Advance tax";
  if (/prepaid|preliminary/.test(text)) return "Prepaids and deferred assets";
  if (/suspense/.test(text)) return "Suspense and control accounts";
  if (/reimbursement/.test(text)) return "Reimbursements receivable";
  return "Other current assets";
}

export function buildBalanceSheetModel(rows: TrialBalanceRow[]): BalanceSheetModel {
  const liabilities: ScheduleLine[] = [];
  const assets: ScheduleLine[] = [];

  rows.forEach((row) => {
    const net = rowNet(row);
    const path = rowPath(row);

    if (row.category === "equity" || row.category === "current-liability" || (isDebtor(row) && net < 0)) {
      const amount = Math.abs(net);
      if (!amount) return;
      liabilities.push({
        section: liabilitySection(row),
        account: row.accountName,
        amount,
        debit: row.debit,
        credit: row.credit,
        path,
        note: isDebtor(row) && net < 0 ? "Credit balance in debtor ledger; present as customer advance/liability, not ordinary debtor revenue." : undefined,
      });
      return;
    }

    if (row.category === "fixed-asset" || row.category === "current-asset") {
      const amount = Math.abs(net);
      if (!amount) return;
      assets.push({
        section: assetSection(row),
        account: row.accountName,
        amount,
        debit: row.debit,
        credit: row.credit,
        path,
      });
    }
  });

  const liabilityTotal = liabilities.reduce((sum, item) => sum + item.amount, 0);
  const assetTotal = assets.reduce((sum, item) => sum + item.amount, 0);
  const bankAndFd = assets.filter((item) => item.section === "Bank, cash and fixed deposits").reduce((sum, item) => sum + item.amount, 0);
  const debtors = assets.filter((item) => item.section === "Trade receivables / debtors").reduce((sum, item) => sum + item.amount, 0);
  const customerAdvances = liabilities.filter((item) => item.section === "Customer advances / credit-balance debtors").reduce((sum, item) => sum + item.amount, 0);

  return {
    liabilities,
    assets,
    totals: {
      liabilities: liabilityTotal,
      assets: assetTotal,
      difference: assetTotal - liabilityTotal,
      bankAndFd,
      debtors,
      customerAdvances,
    },
  };
}

export function buildTaxSchedule(rows: TrialBalanceRow[]) {
  const taxRows = rows.filter((row) => {
    const text = lower(`${row.accountName} ${rowPath(row)}`);
    return /\bgst\b|\btds\b|\bitc\b|input tax|advance tax|pf\b|epf|esi|esic|professional tax|income tax|tax payable|duties/.test(text);
  });

  return taxRows.map<TaxScheduleLine>((row) => {
    const net = rowNet(row);
    return {
      section: net < 0 ? "Payable / statutory liability" : "Recoverable / advance",
      account: row.accountName,
      debit: row.debit,
      credit: row.credit,
      assetAmount: Math.max(net, 0),
      liabilityAmount: Math.max(-net, 0),
      path: rowPath(row),
    };
  });
}

export function buildChartData(rows: TrialBalanceRow[], bankSources: BankSourceFile[] = []) {
  const byCategory = new Map<string, number>();
  rows.forEach((row) => {
    const amount = Math.abs(categorySignedAmount(row));
    byCategory.set(row.category, (byCategory.get(row.category) || 0) + amount);
  });

  const balanceSheet = buildBalanceSheetModel(rows);
  const receipts = bankSources.flatMap((source) => source.summary.monthly).reduce((sum, month) => sum + month.receipts, 0);
  const payments = bankSources.flatMap((source) => source.summary.monthly).reduce((sum, month) => sum + month.payments, 0);

  return {
    revenueMix: rows
      .filter((row) => row.category === "revenue" || row.category === "other-income")
      .map((row) => ({ label: row.accountName, value: Math.abs(categorySignedAmount(row)) })),
    expenseMix: rows
      .filter((row) => ["direct-cost", "people-cost", "operating-expense", "finance-cost", "tax"].includes(row.category))
      .map((row) => ({ label: row.accountName, value: Math.abs(categorySignedAmount(row)) })),
    categoryMix: Array.from(byCategory.entries()).map(([label, value]) => ({ label, value })),
    assetComposition: balanceSheet.assets.map((line) => ({ label: line.section, account: line.account, value: line.amount })),
    liabilityComposition: balanceSheet.liabilities.map((line) => ({ label: line.section, account: line.account, value: line.amount })),
    receiptsVsPayments: [
      { label: "Receipts", value: receipts },
      { label: "Payments", value: payments },
    ],
  };
}

function toMonthlyMap(bankSources: BankSourceFile[]) {
  const map = new Map<string, BankMonthlySummary>();
  bankSources.forEach((source) => {
    source.summary.monthly.forEach((month) => {
      const current = map.get(month.month) || {
        month: month.month,
        receipts: 0,
        payments: 0,
        openingBalance: null,
        closingBalance: null,
        transactionCount: 0,
      };
      current.receipts += month.receipts;
      current.payments += month.payments;
      current.transactionCount += month.transactionCount;
      current.openingBalance = current.openingBalance ?? month.openingBalance;
      current.closingBalance = month.closingBalance ?? current.closingBalance;
      map.set(month.month, current);
    });
  });
  return Array.from(map.values()).sort((a, b) => a.month.localeCompare(b.month));
}

function addMonths(monthKey: string, offset: number) {
  const [year, month] = monthKey.split("-").map((part) => Number.parseInt(part, 10));
  const date = new Date(year, month - 1 + offset, 1);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

export function buildFundFlowModel(profile: BusinessProfile, rows: TrialBalanceRow[], bankSources: BankSourceFile[] = []): FundFlowModel {
  const bankAndFd = rows.filter(isBankOrFd).reduce((sum, row) => sum + Math.max(rowNet(row), 0), 0);
  const monthly = toMonthlyMap(bankSources);
  const hasBankData = monthly.length > 0;
  const basis = profile.fundFlowBasis;
  const actualMonths = monthly.map<FundFlowMonth>((month) => ({
    month: month.month,
    receipts: month.receipts,
    payments: month.payments,
    netMovement: month.receipts - month.payments,
    openingBalance: month.openingBalance,
    closingBalance: month.closingBalance,
    transactionCount: month.transactionCount,
    projection: false,
  }));

  const averageBase = profile.projectionBasis === "past-month" ? actualMonths.slice(-1) : actualMonths.slice(-12);
  const avgReceipts = averageBase.length ? averageBase.reduce((sum, month) => sum + month.receipts, 0) / averageBase.length : 0;
  const avgPayments = averageBase.length ? averageBase.reduce((sum, month) => sum + month.payments, 0) / averageBase.length : 0;
  const projectionStart = actualMonths.length ? addMonths(actualMonths[actualMonths.length - 1].month, 1) : addMonths((profile.periodEnd || "2026-03-31").slice(0, 7), 1);
  let projectedOpening = actualMonths.at(-1)?.closingBalance ?? bankAndFd;
  const projectionMonths: FundFlowMonth[] = Array.from({ length: 12 }, (_, index) => {
    const month = addMonths(projectionStart, index);
    const closing = projectedOpening + avgReceipts - avgPayments;
    const item: FundFlowMonth = {
      month,
      receipts: avgReceipts,
      payments: avgPayments,
      netMovement: avgReceipts - avgPayments,
      openingBalance: projectedOpening,
      closingBalance: closing,
      transactionCount: 0,
      projection: true,
    };
    projectedOpening = closing;
    return item;
  });

  return {
    basis,
    status: hasBankData ? "actual-from-bank-source" : basis === "manual-assumptions" ? "manual-assumptions-required" : "provisional-tb-proxy",
    openingCashBank: actualMonths[0]?.openingBalance ?? null,
    closingCashBank: actualMonths.at(-1)?.closingBalance ?? bankAndFd,
    actualMonths,
    projectionMonths,
    notes: [
      hasBankData
        ? "Fund flow is based on uploaded bank statement / bank ledger transactions."
        : "Fund flow is provisional until a bank statement or bank ledger is uploaded; only TB-derived bank and FD closing balances are available.",
      "Projection rows use the selected historical basis and should be adjusted for known one-off payments, bonus, taxes, capex, loans and collections.",
    ],
  };
}
