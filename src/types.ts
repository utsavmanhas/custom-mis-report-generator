export type Suite = "mis-report" | "bank-recon" | "tds-recon" | "gst-recon";

export interface SuiteDefinition {
  id: Suite;
  label: string;
  description: string;
  availableNow: boolean;
}

export const SUITES: SuiteDefinition[] = [
  { id: "mis-report",  label: "MIS Report",        description: "Monthly management information system from Trial Balance", availableNow: true },
  { id: "bank-recon",  label: "Bank Reconciliation", description: "Reconcile bank statement with ledger balances",           availableNow: false },
  { id: "tds-recon",   label: "TDS Reconciliation",  description: "Match TDS deducted against Form 26AS",                   availableNow: false },
  { id: "gst-recon",   label: "GST Reconciliation",  description: "Reconcile GSTR-2A/2B with purchase register",            availableNow: false },
];

export interface ReferenceStructure {
  sheets: Array<{
    name: string;
    headers: string[];
    sampleRows: string[][];
    rowCount: number;
  }>;
}

export type BusinessType =
  | "consulting"
  | "professional-services"
  | "university"
  | "saas"
  | "d2c"
  | "manufacturing"
  | "custom";

export type PeriodCadence = "monthly" | "quarterly" | "annual";

export type FundFlowBasis = "bank-statement" | "bank-ledger" | "trial-balance-proxy" | "manual-assumptions";

export type AccountCategory =
  | "revenue"
  | "other-income"
  | "direct-cost"
  | "people-cost"
  | "operating-expense"
  | "finance-cost"
  | "tax"
  | "fixed-asset"
  | "current-asset"
  | "current-liability"
  | "equity"
  | "unknown";

export type AllocationBase = "revenue" | "headcount" | "fte" | "equal" | "manual";

export type ProfitCenterKind = "project" | "department" | "batch" | "product" | "vertical" | "location" | "custom";

export interface BusinessProfile {
  businessName: string;
  legalEntity: string;
  businessType: BusinessType;
  cadence: PeriodCadence;
  periodStart: string;
  periodEnd: string;
  currency: string;
  website: string;
  geography: string;
  publicNotes: string;
  sourceUrls: string;
  allocationBase: AllocationBase;
  fundFlowBasis: FundFlowBasis;
  projectionBasis: "past-year" | "past-month" | "manual";
  profileSource?: "user" | "trial-balance" | "default";
  tbSourceFileName?: string;
  tbPeriodText?: string;
  tbTotalDebit?: number;
  tbTotalCredit?: number;
}

export interface TrialBalanceMetadata {
  companyName: string;
  periodText: string;
  periodStart: string;
  periodEnd: string;
  totalDebit: number;
  totalCredit: number;
  sourceFileName: string;
}

export interface TrialBalanceGroupRow {
  sourceSheet: string;
  sourceRowNumber: number;
  accountName: string;
  accountPath: string[];
  debit: number;
  credit: number;
  rowType: "group" | "total";
}

export interface ParsedTrialBalance {
  rows: TrialBalanceRow[];
  metadata: TrialBalanceMetadata;
  groupRows: TrialBalanceGroupRow[];
  warnings: string[];
}

export interface TrialBalanceRow {
  id: string;
  sourceSheet: string;
  sourceRowNumber?: number;
  accountName: string;
  accountGroup: string;
  accountPath?: string[];
  hierarchyLevel?: number;
  debit: number;
  credit: number;
  balance: number;
  category: AccountCategory;
  misGroup: string;
  profitCenterId: string;
  allocationBase: AllocationBase;
  confidence: number;
  riskFlags?: string[];
  raw: Record<string, string | number | null>;
}

export interface BankMonthlySummary {
  month: string;
  receipts: number;
  payments: number;
  openingBalance: number | null;
  closingBalance: number | null;
  transactionCount: number;
}

export interface BankTransaction {
  id: string;
  sourceFileId: string;
  sourceType: "bank-statement" | "bank-ledger";
  sourceSheet: string;
  sourceRowNumber: number;
  date: string;
  description: string;
  debit: number;
  credit: number;
  balance: number | null;
  raw: Record<string, string | number | null>;
}

export interface BankSourceFile {
  id: string;
  sourceType: "bank-statement" | "bank-ledger";
  fileName: string;
  rowsImported: number;
  transactions: BankTransaction[];
  summary: {
    totalReceipts: number;
    totalPayments: number;
    openingBalance: number | null;
    closingBalance: number | null;
    monthly: BankMonthlySummary[];
    warnings: string[];
  };
}

export interface ProfitCenter {
  id: string;
  name: string;
  kind: ProfitCenterKind;
  owner: string;
  segment: string;
  revenueDriver: string;
  manualRevenue: number;
  manualDirectCost: number;
  priorRevenue: number;
  priorDirectCost: number;
  studentCount: number;
  teachingStaffCount: number;
  nonTeachingStaffCount: number;
  averageRevenueRate: number;
  variableCostRate: number;
  utilizationPercent: number;
  allocationWeight: number;
  notes: string;
}

export interface StaffAssignment {
  profitCenterId: string;
  fte: number;
}

export interface StaffMember {
  id: string;
  name: string;
  role: string;
  department: string;
  monthlyCost: number;
  assignments: StaffAssignment[];
}

export interface QuestionAnswer {
  id: string;
  question: string;
  answer: string;
  status: "open" | "answered" | "not-applicable";
}

export interface GeneratedQuestion {
  id: string;
  section: string;
  prompt: string;
  reason: string;
  priority: "high" | "medium" | "low";
}

export interface WorkbookIssue {
  label: string;
  detail: string;
  severity: "high" | "medium" | "low";
}
