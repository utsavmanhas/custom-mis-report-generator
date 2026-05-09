export type BusinessType =
  | "consulting"
  | "professional-services"
  | "university"
  | "saas"
  | "d2c"
  | "manufacturing"
  | "custom";

export type PeriodCadence = "monthly" | "quarterly" | "annual";

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

export type FundFlowBasis = "bank-statement" | "bank-ledger" | "bank-statement-and-ledger" | "trial-balance" | "manual";
export type BankSourceType = "bank-statement" | "bank-ledger";

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
}

export interface TrialBalanceRow {
  id: string;
  sourceSheet: string;
  accountName: string;
  accountGroup: string;
  accountPath?: string[];
  debit: number;
  credit: number;
  balance: number;
  category: AccountCategory;
  misGroup: string;
  profitCenterId: string;
  allocationBase: AllocationBase;
  confidence: number;
  raw: Record<string, string | number | null>;
}

export interface TrialBalanceMetadata {
  businessName: string;
  periodStart: string;
  periodEnd: string;
  sourceFileName: string;
  sourceSheetName: string;
  totalDebit: number;
  totalCredit: number;
}

export interface ParsedTrialBalance {
  rows: TrialBalanceRow[];
  metadata: TrialBalanceMetadata;
  warnings: string[];
}

export interface BankTransaction {
  id: string;
  sourceType: BankSourceType;
  sourceFileName: string;
  sourceSheet: string;
  date: string;
  narration: string;
  debit: number;
  credit: number;
  amount: number;
  balance: number;
  bankName: string;
  accountName: string;
  category: string;
  fundFlowGroup: string;
  isInterAccountTransfer: boolean;
  raw: Record<string, string | number | null>;
}

export interface ParsedBankStatement {
  transactions: BankTransaction[];
  warnings: string[];
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
