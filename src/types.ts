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
}

export interface TrialBalanceRow {
  id: string;
  sourceSheet: string;
  accountName: string;
  accountGroup: string;
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
