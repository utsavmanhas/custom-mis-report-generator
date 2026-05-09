import type { AccountCategory, AllocationBase, BusinessType, ProfitCenterKind, TrialBalanceRow } from "../types";

const categoryRules: Array<{
  category: AccountCategory;
  misGroup: string;
  confidence: number;
  patterns: RegExp[];
}> = [
  {
    category: "revenue",
    misGroup: "Operating Revenue",
    confidence: 0.9,
    patterns: [/sales?/, /revenue/, /turnover/, /subscription/, /tuition/, /fees? received/, /retainer/, /consulting fees?/, /income from operations/],
  },
  {
    category: "other-income",
    misGroup: "Other Income",
    confidence: 0.82,
    patterns: [/interest income/, /dividend/, /other income/, /gain on/, /misc\.? income/],
  },
  {
    category: "direct-cost",
    misGroup: "Direct Costs",
    confidence: 0.84,
    patterns: [/cogs/, /cost of goods/, /purchase/, /raw material/, /subcontract/, /contractor/, /project expense/, /faculty honorarium/, /cloud hosting/, /production cost/, /freight inward/],
  },
  {
    category: "people-cost",
    misGroup: "People Costs",
    confidence: 0.88,
    patterns: [/salary/, /salaries/, /wages/, /payroll/, /stipend/, /bonus/, /gratuity/, /provident fund/, /\bpf\b/, /\besi\b/, /employee benefit/, /staff welfare/],
  },
  {
    category: "operating-expense",
    misGroup: "Operating Expenses",
    confidence: 0.8,
    patterns: [/rent/, /office expense/, /marketing/, /advertising/, /travel/, /conveyance/, /utility/, /electricity/, /software/, /legal/, /professional fee/, /audit/, /repair/, /maintenance/, /insurance/, /telephone/, /internet/, /printing/, /stationery/, /admin/],
  },
  {
    category: "finance-cost",
    misGroup: "Finance Costs",
    confidence: 0.86,
    patterns: [/interest expense/, /bank charges?/, /finance cost/, /loan processing/, /credit card charges?/],
  },
  {
    category: "tax",
    misGroup: "Taxes",
    confidence: 0.84,
    patterns: [/\bgst\b/, /\btax\b/, /\btds\b/, /income tax/, /professional tax/, /cess/],
  },
  {
    category: "fixed-asset",
    misGroup: "Fixed Assets",
    confidence: 0.8,
    patterns: [/fixed asset/, /plant/, /machinery/, /computer/, /furniture/, /vehicle/, /building/, /equipment/],
  },
  {
    category: "current-asset",
    misGroup: "Current Assets",
    confidence: 0.76,
    patterns: [/bank/, /cash/, /debtor/, /receivable/, /inventory/, /stock/, /advance to/, /deposit/, /prepaid/],
  },
  {
    category: "current-liability",
    misGroup: "Current Liabilities",
    confidence: 0.76,
    patterns: [/creditor/, /payable/, /advance from/, /loan/, /overdraft/, /provision/, /accrued/, /duties payable/],
  },
  {
    category: "equity",
    misGroup: "Equity",
    confidence: 0.82,
    patterns: [/capital/, /shareholder/, /partner capital/, /reserve/, /retained earnings/, /drawings?/],
  },
];

export const businessTypeLabels: Record<BusinessType, string> = {
  consulting: "Consulting firm",
  "professional-services": "Professional services",
  university: "University / education",
  saas: "SaaS company",
  d2c: "D2C brand",
  manufacturing: "Manufacturing",
  custom: "Custom organization",
};

export const categoryLabels: Record<AccountCategory, string> = {
  revenue: "Revenue",
  "other-income": "Other income",
  "direct-cost": "Direct cost",
  "people-cost": "People cost",
  "operating-expense": "Operating expense",
  "finance-cost": "Finance cost",
  tax: "Tax",
  "fixed-asset": "Fixed asset",
  "current-asset": "Current asset",
  "current-liability": "Current liability",
  equity: "Equity",
  unknown: "Unknown",
};

export const allocationLabels: Record<AllocationBase, string> = {
  revenue: "Revenue share",
  headcount: "Headcount",
  fte: "Assigned FTE",
  equal: "Equal split",
  manual: "Manual weights",
};

export function classifyAccount(accountName: string, accountGroup = "") {
  const haystack = `${accountName} ${accountGroup}`.toLowerCase();
  const hit = categoryRules.find((rule) => rule.patterns.some((pattern) => pattern.test(haystack)));

  if (!hit) {
    return {
      category: "unknown" as AccountCategory,
      misGroup: "Unclassified",
      confidence: 0.35,
    };
  }

  return {
    category: hit.category,
    misGroup: hit.misGroup,
    confidence: hit.confidence,
  };
}

export function defaultProfitCenterKind(type: BusinessType): ProfitCenterKind {
  if (type === "consulting" || type === "professional-services") return "project";
  if (type === "university") return "department";
  if (type === "saas" || type === "d2c") return "product";
  if (type === "manufacturing") return "vertical";
  return "custom";
}

export function profitCenterLabel(type: BusinessType) {
  if (type === "consulting" || type === "professional-services") return "Project";
  if (type === "university") return "Department / batch";
  if (type === "saas") return "Product / plan";
  if (type === "d2c") return "SKU / channel";
  if (type === "manufacturing") return "Plant / product line";
  return "Profit center";
}

export function categorySignedAmount(row: TrialBalanceRow) {
  if (row.category === "revenue" || row.category === "other-income" || row.category === "current-liability" || row.category === "equity") {
    return row.credit - row.debit;
  }

  return row.debit - row.credit;
}

export function isPnLCategory(category: AccountCategory) {
  return ["revenue", "other-income", "direct-cost", "people-cost", "operating-expense", "finance-cost", "tax"].includes(category);
}
