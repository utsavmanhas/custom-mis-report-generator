import type { AccountCategory, AllocationBase, BusinessType, ProfitCenterKind, TrialBalanceRow } from "../types";

const categoryRules: Array<{
  category: AccountCategory;
  misGroup: string;
  confidence: number;
  patterns: RegExp[];
}> = [
  {
    category: "current-liability",
    misGroup: "PF / ESIC Payable",
    confidence: 0.97,
    patterns: [/pf\s*\/?\s*esic.*payable/, /esi.*payable/, /esic.*payable/],
  },
  {
    category: "current-liability",
    misGroup: "GST Payable",
    confidence: 0.97,
    patterns: [/igst payable/, /cgst payable/, /sgst payable/, /gst payable/, /rcm.*payable/],
  },
  {
    category: "current-liability",
    misGroup: "TDS Payable",
    confidence: 0.97,
    patterns: [/tds .*a\/c/, /tds .*account/, /tds .*fees/, /tds .*rent/, /tds .*salary/, /^tds\b/],
  },
  {
    category: "current-liability",
    misGroup: "Employee Payables / Provisions",
    confidence: 0.96,
    patterns: [/employee payables?/, /salary payable/, /bonus payable/, /provision.*employee/],
  },
  {
    category: "current-liability",
    misGroup: "Audit Fee Provision",
    confidence: 0.96,
    patterns: [/provision.*audit/, /audit fee payable/],
  },
  {
    category: "current-liability",
    misGroup: "Sundry Creditors - Vendors",
    confidence: 0.94,
    patterns: [/abcom/, /guidepoint/, /microsoft corporation/, /profectus/, /sp manhas/, /yellow learning/, /sundry creditors?/],
  },
  {
    category: "current-asset",
    misGroup: "Employee Advances / Reimbursements",
    confidence: 0.92,
    patterns: [/reimbursement a\/c/, /reimbursement account/],
  },
  {
    category: "fixed-asset",
    misGroup: "Computers & Laptops",
    confidence: 0.94,
    patterns: [/computers? & laptop/, /computers? and laptop/, /\blaptop\b.*asset/],
  },
  {
    category: "fixed-asset",
    misGroup: "Office Equipment",
    confidence: 0.94,
    patterns: [/office equipments?/, /office equipment/],
  },
  {
    category: "current-asset",
    misGroup: "Security Deposits",
    confidence: 0.95,
    patterns: [/security deposit/, /deposits? \(?asset\)?/],
  },
  {
    category: "current-asset",
    misGroup: "Client Receivables / Debtors",
    confidence: 0.96,
    patterns: [/alnylam/, /jntl/, /kenvue/, /replimune/, /vifor/, /xoma/, /sundry debtors?/],
  },
  {
    category: "current-liability",
    misGroup: "Customer Advances / Credit Debtors",
    confidence: 0.96,
    patterns: [/reliant ai/],
  },
  {
    category: "current-asset",
    misGroup: "Fixed Deposits / Treasury",
    confidence: 0.97,
    patterns: [/fd account/, /fixed deposit/, /\bfd\b/],
  },
  {
    category: "current-asset",
    misGroup: "Bank Balance",
    confidence: 0.96,
    patterns: [/icici\s+bank/, /bank account/],
  },
  {
    category: "current-asset",
    misGroup: "Input Tax Credit",
    confidence: 0.97,
    patterns: [/input tax credit/, /\bitc\b/],
  },
  {
    category: "current-asset",
    misGroup: "Advance Tax",
    confidence: 0.97,
    patterns: [/advance tax/],
  },
  {
    category: "current-asset",
    misGroup: "Prepaids / Deferred Assets",
    confidence: 0.9,
    patterns: [/prepaid/, /preliminary expenses$/],
  },
  {
    category: "current-asset",
    misGroup: "Suspense / Clearing",
    confidence: 0.88,
    patterns: [/suspense/],
  },
  {
    category: "other-income",
    misGroup: "Interest on Fixed Deposit",
    confidence: 0.96,
    patterns: [/interest on fd/, /interest.*fixed deposit/],
  },
  {
    category: "revenue",
    misGroup: "Professional Research & Analytics Export Fees",
    confidence: 0.97,
    patterns: [/professional.*research.*analytics.*export.*fee/, /research.*analytics.*export/],
  },
  {
    category: "people-cost",
    misGroup: "Bonus / Variable Compensation",
    confidence: 0.96,
    patterns: [/bonus account/, /\bbonus\b/],
  },
  {
    category: "people-cost",
    misGroup: "Salaries & Wages",
    confidence: 0.96,
    patterns: [/salary account/, /\bsalary\b/, /salaries/],
  },
  {
    category: "people-cost",
    misGroup: "Compensation Related Expenses",
    confidence: 0.95,
    patterns: [/comp\.? related/, /employee benefit/],
  },
  {
    category: "people-cost",
    misGroup: "Employer PF / EDLI / Statutory Payroll Cost",
    confidence: 0.95,
    patterns: [/epf employer/, /edli contribution/, /provident fund/],
  },
  {
    category: "people-cost",
    misGroup: "Staff Welfare",
    confidence: 0.95,
    patterns: [/staff welfare/],
  },
  {
    category: "operating-expense",
    misGroup: "Accounting, Audit & Legal",
    confidence: 0.94,
    patterns: [/accounting fee/, /audit fees?$/, /legal fee/],
  },
  {
    category: "operating-expense",
    misGroup: "Professional & Consultancy",
    confidence: 0.93,
    patterns: [/consultancy fees?/, /professional fees?$/],
  },
  {
    category: "operating-expense",
    misGroup: "Office & Co-working Infrastructure",
    confidence: 0.94,
    patterns: [/office rent/, /co-?working/, /office expenses/, /admin\s*\/?\s*insp/],
  },
  {
    category: "operating-expense",
    misGroup: "Travel, Lodging & Conveyance",
    confidence: 0.94,
    patterns: [/air travelling/, /travelling exp/, /travel/, /lodging/, /boarding/, /car hire/, /conveyance/, /meals/],
  },
  {
    category: "operating-expense",
    misGroup: "Business Development & Team",
    confidence: 0.93,
    patterns: [/business development/, /business expenses.*team lunch/, /team lunch/, /team offsite/, /training and movement/, /training & workshop/],
  },
  {
    category: "operating-expense",
    misGroup: "Technology, IT & Subscriptions",
    confidence: 0.94,
    patterns: [/software subscription/, /fees? & subscription/, /fees? and subscription/, /domain fee/, /website.*software/, /it repair/, /computer consumables/, /identification protection/],
  },
  {
    category: "operating-expense",
    misGroup: "Laptop Rentals & Insurance",
    confidence: 0.94,
    patterns: [/laptop rentals?/, /laptops insurance/],
  },
  {
    category: "operating-expense",
    misGroup: "Insurance",
    confidence: 0.92,
    patterns: [/insurance expenses?$/],
  },
  {
    category: "operating-expense",
    misGroup: "Repairs & Maintenance",
    confidence: 0.92,
    patterns: [/repair & maintenance/, /repairs? and maintenance/],
  },
  {
    category: "operating-expense",
    misGroup: "Printing & Stationery",
    confidence: 0.92,
    patterns: [/printing/, /stationery/],
  },
  {
    category: "operating-expense",
    misGroup: "Rates, Taxes & RCM Expenses",
    confidence: 0.9,
    patterns: [/rate and taxes/, /subscription - rcm/],
  },
  {
    category: "operating-expense",
    misGroup: "Depreciation & Amortisation",
    confidence: 0.94,
    patterns: [/depreciation/, /preliminary expenses written off/],
  },
  {
    category: "finance-cost",
    misGroup: "Bank Charges & Forex",
    confidence: 0.94,
    patterns: [/bank charges?/, /currency rate fluctuation/],
  },
  {
    category: "finance-cost",
    misGroup: "Interest on Statutory Dues",
    confidence: 0.94,
    patterns: [/interest paid on gst/, /interest paid on tds/],
  },
  {
    category: "other-income",
    misGroup: "Rounding / Short & Excess",
    confidence: 0.86,
    patterns: [/short & excess/],
  },
  {
    category: "revenue",
    misGroup: "Operating Revenue",
    confidence: 0.9,
    patterns: [/sales?/, /revenue/, /turnover/, /subscription (?:revenue|income)/, /subscription fees? received/, /tuition/, /fees? received/, /retainer/, /consulting fees?/, /export fee/, /analytics.*fee/, /research.*fee/, /income from operations/],
  },
  {
    category: "other-income",
    misGroup: "Other Income",
    confidence: 0.82,
    patterns: [/interest income/, /interest on fd/, /dividend/, /other income/, /gain on/, /misc\.? income/],
  },
  {
    category: "direct-cost",
    misGroup: "Direct Costs",
    confidence: 0.84,
    patterns: [/cogs/, /cost of goods/, /purchase/, /raw material/, /subcontract/, /contractor/, /consultancy fees?/, /project expense/, /faculty honorarium/, /cloud hosting/, /production cost/, /freight inward/],
  },
  {
    category: "people-cost",
    misGroup: "People Costs",
    confidence: 0.88,
    patterns: [/salary/, /salaries/, /wages/, /payroll/, /stipend/, /bonus/, /gratuity/, /provident fund/, /\bpf\b/, /epf/, /\besi\b/, /edli/, /employee benefit/, /comp\.? related/, /staff welfare/],
  },
  {
    category: "operating-expense",
    misGroup: "Operating Expenses",
    confidence: 0.8,
    patterns: [/\brent\b/, /office expense/, /marketing/, /advertising/, /business development/, /business expense/, /team lunch/, /travel/, /travelling/, /conveyance/, /car hire/, /lodging/, /boarding/, /meals/, /relocation/, /training/, /workshop/, /utility/, /electricity/, /software/, /domain fee/, /fees? & subscription/, /fees? and subscription/, /identification protection/, /co-?working/, /laptop rentals?/, /computer consumables/, /legal/, /professional fee/, /accounting fee/, /audit/, /depreciation/, /repair/, /maintenance/, /insurance/, /telephone/, /internet/, /printing/, /stationery/, /admin/, /preliminary expenses written off/, /rate and taxes/, /short & excess/, /team offsite/],
  },
  {
    category: "finance-cost",
    misGroup: "Finance Costs",
    confidence: 0.86,
    patterns: [/interest expense/, /bank charges?/, /finance cost/, /loan processing/, /credit card charges?/, /currency rate fluctuation/],
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
    patterns: [/bank/, /cash/, /debtor/, /receivable/, /inventory/, /stock/, /advance to/, /advance tax/, /input tax credit/, /deposit/, /prepaid/],
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
