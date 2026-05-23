import type { BusinessType, ProfitCenterKind } from "../types";

export interface BusinessTemplate {
  type: BusinessType;
  unitLabel: string;
  unitPlural: string;
  subUnitLabel: string;
  defaultKind: ProfitCenterKind;
  defaultUnitName: string;
  segmentLabel: string;
  ownerLabel: string;
  revenueDriverLabel: string;
  directCostLabel: string;
  metricLabels: {
    primary: string;
    secondary: string;
    tertiary: string;
    averageRate: string;
    variableCostRate: string;
    utilization: string;
  };
  workbookSheets: string[];
  deepDiveSheet: string;
  questionSections: Array<{
    id: string;
    section: string;
    prompt: string;
    reason: string;
    priority: "high" | "medium" | "low";
  }>;
  scheduleRows: Array<{
    section: string;
    requirement: string;
    purpose: string;
  }>;
}

const templates: Record<BusinessType, BusinessTemplate> = {
  consulting: {
    type: "consulting",
    unitLabel: "Project",
    unitPlural: "Projects / client mandates",
    subUnitLabel: "Milestone / workstream",
    defaultKind: "project",
    defaultUnitName: "Client mandate",
    segmentLabel: "Client / practice",
    ownerLabel: "Engagement owner",
    revenueDriverLabel: "Billing model",
    directCostLabel: "Project direct cost",
    metricLabels: {
      primary: "Billable hours",
      secondary: "Assigned FTE",
      tertiary: "Delivery weeks",
      averageRate: "Blended billing rate",
      variableCostRate: "Delivery cost per hour",
      utilization: "Realization %",
    },
    workbookSheets: ["Project Profitability", "Staff Utilization", "Client Recovery", "WIP & Collections"],
    deepDiveSheet: "Project Profitability",
    questionSections: [
      {
        id: "consulting-scope-billing",
        section: "Project economics",
        prompt: "Capture every project or retainer with client, scope, billing basis, contracted value, milestones, and billing status.",
        reason: "Consulting profitability needs a project-level revenue build-up instead of a broad revenue ledger.",
        priority: "high",
      },
      {
        id: "consulting-staffing-equivalent",
        section: "People allocation",
        prompt: "For each project, capture named staff, assignment equivalent, cost-to-company, expected billable hours, and utilization.",
        reason: "Salary cost must follow actual staffing, not a flat overhead percentage.",
        priority: "high",
      },
      {
        id: "consulting-recovery-costs",
        section: "Recoverability",
        prompt: "Identify pass-through costs, subcontractors, travel, write-offs, WIP, unbilled revenue, and collection risk by project.",
        reason: "Real project margin changes materially when recoverability and working capital are included.",
        priority: "medium",
      },
    ],
    scheduleRows: [
      { section: "Revenue", requirement: "Contract value, billing basis, invoices raised, unbilled WIP", purpose: "Separate earned revenue from cash collection." },
      { section: "People", requirement: "Named staff, FTE, billable hours, CTC, realization", purpose: "Calculate actual delivery cost and utilization." },
      { section: "Direct costs", requirement: "Subcontractor, travel, tools, project-specific expenses", purpose: "Calculate contribution margin before shared overhead." },
      { section: "Working capital", requirement: "Receivables ageing, advances, deferred revenue", purpose: "Expose cash strain by client mandate." },
    ],
  },
  "professional-services": {
    type: "professional-services",
    unitLabel: "Engagement",
    unitPlural: "Engagements / service lines",
    subUnitLabel: "Matter / deliverable",
    defaultKind: "project",
    defaultUnitName: "Client engagement",
    segmentLabel: "Client / service line",
    ownerLabel: "Partner / manager",
    revenueDriverLabel: "Fee arrangement",
    directCostLabel: "Engagement direct cost",
    metricLabels: {
      primary: "Billable hours",
      secondary: "Partner hours",
      tertiary: "Open matters",
      averageRate: "Billing rate",
      variableCostRate: "Delivery cost per hour",
      utilization: "Recovery %",
    },
    workbookSheets: ["Engagement Profitability", "Team Leverage", "WIP & Debtors", "Client Concentration"],
    deepDiveSheet: "Engagement Profitability",
    questionSections: [
      {
        id: "ps-engagement-register",
        section: "Engagement register",
        prompt: "List engagements by client, service line, partner, fee arrangement, billing milestone, and status.",
        reason: "Professional-service MIS needs profitability by matter, not only by office or department.",
        priority: "high",
      },
      {
        id: "ps-leverage",
        section: "Team leverage",
        prompt: "Capture partner, manager, associate, and support hours for each engagement.",
        reason: "Margin depends on leverage mix and recovery of senior time.",
        priority: "high",
      },
      {
        id: "ps-wip-ar",
        section: "Working capital",
        prompt: "Capture WIP, billed receivables, retainer advances, write-offs, and collection ageing by client.",
        reason: "Cash conversion is a core MIS view for professional services.",
        priority: "medium",
      },
    ],
    scheduleRows: [
      { section: "Revenue", requirement: "Fees earned, retainers, milestone billing, write-offs", purpose: "Bridge books revenue to engagement economics." },
      { section: "People", requirement: "Partner/manager/associate hours and cost rates", purpose: "Measure leverage and delivery margin." },
      { section: "Working capital", requirement: "WIP, debtors, advances, ageing", purpose: "Show collection risk by client." },
      { section: "Quality", requirement: "Rework hours, scope creep, discounting", purpose: "Separate profitable revenue from noisy revenue." },
    ],
  },
  university: {
    type: "university",
    unitLabel: "Department / batch",
    unitPlural: "Departments / batches / programs",
    subUnitLabel: "Program / cohort",
    defaultKind: "department",
    defaultUnitName: "Academic department",
    segmentLabel: "Faculty / school",
    ownerLabel: "Dean / HOD",
    revenueDriverLabel: "Fee category",
    directCostLabel: "Academic direct cost",
    metricLabels: {
      primary: "Students",
      secondary: "Teaching staff",
      tertiary: "Non-teaching staff",
      averageRate: "Average fee per student",
      variableCostRate: "Academic cost per student",
      utilization: "Capacity utilization %",
    },
    workbookSheets: ["Faculty Wise Fees", "Corporate Centre", "Training Projects", "Research Project", "Hostel", "Transport"],
    deepDiveSheet: "Department Profitability",
    questionSections: [
      {
        id: "university-department-batches",
        section: "Academic granularity",
        prompt: "For every department, program, or batch, capture enrolled students, fee category, scholarships, faculty load, classroom/lab usage, and placement/program costs.",
        reason: "Education MIS needs cross-sectional profitability by program, department, and batch.",
        priority: "high",
      },
      {
        id: "university-shared-services",
        section: "Shared services",
        prompt: "Split corporate centre, admission, examination, library, hostel, transport, and student welfare costs into direct and shared pools.",
        reason: "Academic profitability changes materially when shared institutional costs are attributed correctly.",
        priority: "medium",
      },
    ],
    scheduleRows: [
      { section: "Revenue", requirement: "Tuition, hostel, transport, exam fees, scholarships, refunds", purpose: "Build revenue by department and ancillary vertical." },
      { section: "People", requirement: "Faculty, visiting faculty, non-teaching staff, cross-assignment", purpose: "Attribute salary cost to departments and batches." },
      { section: "Shared assets", requirement: "Labs, classrooms, library, buses, hostel usage", purpose: "Allocate shared resources on measurable drivers." },
      { section: "Projects", requirement: "Research grants and training project inflow/outflow", purpose: "Separate restricted project profitability." },
    ],
  },
  saas: {
    type: "saas",
    unitLabel: "Product / plan",
    unitPlural: "Products / plans / customer segments",
    subUnitLabel: "Customer cohort",
    defaultKind: "product",
    defaultUnitName: "Core product plan",
    segmentLabel: "Segment / geography",
    ownerLabel: "Product owner",
    revenueDriverLabel: "Subscription model",
    directCostLabel: "Delivery cost",
    metricLabels: {
      primary: "Active customers",
      secondary: "Support tickets",
      tertiary: "Product FTE",
      averageRate: "ARPA / MRR per customer",
      variableCostRate: "Hosting cost per customer",
      utilization: "NRR / retention %",
    },
    workbookSheets: ["Product ARR", "Cohort Retention", "Hosting & Support", "CAC Payback"],
    deepDiveSheet: "SaaS Unit Economics",
    questionSections: [
      {
        id: "saas-arr-cohorts",
        section: "Subscription revenue",
        prompt: "Capture ARR/MRR by product, plan, customer segment, new MRR, expansion, contraction, churn, and deferred revenue.",
        reason: "SaaS MIS needs revenue quality and retention, not only accounting revenue.",
        priority: "high",
      },
      {
        id: "saas-delivery-cost",
        section: "Delivery cost",
        prompt: "Capture hosting, third-party API, support tickets, success load, implementation cost, and product-engineering FTE by product or plan.",
        reason: "Gross margin needs product-level delivery cost attribution.",
        priority: "high",
      },
      {
        id: "saas-growth-efficiency",
        section: "Growth efficiency",
        prompt: "Capture CAC by channel, sales cycle, payback period, activation, churn reason, and renewal risk.",
        reason: "Management MIS should connect P&L to growth quality.",
        priority: "medium",
      },
    ],
    scheduleRows: [
      { section: "ARR bridge", requirement: "Opening ARR, new, expansion, contraction, churn, closing ARR", purpose: "Explain recurring revenue movement." },
      { section: "Cohorts", requirement: "Customers by acquisition month, churn, NRR, GRR", purpose: "Measure retention quality by cohort." },
      { section: "Delivery", requirement: "Hosting, support, success, implementation effort", purpose: "Calculate product gross margin." },
      { section: "Growth", requirement: "CAC, channel spend, pipeline, payback", purpose: "Evaluate growth efficiency." },
    ],
  },
  d2c: {
    type: "d2c",
    unitLabel: "SKU / channel",
    unitPlural: "SKUs / channels / categories",
    subUnitLabel: "Variant / campaign",
    defaultKind: "product",
    defaultUnitName: "SKU or sales channel",
    segmentLabel: "Category / marketplace",
    ownerLabel: "Category owner",
    revenueDriverLabel: "Sales channel",
    directCostLabel: "COGS + fulfilment",
    metricLabels: {
      primary: "Units sold",
      secondary: "Orders",
      tertiary: "Returns",
      averageRate: "Average selling price",
      variableCostRate: "COGS + fulfilment per unit",
      utilization: "Return %",
    },
    workbookSheets: ["SKU Contribution", "Channel Margin", "Inventory Turns", "Marketplace Fees"],
    deepDiveSheet: "SKU Channel Contribution",
    questionSections: [
      {
        id: "d2c-sku-channel",
        section: "SKU and channel mix",
        prompt: "Capture revenue, units, discounts, returns, marketplace/channel fees, and GST/tax treatment by SKU and channel.",
        reason: "D2C margin changes sharply by SKU-channel combination.",
        priority: "high",
      },
      {
        id: "d2c-logistics-inventory",
        section: "Fulfilment and inventory",
        prompt: "Capture COGS, packaging, shipping, RTO, warehousing, inventory ageing, damaged stock, and stock-outs by SKU.",
        reason: "Contribution margin needs logistics and inventory leakage.",
        priority: "high",
      },
      {
        id: "d2c-marketing",
        section: "Marketing efficiency",
        prompt: "Capture ad spend, CAC, ROAS, discounts, influencer costs, repeat purchase, and cohort retention by channel.",
        reason: "Marketing spend should be matched to SKU/channel contribution.",
        priority: "medium",
      },
    ],
    scheduleRows: [
      { section: "Revenue", requirement: "Gross sales, discounts, returns, net sales by SKU/channel", purpose: "Move from GMV to net revenue." },
      { section: "Contribution", requirement: "COGS, packaging, freight, gateway, marketplace fees", purpose: "Calculate unit contribution." },
      { section: "Inventory", requirement: "Opening stock, purchases, closing stock, ageing, damage", purpose: "Track cash tied up and margin leakage." },
      { section: "Marketing", requirement: "Ad spend, CAC, ROAS, repeat purchase", purpose: "Connect growth spend to contribution." },
    ],
  },
  manufacturing: {
    type: "manufacturing",
    unitLabel: "Product line / plant",
    unitPlural: "Product lines / plants / jobs",
    subUnitLabel: "SKU / batch / work order",
    defaultKind: "vertical",
    defaultUnitName: "Product line",
    segmentLabel: "Plant / customer segment",
    ownerLabel: "Plant or line owner",
    revenueDriverLabel: "Production / sales model",
    directCostLabel: "BOM + conversion cost",
    metricLabels: {
      primary: "Units produced / sold",
      secondary: "Machine hours",
      tertiary: "Labour hours",
      averageRate: "Selling price per unit",
      variableCostRate: "Material + conversion cost/unit",
      utilization: "Capacity utilization %",
    },
    workbookSheets: ["Product Line Margin", "BOM Variance", "Plant Utilization", "Inventory & WIP"],
    deepDiveSheet: "Product Line Margin",
    questionSections: [
      {
        id: "mfg-product-bom",
        section: "Product economics",
        prompt: "Capture product lines, SKUs, units sold, standard BOM, actual material consumption, conversion cost, scrap, and rework.",
        reason: "Manufacturing MIS needs variance and contribution by product line.",
        priority: "high",
      },
      {
        id: "mfg-plant-capacity",
        section: "Plant utilization",
        prompt: "Capture plant, machine hours, labour hours, shifts, downtime, capacity, batch size, and yield by product line.",
        reason: "Fixed cost absorption depends on utilization and yield.",
        priority: "high",
      },
      {
        id: "mfg-working-capital",
        section: "Inventory and WIP",
        prompt: "Capture raw material, WIP, finished goods, slow-moving inventory, advances, and receivables by product/customer.",
        reason: "Working capital is a major management MIS lens for manufacturing.",
        priority: "medium",
      },
    ],
    scheduleRows: [
      { section: "BOM", requirement: "Standard quantity/rate, actual quantity/rate, variance", purpose: "Track material leakage." },
      { section: "Conversion", requirement: "Machine hours, labour hours, power, repairs, downtime", purpose: "Calculate conversion cost and absorption." },
      { section: "Inventory", requirement: "RM, WIP, FG, ageing, scrap", purpose: "Show cash tied up and obsolescence." },
      { section: "Sales", requirement: "Units, ASP, customer margin, freight", purpose: "Connect plant economics to customer profitability." },
    ],
  },
  custom: {
    type: "custom",
    unitLabel: "Profit center",
    unitPlural: "Profit centers / cost centers",
    subUnitLabel: "Sub-unit",
    defaultKind: "custom",
    defaultUnitName: "Business unit",
    segmentLabel: "Segment",
    ownerLabel: "Owner",
    revenueDriverLabel: "Revenue driver",
    directCostLabel: "Direct cost",
    metricLabels: {
      primary: "Volume driver",
      secondary: "Capacity driver",
      tertiary: "Support driver",
      averageRate: "Revenue per unit",
      variableCostRate: "Variable cost per unit",
      utilization: "Utilization / realization %",
    },
    workbookSheets: ["Unit Economics", "Cost Center Allocation", "Driver Build-Up", "Working Capital"],
    deepDiveSheet: "Unit Economics",
    questionSections: [
      {
        id: "custom-unit-definition",
        section: "Custom structure",
        prompt: "Define the operating units that management actually reviews, plus the measurable revenue, cost, capacity, and support drivers for each.",
        reason: "A custom MIS has to mirror the decision-making structure of the organization.",
        priority: "high",
      },
      {
        id: "custom-driver-basis",
        section: "Driver basis",
        prompt: "Identify which costs are direct, which are shared, and which allocation basis is defensible for each shared cost pool.",
        reason: "The model should avoid arbitrary splits where a measurable driver exists.",
        priority: "high",
      },
    ],
    scheduleRows: [
      { section: "Structure", requirement: "Profit centers, cost centers, sub-units, owners", purpose: "Set the reporting spine." },
      { section: "Drivers", requirement: "Volume, rate, capacity, support load", purpose: "Build formula-driven revenue and cost." },
      { section: "Costs", requirement: "Direct, shared, fixed, variable pools", purpose: "Allocate costs transparently." },
      { section: "Balance sheet", requirement: "Receivables, payables, inventory/advances, assets", purpose: "Connect profitability to cash and capital." },
    ],
  },
};

export function getBusinessTemplate(type: BusinessType) {
  return templates[type] || templates.custom;
}

export function templateForBusinessType(type: BusinessType) {
  return getBusinessTemplate(type);
}
