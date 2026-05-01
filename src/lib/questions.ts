import { categorySignedAmount } from "./classification";
import { getBusinessTemplate } from "./businessTemplates";
import type { BusinessProfile, GeneratedQuestion, ProfitCenter, QuestionAnswer, StaffMember, TrialBalanceRow } from "../types";

function answered(answers: QuestionAnswer[], id: string) {
  return answers.some((answer) => answer.id === id && answer.status !== "open" && answer.answer.trim());
}

function pushIfOpen(list: GeneratedQuestion[], answers: QuestionAnswer[], item: GeneratedQuestion) {
  if (!answered(answers, item.id)) list.push(item);
}

export function buildQuestions(
  profile: BusinessProfile,
  rows: TrialBalanceRow[],
  centers: ProfitCenter[],
  staff: StaffMember[],
  answers: QuestionAnswer[],
): GeneratedQuestion[] {
  const questions: GeneratedQuestion[] = [];
  const template = getBusinessTemplate(profile.businessType);
  const pnlRows = rows.filter((row) => ["revenue", "direct-cost", "people-cost", "operating-expense"].includes(row.category));
  const unknownRows = rows.filter((row) => row.category === "unknown");
  const unassignedPnl = pnlRows.filter((row) => !row.profitCenterId);
  const peopleCost = rows
    .filter((row) => row.category === "people-cost")
    .reduce((sum, row) => sum + Math.abs(categorySignedAmount(row)), 0);
  const revenue = rows.filter((row) => row.category === "revenue").reduce((sum, row) => sum + Math.abs(categorySignedAmount(row)), 0);

  if (!rows.length) {
    pushIfOpen(questions, answers, {
      id: "upload-trial-balance",
      section: "Accounting source",
      prompt: "Upload the trial balance, Tally export, Zoho Books export, or ledger summary for this reporting period.",
      reason: "The MIS workbook needs a ledger-level base before it can allocate revenue, costs, assets, and liabilities.",
      priority: "high",
    });
  }

  if (!profile.website.trim() && !profile.publicNotes.trim()) {
    pushIfOpen(questions, answers, {
      id: "business-public-profile",
      section: "External intelligence",
      prompt: "Add the official website, marketplace profile, prospectus, public deck, or source URLs that describe the business model.",
      reason: "Public context helps the MIS model infer product lines, departments, locations, and operating drivers.",
      priority: "medium",
    });
  }

  if (!centers.length) {
    pushIfOpen(questions, answers, {
      id: "define-profit-centers",
      section: "MIS structure",
      prompt: `Define the ${template.unitPlural} that should appear as separate MIS columns.`,
      reason: "The generated P&L allocates revenue and cost against these centers.",
      priority: "high",
    });
  }

  if (centers.length === 1 && centers[0].name.toLowerCase().includes(template.defaultUnitName.toLowerCase().split(" ")[0])) {
    pushIfOpen(questions, answers, {
      id: `${profile.businessType}-split-units`,
      section: "Granularity",
      prompt: `Confirm whether this business should be split below one ${template.unitLabel.toLowerCase()} into ${template.subUnitLabel.toLowerCase()} level reporting.`,
      reason: "The app should choose the lowest useful operating unit instead of stopping at a generic P&L.",
      priority: "medium",
    });
  }

  if (centers.some((center) => !(center.segment || "").trim())) {
    pushIfOpen(questions, answers, {
      id: `${profile.businessType}-segments`,
      section: "Operating structure",
      prompt: `Map every ${template.unitLabel.toLowerCase()} to its ${template.segmentLabel.toLowerCase()} and owner.`,
      reason: "Segment and owner mapping lets the MIS roll up from unit to management accountability.",
      priority: "medium",
    });
  }

  if (centers.some((center) => center.studentCount <= 0 && center.manualRevenue <= 0)) {
    pushIfOpen(questions, answers, {
      id: `${profile.businessType}-primary-driver`,
      section: "Driver build-up",
      prompt: `Enter ${template.metricLabels.primary.toLowerCase()} and ${template.metricLabels.averageRate.toLowerCase()} for each ${template.unitLabel.toLowerCase()}.`,
      reason: "Revenue should be formula-driven from operating volume and rate wherever the trial balance does not expose the split.",
      priority: "high",
    });
  }

  if (centers.some((center) => (center.variableCostRate || 0) <= 0 && center.manualDirectCost <= 0)) {
    pushIfOpen(questions, answers, {
      id: `${profile.businessType}-variable-cost-driver`,
      section: "Cost build-up",
      prompt: `Enter ${template.metricLabels.variableCostRate.toLowerCase()} or direct costs for each ${template.unitLabel.toLowerCase()}.`,
      reason: "Direct cost should be traced to the unit before shared overhead is allocated.",
      priority: "medium",
    });
  }

  if (unknownRows.length) {
    pushIfOpen(questions, answers, {
      id: "classify-unknown-ledgers",
      section: "Ledger mapping",
      prompt: `Confirm the category for ${unknownRows.length} unclassified ledger accounts.`,
      reason: "Unknown ledger accounts are placed in a data-gap sheet and excluded from precise profitability until mapped.",
      priority: "high",
    });
  }

  if (unassignedPnl.length && centers.length) {
    pushIfOpen(questions, answers, {
      id: "assign-pnl-ledgers",
      section: "Profit center allocation",
      prompt: `Assign ${unassignedPnl.length} P&L ledger accounts to a direct profit center or keep them as shared overhead.`,
      reason: "Direct ledger assignment reduces arbitrary allocation and makes center profitability more reliable.",
      priority: "high",
    });
  }

  if (peopleCost > 0 && !staff.length) {
    pushIfOpen(questions, answers, {
      id: "staffing-cost-roster",
      section: "People allocation",
      prompt: "Add the staffed people, their monthly cost, and their assignment equivalent across projects, departments, or products.",
      reason: "Salary cost is often the largest MIS allocation driver and needs person-level assignment for granular profitability.",
      priority: "high",
    });
  }

  if (profile.allocationBase === "manual" && centers.some((center) => center.allocationWeight <= 0)) {
    pushIfOpen(questions, answers, {
      id: "manual-allocation-weights",
      section: "Shared cost allocation",
      prompt: "Enter manual allocation weights for every profit center.",
      reason: "Shared operating costs selected for manual allocation need weights before the workbook can calculate formulas.",
      priority: "medium",
    });
  }

  if (!revenue && centers.every((center) => center.manualRevenue <= 0)) {
    pushIfOpen(questions, answers, {
      id: "revenue-driver-data",
      section: "Revenue build-up",
      prompt: "Enter center-level revenue, billing, enrollment, subscription, order, or production driver data.",
      reason: "The trial balance does not expose revenue by center, so the MIS needs a driver-based split.",
      priority: "high",
    });
  }

  template.questionSections.forEach((question) => {
    pushIfOpen(questions, answers, {
      ...question,
      priority: question.priority,
    });
  });

  return questions;
}
