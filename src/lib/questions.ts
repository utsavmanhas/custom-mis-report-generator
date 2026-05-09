import { categorySignedAmount } from "./classification";
import { getBusinessTemplate } from "./businessTemplates";
import type { BankTransaction, BusinessProfile, GeneratedQuestion, ProfitCenter, QuestionAnswer, StaffMember, TrialBalanceRow } from "../types";

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
  bankTransactions: BankTransaction[] = [],
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
  const bankCashRows = rows.filter((row) => row.category === "current-asset" && /bank|cash|fd|fixed deposit/i.test(`${row.accountName} ${row.accountGroup}`));
  const taxRows = rows.filter((row) => /gst|tds|pf|esi|esic|advance tax|input tax|itc|tax/i.test(`${row.accountName} ${row.accountGroup}`));
  const capexRows = rows.filter((row) => row.category === "fixed-asset");
  const bonusRows = rows.filter((row) => /bonus/i.test(row.accountName));
  const hasBankStatement = bankTransactions.some((txn) => txn.sourceType === "bank-statement");
  const hasBankLedger = bankTransactions.some((txn) => txn.sourceType === "bank-ledger");
  const hasBankBalances = bankTransactions.some((txn) => txn.balance !== 0);
  const hasUnclassifiedBankRows = bankTransactions.some((txn) => /unclassified/i.test(`${txn.category} ${txn.fundFlowGroup}`));
  const hasBankFlowGroup = (pattern: RegExp) => bankTransactions.some((txn) => pattern.test(`${txn.category} ${txn.fundFlowGroup} ${txn.narration}`));

  if (!rows.length) {
    pushIfOpen(questions, answers, {
      id: "upload-trial-balance",
      section: "Accounting source",
      prompt: "Upload the trial balance, Tally export, Zoho Books export, or ledger summary for this reporting period.",
      reason: "The MIS workbook needs a ledger-level base before it can allocate revenue, costs, assets, and liabilities.",
      priority: "high",
    });
  }

  if (rows.length && !bankTransactions.length) {
    pushIfOpen(questions, answers, {
      id: "upload-bank-statements-ledgers",
      section: "Fund flow",
      prompt: "Upload bank statements and bank ledgers for every bank, cash, and FD account found in the trial balance.",
      reason: "The TB already provides the MIS period and closing bank/FD balances; the fund-flow statement still needs actual cash movement and book-side ledger detail.",
      priority: "high",
    });
  }

  if (bankTransactions.length) {
    if (!hasBankStatement) {
      pushIfOpen(questions, answers, {
        id: "upload-bank-statements",
        section: "Fund flow",
        prompt: "Upload bank statements for the bank/FD accounts already identified from the trial balance.",
        reason: "Bank statements are the cash-movement source for the actual fund flow; the bank ledger alone remains book-side support.",
        priority: "high",
      });
    }
    if (!hasBankLedger) {
      pushIfOpen(questions, answers, {
        id: "upload-bank-ledgers",
        section: "Fund flow",
        prompt: "Upload bank ledgers for the bank/FD accounts already identified from the trial balance.",
        reason: "Bank ledgers help classify bank-statement narrations and reconcile book movement to bank movement.",
        priority: "medium",
      });
    }
    if (hasUnclassifiedBankRows) {
      pushIfOpen(questions, answers, {
        id: "review-unclassified-bank-tags",
        section: "Fund flow",
        prompt: "Review only the unclassified bank movements and tell us whether they are client receipts, opex, salary/bonus, tax, FD, capex, loan, advance, or internal transfer.",
        reason: "Most transaction tags are inferred automatically; only unclear bank narrations need user input.",
        priority: "high",
      });
    }
  }

  if (bankCashRows.length && bankTransactions.length && !hasBankBalances) {
    pushIfOpen(questions, answers, {
      id: "fund-flow-opening-balances",
      section: "Fund flow",
      prompt: "Provide opening bank, cash, and FD balances only because the uploaded bank files do not expose running balances.",
      reason: "The TB gives closing bank/FD balances; opening balances are needed only when the bank source lacks opening/running balances.",
      priority: "high",
    });
  }

  if (rows.length) {
    pushIfOpen(questions, answers, {
      id: "fund-flow-projection-assumptions",
      section: "Fund flow projection",
      prompt: "Confirm next-period projection assumptions: revenue growth, cost inflation, March bonus, tax payment schedule, capex, loan EMI, FD maturity/renewal, and any known one-offs.",
      reason: "The actual MIS period comes from the TB; only the future fund-flow forecast needs management assumptions.",
      priority: "medium",
    });
  }

  if (taxRows.length && !hasBankFlowGroup(/gst|tds|statutory|tax/i)) {
    pushIfOpen(questions, answers, {
      id: "fund-flow-tax-statutory-timing",
      section: "Fund flow",
      prompt: "Provide GST, TDS, PF/ESIC, advance tax, and ITC payment timing only for items not visible in the uploaded bank files.",
      reason: "Tax payable/recoverable is available from the TB, but cash timing needs bank evidence or a specific tax schedule.",
      priority: "high",
    });
  }

  if (rows.length && ((capexRows.length && !hasBankFlowGroup(/capex|fixed assets/i)) || (bankTransactions.length > 0 && !hasBankFlowGroup(/loan|fd|treasury|owner|capital/i)))) {
    pushIfOpen(questions, answers, {
      id: "fund-flow-capex-fd-loans",
      section: "Fund flow",
      prompt: "Confirm only missing or future capex, FD placements/maturities, loan receipts/repayments, owner capital/drawings, and internal transfers.",
      reason: "The app classifies visible bank movements automatically; this question covers items not present in the uploaded sources and next-period plans.",
      priority: "medium",
    });
  }

  if (bonusRows.length && !hasBankFlowGroup(/salary|bonus|payroll/i)) {
    pushIfOpen(questions, answers, {
      id: "fund-flow-bonus-treatment",
      section: "Fund flow",
      prompt: "Confirm whether the TB bonus is paid in March, accrued unpaid, or should be projected as a future one-off.",
      reason: "Bonus is visible in the TB; only its cash timing or projection treatment needs confirmation if the bank source does not show it.",
      priority: "high",
    });
  }

  if (rows.length) {
    pushIfOpen(questions, answers, {
      id: "fund-flow-manual-adjustments",
      section: "Fund flow projection",
      prompt: "List known future one-off cash events to include in the projection, such as bonus, tax challans, capex, loan repayment, FD maturity, client advance, or reimbursement recovery.",
      reason: "Actual-period data is source-led; forecast-period fund flow still needs known future events that are not in historical files.",
      priority: "low",
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

  void unknownRows;

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
