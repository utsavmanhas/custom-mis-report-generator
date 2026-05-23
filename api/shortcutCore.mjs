import ExcelJS from "exceljs";
import { readFileSync } from "node:fs";
import { join } from "node:path";

const DEFAULT_BASE_URL = "https://api.shortcut.ai/api/spreadsheets";
const DEFAULT_TIMEOUT_MS = 900000;
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function getConfig(env = process.env) {
  const apiKey = env.SHORTCUT_API_KEY?.trim();
  return {
    apiKey,
    baseUrl: (env.SHORTCUT_API_BASE || DEFAULT_BASE_URL).replace(/\/+$/, ""),
    timeoutMs: Number(env.SHORTCUT_TIMEOUT_MS || DEFAULT_TIMEOUT_MS),
    skillId: env.SHORTCUT_SKILL_ID?.trim(),
  };
}

export function isShortcutConfigured(env = process.env) {
  return Boolean(getConfig(env).apiKey);
}

function safeFilename(value, fallback = "mis") {
  const slug = String(value || fallback)
    .replace(/[^a-z0-9]+/gi, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase();
  return slug || fallback;
}

function dateRange(profile = {}) {
  const start = profile.periodStart || "period start not set";
  const end = profile.periodEnd || "period end not set";
  return `${start} to ${end}`;
}

function addRows(sheet, rows) {
  rows.forEach((row) => sheet.addRow(row));
  sheet.getRow(1).font = { bold: true };
  sheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFEFF4F7" },
  };
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.columns.forEach((column) => {
    let max = 12;
    column.eachCell({ includeEmpty: true }, (cell) => {
      max = Math.max(max, String(cell.value ?? "").slice(0, 80).length + 2);
    });
    column.width = Math.min(max, 46);
  });
}

function rowsFromObject(value) {
  return Object.entries(value || {}).map(([key, val]) => [key, Array.isArray(val) || typeof val === "object" ? JSON.stringify(val) : val ?? ""]);
}

function safeJson(value) {
  return JSON.stringify(value ?? null, null, 2);
}

function loadDesignLanguage() {
  return readFileSync(join(process.cwd(), "prompts", "design-language.md"), "utf-8");
}

function pathOf(row = {}) {
  return Array.isArray(row.accountPath) && row.accountPath.length
    ? row.accountPath.join(" > ")
    : [row.accountGroup, row.accountName].filter(Boolean).join(" > ");
}

function rowNet(row = {}) {
  return Number(row.debit || 0) - Number(row.credit || 0);
}

function signedAmount(row = {}) {
  if (["revenue", "other-income", "current-liability", "equity"].includes(row.category)) {
    return Number(row.credit || 0) - Number(row.debit || 0);
  }
  return Number(row.debit || 0) - Number(row.credit || 0);
}

function isBankOrFd(row = {}) {
  const text = `${row.accountName || ""} ${pathOf(row)}`;
  return row.category === "current-asset" && (/\bbank\b|bank accounts|cash/i.test(text) || /\bfd\b|fixed deposit/i.test(row.accountName || ""));
}

function isDebtor(row = {}) {
  return /sundry debtors|debtor|receivable/i.test(pathOf(row));
}

function sectionForLiability(row = {}) {
  const text = `${row.accountName || ""} ${pathOf(row)}`.toLowerCase();
  if (/capital|shareholder|reserve|retained/.test(text)) return "Capital & equity";
  if (/gst|tds|pf|epf|esi|esic|edli|duties|tax/.test(text)) return "GST, TDS, PF/ESIC and statutory payables";
  if (/provision|employee payable|bonus payable|audit fee/.test(text)) return "Provisions and employee payables";
  if (/sundry creditors|creditor/.test(text)) return "Sundry creditors";
  if (/reimbursement/.test(text)) return "Reimbursements payable";
  if (isDebtor(row) && rowNet(row) < 0) return "Customer advances / credit-balance debtors";
  return "Other current liabilities";
}

function sectionForAsset(row = {}) {
  const text = `${row.accountName || ""} ${pathOf(row)}`.toLowerCase();
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

function balanceSheetRows(rows = []) {
  const liabilities = [];
  const assets = [];
  rows.forEach((row) => {
    const net = rowNet(row);
    if (row.category === "equity" || row.category === "current-liability" || (isDebtor(row) && net < 0)) {
      liabilities.push([
        sectionForLiability(row),
        row.accountName,
        Math.abs(net),
        row.debit || 0,
        row.credit || 0,
        pathOf(row),
        isDebtor(row) && net < 0 ? "Credit balance debtor; classify as customer advance/liability." : "",
      ]);
      return;
    }
    if (row.category === "fixed-asset" || row.category === "current-asset") {
      assets.push([sectionForAsset(row), row.accountName, Math.abs(net), row.debit || 0, row.credit || 0, pathOf(row), ""]);
    }
  });
  return { liabilities, assets };
}

function taxRows(rows = []) {
  return rows
    .filter((row) => /\bgst\b|\btds\b|\bitc\b|input tax|advance tax|pf\b|epf|esi|esic|professional tax|income tax|tax payable|duties/i.test(`${row.accountName || ""} ${pathOf(row)}`))
    .map((row) => {
      const net = rowNet(row);
      return [
        net < 0 ? "Payable / statutory liability" : "Recoverable / advance",
        row.accountName,
        row.debit || 0,
        row.credit || 0,
        Math.max(net, 0),
        Math.max(-net, 0),
        pathOf(row),
      ];
    });
}

function chartRows(rows = [], bankSources = []) {
  const output = [];
  rows.filter((row) => ["revenue", "other-income"].includes(row.category)).forEach((row) => output.push(["Revenue mix", row.accountName, row.accountName, Math.abs(signedAmount(row))]));
  rows.filter((row) => ["direct-cost", "people-cost", "operating-expense", "finance-cost", "tax"].includes(row.category)).forEach((row) => output.push(["Expense mix", row.accountName, row.accountName, Math.abs(signedAmount(row))]));
  const bs = balanceSheetRows(rows);
  bs.assets.forEach((row) => output.push(["Asset composition", row[0], row[1], row[2]]));
  bs.liabilities.forEach((row) => output.push(["Liability composition", row[0], row[1], row[2]]));
  const receipts = bankSources.flatMap((source) => source.summary?.monthly || []).reduce((sum, month) => sum + Number(month.receipts || 0), 0);
  const payments = bankSources.flatMap((source) => source.summary?.monthly || []).reduce((sum, month) => sum + Number(month.payments || 0), 0);
  output.push(["Receipts vs payments", "Receipts", "Receipts", receipts]);
  output.push(["Receipts vs payments", "Payments", "Payments", payments]);
  return output;
}

export function buildShortcutPrompt({ profile, analysis }) {
  const sheetPlan = analysis?.sheetPlan?.length
    ? analysis.sheetPlan.map((sheet) => `${sheet.name}: ${sheet.purpose}`).join("\n")
    : "Use the attached source workbook and Design Brief sheet to infer the right MIS sheet roster.";

  return [
    "Create a finished, client-ready MIS Excel workbook from the attached source workbook.",
    "",
    "Use the attached workbook as source data. It contains Profile, Trial Balance, Profit Centers, Staff, User Answers, and Design Brief sheets, plus Report Analysis when available.",
    "",
    "Requirements:",
    "- Build a polished formula-driven Excel workbook, not a prose report.",
    "- Use live Excel formulas wherever possible instead of hard-coded derived totals.",
    "- Include a dashboard/cover, consolidated P&L or Income & Expenditure, Balance Sheet, Tax & Statutory, Fund Flow, next-period projection, ChartData, QC Tie-Outs, drill-down sheets by operating segment, supporting schedules, and a Notes & Adjustments sheet.",
    "- Use Indian number formatting for INR amounts and percentage columns for margin/allocation analysis.",
    "- Preserve source-data traceability. Keep source sheets at the back or clearly label assumptions if you rebuild the workbook structure.",
    "- Document missing data, allocation choices, and assumptions in Notes & Adjustments.",
    "- Do not invent client facts. If data is missing, create clearly marked placeholders or assumptions.",
    "- Never use Sample Business when source metadata provides the company name.",
    "- Do not ask for company name or reporting period when they are already present in the Profile sheet.",
    "- Fund flow must use the Bank Transactions and Fund Flow Source sheets when available; otherwise mark it provisional and use only TB-derived bank/FD closing balances.",
    "- Treat debtor balances as receivables/advances only. Credit-balance debtors belong in customer advances/liabilities, not ordinary debtors or client revenue.",
    "- Client profitability needs invoice-wise/client-wise revenue; debtor balances are visible proxy data only.",
    "- PBT margin and all margin percentages must use the intended denominator consistently: PBT margin = Profit Before Tax / Revenue.",
    "",
    `Business: ${profile?.businessName || "Untitled MIS report"}`,
    `Period: ${dateRange(profile)}`,
    `Currency: ${profile?.currency || "INR"}`,
    "",
    "Preferred sheet plan:",
    sheetPlan,
  ].join("\n");
}

export async function buildShortcutSourceWorkbook(input) {
  const { profile = {}, rows = [], bankSources = [], profitCenters = [], staff = [], answers = [], reportAnswers = [], analysis = null } = input || {};
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Custom MIS Generator";
  workbook.created = new Date();

  addRows(workbook.addWorksheet("Profile"), [
    ["Field", "Value"],
    ...rowsFromObject(profile),
  ]);

  addRows(workbook.addWorksheet("Trial Balance"), [
    ["ID", "Source Sheet", "Source Row", "Account", "Group / Path", "Debit", "Credit", "Balance", "Category", "MIS Group", "Profit Center ID", "Confidence", "Risk Flags"],
    ...rows.map((row) => [
      row.id,
      row.sourceSheet,
      row.sourceRowNumber || "",
      row.accountName,
      pathOf(row),
      row.debit,
      row.credit,
      row.balance,
      row.category,
      row.misGroup,
      row.profitCenterId,
      row.confidence,
      Array.isArray(row.riskFlags) ? row.riskFlags.join(", ") : "",
    ]),
  ]);

  const bs = balanceSheetRows(rows);
  addRows(workbook.addWorksheet("Balance Sheet Source"), [
    ["Side", "Section", "Account", "Amount", "Debit", "Credit", "Path", "Note"],
    ...bs.liabilities.map((row) => ["Liability", ...row]),
    ...bs.assets.map((row) => ["Asset", ...row]),
  ]);

  addRows(workbook.addWorksheet("Tax Statutory Source"), [
    ["Section", "Account", "Debit", "Credit", "Recoverable / advance", "Payable", "Path"],
    ...taxRows(rows),
    ["Tax provision prompt", "Ask for expected income-tax provision, deferred tax, assessments and payment timing if not present.", "", "", "", "", ""],
  ]);

  addRows(workbook.addWorksheet("Bank Transactions"), [
    ["Source Type", "File", "Sheet", "Row", "Date", "Description", "Debit / payment", "Credit / receipt", "Balance"],
    ...bankSources.flatMap((source) =>
      (source.transactions || []).map((transaction) => [
        source.sourceType,
        source.fileName,
        transaction.sourceSheet,
        transaction.sourceRowNumber,
        transaction.date,
        transaction.description,
        transaction.debit,
        transaction.credit,
        transaction.balance,
      ]),
    ),
  ]);

  addRows(workbook.addWorksheet("Fund Flow Source"), [
    ["Field", "Value"],
    ["Fund flow basis", profile.fundFlowBasis || "trial-balance-proxy"],
    ["Projection basis", profile.projectionBasis || "past-year"],
    ["Bank source files", bankSources.length],
    ["TB-derived bank and FD closing", rows.filter(isBankOrFd).reduce((sum, row) => sum + Math.max(rowNet(row), 0), 0)],
    [],
    ["Month", "Receipts", "Payments", "Opening Balance", "Closing Balance", "Transaction Count", "Source"],
    ...bankSources.flatMap((source) =>
      (source.summary?.monthly || []).map((month) => [
        month.month,
        month.receipts,
        month.payments,
        month.openingBalance,
        month.closingBalance,
        month.transactionCount,
        source.fileName,
      ]),
    ),
    [],
    ["Projection instruction", "Create the next-year / next-period projection from historical monthly receipt/payment run-rate, then adjust using user assumptions for bonus, tax, capex, loans, advances and collections."],
  ]);

  addRows(workbook.addWorksheet("ChartData"), [
    ["Chart", "Label", "Account", "Value"],
    ...chartRows(rows, bankSources),
  ]);

  addRows(workbook.addWorksheet("QC Tie-Outs"), [
    ["Check", "Value", "Expected / note"],
    ["Source TB debit", profile.tbTotalDebit || rows.reduce((sum, row) => sum + Number(row.debit || 0), 0), "Should equal source TB credit"],
    ["Source TB credit", profile.tbTotalCredit || rows.reduce((sum, row) => sum + Number(row.credit || 0), 0), "Should equal source TB debit"],
    ["TB difference", (profile.tbTotalDebit || rows.reduce((sum, row) => sum + Number(row.debit || 0), 0)) - (profile.tbTotalCredit || rows.reduce((sum, row) => sum + Number(row.credit || 0), 0)), "Should be zero"],
    ["Imported ledger debit", rows.reduce((sum, row) => sum + Number(row.debit || 0), 0), "Ledger-row sum after excluding group subtotals"],
    ["Imported ledger credit", rows.reduce((sum, row) => sum + Number(row.credit || 0), 0), "Ledger-row sum after excluding group subtotals"],
    ["Operating revenue", rows.filter((row) => row.category === "revenue").reduce((sum, row) => sum + Math.abs(signedAmount(row)), 0), "Revenue ledgers only"],
    ["Other income", rows.filter((row) => row.category === "other-income").reduce((sum, row) => sum + Math.abs(signedAmount(row)), 0), "Separate from operating revenue"],
    ["Bank + FD", rows.filter(isBankOrFd).reduce((sum, row) => sum + Math.max(rowNet(row), 0), 0), "TB-derived closing balance"],
    ["Credit-balance debtors", bs.liabilities.filter((row) => row[0] === "Customer advances / credit-balance debtors").reduce((sum, row) => sum + Number(row[2] || 0), 0), "Show as customer advances/liability"],
    [],
    ["Risk Flag", "Account", "Path"],
    ...rows.filter((row) => Array.isArray(row.riskFlags) && row.riskFlags.length).map((row) => [row.riskFlags.join(", "), row.accountName, pathOf(row)]),
  ]);

  addRows(workbook.addWorksheet("Profit Centers"), [
    ["ID", "Name", "Kind", "Segment", "Owner", "Revenue Driver", "Manual Revenue", "Manual Direct Cost", "Prior Revenue", "Prior Direct Cost", "Primary Driver", "Secondary Driver", "Tertiary Driver", "Average Rate", "Variable Cost Rate", "Utilization %", "Allocation Weight", "Notes"],
    ...profitCenters.map((center) => [
      center.id,
      center.name,
      center.kind,
      center.segment,
      center.owner,
      center.revenueDriver,
      center.manualRevenue,
      center.manualDirectCost,
      center.priorRevenue,
      center.priorDirectCost,
      center.studentCount,
      center.teachingStaffCount,
      center.nonTeachingStaffCount,
      center.averageRevenueRate,
      center.variableCostRate,
      center.utilizationPercent,
      center.allocationWeight,
      center.notes,
    ]),
  ]);

  addRows(workbook.addWorksheet("Staff"), [
    ["ID", "Name", "Role", "Department", "Monthly Cost", "Assignments JSON"],
    ...staff.map((person) => [
      person.id,
      person.name,
      person.role,
      person.department,
      person.monthlyCost,
      safeJson(person.assignments),
    ]),
  ]);

  addRows(workbook.addWorksheet("User Answers"), [
    ["ID", "Question", "Answer", "Status"],
    ...answers.map((answer) => [answer.id, answer.question, answer.answer, answer.status]),
    ...reportAnswers.map((answer) => [answer.id, answer.question, answer.answer, "report-question-answer"]),
  ]);

  if (analysis) {
    addRows(workbook.addWorksheet("Report Analysis"), [
      ["Section", "Detail"],
      ["Business Type", analysis.businessType],
      ["Business Type Reasoning", analysis.businessTypeReasoning],
      ["Detected Segments", safeJson(analysis.detectedSegments)],
      ["Allocation Options", safeJson(analysis.allocationLogicOptions)],
      ["Sheet Plan", safeJson(analysis.sheetPlan)],
      ["Clarifying Questions", safeJson(analysis.clarifyingQuestions)],
      ["Unclassified Accounts", safeJson(analysis.unclassifiedAccounts)],
    ]);
  }

  const designLanguage = loadDesignLanguage();
  addRows(workbook.addWorksheet("Design Brief"), [
    ["Field", "Detail"],
    ["Workbook Build Prompt", buildShortcutPrompt({ profile, analysis })],
    ["MIS Design Language", designLanguage],
    ["Data Handling", "This workbook may contain client financials, ledgers, salary or pod costs, receivables, and assumptions. Use only with client/firm approval and appropriate provider settings."],
  ]);

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
}

async function shortcutJson(path, init = {}, env = process.env) {
  const config = getConfig(env);
  if (!config.apiKey) throw new Error("Workbook build credentials missing");
  const response = await fetch(`${config.baseUrl}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${config.apiKey}`,
      ...(init.headers || {}),
    },
  });
  const text = await response.text();
  let data = {};
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = { message: text };
  }
  if (!response.ok) {
    throw new Error(data.message || data.error || `Workbook build HTTP ${response.status}`);
  }
  return data;
}

async function uploadWithField(buffer, filename, fieldName, env) {
  const config = getConfig(env);
  const form = new FormData();
  form.append(fieldName, new Blob([buffer], { type: XLSX_MIME }), filename);
  const response = await fetch(`${config.baseUrl}/upload`, {
    method: "POST",
    headers: { Authorization: `Bearer ${config.apiKey}` },
    body: form,
  });
  const text = await response.text();
  let data = {};
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = { message: text };
  }
  if (!response.ok) {
    throw new Error(data.message || data.error || `Workbook upload HTTP ${response.status}`);
  }
  if (!data.fileId) throw new Error("Workbook upload did not return fileId");
  return data.fileId;
}

export async function verifyShortcut(env = process.env) {
  if (!isShortcutConfigured(env)) {
    return { ok: true, configured: false };
  }
  const data = await shortcutJson("/verify", { method: "GET" }, env);
  return { ok: true, configured: true, valid: data.success === true, message: data.message || "Workbook build credentials verified" };
}

export async function startShortcutJob(input, env = process.env) {
  const config = getConfig(env);
  if (!config.apiKey) throw new Error("Workbook build credentials missing");
  if (!Array.isArray(input?.rows) || !input.rows.length) throw new Error("trialBalance required");

  const filename = `${safeFilename(input.profile?.businessName)}-mis-source.xlsx`;
  const outputFilename = `${safeFilename(input.profile?.businessName)}-mis-report.xlsx`;
  const buffer = await buildShortcutSourceWorkbook(input);

  let fileId;
  try {
    fileId = await uploadWithField(buffer, filename, "file", env);
  } catch (firstError) {
    try {
      fileId = await uploadWithField(buffer, filename, "files", env);
    } catch {
      throw firstError;
    }
  }

  const payload = {
    prompt: buildShortcutPrompt({ profile: input.profile, analysis: input.analysis }),
    initFile: fileId,
    mode: "action",
    timeoutMs: Math.min(Math.max(config.timeoutMs || DEFAULT_TIMEOUT_MS, 1), 1800000),
    ...(config.skillId ? { skills: [config.skillId] } : {}),
  };

  const data = await shortcutJson("", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  }, env);

  return {
    ok: true,
    runId: data.runId,
    status: data.status,
    filename: outputFilename,
    sourceFileId: fileId,
  };
}

export async function getShortcutStatus(runId, env = process.env) {
  if (!runId) throw new Error("runId required");
  return shortcutJson(`/${encodeURIComponent(runId)}`, { method: "GET" }, env);
}

export async function downloadShortcutResult(runId, env = process.env) {
  if (!runId) throw new Error("runId required");
  const config = getConfig(env);
  if (!config.apiKey) throw new Error("Workbook build credentials missing");
  const response = await fetch(`${config.baseUrl}/${encodeURIComponent(runId)}/download`, {
    headers: { Authorization: `Bearer ${config.apiKey}` },
  });
  if (!response.ok) {
    let message = `Workbook download HTTP ${response.status}`;
    try {
      const data = await response.json();
      message = data.message || data.error || message;
    } catch {
      // Keep status fallback.
    }
    throw new Error(message);
  }
  return Buffer.from(await response.arrayBuffer());
}

export { XLSX_MIME };
