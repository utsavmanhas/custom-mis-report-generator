import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  AlertTriangle,
  BarChart3,
  Building2,
  Check,
  ClipboardList,
  Download,
  FileSpreadsheet,
  Globe2,
  Landmark,
  Layers3,
  Plus,
  Sparkles,
  Trash2,
  Upload,
  Users,
} from "lucide-react";
import { allocationLabels, businessTypeLabels, categoryLabels } from "./lib/classification";
import { getBusinessTemplate } from "./lib/businessTemplates";
import { parseBankStatementWorkbook, parseTrialBalanceWorkbook } from "./lib/parser";
import { buildQuestions } from "./lib/questions";
import { buildWorkbookIssues, generateMisWorkbook } from "./lib/reporting";
import type {
  AccountCategory,
  AllocationBase,
  BankSourceType,
  BankTransaction,
  BusinessProfile,
  BusinessType,
  ProfitCenter,
  ProfitCenterKind,
  QuestionAnswer,
  StaffMember,
  TrialBalanceRow,
} from "./types";

const initialProfile: BusinessProfile = {
  businessName: "",
  legalEntity: "",
  businessType: "consulting",
  cadence: "monthly",
  periodStart: "",
  periodEnd: "",
  currency: "INR",
  website: "",
  geography: "India",
  publicNotes: "",
  sourceUrls: "",
  allocationBase: "revenue",
  fundFlowBasis: "trial-balance",
};

const categories = Object.keys(categoryLabels) as AccountCategory[];
const businessTypes = Object.keys(businessTypeLabels) as BusinessType[];
const allocationBases = Object.keys(allocationLabels) as AllocationBase[];
const centerKinds: ProfitCenterKind[] = ["project", "department", "batch", "product", "vertical", "location", "custom"];

function makeId(prefix: string) {
  return `${prefix}-${Math.random().toString(36).slice(2, 9)}`;
}

function currency(value: number, code: string) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: code || "INR",
    maximumFractionDigits: 0,
  }).format(Number.isFinite(value) ? value : 0);
}

function numberValue(value: string) {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function createDefaultProfitCenter(type: BusinessType, index = 1): ProfitCenter {
  const template = getBusinessTemplate(type);
  return {
    id: makeId("pc"),
    name: index === 1 ? template.defaultUnitName : `${template.unitLabel} ${index}`,
    kind: template.defaultKind,
    owner: "",
    segment: "",
    revenueDriver: template.revenueDriverLabel,
    manualRevenue: 0,
    manualDirectCost: 0,
    priorRevenue: 0,
    priorDirectCost: 0,
    studentCount: 0,
    teachingStaffCount: 0,
    nonTeachingStaffCount: 0,
    averageRevenueRate: 0,
    variableCostRate: 0,
    utilizationPercent: 0,
    allocationWeight: 1,
    notes: "",
  };
}

function isUnconfiguredCenter(center: ProfitCenter) {
  return (
    !center.owner &&
    !center.segment &&
    !center.notes &&
    (center.manualRevenue || 0) === 0 &&
    (center.manualDirectCost || 0) === 0 &&
    (center.priorRevenue || 0) === 0 &&
    (center.priorDirectCost || 0) === 0 &&
    (center.studentCount || 0) === 0 &&
    (center.teachingStaffCount || 0) === 0 &&
    (center.nonTeachingStaffCount || 0) === 0 &&
    (center.averageRevenueRate || 0) === 0 &&
    (center.variableCostRate || 0) === 0 &&
    (center.utilizationPercent || 0) === 0
  );
}

const simpleDemoCenters = [
  {
    id: "pc-simple-alnylam",
    name: "ALNYLAM US",
    segment: "Advanced analytics / research",
    owner: "Subash Chander",
    manualRevenue: 7134849,
    manualDirectCost: 270376,
    studentCount: 2613,
    teachingStaffCount: 130,
    allocationWeight: 8.97,
  },
  {
    id: "pc-simple-jntl",
    name: "JNTL Consumer Health",
    segment: "Consumer health analytics",
    owner: "Renuka Rajan",
    manualRevenue: 21802708,
    manualDirectCost: 826215,
    studentCount: 7986,
    teachingStaffCount: 399,
    allocationWeight: 27.42,
  },
  {
    id: "pc-simple-kenvue",
    name: "Kenvue",
    segment: "Consumer intelligence / dashboards",
    owner: "Delivery Lead",
    manualRevenue: 7168835,
    manualDirectCost: 271664,
    studentCount: 2626,
    teachingStaffCount: 131,
    allocationWeight: 9.02,
  },
  {
    id: "pc-simple-replimune",
    name: "Replimune",
    segment: "Research analytics",
    owner: "Delivery Lead",
    manualRevenue: 18160505,
    manualDirectCost: 688194,
    studentCount: 6652,
    teachingStaffCount: 333,
    allocationWeight: 22.84,
  },
  {
    id: "pc-simple-vifor",
    name: "Vifor International",
    segment: "Pharma / healthcare analytics",
    owner: "Delivery Lead",
    manualRevenue: 13116060,
    manualDirectCost: 497034,
    studentCount: 4804,
    teachingStaffCount: 240,
    allocationWeight: 16.5,
  },
  {
    id: "pc-simple-xoma",
    name: "XOMA",
    segment: "Investment / scientific research",
    owner: "Delivery Lead",
    manualRevenue: 12119722,
    manualDirectCost: 459278,
    studentCount: 4439,
    teachingStaffCount: 222,
    allocationWeight: 15.24,
  },
];

function simpleDemoNotes(name: string) {
  return `Demo allocation for ${name} based on client debtor balance as proxy for revenue split. Replace with invoice-wise revenue, project timesheets, and direct vendor tagging for final MIS.`;
}

function createSimpleDemoProfitCenters(): ProfitCenter[] {
  return simpleDemoCenters.map((center) => ({
    ...createDefaultProfitCenter("professional-services"),
    id: center.id,
    name: center.name,
    kind: "project",
    owner: center.owner,
    segment: center.segment,
    revenueDriver: "Retainer / project-based analytics engagement",
    manualRevenue: center.manualRevenue,
    manualDirectCost: center.manualDirectCost,
    priorRevenue: 0,
    priorDirectCost: 0,
    studentCount: center.studentCount,
    teachingStaffCount: center.teachingStaffCount,
    nonTeachingStaffCount: 1,
    averageRevenueRate: 2730,
    variableCostRate: 1331,
    utilizationPercent: 90,
    allocationWeight: center.allocationWeight,
    notes: simpleDemoNotes(center.name),
  }));
}

function createSimpleDemoStaff(): StaffMember[] {
  const assignments = [
    ["Senior delivery lead", "Partner / manager", 600000, [0.1, 0.25, 0.1, 0.25, 0.15, 0.15]],
    ["Healthcare analytics pod", "Analysts", 1200000, [0.15, 0.1, 0.1, 0.3, 0.25, 0.1]],
    ["Consumer analytics pod", "Analysts", 900000, [0.05, 0.45, 0.3, 0.05, 0.05, 0.1]],
    ["Research operations pod", "Research associates", 530000, [0.1, 0.25, 0.05, 0.25, 0.1, 0.25]],
  ] as const;

  return assignments.map(([name, role, monthlyCost, weights], index) => ({
    id: `person-simple-${index + 1}`,
    name,
    role,
    department: "Delivery",
    monthlyCost,
    assignments: simpleDemoCenters.map((center, weightIndex) => ({
      profitCenterId: center.id,
      fte: weights[weightIndex],
    })),
  }));
}

function applySimpleDemoLedgerQc(rows: TrialBalanceRow[]): TrialBalanceRow[] {
  return rows.map((row) => {
    const name = row.accountName.toLowerCase();
    const group = row.accountGroup.toLowerCase();
    let category = row.category;
    let misGroup = row.misGroup;

    if (group.includes("capital account")) {
      category = "equity";
      misGroup = "Equity";
    } else if (["duties", "provisions", "sundry creditors", "current liabilities"].some((term) => group.includes(term))) {
      category = "current-liability";
      misGroup = "Current Liabilities";
    } else if (group.includes("fixed assets")) {
      category = "fixed-asset";
      misGroup = "Fixed Assets";
    } else if (["current assets", "sundry debtors", "bank accounts"].some((term) => group.includes(term))) {
      category = "current-asset";
      misGroup = "Current Assets";
    } else if (name.includes("interest on fd")) {
      category = "other-income";
      misGroup = "Other Income";
    } else if (name.includes("professional") && name.includes("export")) {
      category = "revenue";
      misGroup = "Operating Revenue";
    } else if (["salary", "bonus", "epf", "edli", "staff welfare", "comp. related"].some((term) => name.includes(term))) {
      category = "people-cost";
      misGroup = "People Costs";
    } else if (name.includes("consultancy fees")) {
      category = "direct-cost";
      misGroup = "Direct Costs";
    } else if (name.includes("bank charges") || name.includes("currency rate fluctuation")) {
      category = "finance-cost";
      misGroup = "Finance Costs";
    } else if (name.includes("interest paid on gst") || name.includes("interest paid on tds")) {
      category = "tax";
      misGroup = "Taxes";
    } else if (group.includes("indirect expenses")) {
      category = "operating-expense";
      misGroup = "Operating Expenses";
    }

    return {
      ...row,
      category,
      misGroup,
      confidence: 1,
      profitCenterId: "",
    };
  });
}

export function App() {
  const [profile, setProfile] = useState<BusinessProfile>(initialProfile);
  const [rows, setRows] = useState<TrialBalanceRow[]>([]);
  const [profitCenters, setProfitCenters] = useState<ProfitCenter[]>([createDefaultProfitCenter(initialProfile.businessType)]);
  const [staff, setStaff] = useState<StaffMember[]>([]);
  const [answers, setAnswers] = useState<QuestionAnswer[]>([]);
  const [bankTransactions, setBankTransactions] = useState<BankTransaction[]>([]);
  const [sourceWarnings, setSourceWarnings] = useState<string[]>([]);
  const [bankWarnings, setBankWarnings] = useState<string[]>([]);
  const [activeTab, setActiveTab] = useState("source");
  const [isParsing, setIsParsing] = useState(false);
  const [importError, setImportError] = useState("");
  const inputRef = useRef<HTMLInputElement | null>(null);
  const bankStatementInputRef = useRef<HTMLInputElement | null>(null);
  const bankLedgerInputRef = useRef<HTMLInputElement | null>(null);

  const template = useMemo(() => getBusinessTemplate(profile.businessType), [profile.businessType]);
  const questions = useMemo(() => buildQuestions(profile, rows, profitCenters, staff, answers, bankTransactions), [profile, rows, profitCenters, staff, answers, bankTransactions]);
  const issues = useMemo(() => buildWorkbookIssues(profile, rows, profitCenters, staff, questions, bankTransactions), [profile, rows, profitCenters, staff, questions, bankTransactions]);
  const totalDebit = rows.reduce((sum, row) => sum + row.debit, 0);
  const totalCredit = rows.reduce((sum, row) => sum + row.credit, 0);
  const revenue = rows.filter((row) => row.category === "revenue").reduce((sum, row) => sum + Math.abs(row.credit - row.debit), 0);
  const expenses = rows
    .filter((row) => ["direct-cost", "people-cost", "operating-expense", "finance-cost", "tax"].includes(row.category))
    .reduce((sum, row) => sum + Math.abs(row.debit - row.credit), 0);
  const bankStatementTransactions = bankTransactions.filter((txn) => txn.sourceType === "bank-statement");
  const bankLedgerTransactions = bankTransactions.filter((txn) => txn.sourceType === "bank-ledger");
  const inferredBankAccounts = rows.filter((row) => row.category === "current-asset" && /bank|cash|fd|fixed deposit/i.test(`${row.accountName} ${row.accountGroup}`));

  function loadSimpleDemoData() {
    const demoCenters = createSimpleDemoProfitCenters();
    setProfile({
      businessName: "The Simple",
      legalEntity: "Simple Insights Private Limited",
      businessType: "professional-services",
      cadence: "annual",
      periodStart: "2025-04-01",
      periodEnd: "2026-03-31",
      currency: "INR",
      website: "https://thesimpleintel.com/",
      geography: "India / export services",
      publicNotes:
        "Advanced analytics, research, investment intelligence, financial modelling, dashboards, and AI-enabled decision-support services. Demo allocation uses debtor balances as a proxy for engagement revenue split until invoice-wise data is available.",
      sourceUrls: "https://thesimpleintel.com/\nhttps://thesimpleintel.com/about-us",
      allocationBase: "revenue",
      fundFlowBasis: "trial-balance",
    });
    setProfitCenters(demoCenters);
    setStaff(createSimpleDemoStaff());
    setAnswers([
      {
        id: "assign-pnl-ledgers",
        question: "Assign P&L ledger accounts to profit centers or shared overhead.",
        answer: "Demo preset keeps trial-balance P&L ledgers shared and uses manual engagement revenue/direct-cost allocations based on client debtor balances to avoid double-counting ledger revenue.",
        status: "answered",
      },
      {
        id: "ps-engagement-register",
        question: "List engagements by client, service line, partner, fee arrangement, billing milestone, and status.",
        answer: "Six demo engagements loaded: ALNYLAM US, JNTL Consumer Health, Kenvue, Replimune, Vifor International, and XOMA, with client/service line, owner, fee arrangement, revenue, direct cost, hours, and notes.",
        status: "answered",
      },
      {
        id: "ps-leverage",
        question: "Capture partner, manager, associate, and support hours for each engagement.",
        answer: "Demo staffing matrix loaded with senior delivery lead, healthcare analytics pod, consumer analytics pod, and research operations pod. Assignment weights represent FTE/time split across engagements.",
        status: "answered",
      },
      {
        id: "ps-wip-ar",
        question: "Capture WIP, billed receivables, retainer advances, write-offs, and collection ageing by client.",
        answer: "Client debtor balances in the trial balance are used as a demo receivables proxy. Final MIS should replace this with invoice-wise ageing and WIP data.",
        status: "answered",
      },
    ]);
    setRows((current) => applySimpleDemoLedgerQc(current));
    setActiveTab("allocation");
  }

  async function handleFile(file: File) {
    setIsParsing(true);
    setImportError("");
    try {
      const parsed = parseTrialBalanceWorkbook(await file.arrayBuffer(), file.name);
      const parsedRows = parsed.rows;
      if (!parsedRows.length) {
        setImportError("No ledger rows were detected. Use a CSV/XLSX file with account, debit, credit, or balance columns.");
      }
      setSourceWarnings(parsed.warnings);
      setProfile((current) => ({
        ...current,
        businessName: parsed.metadata.businessName || current.businessName,
        legalEntity: parsed.metadata.businessName || current.legalEntity,
        cadence: parsed.metadata.periodStart && parsed.metadata.periodEnd ? "annual" : current.cadence,
        periodStart: parsed.metadata.periodStart || current.periodStart,
        periodEnd: parsed.metadata.periodEnd || current.periodEnd,
      }));
      setRows(profile.businessName === "The Simple" ? applySimpleDemoLedgerQc(parsedRows) : parsedRows);
      setActiveTab("mapping");
    } catch (error) {
      setImportError(error instanceof Error ? error.message : "Could not parse the selected file.");
    } finally {
      setIsParsing(false);
    }
  }

  function inferredFundFlowBasis(transactions: BankTransaction[]): BusinessProfile["fundFlowBasis"] {
    const hasStatement = transactions.some((txn) => txn.sourceType === "bank-statement");
    const hasLedger = transactions.some((txn) => txn.sourceType === "bank-ledger");
    if (hasStatement && hasLedger) return "bank-statement-and-ledger";
    if (hasStatement) return "bank-statement";
    if (hasLedger) return "bank-ledger";
    return profile.fundFlowBasis;
  }

  async function handleBankFile(file: File, sourceType: BankSourceType) {
    setIsParsing(true);
    setImportError("");
    try {
      const parsed = parseBankStatementWorkbook(await file.arrayBuffer(), sourceType, file.name);
      setBankTransactions((current) => {
        const withoutSameFile = current.filter((txn) => !(txn.sourceType === sourceType && txn.sourceFileName === file.name));
        const next = [...withoutSameFile, ...parsed.transactions];
        setProfile((profileCurrent) => ({ ...profileCurrent, fundFlowBasis: next.length ? inferredFundFlowBasis(next) : profileCurrent.fundFlowBasis }));
        return next;
      });
      setBankWarnings((current) => [...current.filter((warning) => !warning.includes(file.name)), ...parsed.warnings.map((warning) => `${file.name}: ${warning}`)]);
      setActiveTab("bank");
    } catch (error) {
      setImportError(error instanceof Error ? error.message : "Could not parse the selected bank source.");
    } finally {
      setIsParsing(false);
    }
  }

  function handleBankFiles(fileList: FileList | null, sourceType: BankSourceType) {
    Array.from(fileList || []).forEach((file) => {
      void handleBankFile(file, sourceType);
    });
  }

  function handleBusinessTypeChange(type: BusinessType) {
    setProfile((current) => ({
      ...current,
      businessType: type,
    }));
    setProfitCenters((current) => {
      if (!current.length || (current.length === 1 && isUnconfiguredCenter(current[0]) && !rows.some((row) => row.profitCenterId))) {
        return [createDefaultProfitCenter(type)];
      }
      return current.map((center) => ({ ...center, kind: center.kind || getBusinessTemplate(type).defaultKind }));
    });
  }

  function addProfitCenter() {
    setProfitCenters((current) => [
      ...current,
      createDefaultProfitCenter(profile.businessType, current.length + 1),
    ]);
  }

  function updateProfitCenter(id: string, patch: Partial<ProfitCenter>) {
    setProfitCenters((current) => current.map((center) => (center.id === id ? { ...center, ...patch } : center)));
  }

  function addStaffMember() {
    setStaff((current) => [
      ...current,
      {
        id: makeId("person"),
        name: `Team member ${current.length + 1}`,
        role: "",
        department: "",
        monthlyCost: 0,
        assignments: profitCenters.map((center, index) => ({ profitCenterId: center.id, fte: index === 0 ? 1 : 0 })),
      },
    ]);
  }

  function updateStaff(id: string, patch: Partial<StaffMember>) {
    setStaff((current) => current.map((person) => (person.id === id ? { ...person, ...patch } : person)));
  }

  function updateAssignment(personId: string, profitCenterId: string, fte: number) {
    setStaff((current) =>
      current.map((person) => {
        if (person.id !== personId) return person;
        const existing = person.assignments.some((assignment) => assignment.profitCenterId === profitCenterId);
        const assignments = existing
          ? person.assignments.map((assignment) => (assignment.profitCenterId === profitCenterId ? { ...assignment, fte } : assignment))
          : [...person.assignments, { profitCenterId, fte }];
        return { ...person, assignments };
      }),
    );
  }

  function updateAnswer(questionId: string, question: string, answer: string) {
    setAnswers((current) => {
      const existing = current.find((item) => item.id === questionId);
      if (existing) {
        return current.map((item) => (item.id === questionId ? { ...item, answer, status: answer.trim() ? "answered" : "open" } : item));
      }
      return [...current, { id: questionId, question, answer, status: answer.trim() ? "answered" : "open" }];
    });
  }

  function exportWorkbook() {
    if (!rows.length) {
      const confirmed = window.confirm(
        "No trial balance is loaded. This will export a demo-only MIS production pack with missing-ledger warnings, blank raw ledger schedules, and zero working-capital schedules. Continue?",
      );
      if (!confirmed) {
        setActiveTab("source");
        return;
      }
    }

    const output = generateMisWorkbook({ profile, rows, centers: profitCenters, staff, questions, answers, bankTransactions });
    XLSX.writeFile(output.workbook, output.filename, { compression: true });
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark">
            <BarChart3 size={22} />
          </div>
          <div>
            <h1>Custom MIS</h1>
            <p>{businessTypeLabels[profile.businessType]}</p>
          </div>
        </div>

        <nav className="step-nav" aria-label="MIS sections">
          <TabButton icon={<Upload size={18} />} label="Source" value="source" active={activeTab} onClick={setActiveTab} />
          <TabButton icon={<Landmark size={18} />} label="Bank & Fund Flow" value="bank" active={activeTab} onClick={setActiveTab} />
          <TabButton icon={<Building2 size={18} />} label="Profile" value="profile" active={activeTab} onClick={setActiveTab} />
          <TabButton icon={<Layers3 size={18} />} label="Mapping" value="mapping" active={activeTab} onClick={setActiveTab} />
          <TabButton icon={<Users size={18} />} label="Allocation" value="allocation" active={activeTab} onClick={setActiveTab} />
          <TabButton icon={<ClipboardList size={18} />} label="Questions" value="questions" active={activeTab} onClick={setActiveTab} />
          <TabButton icon={<Download size={18} />} label="Workbook" value="workbook" active={activeTab} onClick={setActiveTab} />
        </nav>

        <div className="side-panel">
          <span className="eyebrow">Readiness</span>
          <div className="readiness">
            <strong>{Math.max(0, 100 - issues.length * 15)}%</strong>
            <span>{issues.length ? `${issues.length} open checks` : "ready to export"}</span>
          </div>
          <div className="meter">
            <span style={{ width: `${Math.max(6, 100 - issues.length * 15)}%` }} />
          </div>
        </div>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div>
            <p className="eyebrow">{profile.cadence} report</p>
            <h2>{profile.businessName || "Untitled MIS report"}</h2>
          </div>
          <div className="topbar-actions">
            <button className="secondary-action" type="button" onClick={loadSimpleDemoData}>
              <Sparkles size={18} />
              Load The Simple demo
            </button>
            <button className="primary-action" type="button" onClick={exportWorkbook}>
              <Download size={18} />
              Export Excel
            </button>
          </div>
        </header>

        <section className="metric-grid" aria-label="MIS summary">
          <Metric label="Ledger rows" value={rows.length.toString()} />
          <Metric label="Debit" value={currency(totalDebit, profile.currency)} />
          <Metric label="Credit" value={currency(totalCredit, profile.currency)} />
          <Metric label="Revenue" value={currency(revenue, profile.currency)} />
          <Metric label="P&L cost" value={currency(expenses, profile.currency)} />
          <Metric label="Bank rows" value={bankTransactions.length.toString()} />
        </section>

        {activeTab === "source" && (
          <Panel title="Accounting Source" icon={<FileSpreadsheet size={20} />}>
            <div className="upload-zone" onClick={() => inputRef.current?.click()} onKeyDown={() => inputRef.current?.click()} role="button" tabIndex={0}>
              <input
                ref={inputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={(event) => {
                  const file = event.target.files?.[0];
                  if (file) void handleFile(file);
                }}
              />
              <Upload size={26} />
              <strong>{isParsing ? "Reading file..." : "Upload trial balance or ledger export"}</strong>
              <span>Tally, Zoho Books, Busy, QuickBooks, CSV, XLS, or XLSX with account and amount columns.</span>
            </div>
            <div className="demo-callout">
              <div>
                <span className="eyebrow">Demo shortcut</span>
                <strong>The Simple engagement MIS</strong>
                <p>Pre-fill the business profile, six client engagements, allocation notes, and staffing matrix. Upload the trial balance before or after this step.</p>
              </div>
              <button className="secondary-action" type="button" onClick={loadSimpleDemoData}>
                <Sparkles size={16} />
                Load demo data
              </button>
            </div>
            <div className="demo-callout">
              <div>
                <span className="eyebrow">Fund flow source</span>
                <strong>{bankTransactions.length ? `${bankTransactions.length} bank transactions loaded` : "Bank & Fund Flow has its own upload step"}</strong>
                <p>Use the dedicated bank section for separate Bank Statement and Bank Ledger uploads. The TB period and account categories are reused automatically.</p>
              </div>
              <button className="secondary-action" type="button" onClick={() => setActiveTab("bank")}>
                <Landmark size={16} />
                Open bank section
              </button>
            </div>
            {importError && <div className="alert danger">{importError}</div>}
            {[...sourceWarnings, ...bankWarnings].map((warning) => (
              <div className="alert warning" key={warning}>{warning}</div>
            ))}
            <div className="split">
              <div className="quiet-box">
                <span className="eyebrow">Detected structure</span>
                <strong>{rows.length ? `${rows.length} accounts imported` : "Awaiting source file"}</strong>
                <p>Debit-credit imbalance: {currency(Math.abs(totalDebit - totalCredit), profile.currency)}</p>
              </div>
              <div className="quiet-box">
                <span className="eyebrow">Classification</span>
                <strong>{rows.filter((row) => row.category !== "unknown").length} mapped</strong>
                <p>{rows.filter((row) => row.category === "unknown").length} rows need review.</p>
              </div>
              <div className="quiet-box">
                <span className="eyebrow">Inferred profile</span>
                <strong>{profile.businessName || "Not inferred yet"}</strong>
                <p>{profile.periodStart && profile.periodEnd ? `${profile.periodStart} to ${profile.periodEnd}` : "Upload a TB with period headers."}</p>
              </div>
            </div>
          </Panel>
        )}

        {activeTab === "bank" && (
          <Panel title="Bank & Fund Flow" icon={<Landmark size={20} />}>
            <div className="bank-upload-grid">
              <article className="bank-upload-card">
                <div>
                  <span className="eyebrow">Bank statement</span>
                  <strong>{bankStatementTransactions.length ? `${bankStatementTransactions.length} statement transactions` : "Upload bank statements"}</strong>
                  <p>Use statements as the cash-movement truth for actual receipts, payments, closing balances, and reconciliation.</p>
                </div>
                <input
                  ref={bankStatementInputRef}
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  multiple
                  hidden
                  onChange={(event) => {
                    handleBankFiles(event.target.files, "bank-statement");
                    event.currentTarget.value = "";
                  }}
                />
                <button className="secondary-action" type="button" onClick={() => bankStatementInputRef.current?.click()}>
                  <Upload size={16} />
                  Upload statements
                </button>
              </article>

              <article className="bank-upload-card">
                <div>
                  <span className="eyebrow">Bank ledger</span>
                  <strong>{bankLedgerTransactions.length ? `${bankLedgerTransactions.length} ledger transactions` : "Upload bank ledgers"}</strong>
                  <p>Use ledgers for book-side classification, timing support, and reconciliation against bank statement movement.</p>
                </div>
                <input
                  ref={bankLedgerInputRef}
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  multiple
                  hidden
                  onChange={(event) => {
                    handleBankFiles(event.target.files, "bank-ledger");
                    event.currentTarget.value = "";
                  }}
                />
                <button className="secondary-action" type="button" onClick={() => bankLedgerInputRef.current?.click()}>
                  <Upload size={16} />
                  Upload ledgers
                </button>
              </article>
            </div>

            <div className="split">
              <div className="quiet-box">
                <span className="eyebrow">Inferred from TB</span>
                <strong>{inferredBankAccounts.length ? `${inferredBankAccounts.length} bank/cash/FD accounts` : "No bank accounts inferred yet"}</strong>
                <p>{inferredBankAccounts.slice(0, 4).map((row) => row.accountName).join(", ") || "Upload the TB first so the app can match bank sources to book balances."}</p>
              </div>
              <div className="quiet-box">
                <span className="eyebrow">Fund flow basis</span>
                <strong>{profile.fundFlowBasis.replace(/-/g, " ")}</strong>
                <p>{bankStatementTransactions.length ? "Actual fund flow will use statements first, with ledgers as support." : bankLedgerTransactions.length ? "Ledger-only fund flow is provisional until statements are uploaded." : "Upload bank sources to complete fund flow."}</p>
              </div>
            </div>

            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Date</th>
                    <th>Source</th>
                    <th>Narration</th>
                    <th>Payment</th>
                    <th>Receipt</th>
                    <th>Fund-flow group</th>
                  </tr>
                </thead>
                <tbody>
                  {bankTransactions.slice(0, 120).map((txn) => (
                    <tr key={txn.id}>
                      <td>{txn.date}</td>
                      <td>
                        <strong>{txn.sourceType === "bank-statement" ? "Statement" : "Ledger"}</strong>
                        <span>{txn.sourceFileName || txn.sourceSheet}</span>
                      </td>
                      <td>{txn.narration}</td>
                      <td>{currency(txn.debit, profile.currency)}</td>
                      <td>{currency(txn.credit, profile.currency)}</td>
                      <td>
                        <strong>{txn.fundFlowGroup}</strong>
                        <span>{txn.isInterAccountTransfer ? "Internal transfer excluded" : txn.category}</span>
                      </td>
                    </tr>
                  ))}
                  {!bankTransactions.length && (
                    <tr>
                      <td colSpan={6}>Upload statements and ledgers to build the AKEV-style fund flow and next-period projection.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            {bankTransactions.length > 120 && <p className="table-note">Showing first 120 bank rows. All imported rows are included in the workbook.</p>}
          </Panel>
        )}

        {activeTab === "profile" && (
          <Panel title="Business Profile" icon={<Building2 size={20} />}>
            <div className="form-grid">
              <Field label="Business name" value={profile.businessName} onChange={(value) => setProfile({ ...profile, businessName: value })} />
              <Field label="Legal entity" value={profile.legalEntity} onChange={(value) => setProfile({ ...profile, legalEntity: value })} />
              <label>
                Business type
                <select value={profile.businessType} onChange={(event) => handleBusinessTypeChange(event.target.value as BusinessType)}>
                  {businessTypes.map((type) => (
                    <option value={type} key={type}>
                      {businessTypeLabels[type]}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                Cadence
                <select value={profile.cadence} onChange={(event) => setProfile({ ...profile, cadence: event.target.value as BusinessProfile["cadence"] })}>
                  <option value="monthly">Monthly</option>
                  <option value="quarterly">Quarterly</option>
                  <option value="annual">Annual</option>
                </select>
              </label>
              <Field label="Period start" type="date" value={profile.periodStart} onChange={(value) => setProfile({ ...profile, periodStart: value })} />
              <Field label="Period end" type="date" value={profile.periodEnd} onChange={(value) => setProfile({ ...profile, periodEnd: value })} />
              <Field label="Currency" value={profile.currency} onChange={(value) => setProfile({ ...profile, currency: value.toUpperCase() })} />
              <Field label="Geography" value={profile.geography} onChange={(value) => setProfile({ ...profile, geography: value })} />
              <label>
                Fund flow basis
                <select value={profile.fundFlowBasis} onChange={(event) => setProfile({ ...profile, fundFlowBasis: event.target.value as BusinessProfile["fundFlowBasis"] })}>
                  <option value="bank-statement-and-ledger">Bank statement + ledger</option>
                  <option value="bank-statement">Bank statement</option>
                  <option value="bank-ledger">Bank ledger</option>
                  <option value="trial-balance">Trial balance proxy</option>
                  <option value="manual">Manual assumptions</option>
                </select>
              </label>
            </div>
            <div className="form-grid single">
              <Field label="Website / public profile" value={profile.website} onChange={(value) => setProfile({ ...profile, website: value })} icon={<Globe2 size={16} />} />
              <label>
                Public research notes
                <textarea value={profile.publicNotes} onChange={(event) => setProfile({ ...profile, publicNotes: event.target.value })} rows={4} />
              </label>
              <label>
                Source URLs
                <textarea value={profile.sourceUrls} onChange={(event) => setProfile({ ...profile, sourceUrls: event.target.value })} rows={3} />
              </label>
            </div>
            <div className="template-band">
              <div>
                <span className="eyebrow">Unit</span>
                <strong>{template.unitPlural}</strong>
              </div>
              <div>
                <span className="eyebrow">Deeper cut</span>
                <strong>{template.subUnitLabel}</strong>
              </div>
              <div>
                <span className="eyebrow">Core driver</span>
                <strong>{template.metricLabels.primary} x {template.metricLabels.averageRate}</strong>
              </div>
              <div>
                <span className="eyebrow">Deep sheet</span>
                <strong>{template.deepDiveSheet}</strong>
              </div>
            </div>
          </Panel>
        )}

        {activeTab === "mapping" && (
          <Panel title="Ledger Mapping" icon={<Layers3 size={20} />}>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Account</th>
                    <th>Debit</th>
                    <th>Credit</th>
                    <th>Category</th>
                    <th>MIS group</th>
                    <th>Profit center</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.slice(0, 200).map((row) => (
                    <tr key={row.id}>
                      <td>
                        <strong>{row.accountName}</strong>
                        <span>{row.accountGroup || row.sourceSheet}</span>
                      </td>
                      <td>{currency(row.debit, profile.currency)}</td>
                      <td>{currency(row.credit, profile.currency)}</td>
                      <td>
                        <select
                          value={row.category}
                          onChange={(event) =>
                            setRows((current) =>
                              current.map((item) => (item.id === row.id ? { ...item, category: event.target.value as AccountCategory, confidence: 1 } : item)),
                            )
                          }
                        >
                          {categories.map((category) => (
                            <option key={category} value={category}>
                              {categoryLabels[category]}
                            </option>
                          ))}
                        </select>
                      </td>
                      <td>{row.misGroup}</td>
                      <td>
                        <select
                          value={row.profitCenterId}
                          onChange={(event) =>
                            setRows((current) => current.map((item) => (item.id === row.id ? { ...item, profitCenterId: event.target.value } : item)))
                          }
                        >
                          <option value="">Shared / unassigned</option>
                          {profitCenters.map((center) => (
                            <option key={center.id} value={center.id}>
                              {center.name}
                            </option>
                          ))}
                        </select>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {rows.length > 200 && <p className="table-note">Showing first 200 ledger rows. All imported rows are included in the workbook.</p>}
          </Panel>
        )}

        {activeTab === "allocation" && (
          <Panel title="Profit Center Allocation" icon={<Users size={20} />}>
            <div className="toolbar">
              <label>
                Shared cost basis
                <select value={profile.allocationBase} onChange={(event) => setProfile({ ...profile, allocationBase: event.target.value as AllocationBase })}>
                  {allocationBases.map((base) => (
                    <option value={base} key={base}>
                      {allocationLabels[base]}
                    </option>
                  ))}
                </select>
              </label>
              <button type="button" className="secondary-action" onClick={addProfitCenter}>
                <Plus size={16} />
                Add {template.unitLabel}
              </button>
            </div>

            <div className="center-grid">
              {profitCenters.map((center) => (
                <article className="unit-card" key={center.id}>
                  <div className="card-head">
                    <input value={center.name} onChange={(event) => updateProfitCenter(center.id, { name: event.target.value })} />
                    <button
                      type="button"
                      className="icon-button"
                      aria-label="Remove profit center"
                      onClick={() => setProfitCenters((current) => current.filter((item) => item.id !== center.id))}
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>
                  <div className="mini-grid">
                    <label>
                      Kind
                      <select value={center.kind} onChange={(event) => updateProfitCenter(center.id, { kind: event.target.value as ProfitCenterKind })}>
                        {centerKinds.map((kind) => (
                          <option value={kind} key={kind}>
                            {kind}
                          </option>
                      ))}
                      </select>
                    </label>
                    <Field label={template.segmentLabel} value={center.segment || ""} onChange={(value) => updateProfitCenter(center.id, { segment: value })} />
                    <Field label={template.ownerLabel} value={center.owner || ""} onChange={(value) => updateProfitCenter(center.id, { owner: value })} />
                    <Field label={template.revenueDriverLabel} value={center.revenueDriver || ""} onChange={(value) => updateProfitCenter(center.id, { revenueDriver: value })} />
                    <Field
                      label="Manual revenue"
                      type="number"
                      value={String(center.manualRevenue ?? 0)}
                      onChange={(value) => updateProfitCenter(center.id, { manualRevenue: numberValue(value) })}
                    />
                    <Field
                      label={template.directCostLabel}
                      type="number"
                      value={String(center.manualDirectCost ?? 0)}
                      onChange={(value) => updateProfitCenter(center.id, { manualDirectCost: numberValue(value) })}
                    />
                    <Field
                      label="Prior revenue"
                      type="number"
                      value={String(center.priorRevenue ?? 0)}
                      onChange={(value) => updateProfitCenter(center.id, { priorRevenue: numberValue(value) })}
                    />
                    <Field
                      label="Prior direct cost"
                      type="number"
                      value={String(center.priorDirectCost ?? 0)}
                      onChange={(value) => updateProfitCenter(center.id, { priorDirectCost: numberValue(value) })}
                    />
                    <Field
                      label={template.metricLabels.primary}
                      type="number"
                      value={String(center.studentCount ?? 0)}
                      onChange={(value) => updateProfitCenter(center.id, { studentCount: numberValue(value) })}
                    />
                    <Field
                      label={template.metricLabels.secondary}
                      type="number"
                      value={String(center.teachingStaffCount ?? 0)}
                      onChange={(value) => updateProfitCenter(center.id, { teachingStaffCount: numberValue(value) })}
                    />
                    <Field
                      label={template.metricLabels.tertiary}
                      type="number"
                      value={String(center.nonTeachingStaffCount ?? 0)}
                      onChange={(value) => updateProfitCenter(center.id, { nonTeachingStaffCount: numberValue(value) })}
                    />
                    <Field
                      label={template.metricLabels.averageRate}
                      type="number"
                      value={String(center.averageRevenueRate ?? 0)}
                      onChange={(value) => updateProfitCenter(center.id, { averageRevenueRate: numberValue(value) })}
                    />
                    <Field
                      label={template.metricLabels.variableCostRate}
                      type="number"
                      value={String(center.variableCostRate ?? 0)}
                      onChange={(value) => updateProfitCenter(center.id, { variableCostRate: numberValue(value) })}
                    />
                    <Field
                      label={template.metricLabels.utilization}
                      type="number"
                      value={String(center.utilizationPercent ?? 0)}
                      onChange={(value) => updateProfitCenter(center.id, { utilizationPercent: numberValue(value) })}
                    />
                    <Field
                      label="Allocation weight"
                      type="number"
                      value={String(center.allocationWeight ?? 0)}
                      onChange={(value) => updateProfitCenter(center.id, { allocationWeight: numberValue(value) })}
                    />
                  </div>
                  <label>
                    Notes
                    <textarea value={center.notes || ""} onChange={(event) => updateProfitCenter(center.id, { notes: event.target.value })} rows={2} />
                  </label>
                </article>
              ))}
            </div>

            <div className="section-head">
              <div>
                <span className="eyebrow">People cost</span>
                <h3>Staffing matrix</h3>
              </div>
              <button type="button" className="secondary-action" onClick={addStaffMember}>
                <Plus size={16} />
                Add person
              </button>
            </div>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Name</th>
                    <th>Role</th>
                    <th>Monthly cost</th>
                    {profitCenters.map((center) => (
                      <th key={center.id}>{center.name}</th>
                    ))}
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {staff.map((person) => (
                    <tr key={person.id}>
                      <td>
                        <input value={person.name} onChange={(event) => updateStaff(person.id, { name: event.target.value })} />
                      </td>
                      <td>
                        <input value={person.role} onChange={(event) => updateStaff(person.id, { role: event.target.value })} />
                      </td>
                      <td>
                        <input type="number" value={person.monthlyCost} onChange={(event) => updateStaff(person.id, { monthlyCost: numberValue(event.target.value) })} />
                      </td>
                      {profitCenters.map((center) => (
                        <td key={center.id}>
                          <input
                            type="number"
                            min="0"
                            step="0.1"
                            value={person.assignments.find((assignment) => assignment.profitCenterId === center.id)?.fte || 0}
                            onChange={(event) => updateAssignment(person.id, center.id, numberValue(event.target.value))}
                          />
                        </td>
                      ))}
                      <td>
                        <button type="button" className="icon-button" aria-label="Remove person" onClick={() => setStaff((current) => current.filter((item) => item.id !== person.id))}>
                          <Trash2 size={16} />
                        </button>
                      </td>
                    </tr>
                  ))}
                  {!staff.length && (
                    <tr>
                      <td colSpan={profitCenters.length + 4}>No staff rows yet.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </Panel>
        )}

        {activeTab === "questions" && (
          <Panel title="Question Queue" icon={<ClipboardList size={20} />}>
            <div className="question-list">
              {questions.map((question) => {
                const answer = answers.find((item) => item.id === question.id);
                return (
                  <article className={`question ${question.priority}`} key={question.id}>
                    <div>
                      <span>{question.section}</span>
                      <strong>{question.prompt}</strong>
                      <p>{question.reason}</p>
                    </div>
                    <textarea
                      value={answer?.answer || ""}
                      rows={3}
                      onChange={(event) => updateAnswer(question.id, question.prompt, event.target.value)}
                      placeholder="Answer or assumption"
                    />
                  </article>
                );
              })}
              {!questions.length && (
                <div className="success-state">
                  <Check size={24} />
                  <strong>No open generated questions</strong>
                  <span>The workbook will still include the full question register.</span>
                </div>
              )}
            </div>
          </Panel>
        )}

        {activeTab === "workbook" && (
          <Panel title="Workbook Export" icon={<Download size={20} />}>
            <div className="export-grid">
              <div className="export-card">
                <span className="eyebrow">Sheets generated</span>
                <strong>{["Control Panel", "Balance Sheet", "Tax & Statutory", "Fund Flow", "Fund Flow Projection", "Bank Reconciliation", "ChartData", "Executive MIS", "Unit Profitability", ...template.workbookSheets, "Question Register", "QC Tie-Outs"].join(", ")}</strong>
                <button className="primary-action" type="button" onClick={exportWorkbook}>
                  <Download size={18} />
                  Generate MIS workbook
                </button>
              </div>
              <div className="issue-list">
                {issues.map((issue) => (
                  <div className={`issue ${issue.severity}`} key={`${issue.label}-${issue.detail}`}>
                    <AlertTriangle size={18} />
                    <div>
                      <strong>{issue.label}</strong>
                      <span>{issue.detail}</span>
                    </div>
                  </div>
                ))}
                {!issues.length && (
                  <div className="issue low">
                    <Check size={18} />
                    <div>
                      <strong>Ready</strong>
                      <span>The MIS model has no blocking checks.</span>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </Panel>
        )}
      </section>
    </main>
  );
}

function TabButton({
  icon,
  label,
  value,
  active,
  onClick,
}: {
  icon: React.ReactNode;
  label: string;
  value: string;
  active: string;
  onClick: (value: string) => void;
}) {
  return (
    <button type="button" className={active === value ? "active" : ""} onClick={() => onClick(value)}>
      {icon}
      {label}
    </button>
  );
}

function Metric({ label, value }: { label: string; value: string }) {
  return (
    <article className="metric">
      <span>{label}</span>
      <strong>{value}</strong>
    </article>
  );
}

function Panel({ title, icon, children }: { title: string; icon: React.ReactNode; children: React.ReactNode }) {
  return (
    <section className="panel">
      <div className="panel-title">
        {icon}
        <h3>{title}</h3>
      </div>
      {children}
    </section>
  );
}

function Field({
  label,
  value,
  onChange,
  type = "text",
  icon,
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  type?: string;
  icon?: React.ReactNode;
}) {
  return (
    <label>
      {label}
      <span className={icon ? "input-icon" : undefined}>
        {icon}
        <input type={type} value={value} onChange={(event) => onChange(event.target.value)} />
      </span>
    </label>
  );
}
