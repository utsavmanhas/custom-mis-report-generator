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
  Layers3,
  Plus,
  Trash2,
  Upload,
  Users,
} from "lucide-react";
import { allocationLabels, businessTypeLabels, categoryLabels } from "./lib/classification";
import { getBusinessTemplate } from "./lib/businessTemplates";
import { parseTrialBalanceFile } from "./lib/parser";
import { buildQuestions } from "./lib/questions";
import { buildWorkbookIssues, generateMisWorkbook } from "./lib/reporting";
import type {
  AccountCategory,
  AllocationBase,
  BusinessProfile,
  BusinessType,
  ProfitCenter,
  ProfitCenterKind,
  QuestionAnswer,
  StaffMember,
  TrialBalanceRow,
} from "./types";

const initialProfile: BusinessProfile = {
  businessName: "Sample Business",
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

export function App() {
  const [profile, setProfile] = useState<BusinessProfile>(initialProfile);
  const [rows, setRows] = useState<TrialBalanceRow[]>([]);
  const [profitCenters, setProfitCenters] = useState<ProfitCenter[]>([createDefaultProfitCenter(initialProfile.businessType)]);
  const [staff, setStaff] = useState<StaffMember[]>([]);
  const [answers, setAnswers] = useState<QuestionAnswer[]>([]);
  const [activeTab, setActiveTab] = useState("source");
  const [isParsing, setIsParsing] = useState(false);
  const [importError, setImportError] = useState("");
  const inputRef = useRef<HTMLInputElement | null>(null);

  const template = useMemo(() => getBusinessTemplate(profile.businessType), [profile.businessType]);
  const questions = useMemo(() => buildQuestions(profile, rows, profitCenters, staff, answers), [profile, rows, profitCenters, staff, answers]);
  const issues = useMemo(() => buildWorkbookIssues(profile, rows, profitCenters, staff, questions), [profile, rows, profitCenters, staff, questions]);
  const totalDebit = rows.reduce((sum, row) => sum + row.debit, 0);
  const totalCredit = rows.reduce((sum, row) => sum + row.credit, 0);
  const revenue = rows.filter((row) => row.category === "revenue").reduce((sum, row) => sum + Math.abs(row.credit - row.debit), 0);
  const expenses = rows
    .filter((row) => ["direct-cost", "people-cost", "operating-expense", "finance-cost", "tax"].includes(row.category))
    .reduce((sum, row) => sum + Math.abs(row.debit - row.credit), 0);

  async function handleFile(file: File) {
    setIsParsing(true);
    setImportError("");
    try {
      const parsedRows = await parseTrialBalanceFile(file);
      if (!parsedRows.length) {
        setImportError("No ledger rows were detected. Use a CSV/XLSX file with account, debit, credit, or balance columns.");
      }
      setRows(parsedRows);
      setActiveTab("mapping");
    } catch (error) {
      setImportError(error instanceof Error ? error.message : "Could not parse the selected file.");
    } finally {
      setIsParsing(false);
    }
  }

  function handleBusinessTypeChange(type: BusinessType) {
    setProfile((current) => ({
      ...current,
      businessType: type,
      businessName: current.businessName === "Sample University" || current.businessName === "Sample Business" ? "Sample Business" : current.businessName,
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
    const output = generateMisWorkbook({ profile, rows, centers: profitCenters, staff, questions, answers });
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
          <button className="primary-action" type="button" onClick={exportWorkbook}>
            <Download size={18} />
            Export Excel
          </button>
        </header>

        <section className="metric-grid" aria-label="MIS summary">
          <Metric label="Ledger rows" value={rows.length.toString()} />
          <Metric label="Debit" value={currency(totalDebit, profile.currency)} />
          <Metric label="Credit" value={currency(totalCredit, profile.currency)} />
          <Metric label="Revenue" value={currency(revenue, profile.currency)} />
          <Metric label="P&L cost" value={currency(expenses, profile.currency)} />
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
            {importError && <div className="alert danger">{importError}</div>}
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
            </div>
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
                <strong>{["Methodology", "Raw Trial Balance", "Executive MIS", "Driver Build-Up", "Unit Profitability", "People Allocation", "Shared Cost Allocation", ...template.workbookSheets, "Question Register", "Data Gaps"].join(", ")}</strong>
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
