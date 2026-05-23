import { useRef, useState } from "react";
import { AlertTriangle, BarChart3, Check, ChevronDown, Download, FileSpreadsheet, Landmark, Sparkles, Upload } from "lucide-react";
import { parseBankSourceFile, parseReferenceStructure, parseTrialBalanceWorkbook } from "./lib/parser";
import {
  analyzeWithClaude,
  buildSheetWithClaude,
  revisePlanWithClaude,
  buildWorkbookWithShortcut,
  downloadBlob,
  type AnalyzeResult,
  type BuildSheetResult,
} from "./lib/claudeApi";
import { renderClaudeWorkbook } from "./lib/claudeRenderer";
import type { BankSourceFile, BusinessProfile, ReferenceStructure, Suite, TrialBalanceRow } from "./types";
import { SUITES } from "./types";

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
  fundFlowBasis: "trial-balance-proxy",
  projectionBasis: "past-year",
  profileSource: "default",
};

function inferCadence(start: string, end: string): BusinessProfile["cadence"] {
  if (!start || !end) return "monthly";
  const s = new Date(`${start}T00:00:00`);
  const e = new Date(`${end}T00:00:00`);
  if (Number.isNaN(s.getTime()) || Number.isNaN(e.getTime())) return "monthly";
  const months = (e.getFullYear() - s.getFullYear()) * 12 + e.getMonth() - s.getMonth() + 1;
  if (months >= 10) return "annual";
  if (months >= 3) return "quarterly";
  return "monthly";
}

// ── Upload drop zone ─────────────────────────────────────────────────────────

interface UploadZoneProps {
  label: string;
  icon: React.ReactNode;
  accepts?: string;
  stats: string | null;
  warning?: string;
  onFile: (f: File) => void;
  loading?: boolean;
}

function UploadZone({ label, icon, accepts, stats, warning, onFile, loading }: UploadZoneProps) {
  const ref = useRef<HTMLInputElement>(null);
  function handleDrop(e: React.DragEvent) {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) onFile(file);
  }
  return (
    <div
      className="upload-zone"
      onDragOver={(e) => e.preventDefault()}
      onDrop={handleDrop}
      onClick={() => ref.current?.click()}
      role="button"
      tabIndex={0}
      onKeyDown={(e) => e.key === "Enter" && ref.current?.click()}
    >
      <input ref={ref} type="file" accept={accepts} style={{ display: "none" }} onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); }} />
      <span className="upload-icon">{icon}</span>
      <div className="upload-label">
        {loading ? <span className="upload-hint">Parsing…</span> : (
          <>
            <span className="upload-name">{stats ? stats : label}</span>
            {stats ? <span className="upload-hint upload-ok"><Check size={12} /> Parsed</span> : <span className="upload-hint">drag or click</span>}
          </>
        )}
      </div>
      {warning && <span className="upload-warn"><AlertTriangle size={12} /> {warning}</span>}
    </div>
  );
}

// ── Suite selector ───────────────────────────────────────────────────────────

function SuiteSelector({ value, onChange }: { value: Suite; onChange: (s: Suite) => void }) {
  const [open, setOpen] = useState(false);
  const current = SUITES.find((s) => s.id === value)!;
  return (
    <div className="suite-selector" style={{ position: "relative" }}>
      <button type="button" className="suite-btn" onClick={() => setOpen(!open)}>
        <BarChart3 size={15} />
        <span>{current.label}</span>
        <ChevronDown size={14} />
      </button>
      {open && (
        <div className="suite-dropdown" onMouseLeave={() => setOpen(false)}>
          {SUITES.map((s) => (
            <button
              key={s.id}
              type="button"
              className={`suite-option${s.id === value ? " active" : ""}${!s.availableNow ? " disabled" : ""}`}
              onClick={() => { if (s.availableNow) { onChange(s.id); setOpen(false); } }}
              disabled={!s.availableNow}
            >
              <span className="suite-opt-label">{s.label}</span>
              {!s.availableNow && <span className="suite-soon">Soon</span>}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

// ── Step card ────────────────────────────────────────────────────────────────

function StepCard({ n, title, locked, children }: { n: number; title: string; locked?: boolean; children: React.ReactNode }) {
  return (
    <div className={`step-card${locked ? " step-locked" : ""}`}>
      <div className="step-head">
        <span className="step-num">{n}</span>
        <h2 className="step-title">{title}</h2>
      </div>
      <div className="step-body">{children}</div>
    </div>
  );
}

// ── Main app ─────────────────────────────────────────────────────────────────

export function App() {
  const [suite, setSuite] = useState<Suite>("mis-report");

  // Upload state
  const [rows, setRows] = useState<TrialBalanceRow[]>([]);
  const [profile, setProfile] = useState<BusinessProfile>(initialProfile);
  const [bankSources, setBankSources] = useState<BankSourceFile[]>([]);
  const [tbStats, setTbStats] = useState<string | null>(null);
  const [tbWarning, setTbWarning] = useState("");
  const [bsStats, setBsStats] = useState<string | null>(null);
  const [bsWarning, setBsWarning] = useState("");
  const [blStats, setBlStats] = useState<string | null>(null);
  const [blWarning, setBlWarning] = useState("");
  const [parsing, setParsing] = useState<"tb" | "bs" | "bl" | "ref" | null>(null);
  const [clientBrief, setClientBrief] = useState("");
  const [referenceStructure, setReferenceStructure] = useState<ReferenceStructure | null>(null);
  const [refStats, setRefStats] = useState<string | null>(null);
  const [refWarning, setRefWarning] = useState("");
  const [prefAnswers, setPrefAnswers] = useState<Record<string, string>>({});
  const [questionAnswers, setQuestionAnswers] = useState<Record<string, string>>({});
  const [sheetFeedback, setSheetFeedback] = useState<Record<string, string>>({});
  const [activeSheetFeedback, setActiveSheetFeedback] = useState<string | null>(null);
  const [regeneratingSheet, setRegeneratingSheet] = useState<string | null>(null);

  // Analysis state
  const [analysis, setAnalysis] = useState<AnalyzeResult | null>(null);
  const [analyzeLoading, setAnalyzeLoading] = useState(false);
  const [analyzeError, setAnalyzeError] = useState("");

  // Plan revision state
  const [revisePrompt, setRevisePrompt] = useState("");
  const [reviseLoading, setReviseLoading] = useState(false);
  const [revisionHistory, setRevisionHistory] = useState<Array<{ prompt: string; ack: string }>>([]);
  const [clientInstructions, setClientInstructions] = useState("");

  // Build state
  const [buildStarted, setBuildStarted] = useState(false);
  const [buildLoading, setBuildLoading] = useState(false);
  const [buildError, setBuildError] = useState("");
  const [builtSheets, setBuiltSheets] = useState<BuildSheetResult[]>([]);
  const [buildProgress, setBuildProgress] = useState({ done: 0, total: 0, current: "" });
  const [tokenTotals, setTokenTotals] = useState({ input: 0, output: 0 });
  const [downloadLoading, setDownloadLoading] = useState(false);
  const [chartWarning, setChartWarning] = useState<string | null>(null);
  const [shortcutLoading, setShortcutLoading] = useState(false);
  const [shortcutError, setShortcutError] = useState("");
  const [shortcutStatus, setShortcutStatus] = useState("");

  // ── File handlers ──────────────────────────────────────────────────────────

  async function handleTb(file: File) {
    setParsing("tb");
    setTbWarning("");
    try {
      const parsed = await parseTrialBalanceWorkbook(file);
      setRows(parsed.rows);
      setProfile((p) => ({
        ...p,
        businessName: parsed.metadata.companyName || p.businessName,
        legalEntity: p.legalEntity || parsed.metadata.companyName,
        periodStart: parsed.metadata.periodStart || p.periodStart,
        periodEnd: parsed.metadata.periodEnd || p.periodEnd,
        cadence: parsed.metadata.periodStart && parsed.metadata.periodEnd
          ? inferCadence(parsed.metadata.periodStart, parsed.metadata.periodEnd) : p.cadence,
        profileSource: parsed.metadata.companyName || parsed.metadata.periodStart ? "trial-balance" : p.profileSource,
        tbSourceFileName: parsed.metadata.sourceFileName,
        tbPeriodText: parsed.metadata.periodText,
        tbTotalDebit: parsed.metadata.totalDebit,
        tbTotalCredit: parsed.metadata.totalCredit,
      }));
      setTbStats(`${parsed.rows.length} ledger rows · ${file.name}`);
      if (parsed.warnings.length) setTbWarning(parsed.warnings.join(" "));
    } catch (e) {
      setTbWarning(e instanceof Error ? e.message : "Parse failed");
    } finally {
      setParsing(null);
    }
  }

  async function handleBank(file: File, sourceType: BankSourceFile["sourceType"]) {
    setParsing(sourceType === "bank-statement" ? "bs" : "bl");
    const setStats = sourceType === "bank-statement" ? setBsStats : setBlStats;
    const setWarn  = sourceType === "bank-statement" ? setBsWarning : setBlWarning;
    setWarn("");
    try {
      const parsed = await parseBankSourceFile(file, sourceType);
      setBankSources((prev) => [...prev.filter((s) => s.sourceType !== sourceType), parsed]);
      setStats(`${parsed.rowsImported} transactions · ${file.name}`);
      if (parsed.summary.warnings.length) setWarn(parsed.summary.warnings.join(" "));
    } catch (e) {
      setWarn(e instanceof Error ? e.message : "Parse failed");
    } finally {
      setParsing(null);
    }
  }

  async function handleRef(file: File) {
    setParsing("ref");
    setRefWarning("");
    try {
      const ref = await parseReferenceStructure(file);
      setReferenceStructure(ref);
      setRefStats(`${ref.sheets.length} sheets · ${file.name}`);
    } catch (e) {
      setRefWarning(e instanceof Error ? e.message : "Parse failed");
    } finally {
      setParsing(null);
    }
  }

  async function regenerateSingleSheet(sheetName: string, feedback: string) {
    if (!analysis) return;
    setRegeneratingSheet(sheetName);
    const sheet = analysis.sheetPlan.find((s) => s.name === sheetName);
    if (!sheet) return;
    const allocationChoice = analysis.allocationLogicOptions.find((o) => o.recommended) ?? analysis.allocationLogicOptions[0] ?? null;
    const allSheetNames = analysis.sheetPlan.map((s) => s.name);
    const answers = analysis.clarifyingQuestions.map((q) => ({ id: q.id, question: q.prompt, answer: questionAnswers[q.id] ?? "" }));
    const combinedInstructions = [clientInstructions, feedback].filter(Boolean).join("\n");
    try {
      const r = await buildSheetWithClaude({
        profile, rows, segments: analysis.detectedSegments, allocationChoice, answers, bankSources,
        sheet, allSheetNames,
        clientInstructions: combinedInstructions,
        clientBrief,
        referenceStructure,
      });
      setBuiltSheets((prev) => prev.map((s) => s.sheet.name === sheetName ? r.result : s));
      setSheetFeedback((prev) => { const n = { ...prev }; delete n[sheetName]; return n; });
      setActiveSheetFeedback(null);
    } catch (e) {
      setBuildError(e instanceof Error ? e.message : "Regeneration failed.");
    } finally {
      setRegeneratingSheet(null);
    }
  }

  // ── Analyze (Phase 1) ──────────────────────────────────────────────────────

  async function analyzeFlow() {
    if (!rows.length) { setAnalyzeError("Upload a Trial Balance first."); return; }
    setAnalyzeLoading(true);
    setAnalyzeError("");
    setAnalysis(null);
    setBuiltSheets([]);
    setBuildStarted(false);
    setRevisionHistory([]);
    setClientInstructions("");
    setRevisePrompt("");
    setQuestionAnswers({});
    try {
      const r = await analyzeWithClaude({ profile, rows, bankSources, clientBrief, referenceStructure });
      setAnalysis(r.result);
      if (r.usage) setTokenTotals({ input: r.usage.input_tokens, output: r.usage.output_tokens });
    } catch (e) {
      setAnalyzeError(e instanceof Error ? e.message : "Analysis failed.");
    } finally {
      setAnalyzeLoading(false);
    }
  }

  // ── Revise plan (Phase 1.5) ────────────────────────────────────────────────

  async function revisePlan() {
    if (!analysis || !revisePrompt.trim()) return;
    setReviseLoading(true);
    try {
      const r = await revisePlanWithClaude({ currentPlan: analysis, userPrompt: revisePrompt });
      setRevisionHistory((h) => [...h, { prompt: revisePrompt, ack: r.result.acknowledgement }]);
      setClientInstructions((prev) => prev ? `${prev}\n---\n${revisePrompt}` : revisePrompt);
      setAnalysis(r.result.plan);
      setRevisePrompt("");
    } catch (e) {
      setAnalyzeError(e instanceof Error ? e.message : "Plan revision failed.");
    } finally {
      setReviseLoading(false);
    }
  }

  // ── Build workbook (Phase 2) ───────────────────────────────────────────────

  async function buildWorkbookFlow() {
    if (!analysis) return;
    setBuildStarted(true);
    setBuildLoading(true);
    setBuildError("");
    setBuiltSheets([]);
    setChartWarning(null);
    const allocationChoice = analysis.allocationLogicOptions.find((o) => o.recommended) ?? analysis.allocationLogicOptions[0] ?? null;
    const allSheetNames = analysis.sheetPlan.map((s) => s.name);
    const answers = analysis.clarifyingQuestions.map((q) => ({ id: q.id, question: q.prompt, answer: questionAnswers[q.id] ?? "" }));
    const prefInstructions = (analysis.clientPreferenceQuestions ?? [])
      .map((q) => prefAnswers[q.id] ? `${q.question}: ${prefAnswers[q.id]}` : null)
      .filter(Boolean).join("\n");
    const effectiveInstructions = [clientInstructions, prefInstructions].filter(Boolean).join("\n");
    setBuildProgress({ done: 0, total: analysis.sheetPlan.length, current: "" });
    try {
      const results = await Promise.all(
        analysis.sheetPlan.map(async (sheet) => {
          const r = await buildSheetWithClaude({
            profile, rows, segments: analysis.detectedSegments, allocationChoice, answers,
            bankSources, sheet, allSheetNames,
            clientInstructions: effectiveInstructions, clientBrief, referenceStructure,
          });
          if (r.usage) {
            setTokenTotals((prev) => ({
              input: prev.input + r.usage!.input_tokens,
              output: prev.output + r.usage!.output_tokens,
            }));
          }
          setBuiltSheets((prev) => [...prev, r.result]);
          setBuildProgress((prev) => ({ ...prev, done: prev.done + 1 }));
          return r.result;
        })
      );
      // Lock in the plan-ordered sequence once all sheets arrive
      setBuiltSheets(results);
      setBuildProgress({ done: analysis.sheetPlan.length, total: analysis.sheetPlan.length, current: "" });
    } catch (e) {
      setBuildError(e instanceof Error ? e.message : "Build failed.");
    } finally {
      setBuildLoading(false);
    }
  }

  // ── Download (render + xlsx) ───────────────────────────────────────────────

  async function downloadClaudeWorkbook() {
    if (!analysis || !builtSheets.length) return;
    setDownloadLoading(true);
    setBuildError("");
    try {
      const result = await renderClaudeWorkbook({ businessName: profile.businessName, currency: profile.currency, analyzeResult: analysis, builtSheets, rows, bankSources });
      const safe = (profile.businessName || "mis").replace(/[^a-z0-9]+/gi, "-").toLowerCase();
      downloadBlob(result.blob, `${safe}-mis-claude.xlsx`);
      if (result.chartWarning) setChartWarning(result.chartWarning);
    } catch (e) {
      setBuildError(e instanceof Error ? `Download failed — ${e.message}` : "Download failed.");
    } finally {
      setDownloadLoading(false);
    }
  }

  // ── Shortcut path ──────────────────────────────────────────────────────────

  async function buildWithShortcut() {
    if (!analysis) return;
    setShortcutLoading(true);
    setShortcutError("");
    setShortcutStatus("Starting…");
    try {
      const allocationChoice = analysis.allocationLogicOptions.find((o) => o.recommended) ?? analysis.allocationLogicOptions[0] ?? null;
      const answers = clientInstructions.trim()
        ? [{ id: "client-instructions", question: "Client instructions", answer: clientInstructions }]
        : [];
      setShortcutStatus("Building with Shortcut…");
      const blob = await buildWorkbookWithShortcut({ profile, rows, bankSources, segments: analysis.detectedSegments, allocationChoice, answers, sheetPlan: analysis.sheetPlan });
      const safe = (profile.businessName || "mis").replace(/[^a-z0-9]+/gi, "-").toLowerCase();
      downloadBlob(blob, `${safe}-mis-shortcut.xlsx`);
      setShortcutStatus("Done.");
    } catch (e) {
      setShortcutError(e instanceof Error ? e.message : "Shortcut build failed.");
    } finally {
      setShortcutLoading(false);
      setShortcutStatus("");
    }
  }

  // ── Render ─────────────────────────────────────────────────────────────────

  const pct = buildProgress.total ? Math.round((buildProgress.done / buildProgress.total) * 100) : 0;

  return (
    <div className="mis-shell">

      {/* ── App header ── */}
      <header className="mis-header">
        <div className="mis-brand">
          <BarChart3 size={20} className="mis-brand-icon" />
          <span className="mis-brand-name">Auren</span>
          <span className="mis-brand-tag">Finance Pack Generator</span>
        </div>
        <SuiteSelector value={suite} onChange={setSuite} />
      </header>

      {/* ── Three-step flow ── */}
      <div className="mis-flow">

        {/* ── STEP 1: Upload ── */}
        <StepCard n={1} title="Upload your files">
          <div className="upload-grid">
            <UploadZone
              label="Trial Balance / Ledger"
              icon={<FileSpreadsheet size={22} />}
              accepts=".xlsx,.xls,.csv"
              stats={tbStats}
              warning={tbWarning}
              loading={parsing === "tb"}
              onFile={handleTb}
            />
            <UploadZone
              label="Bank Statement (optional)"
              icon={<Landmark size={22} />}
              accepts=".xlsx,.xls,.csv"
              stats={bsStats}
              warning={bsWarning}
              loading={parsing === "bs"}
              onFile={(f) => handleBank(f, "bank-statement")}
            />
            <UploadZone
              label="Bank Ledger (optional)"
              icon={<Landmark size={22} />}
              accepts=".xlsx,.xls,.csv"
              stats={blStats}
              warning={blWarning}
              loading={parsing === "bl"}
              onFile={(f) => handleBank(f, "bank-ledger")}
            />
            <UploadZone
              label="Reference MIS (optional)"
              icon={<FileSpreadsheet size={22} />}
              accepts=".xlsx,.xls"
              stats={refStats}
              warning={refWarning}
              loading={parsing === "ref"}
              onFile={handleRef}
            />
          </div>

          <textarea
            className="client-brief-box"
            rows={2}
            placeholder="Tell Claude what your client cares about: industry, how they review numbers, key KPIs, number scale. e.g. 'Series B SaaS, board tracks ARR and burn, numbers in lakhs, monthly columns preferred.'"
            value={clientBrief}
            onChange={(e) => setClientBrief(e.target.value)}
          />

          {profile.businessName && (
            <div className="tb-meta">
              <strong>{profile.businessName}</strong>
              {profile.periodStart && profile.periodEnd && <span> · {profile.periodStart} to {profile.periodEnd}</span>}
              {profile.currency && <span> · {profile.currency}</span>}
            </div>
          )}

          <div className="step-actions">
            <button
              className="btn-primary"
              type="button"
              onClick={analyzeFlow}
              disabled={analyzeLoading || !rows.length}
            >
              <Sparkles size={16} />
              {analyzeLoading ? "Analysing…" : analysis ? "Re-analyse" : "Analyse with Claude"}
            </button>
            {analyzeError && <p className="error-line"><AlertTriangle size={14} /> {analyzeError}</p>}
          </div>
        </StepCard>

        {/* ── STEP 2: Plan Review ── */}
        <StepCard n={2} title="Review & refine the plan" locked={!analysis}>
          {!analysis ? (
            <p className="step-placeholder">Upload files and click Analyse to see the plan here.</p>
          ) : (
            <>
              <div className="plan-summary">
                <span className="plan-pill plan-pill-type">{analysis.businessType}</span>
                <span className="plan-pill">{analysis.detectedSegments.length} segments</span>
                <span className="plan-pill">{analysis.sheetPlan.length} sheets</span>
                {profile.periodStart && <span className="plan-pill">{profile.periodStart} → {profile.periodEnd}</span>}
              </div>

              <div className="plan-sheet-list">
                {analysis.sheetPlan.map((s, i) => (
                  <div key={i} className="plan-sheet-row">
                    {s.tabColor && <span className="tab-swatch" style={{ background: `#${s.tabColor}` }} />}
                    <span className="plan-sheet-name">{s.name}</span>
                    <span className="plan-sheet-purpose">{s.purpose}</span>
                  </div>
                ))}
              </div>

              {analysis.clarifyingQuestions.length > 0 && (
                <div className="claude-qs">
                  <p className="claude-qs-label">Answer Claude's questions:</p>
                  <div className="claude-qa-list">
                    {analysis.clarifyingQuestions.map((q) => (
                      <div key={q.id} className="claude-qa-row">
                        <label className="claude-qa-label">
                          <span className={`claude-qa-priority claude-qa-priority--${q.priority}`}>{q.priority}</span>
                          {q.prompt}
                        </label>
                        {q.answerKind === "choice" && q.choices?.length ? (
                          <select
                            className="pref-select"
                            value={questionAnswers[q.id] ?? ""}
                            onChange={(e) => setQuestionAnswers((prev) => ({ ...prev, [q.id]: e.target.value }))}
                          >
                            <option value="">— choose —</option>
                            {q.choices.map((c) => <option key={c} value={c}>{c}</option>)}
                          </select>
                        ) : (
                          <textarea
                            className="claude-qa-input"
                            rows={2}
                            placeholder={q.answerKind === "number" ? "Enter a number…" : "Type your answer…"}
                            value={questionAnswers[q.id] ?? ""}
                            onChange={(e) => setQuestionAnswers((prev) => ({ ...prev, [q.id]: e.target.value }))}
                          />
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {(analysis.clientPreferenceQuestions ?? []).length > 0 && (
                <div className="pref-questions">
                  <p className="claude-qs-label">Client presentation preferences:</p>
                  <div className="pref-questions-grid">
                    {(analysis.clientPreferenceQuestions ?? []).map((q) => (
                      <div key={q.id} className="pref-question-row">
                        <label className="pref-q-label">{q.question}</label>
                        {q.answerKind === "choice" && q.choices?.length ? (
                          <select
                            className="pref-select"
                            value={prefAnswers[q.id] ?? q.defaultAnswer ?? ""}
                            onChange={(e) => setPrefAnswers((prev) => ({ ...prev, [q.id]: e.target.value }))}
                          >
                            <option value="">— choose —</option>
                            {q.choices.map((c) => <option key={c} value={c}>{c}</option>)}
                          </select>
                        ) : (
                          <input
                            type="text"
                            className="pref-input"
                            placeholder={q.defaultAnswer ?? ""}
                            value={prefAnswers[q.id] ?? ""}
                            onChange={(e) => setPrefAnswers((prev) => ({ ...prev, [q.id]: e.target.value }))}
                          />
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {revisionHistory.length > 0 && (
                <div className="revision-history">
                  {revisionHistory.map((r, i) => (
                    <div key={i} className="revision-entry">
                      <p className="revision-prompt">You: {r.prompt}</p>
                      <p className="revision-ack">Claude: {r.ack}</p>
                    </div>
                  ))}
                </div>
              )}

              <div className="prompt-area">
                <textarea
                  className="prompt-box"
                  rows={3}
                  placeholder="Type changes or answer Claude's questions… e.g. 'Show amounts in lakhs. Remove the hostel sheet. Faculty headcount is 120.'"
                  value={revisePrompt}
                  onChange={(e) => setRevisePrompt(e.target.value)}
                  onKeyDown={(e) => { if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) revisePlan(); }}
                />
                <div className="prompt-actions">
                  <button className="btn-secondary" type="button" onClick={revisePlan} disabled={reviseLoading || !revisePrompt.trim()}>
                    {reviseLoading ? "Revising…" : "Revise Plan"}
                  </button>
                  <button className="btn-primary" type="button" onClick={buildWorkbookFlow} disabled={buildLoading || shortcutLoading}>
                    <Sparkles size={15} />
                    {buildLoading ? `Building ${buildProgress.done}/${buildProgress.total}…` : "Build with Claude"}
                  </button>
                  <button className="btn-secondary" type="button" onClick={buildWithShortcut} disabled={shortcutLoading || buildLoading}>
                    {shortcutLoading ? shortcutStatus || "Shortcut building…" : "Build with Shortcut"}
                  </button>
                </div>
              </div>
            </>
          )}
        </StepCard>

        {/* ── STEP 3: Build ── */}
        <StepCard n={3} title="Build & download" locked={!buildStarted}>
          {!buildStarted ? (
            <p className="step-placeholder">Click "Build with Claude" above to start.</p>
          ) : (
            <>
              {buildLoading && (
                <div className="build-progress">
                  <div className="progress-bar"><div className="progress-fill" style={{ width: `${pct}%` }} /></div>
                  <p className="progress-label">{buildProgress.done === 0 ? "Starting…" : `${buildProgress.done} of ${buildProgress.total} sheets ready`}</p>
                </div>
              )}

              {builtSheets.length > 0 && (
                <div className="sheet-list">
                  {builtSheets.map((s, i) => {
                    const isActive = activeSheetFeedback === s.sheet.name;
                    const isRegen = regeneratingSheet === s.sheet.name;
                    return (
                      <div key={i} className="sheet-row-wrap">
                        <div className="sheet-row">
                          <Check size={14} className="sheet-check" />
                          <span className="sheet-row-name">{s.sheet.name}</span>
                          <button
                            type="button"
                            className="btn-regen"
                            onClick={() => setActiveSheetFeedback(isActive ? null : s.sheet.name)}
                            disabled={isRegen || buildLoading}
                          >
                            {isRegen ? "Rebuilding…" : "Regenerate"}
                          </button>
                        </div>
                        {isActive && (
                          <div className="regen-panel">
                            <textarea
                              className="regen-box"
                              rows={2}
                              placeholder={`What should change in "${s.sheet.name}"? e.g. "Add variance vs budget column"`}
                              value={sheetFeedback[s.sheet.name] ?? ""}
                              onChange={(e) => setSheetFeedback((prev) => ({ ...prev, [s.sheet.name]: e.target.value }))}
                            />
                            <button
                              type="button"
                              className="btn-secondary"
                              disabled={isRegen || !sheetFeedback[s.sheet.name]?.trim()}
                              onClick={() => regenerateSingleSheet(s.sheet.name, sheetFeedback[s.sheet.name] ?? "")}
                            >
                              {isRegen ? "Rebuilding…" : "Rebuild this sheet"}
                            </button>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}

              {buildError && <p className="error-line"><AlertTriangle size={14} /> {buildError}</p>}
              {shortcutError && <p className="error-line"><AlertTriangle size={14} /> {shortcutError}</p>}

              {builtSheets.length > 0 && !buildLoading && (
                <div className="download-area">
                  <button className="btn-download" type="button" onClick={downloadClaudeWorkbook} disabled={downloadLoading}>
                    <Download size={16} />
                    {downloadLoading ? "Preparing…" : "Download MIS Report (.xlsx)"}
                  </button>
                  {chartWarning && <p className="chart-warn"><AlertTriangle size={13} /> {chartWarning}</p>}
                </div>
              )}

              {tokenTotals.input > 0 && (
                <p className="token-usage">
                  Tokens used: {(tokenTotals.input / 1000).toFixed(1)}K in · {(tokenTotals.output / 1000).toFixed(1)}K out
                </p>
              )}
            </>
          )}
        </StepCard>
      </div>
    </div>
  );
}
