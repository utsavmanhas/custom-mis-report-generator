// Renders Claude's per-sheet specs into a single .xlsx workbook.
// Stays "dumb on purpose" - all design decisions live in the spec,
// this file just translates faithfully into ExcelJS calls.

import ExcelJS from "exceljs";
import type { SheetSpec, BuildSheetResult, AnalyzeResult } from "./claudeApi";
import type { TrialBalanceRow, BankSourceFile } from "../types";
import { buildChartData } from "./financialModel";

function nfString(code: string | undefined, currency: string) {
  if (code === "currency") return `${currency} #,##,##0;[Red](${currency} #,##,##0);-`;
  if (code === "percent") return "0.0%";
  if (code === "integer") return "#,##0";
  if (code === "number") return "#,##,##0.00";
  if (code === "date") return "dd-mmm-yyyy";
  return undefined;
}

function argb(value: unknown, fallback = "FFFFFF") {
  const raw = String(value ?? "").replace("#", "").trim().toUpperCase();
  if (/^[0-9A-F]{6}$/.test(raw)) return `FF${raw}`;
  if (/^[0-9A-F]{8}$/.test(raw)) return raw;
  return `FF${fallback}`;
}

function safeSheetName(name: string, used: Set<string>) {
  const base = (name || "Sheet").replace(/[\\/?*[\]:]/g, " ").slice(0, 31).trim() || "Sheet";
  let candidate = base;
  let index = 2;
  while (used.has(candidate.toLowerCase())) {
    const suffix = ` ${index}`;
    candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`.trim();
    index++;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

function applyFormat(cell: ExcelJS.Cell, fmt: any, currency: string) {
  if (!fmt) return;
  const font: Partial<ExcelJS.Font> = {};
  if (fmt.bold) font.bold = true;
  if (fmt.italic) font.italic = true;
  if (fmt.fontSize) font.size = fmt.fontSize;
  if (fmt.color) font.color = { argb: argb(fmt.color, "1F3864") };
  if (Object.keys(font).length) cell.font = font;
  if (fmt.fill) {
    cell.fill = {
      type: "pattern", pattern: "solid",
      fgColor: { argb: argb(fmt.fill, "FFFFFF") },
    };
  }
  if (fmt.align || fmt.indent !== undefined) {
    cell.alignment = {
      horizontal: fmt.align ?? "left",
      vertical: "middle",
      wrapText: true,
      indent: fmt.indent ?? 0,
    };
  }
  const nf = nfString(fmt.numberFormat, currency);
  if (nf) cell.numFmt = nf;
  if (fmt.border && fmt.border !== "none") {
    const side = { style: fmt.border } as ExcelJS.Border;
    cell.border = { top: side, bottom: side, left: side, right: side };
  }
}

// ─── Chart rendering via HTML5 Canvas ───────────────────────────────────────

const PALETTE_BLUE = [
  "#1F3864", "#2E5CA5", "#4472C4", "#70A0D4",
  "#9DC3E6", "#BDD7EE", "#D0E8F8", "#E8F4FF",
];
const PALETTE_RED = [
  "#C00000", "#D93030", "#E85050", "#F07070",
  "#F8A0A0", "#FFCCCC", "#FFE0E0", "#FFF0F0",
];

function shortAmount(val: number, currency: string): string {
  const abs = Math.abs(val);
  if (abs >= 10_000_000) return `${currency} ${(abs / 10_000_000).toFixed(1)}Cr`;
  if (abs >= 100_000) return `${currency} ${(abs / 100_000).toFixed(1)}L`;
  if (abs >= 1_000) return `${currency} ${(abs / 1_000).toFixed(0)}K`;
  return `${currency} ${Math.round(abs).toLocaleString("en-IN")}`;
}

function drawDonutChart(
  ctx: CanvasRenderingContext2D,
  w: number,
  h: number,
  title: string,
  rawItems: { label: string; value: number }[],
  palette: string[],
  currency: string,
): void {
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 0, w, h);

  // Title
  ctx.fillStyle = "#1F3864";
  ctx.font = "bold 13px Arial, sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "top";
  ctx.fillText(title, w / 2, 10);

  const positive = rawItems.filter((x) => x.value > 0).sort((a, b) => b.value - a.value);
  const top = positive.slice(0, 6);
  const othersVal = positive.slice(6).reduce((s, x) => s + x.value, 0);
  const items = othersVal > 0 ? [...top, { label: "Others", value: othersVal }] : top;
  const total = items.reduce((s, x) => s + x.value, 0);

  if (!total) {
    ctx.fillStyle = "#888888";
    ctx.font = "11px Arial, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText("No data available", w / 2, h / 2);
    return;
  }

  const cx = w * 0.36;
  const cy = h * 0.56;
  const outerR = Math.min(cx - 4, cy - 36) * 0.9;
  const innerR = outerR * 0.52;

  let angle = -Math.PI / 2;
  for (let i = 0; i < items.length; i++) {
    const slice = (items[i].value / total) * 2 * Math.PI;
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, outerR, angle, angle + slice);
    ctx.closePath();
    ctx.fillStyle = palette[i % palette.length];
    ctx.fill();
    ctx.strokeStyle = "#FFFFFF";
    ctx.lineWidth = 2;
    ctx.stroke();
    angle += slice;
  }

  // Hole
  ctx.beginPath();
  ctx.arc(cx, cy, innerR, 0, 2 * Math.PI);
  ctx.fillStyle = "#FFFFFF";
  ctx.fill();

  // Center label
  ctx.fillStyle = "#1F3864";
  ctx.font = "bold 9px Arial, sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText("TOTAL", cx, cy - 9);
  ctx.font = "bold 10px Arial, sans-serif";
  ctx.fillText(shortAmount(total, currency), cx, cy + 8);
  ctx.textBaseline = "alphabetic";

  // Legend (right side)
  const legX = w * 0.66;
  const legY0 = 34;
  const rowH = Math.min(24, (h - legY0 - 8) / Math.max(items.length, 1));

  for (let i = 0; i < items.length; i++) {
    const y = legY0 + i * rowH;
    ctx.fillStyle = palette[i % palette.length];
    ctx.fillRect(legX, y, 11, 11);
    ctx.fillStyle = "#333333";
    ctx.font = "9px Arial, sans-serif";
    ctx.textAlign = "left";
    ctx.textBaseline = "top";
    const pct = ((items[i].value / total) * 100).toFixed(1);
    const lbl = items[i].label.length > 20 ? `${items[i].label.slice(0, 19)}…` : items[i].label;
    ctx.fillText(`${lbl}  ${pct}%`, legX + 14, y + 1);
  }
}

function drawBarChart(
  ctx: CanvasRenderingContext2D,
  w: number,
  h: number,
  title: string,
  bars: { label: string; value: number; color: string }[],
  currency: string,
): void {
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 0, w, h);

  ctx.fillStyle = "#1F3864";
  ctx.font = "bold 13px Arial, sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "top";
  ctx.fillText(title, w / 2, 10);

  if (!bars.length) return;
  const maxVal = Math.max(...bars.map((b) => Math.abs(b.value)));
  if (!maxVal) return;

  const pad = { top: 40, bottom: 46, left: 20, right: 20 };
  const chartW = w - pad.left - pad.right;
  const chartH = h - pad.top - pad.bottom;
  const barW = Math.min(70, (chartW / bars.length) * 0.55);
  const slotW = chartW / bars.length;

  bars.forEach((bar, i) => {
    const barH = (Math.abs(bar.value) / maxVal) * chartH;
    const x = pad.left + slotW * i + (slotW - barW) / 2;
    const y = pad.top + chartH - barH;

    // Bar shadow
    ctx.fillStyle = "rgba(0,0,0,0.06)";
    ctx.fillRect(x + 3, y + 3, barW, barH);

    ctx.fillStyle = bar.color;
    ctx.fillRect(x, y, barW, barH);

    // Value label above bar
    ctx.fillStyle = "#1F3864";
    ctx.font = "bold 9px Arial, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "bottom";
    ctx.fillText(shortAmount(bar.value, currency), x + barW / 2, y - 3);

    // X-axis label
    ctx.fillStyle = "#555555";
    ctx.font = "9px Arial, sans-serif";
    ctx.textBaseline = "top";
    ctx.fillText(bar.label, x + barW / 2, pad.top + chartH + 8);
  });

  // Baseline
  ctx.beginPath();
  ctx.moveTo(pad.left, pad.top + chartH);
  ctx.lineTo(w - pad.right, pad.top + chartH);
  ctx.strokeStyle = "#CCCCCC";
  ctx.lineWidth = 1;
  ctx.stroke();
}

function canvasToBase64(canvas: HTMLCanvasElement): string {
  return canvas.toDataURL("image/png").split(",")[1];
}

async function embedDashboardCharts(
  wb: ExcelJS.Workbook,
  rows: TrialBalanceRow[],
  bankSources: BankSourceFile[],
  currency: string,
): Promise<string | null> {
  const dashWs = wb.worksheets.find((ws) => {
    const n = ws.name.toLowerCase();
    return n.includes("dashboard") || n.includes("cover") || n.includes("executive");
  }) ?? wb.worksheets[0];
  if (!dashWs) return "No sheets were built — charts not embedded.";

  const cd = buildChartData(rows, bankSources);

  // Place charts 3 rows below the last data row (min row 50 to clear typical dashboards)
  const lastDataRow = dashWs.lastRow?.number ?? 10;
  const chartTopRow = Math.max(lastDataRow + 3, 50);

  // Pixel dimensions for each chart type
  const CW = 440; // chart width
  const DH = 290; // donut height
  const BH = 250; // bar height

  // Approximate rows consumed by each chart height (Excel default row = ~15pt ≈ 20px)
  const PX_PER_ROW = 20;
  const donutRowSpan = Math.ceil(DH / PX_PER_ROW) + 1;
  const barTopRow = chartTopRow + donutRowSpan + 2;

  const canvas = document.createElement("canvas");

  // ── Revenue Mix donut (col A) ────────────────────────────────────────────
  if (cd.revenueMix.some((x) => x.value > 0)) {
    canvas.width = CW;
    canvas.height = DH;
    const ctx = canvas.getContext("2d")!;
    ctx.clearRect(0, 0, CW, DH);
    drawDonutChart(ctx, CW, DH, "Revenue Mix", cd.revenueMix, PALETTE_BLUE, currency);
    const id = wb.addImage({ base64: canvasToBase64(canvas), extension: "png" });
    dashWs.addImage(id, {
      tl: { col: 0, row: chartTopRow - 1 },
      ext: { width: CW, height: DH },
    });
  }

  // ── Expense Breakdown donut (col H) ─────────────────────────────────────
  if (cd.expenseMix.some((x) => x.value > 0)) {
    canvas.width = CW;
    canvas.height = DH;
    const ctx = canvas.getContext("2d")!;
    ctx.clearRect(0, 0, CW, DH);
    drawDonutChart(ctx, CW, DH, "Expense Breakdown", cd.expenseMix, PALETTE_RED, currency);
    const id = wb.addImage({ base64: canvasToBase64(canvas), extension: "png" });
    dashWs.addImage(id, {
      tl: { col: 7, row: chartTopRow - 1 },
      ext: { width: CW, height: DH },
    });
  }

  // ── P&L Summary bar (col A, below donuts) ───────────────────────────────
  const totalRevenue = cd.revenueMix.reduce((s, x) => s + x.value, 0);
  const totalExpenses = cd.expenseMix.reduce((s, x) => s + x.value, 0);
  const netPl = totalRevenue - totalExpenses;
  if (totalRevenue > 0 || totalExpenses > 0) {
    canvas.width = CW;
    canvas.height = BH;
    const ctx = canvas.getContext("2d")!;
    ctx.clearRect(0, 0, CW, BH);
    drawBarChart(ctx, CW, BH, "P&L Summary", [
      { label: "Revenue", value: totalRevenue, color: "#00B050" },
      { label: "Total Cost", value: totalExpenses, color: "#C00000" },
      {
        label: netPl >= 0 ? "Surplus" : "Deficit",
        value: Math.abs(netPl),
        color: netPl >= 0 ? "#1F3864" : "#FF6B6B",
      },
    ], currency);
    const id = wb.addImage({ base64: canvasToBase64(canvas), extension: "png" });
    dashWs.addImage(id, {
      tl: { col: 0, row: barTopRow - 1 },
      ext: { width: CW, height: BH },
    });
  }

  // ── Balance Sheet Snapshot bar (col H, below donuts) ────────────────────
  const totalAssets = cd.assetComposition.reduce((s, x) => s + x.value, 0);
  const totalLiabilities = cd.liabilityComposition.reduce((s, x) => s + x.value, 0);
  if (totalAssets > 0 || totalLiabilities > 0) {
    canvas.width = CW;
    canvas.height = BH;
    const ctx = canvas.getContext("2d")!;
    ctx.clearRect(0, 0, CW, BH);
    drawBarChart(ctx, CW, BH, "Balance Sheet Snapshot", [
      { label: "Total Assets", value: totalAssets, color: "#4472C4" },
      { label: "Total Liab.", value: totalLiabilities, color: "#FFC000" },
    ], currency);
    const id = wb.addImage({ base64: canvasToBase64(canvas), extension: "png" });
    dashWs.addImage(id, {
      tl: { col: 7, row: barTopRow - 1 },
      ext: { width: CW, height: BH },
    });
  }
  const hasAnyChart = cd.revenueMix.some((x) => x.value > 0) || cd.expenseMix.some((x) => x.value > 0) || totalRevenue > 0 || totalAssets > 0;
  return hasAnyChart ? null : "Charts not embedded — no classifiable revenue or expense rows were found in the trial balance.";
}

// ─── Public interface ────────────────────────────────────────────────────────

export interface RenderOpts {
  businessName: string;
  currency: string;
  analyzeResult: AnalyzeResult;
  builtSheets: BuildSheetResult[];
  notes?: string[];
  rows?: TrialBalanceRow[];
  bankSources?: BankSourceFile[];
}

export interface RenderResult {
  blob: Blob;
  chartWarning: string | null;
}

export async function renderClaudeWorkbook(opts: RenderOpts): Promise<RenderResult> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Claude-powered MIS Generator";
  wb.created = new Date();
  const usedNames = new Set<string>();

  for (const built of opts.builtSheets) {
    addSheet(wb, built.sheet, opts.currency, usedNames);
  }

  // Embed charts into the Dashboard/Cover sheet when TB rows are available
  let chartWarning: string | null = null;
  if (opts.rows?.length) {
    chartWarning = await embedDashboardCharts(wb, opts.rows, opts.bankSources ?? [], opts.currency);
  }

  // Synthetic "Notes & Adjustments" sheet aggregating Claude's notes from
  // every sheet plus the analysis's open questions.
  const notes = wb.addWorksheet(safeSheetName("Notes & Adjustments", usedNames), {
    properties: { tabColor: { argb: "FFAAAAAA" } },
  });
  notes.columns = [
    { header: "Section", width: 28 },
    { header: "Detail", width: 100 },
  ];
  notes.getRow(1).font = { bold: true };

  notes.addRow(["Business type", `${opts.analyzeResult.businessType} - ${opts.analyzeResult.businessTypeReasoning}`]);
  notes.addRow([]);
  notes.addRow(["Detected segments", ""]);
  for (const s of opts.analyzeResult.detectedSegments) {
    notes.addRow([s.name, `${s.kind} - ${s.rationale}`]);
  }
  notes.addRow([]);
  notes.addRow(["Per-sheet notes from Claude", ""]);
  for (const built of opts.builtSheets) {
    for (const n of built.notes ?? []) {
      notes.addRow([built.sheet.name, n]);
    }
  }
  if (opts.analyzeResult.unclassifiedAccounts?.length) {
    notes.addRow([]);
    notes.addRow(["Unclassified accounts", opts.analyzeResult.unclassifiedAccounts.join(", ")]);
  }
  if (opts.analyzeResult.clarifyingQuestions.length) {
    notes.addRow([]);
    notes.addRow(["Open clarifying questions", ""]);
    for (const q of opts.analyzeResult.clarifyingQuestions) {
      notes.addRow([`${q.priority.toUpperCase()} - ${q.section}`, `${q.prompt} (${q.reason})`]);
    }
  }

  const buffer = await wb.xlsx.writeBuffer();
  return {
    blob: new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    chartWarning,
  };
}

function addSheet(wb: ExcelJS.Workbook, spec: SheetSpec, currency: string, usedNames: Set<string>) {
  const safeName = safeSheetName(spec.name, usedNames);
  const sheet = wb.addWorksheet(safeName, {
    properties: spec.tabColor
      ? { tabColor: { argb: argb(spec.tabColor, "AAAAAA") } }
      : undefined,
  });

  if (spec.columnWidths?.length) {
    sheet.columns = spec.columnWidths.map((w) => ({ width: Math.min(Math.max(Number(w) || 12, 6), 60) }));
  }

  (Array.isArray(spec.rows) ? spec.rows : []).forEach((rowCells, rowIdx) => {
    const row = sheet.getRow(rowIdx + 1);
    (Array.isArray(rowCells) ? rowCells : []).forEach((c, colIdx) => {
      const cell = row.getCell(colIdx + 1);
      if (c.formula) {
        cell.value = { formula: c.formula.replace(/^=/, ""), date1904: false };
      } else if (c.value !== undefined && c.value !== null) {
        cell.value = c.value as ExcelJS.CellValue;
      }
      applyFormat(cell, c.format, currency);
    });
    row.commit();
  });

  if (spec.merges?.length) {
    for (const m of spec.merges) {
      try { sheet.mergeCells(m); } catch { /* ignore */ }
    }
  }

  if (spec.freezePanes) {
    try {
      const m = spec.freezePanes.match(/^([A-Z]+)(\d+)$/);
      if (m) {
        const col = m[1];
        const row = Number(m[2]);
        const colNum = col.split("").reduce((a, ch) => a * 26 + (ch.charCodeAt(0) - 64), 0);
        sheet.views = [{
          state: "frozen",
          xSplit: colNum - 1,
          ySplit: row - 1,
        }];
      }
    } catch { /* ignore */ }
  }
}
