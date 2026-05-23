# MIS Workbook Design Language

This document is loaded into every Claude prompt as ground truth for what
"a good MIS report" looks like. It is distilled from a real CA-firm output
for a university client and generalised to any business type.

## 1. Time-series spine
Every financial sheet shows monthly columns from period start to period end,
followed by a YTD-current column and a YTD-prior column for YoY comparison.
Single-period TBs show a single column with a YoY column.

## 2. Profit-center spine
The main P&L (Income & Expenditure Account) has paired columns per segment:
amount + % of that segment's revenue. Always include a Total column pair and
a Prior-Year Total column pair.

  | Particulars | Seg A amt | Seg A % | Seg B amt | Seg B % | ... | Total amt | Total % | PY Total |

## 3. Cross-sheet formulas
The main I&E pulls from segment sheets via formulas like
`='Faculty Wise Fees'!L27`. Segment sheets pull from supporting
schedules (e.g. salary). Never hard-code a number that should be live.

## 4. Hierarchical drill-downs
Section header (bold) -> sub-totals (=SUM) -> indented line items.
Indentation done with leading spaces in the Particulars column
(" - Tuition Fees", "  - Marketing").

## 5. Annualisation formulas
For partial-year data: `=annual_amount/12*months_in_period`. Document the
months count and source year in the Notes sheet.

## 6. Driver-based allocation
Shared costs are allocated to segments using a measurable driver:
- Provisions, building maintenance -> by student count / headcount / sqft
- Salary of shared staff -> by FTE assignment %
- Hostel cost -> by occupancy nights
- Marketing -> by revenue weight
The formula must reference the driver cell, e.g. `=L68/L53*B53` allocates
total provision (L68) to segment by ratio of segment students (B53) to
total students (L53). Never use arbitrary equal splits unless explicitly
asked for.

## 7. Benchmark vs actual
Where the firm has a target metric (e.g. required gross profit %, target
utilisation %), include a row for the benchmark, a row for actual, and
a row for excess/shortfall. Make the shortfall conditionally red.

## 8. Navigation
Every drill-down sheet has a BACK cell in the top-right area:
- value: "BACK"
- format: bold, white text on red fill (#C00000)
- merged cell, bordered
This visually signals it's a drill-down. (The actual hyperlink can be added
manually later; the visual cue is what matters.)

## 9. Tab color coding
- `#C00000` (red): main I&E / P&L / Surplus
- `#FFC000` (orange): Cash Flow, Bank Reconciliation
- `#00B050` (green): Salary schedules
- `#1F3864` (navy): Dashboard / Cover
- (none): supporting schedules, drill-downs

## 10. Mandatory sheet roster (in order)
Every workbook must include, when applicable:
1. Dashboard / Cover (navy) - KPI tiles + headline
2. Income & Expenditure Account (red) - segment-wise consolidated P&L
3. Cash Flow Statement (orange) - if cash data is meaningful
4. One drill-down sheet per detected segment
5. One Salary schedule (green) - employee-level if data permits
6. One Bank Reconciliation per bank account (orange) - if bank data exists
7. Supporting schedules: Sundry Creditors, Employee Advances, Cash Expenses,
   Asset Purchases, Other Receipts (only if data implies them)
8. Notes & Adjustments - always last

Don't pad with empty sheets. 6-15 sheets is normal.

## 11. Provisions and funds (entity-specific examples)
- University: Security refund (1/3 of closing student deposits), IT/equipment
  fund (25% of fixed assets ex-building), Building refresh (5% of construction)
- SaaS: Deferred revenue release, churn provision
- D2C: Returns provision (avg return rate * forward sales)
- Manufacturing: Warranty provision, slow-moving inventory writedown
- Trust: Corpus accretion, restricted-grant carryover
Pick what's defensible for the business type; document the formula in Notes.

## 12. Notes sheet content
Every override or assumption gets one row:
S.No | Particulars (employee/account/policy) | Adjustment description
e.g. "50% of X's salary allocated to VC office", "Hostel deposit refund
provisioned at 33% of closing balance".

## 13. Formula conventions
- Use absolute references for benchmarks: `=$L$60`
- Use SUMIF for category subtotals when categories are inline
- Use INDIRECT only when a sheet name is dynamic (rare)
- Column totals at bottom: `=SUM(B7:B25)` not `=B7+B8+B9...`

## 14. Indian number formatting
Use `#,##,##0;[Red](#,##,##0);-` for INR amounts (lakhs/crores comma style).
Use `0.0%` for percentages. Date format `DD-MMM-YYYY`.

## 15. Print layout (nice-to-have)
- Freeze panes at first data row + first label column
- Print area covers the data block
- Page setup: A4 landscape, fit-to-1-page-wide
- Header: business name + period
- Footer: page numbers + "CONFIDENTIAL"
