# Custom MIS Report Generator - Claude Edition

A web app that turns a Trial Balance into a client-ready Management Information
System (MIS) workbook by piping the data through **Claude Opus 4.7**.

## What's different from Codex's original build

The original used hard-coded categorisation rules and per-industry templates.
This version:

1. **Detects the business type** from account names (no preset list)
2. **Detects operating segments** (faculties, departments, projects, SKUs) from the data
3. **Asks clarifying questions** in a wizard before building - covers things only
   the human can answer (cost-allocation logic, period dates, headcount per segment, etc.)
4. **Builds each sheet individually** - one Claude call per sheet, focused, parallelisable
5. **Speaks the universal MIS design language** distilled from a real CA-firm
   output (Shoolini University 24-sheet workbook):
   - Monthly time-series columns + YTD-current + YTD-prior
   - Profit-center spine: each segment gets paired (amount, %) columns
   - Cross-sheet formulas (`='Faculty Wise Fees'!L27`)
   - Driver-based allocation (`=L68/L53*B53` distributes provision by student count)
   - Annualisation formulas (`=annual/12*period_months`)
   - BACK navigation cells, tab color coding, freeze panes
   - Notes & Adjustments sheet documenting every override

The legacy hard-coded path stays wired up as the **Legacy export** button so the
app keeps working even when no API key is set.

---

## Architecture

```
[Browser]                    [Server (/api on Vercel or local Express)]
   |
   1. Upload TB
   2. Parse client-side
   |
   3. POST /api/analyze ----> Claude (Phase 1)
                              returns: businessType, segments, sheet plan,
                                       allocation options, clarifying questions
   |
   4. Wizard UI: user answers questions, picks allocation logic
   |
   5. POST /api/build-sheet ---> Claude (Phase 2, called once per sheet)
   .  (one call per sheet)        returns: ClaudeSheetSpec for that sheet
   .                              (cells + formulas + formatting)
   .
   6. Frontend renders all SheetSpecs into one .xlsx via ExcelJS
   7. Download
```

The two-phase split means: Phase 1 is cheap (~3K out tokens) and lets the user
see Claude's interpretation before spending tokens on workbook generation.
Phase 2 is per-sheet, so 8 sheets at ~3-5K out each is reliable - much more
than asking for the whole workbook in one shot.

---

## Quick start (local)

You need **Node 18+** and an Anthropic API key.

```cmd
cd custom-mis-report-generator

REM 1. Install
npm install

REM 2. Create .env.local with your key
copy .env.example .env.local
notepad .env.local
REM   paste your key after  ANTHROPIC_API_KEY=
REM   save, close

REM 3. Run
npm run dev
```

Open http://localhost:3000. Upload a Trial Balance, click **Analyse with Claude**,
answer the questions, then click **Build workbook**.

Or just double-click **start.bat** for a one-click setup + launch.

---

## Deploy to Vercel

```cmd
npm install -g vercel
vercel login
vercel
vercel env add ANTHROPIC_API_KEY production
vercel --prod
```

Vercel auto-detects the `api/*.ts` files as serverless functions.

To swap models for cheaper runs while iterating:
```cmd
vercel env add ANTHROPIC_MODEL  production
REM enter:  claude-sonnet-4-6
```

---

## File map

```
prompts/design-language.md     The universal MIS design rules Claude reads each call
api/analyze.ts                 Phase 1 - Vercel function
api/build-sheet.ts             Phase 2 - Vercel function (one call per sheet)
dev-server.mjs                 Local Express + Vite dev server (no Vercel CLI needed)
src/lib/claudeApi.ts           Frontend client for both phases
src/lib/claudeRenderer.ts      ExcelJS renderer for SheetSpecs
src/lib/parser.ts              TB parser (unchanged from Codex)
src/lib/classification.ts      (kept) hard-coded fallback rules
src/lib/customReport.ts        (kept) legacy template-driven builder
src/App.tsx                    Single-page UI with the wizard
.env.example                   Where the API key goes
vercel.json                    Function timeout config
start.bat                      Windows one-click launcher
```

---

## Cost guidance

A typical run on a 35-row consulting TB:
- Phase 1: ~5K input + ~3K output = ~$0.05 at Opus
- Phase 2: 8 sheets * (~6K in + ~4K out each) = ~$0.50 at Opus

Total ~$0.55 per report. Switch to Sonnet via `ANTHROPIC_MODEL=claude-sonnet-4-6`
to drop that to ~$0.06 per report while iterating.

---

## Roadmap

- [ ] Stream sheet builds so the user sees sheets appear progressively
- [ ] Cache Phase-1 results by TB hash so re-runs skip analysis
- [ ] Multi-period support (current + prior TB upload -> filled YoY columns)
- [ ] Re-run a single sheet (when one sheet looks wrong, regenerate just that one)
- [ ] Add sheet templates as few-shot examples in the design language
