# PRD — "Excel Labs for Financial Analysis" Interactive Web App

**Document type:** Product Requirements Document, written for implementation by **Claude Code**
**Version:** 1.6 · September 30, 2026 (v1.1 added Section 17, public deployment and portfolio; v1.2 records the owner's decisions: GPLv3, repo `finance-excel-labs`; v1.3 applies the Milestone 0 resolutions; v1.4 records the Milestone 1 engine findings; v1.5 records the Milestone 2 scope split; v1.6 records the Milestone 3 split into two PRs, see the [change log](#change-log))
**Owner:** repository owner ([@jiaxingxue](https://github.com/jiaxingxue))
**Repository:** `jiaxingxue/finance-excel-labs` · **Live site (after first deploy):** https://jiaxingxue.github.io/finance-excel-labs/
**Source content:** `Excel_Implementation_Labs_Financial_Analysis.md` (the "Labs document")

---

## 0. Instructions for Claude Code (read first)

1. **The Labs document is the single source of truth** for lesson text, workbook cells, expected values, and experiments. Do not invent formulas, numbers, or lesson content. If something is missing or contradictory, stop and list it in `docs/OPEN_ISSUES.md` rather than guessing.
2. **Build in milestone order (Section 12).** Do not start UI work until **Milestone 1 (engine conformance)** passes: all 277 assertions in Appendix D must evaluate correctly in the chosen formula engine.
3. **Never hard-code expected values in application logic.** Expected values live only in the generated data files extracted from the Labs document.
4. **Owner decisions live in `docs/DECISIONS.md`** and override this PRD. For anything still PENDING there, use its default; ask the owner before any other product decision that isn't covered.
5. Work on a feature branch per milestone, write tests alongside code, and keep `README.md` current with run/test/build commands.
6. When a formula-engine function is missing or behaves differently from Excel, implement it as a **custom function plugin** (Section 7.4). Do **not** rewrite the lab formulas to dodge the gap. The lab formulas must run exactly as written.
7. **Deploy early.** This app is also a public portfolio piece (Section 17). Set up the GitHub Pages deployment pipeline in Milestone 0, so every milestone after that is live at a public URL behind passing tests.

---

## 1. Summary

Build a browser-based learning app where finance learners work through 12 Excel labs on variance analysis, reconciliations, and forecasting. The app has three panels:

- **Lesson panel:** instructional content from the Labs document.
- **Live spreadsheet grid:** real formulas recalculate instantly as the learner edits cells.
- **Checks panel:** auto-grades the workbook against the known-correct expected values.

Learners can explore finished models, rebuild formulas themselves, run guided what-if experiments, and download their workbook as a real `.xlsx` file to open in Microsoft Excel.

"Live Excel execution" in v1 means **an Excel-compatible formula engine running entirely in the browser**, not Microsoft Excel itself. Embedding genuine Microsoft Excel is a possible later phase (Section 15, Q2).

---

## 2. Problem and Goals

### 2.1 Problem
Reading about finance formulas is not the same as building them. Learners need a safe place to type formulas, see results recalculate, break things deliberately, and get immediate, specific feedback. They should not need an Excel license or have to set up files.

### 2.2 Goals (v1)
| # | Goal | Measure |
|---|---|---|
| G1 | Faithful execution of every classic (Tier A/B) lab formula | 277/277 assertions pass in CI |
| G2 | Learners can rebuild each lab's formulas and get graded feedback | Build mode available for Labs 2–10 and 12 |
| G3 | What-if learning | All 12 experiments from Appendix E runnable, with results matching the document |
| G4 | Portability | Export produces an `.xlsx` that opens in Excel with formulas intact |
| G5 | Zero-setup | Static site, no login, no backend, works offline after first load |
| G6 | Public, interview-ready demo | Live on GitHub Pages; README with proof of verification, screenshots, and demo script (Section 17) |

### 2.3 Non-goals (v1)
- Running Microsoft 365–only formulas (`LET`, `LAMBDA`, `SCAN`, `GROUPBY`, `PIVOTBY`, `HSTACK`, `VSTACK`) as live formulas. These are shown in a read-only **Modern Excel view**. (Lab 11's LAMBDA library is *simulated* as custom functions; see Section 7.5.)
- Excel UI features that aren't formulas: real PivotTables, Goal Seek dialog, What-If Data Tables, Power Query, Forecast Sheet, native Excel charts. The app explains them in the lesson text and shows app-native equivalents where specified.
- User accounts, multi-user collaboration, instructor dashboards, payments.
- Uploading arbitrary user workbooks.

---

## 3. Users and Key Scenarios

| Persona | Description | Needs |
|---|---|---|
| **Career learner** | Analyst or accountant building FP&A/controller skills | Guided lessons, clear feedback, confidence the numbers are right |
| **Self-studier with data background** | Knows data/SQL, new to finance modeling conventions | Formula explanations, "why" behind finance logic |
| **Instructor (secondary)** | Uses the app in a workshop | Stable URLs per lab, reset button, exportable workbook |

**Key scenarios**
1. *Explore:* the learner opens Lab 7, clicks `BankRec!B16`, and reads the formula plus a plain-English explanation.
2. *Build:* the learner blanks Lab 2, types `=SUMIFS(...)` into `BvA!C6`, fills down, and sees 6 checks turn green.
3. *What-if:* the learner drags the tolerance slider to 1.00, watches the reconciliation still tie, and reads why it's still wrong (Experiment E7.1).
4. *Export:* the learner downloads their workbook and opens it in Excel.

---

## 4. Source Content and Data Pipeline

### 4.1 Inputs
Place the Labs document at `content/Excel_Implementation_Labs_Financial_Analysis.md`. Everything else is **generated** from it by a build-time script (`scripts/extract-content.ts`, run with `npm run extract`).

### 4.2 What to extract

| Output file | Extracted from | Rule |
|---|---|---|
| `src/data/workbook.json` | Appendix C | Each `#### Sheet \`NAME\` (N cells)` heading is followed by one fenced ```` ```json ```` block, an object mapping A1 address → cell. Build `{ sheetOrder: [...headings in order], sheets: { NAME: {...} } }`. Validate that each sheet's cell count equals N. |
| `src/data/assertions.json` | Appendix D | The single fenced JSON array **within the Appendix D section** (other sections, such as Appendix B.2, contain JSON examples that must not match). Validate count = 277. |
| `src/data/experiments.json` | Appendix E | The single fenced JSON array **within the Appendix E section**. Validate count = 12. |
| `src/data/lessons/*.json` | Part 0, Labs 1–12, Appendices A–B | Split at `## ` headings. Each lab is identified by `## Lab N — Title`. Store the title, slug, and raw markdown. |
| `src/data/exercises.json` | Formula tables in each lab | Every markdown table row matching `` | `SHEET!ADDR` | `=FORMULA` | RESULT | `` marks `SHEET!ADDR` as a **pattern cell** of that lab (used by Build mode, Section 6.3). Validate that each formula equals the Appendix C formula for that cell. |
| `src/data/notice.json` | Blockquote before the Table of Contents | The verification notice, verbatim (LS-4). |
| `src/data/manifest.json` | All of the above | Counts, the lesson index (id, title, slug), and a SHA-256 of the Labs document. |

`src/data/` is **entirely generated**: `npm run extract` deletes and rewrites it, and `npm run extract -- --check` fails on any missing, stale, or extra file. Hand-authored configuration lives in `src/config/`.

### 4.3 Cell format (from Appendix C)
```ts
type CellSpec =
  | { v: string | number | boolean; fmt?: string }   // constant
  | { v: { date: string }; fmt?: string }             // date "YYYY-MM-DD"
  | { f: string; fmt?: string };                      // formula, starts with "="
```
- **Dates** → Excel serial number using the 1900 date system (serial = days since 1899-12-30). The engine must use the same epoch. `workbook.json` keeps dates as `{ "date": "YYYY-MM-DD" }` exactly as in Appendix C (so it deep-equals `reference/golden/spec.json`); the conversion to a serial happens when the engine adapter loads the workbook.
- **`fmt`** is an Excel number-format code. Display with an Excel-format library such as SheetJS **SSF** (`ssf` on npm), not hand-written formatting.
- Blank cells are absent from the JSON.

### 4.4 Assertion and experiment formats
```ts
interface Assertion { sheet: string; cell: string; expected: number | string | boolean; tolerance: number }
// numbers pass if |actual - expected| <= max(tolerance, |expected| * 1e-12);
// strings/booleans pass on strict equality.
// Special case: expected "" also passes when the cell evaluates to an empty value.

interface Experiment {
  id: string; lab: number; title: string;
  changes: Record<string /* "Sheet!A1" */, number | string>;  // only a full "YYYY-MM-DD" string → serial;
                                                              // other strings (e.g. text period "2026-07") stay text
  expect: Record<string /* "Sheet!A1" */, number | string | boolean | null>;
}
// Experiment numbers use tolerance 1e-6 relative, or 0.005 absolute (whichever is larger).
// expect null = the cell is empty: it passes when the cell evaluates to an empty value or "".
```

### 4.5 Lab → sheet → assertion mapping
Assertions belong to the lab that owns their sheet:

| Lab | Title (short) | Sheets shown (first tab is the primary sheet) | Owns assertions on |
|---|---|---|---|
| 1 | Data foundation | Map, GL | — (no assertions; completeness is graded via `BvA!D14` in Lab 2) |
| 2 | Budget vs. actual | BvA, GL, Map | BvA |
| 3 | Flux | Flux, GL | Flux |
| 4 | Price–volume–mix | PVM | PVM |
| 5 | Cost variance / flex budget | CostVar | CostVar |
| 6 | Pivot analysis | PivotLab, GL | PivotLab |
| 7 | Bank reconciliation | BankRec, Book, Bank | Book, Bank, BankRec |
| 8 | AR aging | AR, ARSummary | AR, ARSummary |
| 9 | Forecasting | Forecast | Forecast |
| 10 | Driver-based forecast | Drivers | Drivers |
| 11 | LAMBDA library | (lesson + simulated functions sandbox) | — |
| 12 | Checks dashboard | Checks | Checks |

Store this mapping as `src/config/labs.config.ts` (outside the generated `src/data/`). It is hand-authored because it's editorial, but validate it at build time: every sheet in `workbook.json` must appear in at least one lab.

---

## 5. Information Architecture and Layout

### 5.1 Routes
| Route | Content |
|---|---|
| `/` | Home: course overview, lab cards with progress, "Start Lab 1" |
| `/lab/:n` | Lab workspace (default mode: Explore) |
| `/lab/:n?mode=build\|whatif\|challenge\|modern` | Mode deep links |
| `/reference/functions` | Appendix A rendered (function tiers and fallbacks) |
| `/about` | Verification statement, tier explanation, licenses, and a "Further reading" placeholder for the two companion study guides (financial analysis concepts; CFO controls and COSO), shown as "coming soon" per D10 |
| `/verification` | CI verification summary + in-browser re-verification (Section 17.5) |

Use hash routing or a static-host–friendly configuration so the site works on any static host.

### 5.2 Lab workspace layout (desktop ≥ 1200 px)
```
┌──────────────────────────────────────────────────────────────────────────────┐
│ Top bar: Lab selector ▾ | Mode tabs [Explore][Build][What-if][Challenge][Modern] │
│          Progress: 14/22 checks ✓ | Reset lab | Export .xlsx                  │
├───────────────────────┬──────────────────────────────────┬───────────────────┤
│ LESSON PANEL (30%)    │ SPREADSHEET (45%)                 │ CHECKS PANEL (25%)│
│ Markdown for the lab; │ Formula bar: [BvA!C6] =SUMIFS(... │ Per-assertion rows│
│ cell refs like        │ Sheet tabs: BvA | GL | Map         │  ✓ BvA!E6 = -32,000│
│ `BvA!C6` are links →  │ Grid with row/col headers          │  ✗ BvA!H9 expected │
│ select that cell      │ Input cells tinted blue            │    TRUE, got FALSE│
│                       │ Formula cells normal               │ What-if controls  │
│                       │ Changed cells flash                │ (in What-if mode) │
│                       │ Chart area (Labs 4, 9, 10)         │                   │
└───────────────────────┴──────────────────────────────────┴───────────────────┘
```
- **Tablet (768–1199 px):** the lesson panel collapses into a drawer; the grid and checks stack.
- **Phone (< 768 px):** single column with tabs [Lesson | Sheet | Checks]. The grid scrolls horizontally inside its container; the page itself must not scroll horizontally.

---

## 6. Functional Requirements

### 6.1 Spreadsheet grid (all modes)
| ID | Requirement |
|---|---|
| GR-1 | Render sheets from the engine, with column letters, row numbers, and sheet tabs limited to the lab's sheets (Section 4.5). A "Show all sheets" toggle reveals every sheet (read-only outside the lab). |
| GR-2 | Display calculated values using each cell's `fmt` (Excel format codes via SSF). Booleans display as `TRUE`/`FALSE`; errors display as Excel error text (`#N/A`, `#VALUE!`, `#DIV/0!`, `#NAME?`, `#REF!`). |
| GR-3 | **Formula bar** shows the selected cell's address and raw contents (formula or constant). |
| GR-4 | Editing: double-click, Enter, or start typing to edit; Enter/Tab commit; Esc cancels. Typing `=` starts a formula. Commit triggers recalculation. |
| GR-5 | **Recalculation < 50 ms** (p95) for a single-cell edit in the full workbook (~670 formulas) on a mid-range laptop. |
| GR-6 | **Visual roles:** input (constant) cells get a light blue tint; formula cells are untinted; cells that changed value in the last recalculation flash (≈600 ms highlight, and nothing flashes when `prefers-reduced-motion` is set). |
| GR-7 | **Precedent/dependent highlighting:** selecting a formula cell outlines its precedent ranges (color-coded like Excel). A "Trace dependents" button lists the cells that depend on the selection. Use the engine's dependency API if available; otherwise parse references from the formula. |
| GR-8 | **Fill Down / Fill Right** (Ctrl+D / Ctrl+R, plus toolbar buttons) over a selected range, adjusting relative references exactly as Excel does (`$` locks respected). Use the engine's copy/paste so the adjustment is engine-native. |
| GR-9 | Copy/paste within the grid; undo/redo (Ctrl+Z / Ctrl+Y) of at least 50 steps. |
| GR-10 | **Formula autocomplete:** while typing a function name, show matching functions with a one-line signature. Clicking a cell or dragging a range while editing inserts the reference. |
| GR-11 | **Cell-link navigation:** any `Sheet!A1` or `Sheet!A1:B2` reference inside lesson text is a clickable link that switches tab and selects that cell or range. |
| GR-12 | **Conditional formatting** (display only): support the rule types used in the Labs document: formula-based fill ("`=$G6="Unfavorable"` → light red") and simple threshold fills. Define them in `labs.config.ts` per lab. |
| GR-13 | Keyboard navigation: arrow keys, Home/End, Ctrl+arrow jumps, Page Up/Down, and Ctrl+Page Up/Down between sheets. |

### 6.2 Explore mode (default)
| ID | Requirement |
|---|---|
| EX-1 | Load the complete workbook: all Appendix C cells. |
| EX-2 | Cell inspector: selecting a formula cell shows (a) the formula, (b) its value, (c) its explanation if the lesson has one for that pattern, and (d) its tier badge (A/B). |
| EX-3 | Edits are allowed and are sandboxed to the session; "Reset lab" restores the original sheets. |
| EX-4 | The Checks panel shows the lab's assertions live; all should be green on load. |

### 6.3 Build mode
| ID | Requirement |
|---|---|
| BD-1 | On entering Build mode, **clear every formula cell in the lab's primary sheet(s)** (the sheets that own assertions for that lab). Keep all constants. Keep the other sheets fully calculated so cross-sheet references work. |
| BD-2 | Mark **pattern cells** (from `exercises.json`) with a dashed outline and a "✎ build me" affordance. Other blank formula cells are "fill targets," shown with a lighter dotted outline. |
| BD-3 | **Grading is value-based**, never text-based: after each commit, re-evaluate the lab's assertions. A learner may use any valid formula (e.g., `INDEX/MATCH` or `XLOOKUP` if the engine supports it). |
| BD-4 | **Anti-hard-coding check:** a cell passes only if (a) it contains a formula (starts with `=`) and (b) where an experiment for the lab changes one of its precedents, the cell also matches that experiment's expected value. Hidden re-grading runs the lab's experiments against a cloned engine instance and never disturbs the learner's grid. |
| BD-5 | **Three-level hints** per pattern cell: (1) the functions used (parsed from the reference formula, e.g., "SUMIFS"); (2) a skeleton with references masked (`=SUMIFS(GL!$E$2:$E$37, ___, $A6, ___, "BUD", ___, ___)`); (3) reveal the full formula. Record which level was used. |
| BD-6 | "Fill for me" button on a pattern cell: fills its row or column from the learner's pattern formula, using the same fill-direction rules as the Labs document ("fill down to row 11"). The fill ranges come from the lesson text; encode them in `labs.config.ts` as `fills: [{from:"BvA!B6:H6", to:"BvA!B7:H11"}]`. |
| BD-7 | Lab completion = all of the lab's assertions pass with BD-4 satisfied. Show a completion state and store it in progress. |

### 6.4 What-if mode
| ID | Requirement |
|---|---|
| WI-1 | Show input controls bound to cells as listed in Labs document **Appendix B.4** (sliders, dropdowns, date pickers, numeric inputs, editable tables). Define each control in `labs.config.ts` with min, max, step, and format. |
| WI-2 | Moving a control writes to the cell and recalculates live (debounce ≤ 50 ms). Changed cells flash (GR-6). A "delta" column in the Checks panel shows the before → after values of key output cells. |
| WI-3 | Special action for Lab 9: an "Inject outlier" button sets `Forecast!C8` to 250000 (Experiment E9.1), with an undo. |
| WI-4 | "Reset inputs" restores the base values. |

### 6.5 Challenge mode
| ID | Requirement |
|---|---|
| CH-1 | Present the lab's experiments (Appendix E) as challenge cards: title, the change to make, and a question. The question text is drawn from the lesson's "Try it" section and "Lesson from …" paragraph where one exists. |
| CH-2 | The learner applies the change manually (or clicks "Apply for me"). The app compares every `expect` cell and marks the challenge complete when all match. |
| CH-3 | After completion, reveal the lesson paragraph that explains the result (e.g., E7.1: "balanced is not the same as correct"). |
| CH-4 | A prediction step (optional per challenge): before applying the change, the learner picks what they think happens (multiple choice auto-generated from the watch cells: "increases / decreases / unchanged"). Record correctness. |

### 6.6 Modern Excel view
| ID | Requirement |
|---|---|
| MV-1 | For each lab, list the Tier C/D formulas from the Labs document's "Modern equivalent" blocks, side by side with the Tier A formula they replace. |
| MV-2 | Label them clearly: "Microsoft 365 / Excel 2021+ — shown for reference; not executed in this app" (unless Section 7.5 simulation applies). |
| MV-3 | "Copy formula" button; "Try in real Excel" opens the export dialog (Section 6.8) with the option to include these formulas in a `Modern` sheet as text. |

### 6.7 Charts
| ID | Lab | Chart | Data |
|---|---|---|---|
| CT-1 | 4 | Waterfall (variance bridge) | `PVM!A11:B15`. The first and last bars are totals; the others float |
| CT-2 | 9 | Line chart: history (`A5:C34`) + trend × season forecast (`A38:A43`, `E38:E43`, dashed) + backtest actual vs. model (`A47:C52`) | Forecast sheet |
| CT-3 | 10 | Heatmap of the sensitivity grid (`Drivers!B29:F33`) with a diverging scale centered on the Base cell (`D31`) | Drivers sheet |
| CT-4 | 8 | Stacked bar by customer × bucket (`ARSummary!A11:F14`) | ARSummary |

Charts read live engine values and re-render on recalculation. Use an MIT-licensed chart library (e.g., Recharts) or hand-rolled SVG. Axis labels must state units, and color alone must never carry meaning (add labels or patterns).

### 6.8 Export and import
| ID | Requirement |
|---|---|
| XP-1 | "Export .xlsx" writes the learner's **current** workbook: every sheet, every constant and formula, number formats, and cached values, using an MIT-licensed writer (e.g., ExcelJS). |
| XP-2 | Formulas using functions introduced in Excel 2010+ must be written with the `_xlfn.` prefix where the file format requires it (at minimum `FORECAST.LINEAR` → `_xlfn.FORECAST.LINEAR`). Keep a mapping table in code. |
| XP-3 | Options: (a) current state vs. original solution; (b) include a `Modern` sheet listing the Tier C/D formulas as text; (c) include a `Checks` sheet with assertions and pass/fail. |
| XP-4 | **Round-trip test:** exporting the original workbook and re-importing it into the engine must pass 277/277 assertions. |
| XP-5 | (v1.1, optional) "Import my workbook": only re-import files this app exported (same sheet structure), to restore progress. |

### 6.9 Progress and persistence
| ID | Requirement |
|---|---|
| PR-1 | Store per-lab progress in `localStorage`, wrapped in try/catch. The app must work when storage is unavailable (private mode) and simply not persist. |
| PR-2 | Saved state: per-lab completion, per-cell learner formulas in Build mode, hint levels used, challenges completed, last mode, last lab. |
| PR-3 | "Reset lab" and "Reset all progress" (with a confirmation). |
| PR-4 | Schema version the saved state; on version mismatch, migrate or reset with a notice. |

### 6.10 Lessons
| ID | Requirement |
|---|---|
| LS-1 | Render each lab's markdown (GitHub-flavored: tables, code blocks, blockquotes). Code blocks tagged `excel` get formula syntax highlighting. |
| LS-2 | Replace static **values grids** in the lesson (tables whose header row is `| | A | B | …`) with a "Show live cells" button that jumps to that range in the grid. Keep the static table visible for reading. |
| LS-3 | In formula tables (`| Cell | Formula | Result |`), make the Result column **live**: it shows the current engine value, plus a ✓/✗ against the original. |
| LS-4 | The verification notice from the top of the Labs document appears on `/about` and in a collapsible banner on each lab. |

### 6.11 Home and navigation
- Lab cards (1–12) show: title, learning goals (from each lab's "Learning goals" list where present), estimated time, and progress ring.
- Recommended order is 1 → 12, but all labs are unlocked.

---

## 7. Formula Engine

### 7.1 Requirements
| ID | Requirement |
|---|---|
| EN-1 | Runs fully client-side in JavaScript/WebAssembly. No server calls for calculation. |
| EN-2 | Supports multiple sheets, cross-sheet references (`GL!$E$2:$E$37`), absolute/relative references, and ranges. |
| EN-3 | Supports **array arithmetic inside functions**, as used by `SUMPRODUCT(...^range...)` and `LOOKUP(2,1/((cond)*(cond)*(cond)),range)`. |
| EN-4 | Excel-compatible semantics for: empty string vs. blank, TRUE/FALSE results, error propagation, the 1900 date system, text-number coercion in criteria (`"<="&$B$2`, `">0"`). |
| EN-5 | Custom-function plugin mechanism for filling gaps (Section 7.4). |
| EN-6 | Dependency graph access (precedents/dependents) for GR-7, or a documented fallback. |
| EN-7 | Deterministic results; floating-point differences ≤ 1e-9 relative versus the assertions. |

### 7.2 Recommended engine: HyperFormula (behind an adapter)
- **Why:** mature headless JavaScript spreadsheet engine, 400+ built-in functions, multi-sheet support, an array-arithmetic mode (`useArrayArithmetic: true`), a custom-function plugin API, undo/redo, and copy/paste with reference adjustment.
- **License (RESOLVED: GPLv3, see `docs/DECISIONS.md` D1):** HyperFormula is dual-licensed: **GPLv3** (config `licenseKey: 'gpl-v3'`), or a paid proprietary license from Handsontable. Using the GPL key means the app's source must be distributed under GPLv3-compatible terms. The owner has confirmed the app is open source under GPLv3, so configure `licenseKey: 'gpl-v3'`. (If this ever changes to a commercial key, supply it through an environment variable, `VITE_HF_LICENSE_KEY`.)
- **Known function gaps to verify in Milestone 1:** a check of HyperFormula's published built-in function list suggests that some functions used by the labs are **not built in**: `AVERAGEIFS`, `LOOKUP`, `TEXT`, `INTERCEPT`, `TREND`, and `FORECAST.LINEAR`. Treat this list as a hypothesis. The conformance test (Section 7.3) is authoritative. Each confirmed gap gets a custom function plugin (Section 7.4).
- **Milestone 1 result (HyperFormula 3.4.0):** `AVERAGEIFS`, `LOOKUP`, `INTERCEPT`, `TREND`, and `FORECAST.LINEAR` are missing. `TEXT` exists but ignores number-format codes, and `INDEX` on a one-row range with one index returns `#NUM!`; both are overridden by plugins. Two gaps are fixed in engine configuration instead: bare `TRUE`/`FALSE` literals (registered as named expressions) and output rounding (`smartRounding: false`). Details are in `docs/milestones/M1.md`.

### 7.3 Engine adapter and conformance suite
Define an interface so the engine can be swapped:
```ts
interface FormulaEngine {
  load(workbook: WorkbookSpec): void;
  setCell(ref: CellRef, content: string | number | boolean | null): ChangedCell[];
  getValue(ref: CellRef): CellValue;          // number | string | boolean | ErrorValue | null
  getFormula(ref: CellRef): string | null;
  fill(source: RangeRef, target: RangeRef): ChangedCell[];
  precedents(ref: CellRef): RangeRef[];
  dependents(ref: CellRef): CellRef[];
  clone(): FormulaEngine;                     // for hidden re-grading (BD-4)
  undo(): void; redo(): void;
  exportCells(): WorkbookSpec;                // current state, for export
}
```
**Conformance suite (`tests/conformance.test.ts`)** — must pass before any UI work:
1. Load `workbook.json` → evaluate all 277 assertions → **277/277 pass**.
2. Run all 12 experiments on fresh engine clones → every `expect` cell matches.
3. Error-free check: no formula cell in the base workbook evaluates to an error, **except** cells whose expected value is an error (currently none). Cells that legitimately evaluate to `""` (e.g., `Book!F5`) must be empty strings, not errors.
4. Targeted semantic tests (each a small standalone sheet):

| Test | Formula | Expected |
|---|---|---|
| LOOKUP last-match trick | `=LOOKUP(2,1/({0,1,0,1}),{"a","b","c","d"})` | `"d"` |
| LOOKUP no match → error | `=IFERROR(LOOKUP(2,1/({0,0}),{"a","b"}),"")` | `""` |
| Approximate LOOKUP buckets | `=LOOKUP(x,{-99999,1,31,61,91},{"Current","1-30","31-60","61-90","90+"})` for x = −10, 0, 1, 30, 31, 60, 61, 90, 91, 108 | Current, Current, 1-30, 1-30, 31-60, 31-60, 61-90, 61-90, 90+, 90+ |
| Boolean equality | `=(TRUE)=(FALSE)` | `FALSE` |
| N() on text | `=N("n/a")` | `0` |
| SUMIFS text-period criteria | lab formula `BvA!C6` | 500000 |
| AVERAGEIFS with concatenated criterion | lab formula `Forecast!N5` | matches assertion |
| SUMPRODUCT with array power | lab formula `Drivers!C21` | matches assertion |
| TEXT format | `=TEXT(100,"#,##0.00")` | `"100.00"` |
| EDATE on date serial | `=EDATE(DATE(2026,6,1),1)` | serial of 2026-07-01 |
| Divide-by-zero guard | `=IF(0=0,"n/a",1/0)` | `"n/a"` |

### 7.4 Custom function plugins (gap fillers)
Implement each confirmed gap as a plugin under `src/engine/plugins/`, following Excel's documented semantics, with unit tests comparing against the assertions:

| Function | Semantics to implement |
|---|---|
| `AVERAGEIFS(avg_range, crit_range1, crit1, …)` | Average of numeric cells in `avg_range` where all criteria match. Criteria support operators (`<`, `<=`, `>`, `>=`, `<>`, `=`) and wildcards `*`/`?`. Returns `#DIV/0!` when nothing matches. |
| `LOOKUP(lookup, lookup_vector, result_vector)` | Vector form. Binary-search approximate match: returns the result for the last position whose value is ≤ lookup, **ignoring error values** in `lookup_vector`. Returns `#N/A` if none. Must accept array expressions and inline arrays. |
| `INTERCEPT(known_y, known_x)` | Ordinary least-squares intercept. |
| `TREND(known_y, known_x, new_x)` | OLS prediction (the scalar `new_x` case is required; the array case is nice to have). |
| `FORECAST.LINEAR(x, known_y, known_x)` and alias `FORECAST` | OLS prediction at x. |
| `TEXT(value, format)` | Delegate to SSF formatting. Overrides the built-in (M1). |
| `INDEX(array, row, [col])` | With `col` omitted, a one-row array is indexed by column. Out-of-range indexes return `#REF!`. Overrides the built-in (M1). |

If the chosen engine implements any of these natively and passes conformance, don't override it.

### 7.5 Lab 11 — simulated LAMBDA library
The engine cannot execute `LAMBDA`, so register the Lab 11 library as **named custom functions** with identical behavior: `VARPCT`, `FAVFLAG`, `NEEDSCOMMENT`, `AGEBUCKET`, `MAPE`, `BIAS`, `CUSTROLL` (returns a spilled array if the engine supports it; otherwise add an optional 5th argument `k` that returns the k-th value), and `PVMCHECK`.

Lab 11's workspace is a sandbox sheet where the learner calls these functions against the other lab sheets. The "Test call / Should equal" table from the Labs document becomes its checks. Label these functions clearly: "Simulated: in real Excel you'd define these with LAMBDA in Name Manager."

### 7.6 Alternatives considered
| Option | Pros | Cons | Decision |
|---|---|---|---|
| HyperFormula | Mature, rich API, plugins | GPLv3 or paid license; function gaps | **Default**, behind the adapter |
| IronCalc (Rust → WASM) | MIT/Apache license; reads/writes xlsx | Early-stage/work-in-progress per its own README | Implement as a second adapter **only if** the owner rejects the HyperFormula license. It must pass the same conformance suite |
| Formula.js | MIT; function implementations | Functions only, no dependency graph or multi-sheet engine | Possible source of reference implementations for plugins |
| Microsoft Excel (Office.js / Excel for the web embed) | Genuine Excel, including LET/LAMBDA | Requires a Microsoft account and online files; not zero-setup | Out of scope for v1 (Section 15 Q2) |

---

## 8. Non-Functional Requirements

| Area | Requirement |
|---|---|
| Performance | First meaningful paint < 2 s on a 4G connection; full workbook load + calculation < 300 ms; single-edit recalculation < 50 ms (p95) |
| Bundle | Initial JavaScript ≤ 600 KB gzipped (engine may lazy-load after the lesson renders) |
| Offline | Works offline after first load (service worker caches static assets and data) |
| Privacy | No analytics, cookies, or network calls in v1 beyond static assets. Nothing the learner types leaves the browser |
| Accessibility | WCAG 2.1 AA: grid navigable by keyboard, `role="grid"` semantics, screen-reader announcement of the selected cell's address/value/formula, 4.5:1 contrast, focus rings, reduced-motion support |
| Theming | Light and dark modes following the system setting, with a manual toggle. Input-cell tint and conditional formats have dark-mode variants |
| Browsers | Latest 2 versions of Chrome, Edge, Firefox, Safari; iOS Safari 17+ |
| Numbers | Display always uses Excel format codes; internal comparisons use raw values, never formatted strings |
| Robustness | A formula error in a learner cell never crashes the app. Engine exceptions are caught and shown as `#ERROR!` with a message |

---

## 9. Technical Architecture

### 9.1 Stack (recommended)
- **Vite + React + TypeScript** (strict mode)
- **State:** Zustand (UI state and progress) from M5, when persistence arrives; until then one small store read with React's `useSyncExternalStore`. The engine instance lives outside React state and is accessed through the adapter
- **Grid:** a custom React grid component. The sheets are small (≤ 60 rows × 16 columns), so no virtualization is needed. Avoid commercial grid libraries.
- **Markdown:** `react-markdown` + `remark-gfm`; a custom renderer for cell-reference links and live Result columns
- **Formatting:** `ssf` (Excel number formats)
- **Charts:** Recharts or hand-rolled SVG
- **Export:** ExcelJS
- **Tests:** Vitest (unit, conformance), Playwright (end-to-end), axe-core (accessibility)
- **Lint/format:** ESLint + Prettier

### 9.2 Repository structure
```
/
├─ content/Excel_Implementation_Labs_Financial_Analysis.md
├─ scripts/extract-content.ts          # Section 4 pipeline (npm run extract)
├─ src/
│  ├─ data/                            # generated only: workbook.json, assertions.json, experiments.json,
│  │                                   #   exercises.json, notice.json, manifest.json, lessons/*.json
│  ├─ config/labs.config.ts            # hand-authored: sheets, fills, controls, CF rules, charts
│  ├─ engine/
│  │  ├─ FormulaEngine.ts              # interface
│  │  ├─ HyperFormulaEngine.ts         # adapter
│  │  └─ plugins/                      # AVERAGEIFS, LOOKUP, INTERCEPT, TREND, FORECAST.LINEAR, TEXT,
│  │                                   # lab11 simulated functions
│  ├─ grading/                         # assertions, anti-hard-coding, experiments, hints
│  ├─ components/                      # Grid, FormulaBar, SheetTabs, ChecksPanel, LessonPanel,
│  │                                   # ModeTabs, WhatIfControls, ChallengeCard, Charts, ExportDialog
│  ├─ routes/                          # Home, Lab, Reference, About
│  ├─ state/                           # stores, persistence, schema migrations
│  └─ export/                          # ExcelJS writer, _xlfn mapping
├─ tests/
│  ├─ conformance.test.ts
│  ├─ plugins/*.test.ts
│  ├─ grading.test.ts
│  └─ e2e/*.spec.ts
├─ public/  (favicon.svg, og-image.png, robots.txt)
├─ .github/workflows/deploy.yml   # Section 17.4
├─ LICENSE                        # GPLv3 (D1, confirmed)
├─ THIRD_PARTY_NOTICES.md         # licenses of bundled dependencies
└─ docs/  (ARCHITECTURE.md, OPEN_ISSUES.md, CONTENT_PIPELINE.md, DEMO.md, INTERVIEW_NOTES.md)
```

### 9.3 `labs.config.ts` shape (example for Lab 2)
```ts
export const lab2: LabConfig = {
  n: 2, slug: 'budget-vs-actual', title: 'Budget vs. Actual Variance Report',
  sheets: ['BvA', 'GL', 'Map'], primarySheets: ['BvA'],
  fills: [{ from: 'BvA!B6:H6', to: 'BvA!B7:H11' }, { from: 'BvA!E6:H6', to: 'BvA!E12:H12' }],
  controls: [
    { cell: 'BvA!B1', type: 'select', options: ['2025-08', '2026-07', '2026-08'], label: 'Report period' },
    { cell: 'BvA!B2', type: 'slider', min: 0.01, max: 0.20, step: 0.01, fmt: '0%', label: 'Pct threshold' },
    { cell: 'BvA!B3', type: 'slider', min: 0, max: 25000, step: 500, fmt: '#,##0', label: 'Abs threshold' },
  ],
  conditionalFormats: [
    { range: 'BvA!A6:H12', formula: '=$G6="Unfavorable"', style: 'negativeFill' },
    { range: 'BvA!A6:H12', formula: '=$H6=TRUE', style: 'bold' },
  ],
  charts: [],
  estMinutes: 30,
};
```
Conditional-format formulas are evaluated **by the engine** (in a hidden scratch sheet, or by evaluating relative to each row), not with ad-hoc JavaScript.

### 9.4 Grading engine
```ts
gradeLab(lab): { assertions: AssertionResult[]; complete: boolean }
antiHardcode(lab, cell): boolean   // BD-4: formula present + experiment re-check on a clone
applyExperiment(engine, exp): ExperimentResult
```
Hidden re-grading clones run in a Web Worker if they take longer than 16 ms, so the UI never stutters.

---

## 10. Content Rules
- Lesson text comes verbatim from the Labs document (rendering transforms only: links, live results).
- Use the document's terminology throughout: "Tier A/B/C/D," "pattern cell," "fill down," "assertion," "experiment."
- The verification statement must say the same thing as the document: classic formulas were engine-verified, Microsoft 365–only formulas are reference-only, and ETS outputs vary by spreadsheet program. In the app, extend the note to say calculation uses an Excel-compatible engine, not Microsoft Excel.
- `FORECAST.ETS` formulas appear only in the lesson and Modern view, never as live graded cells.

---

## 11. Acceptance Criteria (Definition of Done for v1)

1. `npm run extract` regenerates all data files from the Labs document with the validation counts shown in Section 4.2 (15 sheets, 277 assertions, 12 experiments).
2. **Conformance:** 277/277 assertions and 12/12 experiments pass in CI (`npm test`).
3. **Explore:** every lab loads with all checks green; clicking any cell link in a lesson selects it; formula bar and inspector work.
4. **Build:** for each of Labs 2–10 and 12, an automated Playwright test enters the reference pattern formulas (typed through the UI), uses Fill actions, and reaches lab completion. A second test enters hard-coded numbers and is **rejected** by BD-4 wherever an experiment changes a precedent of the cell.
5. **What-if:** every Appendix B.4 control exists and moves its cell. Experiments E2.1, E7.1, E9.1, and E10.1 reproduce their documented values when driven through controls.
6. **Challenge:** all 12 challenges can be completed and reveal their explanation text.
7. **Export:** the exported `.xlsx` re-imports with 277/277 passing (XP-4). A manual check in Microsoft Excel shows no `#NAME?` errors and matching values (record the result in `docs/OPEN_ISSUES.md` if it can't be automated).
8. **Lab 11:** simulated functions return the "Should equal" values from the Lab 11 table.
9. **Non-functional:** Lighthouse performance ≥ 90 and accessibility ≥ 95 on `/lab/2`; axe-core reports zero serious or critical issues; recalculation p95 < 50 ms measured by a benchmark test.
10. Works offline after first load; no network requests other than static assets (verified in a Playwright test).
11. README documents setup, scripts, the license decision, and how to update content.
12. **Public deployment** (to https://jiaxingxue.github.io/finance-excel-labs/): pushing to `main` runs lint, unit/conformance, build, and e2e tests, and deploys to GitHub Pages only when all pass. A deep link such as `/#/lab/7?mode=whatif` loads correctly after a hard refresh on the deployed site.
13. **Verification page:** `/#/verification` shows the CI result (from `verification.json`: passed/total, commit, date) **and** a "Run verification in this browser" button that re-evaluates all 277 assertions and 12 experiments in the visitor's browser in under 5 seconds, with a pass/fail table.
14. **Portfolio README:** matches the Section 17.6 template, with a working live-demo link, build badge, verification badge, and screenshots auto-generated from the deployed build.
15. **Demo readiness:** the 3-minute script in `docs/DEMO.md` can be performed end-to-end on the deployed site, including offline (after one visit) with Wi-Fi turned off. A Playwright test runs the same click path.

---

## 12. Milestones (build in this order)

| # | Milestone | Deliverables | Exit criteria |
|---|---|---|---|
| **M0** | Scaffold + content pipeline + deploy skeleton | Vite/React/TS app, lint/test setup, `extract-content.ts`, generated data, GitHub Actions workflow (Section 17.4), `LICENSE` (GPLv3), `README.md` kept current, placeholder page live on GitHub Pages once the repo is public | Section 11 #1; the workflow runs green |
| **M1** | Engine + conformance | Adapter, plugins for confirmed gaps, conformance suite | Section 11 #2, **blocking gate** |
| **M2** | Grid + Explore | Grid, formula bar, tabs, formats, lesson panel, cell links, checks panel: GR-1–GR-4, GR-6, GR-11, GR-13 (+ undo/redo), EX-1–EX-4 (EX-2c deferred), LS-1, LS-4; axe-core check in e2e | Section 11 #3 |
| **M3** | What-if + Challenge, in two PRs | **M3a:** What-if controls (WI-1–WI-4), GR-12 (conditional formatting), formula-bar commit on blur. **M3b:** experiments as challenge cards, charts CT-1…CT-4, GR-7 (precedent outlines), LS-2, LS-3, EX-2(c) | Section 11 #5 (M3a), #6 (M3b) |
| **M4** | Build mode | Blanking, fills, hints, anti-hard-coding, completion; also GR-8 (fill), GR-9 (copy/paste), GR-10 (autocomplete) | Section 11 #4 |
| **M5** | Export, Modern view, Lab 11, persistence | ExcelJS export, Modern view, simulated LAMBDA functions, localStorage (introduces Zustand) | Section 11 #7, #8 |
| **M6** | Polish + quality | Accessibility, dark mode, responsive layouts, service worker, performance | Section 11 #9, #10, #11 |
| **M7** | Portfolio release | Verification page, `verification.json`, README (Section 17.6), automated screenshots/GIF, social preview, demo script in `docs/DEMO.md`, v1.0.0 tag and GitHub Release | Section 11 #12–#15 |

Each milestone ends with a short demo note in `docs/` (what works, known issues).

---

## 13. Analytics and Success Metrics

v1 has **no telemetry** (privacy requirement). Success is evaluated qualitatively and locally:
- Labs completed per learner (visible on the local progress page)
- Hint level distribution (local)
- Owner usability test: 5 learners complete Labs 2 and 7 in Build mode unaided in under 45 minutes each

Opt-in, privacy-preserving analytics may be considered for v2.

---

## 14. Risks and Mitigations

| Risk | Impact | Mitigation |
|---|---|---|
| Engine semantics differ from Excel (LOOKUP trick, blanks vs. `""`, boolean comparisons) | Wrong grades, erosion of trust | Conformance suite + targeted tests (Section 7.3); plugins; fail CI on any mismatch |
| HyperFormula licensing incompatible with the owner's plans | Legal exposure | Resolved: GPLv3 confirmed (D1). The repo ships a GPLv3 `LICENSE`; the adapter still allows a later swap to IronCalc |
| Learners use valid alternative formulas that the engine doesn't support (e.g., `XLOOKUP` in an engine without it) | Frustration | Autocomplete only offers supported functions. On `#NAME?`, show "This function isn't supported here. Try the Tier A equivalent: …" (from Appendix A) |
| Hard-coded answers pass grading | False completion | BD-4 anti-hard-coding using experiments |
| Floating-point noise | Spurious failures | Tolerances from assertions; the `ROUND` checks in lab formulas |
| Content drift between the Labs document and the app | Mismatched lessons | Single-source pipeline; CI regenerates data and fails on a diff |
| Export not opening cleanly in Excel | Broken promise | `_xlfn` mapping, round-trip test, manual Excel check |

---

## 15. Open Questions for the Owner

Owner answers are recorded in **`docs/DECISIONS.md`**, which overrides this table.

| # | Question | Status / default |
|---|---|---|
| Q1 | License: open source under GPLv3, or a HyperFormula commercial license, or an MIT/Apache engine? | **Resolved (D1): GPLv3.** Use `licenseKey: 'gpl-v3'`; the repo ships a GPLv3 `LICENSE` |
| Q2 | A later phase with **genuine Microsoft Excel** (Office.js / Excel for the web) to run Tier C/D formulas live? | **Resolved (D9): not in v1** |
| Q3 | Hosting target? | **Resolved: GitHub Pages** via GitHub Actions (Section 17). Keep the build host-agnostic |
| Q4 | Branding: app name, colors, logo? | Pending (D11). Default: neutral design; name "Excel Labs for Financial Analysis" |
| Q5 | Allow Tier C functions (`XLOOKUP`, `FILTER`) in Build mode when the engine supports them? | Pending (D8). Default: yes, if they compute correctly |
| Q6 | Include the two earlier study guides as reading content? | **Resolved (D10): not in v1; placeholder link on the About page** |
| Q7 | GitHub username / repo / custom domain? | **Resolved (D2, D3):** `jiaxingxue/finance-excel-labs` → `https://jiaxingxue.github.io/finance-excel-labs/`. Custom domain pending (D4); default none |
| Q8 | Name and links for the footer/README? | Pending (D5, D6). Default: GitHub handle `@jiaxingxue` only; add name/links when provided. **Never** commit an email address or phone number |

---

## 16. Appendix — Traceability to the Labs Document

| Labs document section | App feature |
|---|---|
| Part 0.2 conventions (inputs blue, bounded ranges, text periods) | GR-6 tinting; engine loads exactly as specified |
| Part 0.3 function tiers | Tier badges (EX-2), Modern view (MV-1), autocomplete limits |
| Labs 1–12 "Formulas" tables | `exercises.json` pattern cells (BD-2), live Result column (LS-3) |
| Labs "Result" value grids | "Show live cells" (LS-2) |
| Labs "Try it" + "Lesson from…" | Challenge mode (CH-1…CH-3) |
| Lab 4 waterfall, Lab 9 chart, Lab 10 grid | CT-1…CT-3 |
| Lab 11 LAMBDA table | Simulated functions (Section 7.5) |
| Lab 12 Checks | Checks panel + Lab 12 grading |
| Appendix A fallbacks | `#NAME?` guidance (Risk table), `/reference/functions` |
| Appendix B blueprint | Sections 5–6 (this PRD refines it) |
| Appendix C / D / E | `workbook.json` / `assertions.json` / `experiments.json` |


---

## 17. Public Deployment and Portfolio Readiness

**Goal:** a stable public URL that a hiring manager can open on any device. Within 30 seconds it should show *what the app does*, *proof that it's correct*, and *who built it*. It must also support a confident live demo.

### 17.1 Hosting decision
| Item | Decision |
|---|---|
| Host | **GitHub Pages**, published by **GitHub Actions** (repo Settings → Pages → Source: *GitHub Actions*) |
| Repository | **Public** (GitHub Pages on the free plan publishes only from public repositories) |
| URL | `https://jiaxingxue.github.io/finance-excel-labs/` (project site; `BASE_PATH=/finance-excel-labs/`) |
| Custom domain | Optional (D4). If used: add `public/CNAME`, configure DNS, and enable "Enforce HTTPS" |
| Alternatives | Netlify, Vercel, Cloudflare Pages: same static `dist/`, no code changes except `BASE_PATH=/` |

### 17.2 Repository setup (Claude Code does these, the owner clicks the settings)
1. Default branch `main`; protect it (require the `ci` job to pass before merging) if the owner wants a PR workflow.
2. Settings → Pages → Build and deployment → Source = **GitHub Actions**. (Owner action; Claude Code documents it in the README "Deploy" section.)
3. Repo "About" box: description "Interactive Excel labs for FP&A: variance analysis, reconciliations, forecasting. Live formula engine, auto-graded.", website = the Pages URL, topics: `excel`, `fpa`, `financial-analysis`, `accounting`, `spreadsheet`, `react`, `typescript`, `education`.
4. Add `LICENSE` (full GPLv3 text from gnu.org; D1 confirmed) and `THIRD_PARTY_NOTICES.md` (generated with a license-listing tool from `package.json` dependencies).
5. Add `.github/ISSUE_TEMPLATE/bug.md` (optional; signals maintainability).

### 17.3 Build configuration for a sub-path site
| Concern | Requirement |
|---|---|
| Base path | `vite.config.ts`: `base: process.env.BASE_PATH ?? '/'`. The workflow sets `BASE_PATH=/${{ github.event.repository.name }}/`. Locally, use `npm run build:pages` (sets `BASE_PATH=/finance-excel-labs/` through `cross-env`), never an inline `BASE_PATH=… npm run build`: that syntax fails in PowerShell, and Git Bash rewrites `/finance-excel-labs/` into a Windows path |
| Routing | **Hash routing** (`/#/lab/7`), so deep links survive refreshes on a static host with no rewrite rules. Also copy `index.html` → `404.html` at build time as a safety net |
| Asset URLs | Every runtime fetch of data files uses `import.meta.env.BASE_URL` as the prefix, never a leading `/` |
| Service worker | Register at `${import.meta.env.BASE_URL}sw.js` with scope = `BASE_URL`. Precache the app shell and `src/data/*`. Use a versioned cache name with the commit SHA, and show a "New version available — reload" toast on update |
| Environment | No secrets are needed for the GPL route. If a commercial HyperFormula key is used instead, **do not** commit it. Note that any key shipped to a public static site is visible to visitors; discuss with the owner |
| Build metadata | Inject `VITE_COMMIT_SHA`, `VITE_BUILD_DATE`, and `VITE_APP_VERSION` at build time; show them on `/about` and `/verification` |

### 17.4 CI/CD workflow (`.github/workflows/deploy.yml`)
Use the current major versions of the official actions at implementation time. The versions below were verified against each action’s latest release on 2026-09-29. Node comes from `.nvmrc` (Node 24; `package.json` has `engines.node: ">=24"`).

```yaml
name: CI and Deploy

on:
  push:
    branches: [main]
  pull_request:
  workflow_dispatch:

# Least privilege by default; only the deploy job can write to Pages (OPEN_ISSUES #5).
permissions:
  contents: read

concurrency:
  group: pages-${{ github.ref }}
  # Superseded PR runs are cancelled; a deploy on main is never interrupted mid-flight.
  cancel-in-progress: ${{ github.event_name == 'pull_request' }}

jobs:
  ci:
    runs-on: ubuntu-latest
    env:
      BASE_PATH: /${{ github.event.repository.name }}/
      VITE_COMMIT_SHA: ${{ github.sha }}
    steps:
      - uses: actions/checkout@v7
      - uses: actions/setup-node@v7
        with:
          node-version-file: .nvmrc
          cache: npm
      - run: npm ci
      - name: Content is in sync with the Labs document
        run: npm run extract -- --check
      - name: Third-party notices are in sync with package-lock.json
        run: npm run notices -- --check
      - run: npm run lint
      - run: npm run format:check
      - run: npm run typecheck
      - name: Unit + golden cross-check tests
        run: npm test -- --reporter=default --reporter=json --outputFile=reports/vitest.json
      - name: Write public/verification.json from the conformance results
        run: npm run verification:write
      - run: npm run build
      - run: npx playwright install --with-deps chromium
      - name: E2E tests against the production build under BASE_PATH
        run: npm run test:e2e
      - uses: actions/upload-artifact@v7
        if: always()
        with:
          name: test-reports
          path: reports/
          if-no-files-found: ignore
      - name: Upload Pages artifact
        if: github.ref == 'refs/heads/main' && github.event_name != 'pull_request'
        uses: actions/upload-pages-artifact@v5
        with:
          path: dist

  deploy:
    needs: ci
    if: github.ref == 'refs/heads/main' && github.event_name != 'pull_request'
    runs-on: ubuntu-latest
    permissions:
      pages: write
      id-token: write
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    steps:
      - uses: actions/configure-pages@v6
      - id: deployment
        uses: actions/deploy-pages@v5
```
Rules:
- **Deploy only if every test passes.** A failing conformance test must block publishing.
- Playwright e2e tests run against the **built** site served under `BASE_PATH`, to catch sub-path bugs before they reach production.
- **Least privilege:** the workflow default is `contents: read`; only the `deploy` job gets `pages: write` and `id-token: write`, so pull-request runs never hold Pages credentials.
- **Concurrency:** superseded pull-request runs are cancelled, but a run on `main` is never cancelled, so a Pages deployment is never interrupted mid-flight.
- The `verification:write` step (Section 17.5) runs after the tests and before the build, so `verification.json` ships with the site.
- Add `npm run smoke:prod`: a scheduled workflow (weekly, plus manual) that loads the live URL, opens Lab 2, and confirms the checks are green. It alerts through a failed run if the public site breaks.

### 17.5 Verification artifact and page
**`public/verification.json`** (written in CI by `scripts/write-verification.ts`):
```json
{
  "assertions": { "passed": 277, "total": 277 },
  "experiments": { "passed": 12, "total": 12 },
  "engine": { "name": "HyperFormula", "version": "x.y.z", "plugins": ["AVERAGEIFS", "LOOKUP", "..."] },
  "commit": "abc1234",
  "builtAt": "2026-10-15T18:00:00Z",
  "label": "277/277 passing"
}
```
**`/#/verification` page:**
- **Headline:** "277 of 277 expected values reproduced · 12 of 12 what-if experiments reproduced", with the commit and date.
- **"Run verification in this browser" button:** runs the conformance suite client-side (Web Worker), showing a progress bar and a results table filterable by lab and pass/fail. Target: under 5 seconds on a laptop.
- **How the expected values were produced:** a short explanation that the workbook was independently recalculated in a separate spreadsheet engine and cross-checked against Python calculations (quote the Labs document's verification note), and that this app re-derives them with its own engine.
- **Link to the conformance test source** on GitHub.

**README badge:** a Shields.io dynamic JSON badge reading `verification.json`:
`https://img.shields.io/badge/dynamic/json?url=https%3A%2F%2Fjiaxingxue.github.io%2Ffinance-excel-labs%2Fverification.json&query=%24.label&label=verified`

### 17.6 README template (`README.md`)
Claude Code fills the placeholders and keeps this structure. Screenshots come from 17.7.

```markdown
# Excel Labs for Financial Analysis

[![CI and Deploy](https://github.com/jiaxingxue/finance-excel-labs/actions/workflows/deploy.yml/badge.svg)](https://github.com/jiaxingxue/finance-excel-labs/actions/workflows/deploy.yml)
[![verified](https://img.shields.io/badge/dynamic/json?url=https%3A%2F%2Fjiaxingxue.github.io%2Ffinance-excel-labs%2Fverification.json&query=%24.label&label=verified)](https://jiaxingxue.github.io/finance-excel-labs/#/verification)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](LICENSE)

**Live demo → https://jiaxingxue.github.io/finance-excel-labs/**

Interactive, auto-graded Excel labs for FP&A and accounting: budget-vs-actual,
price-volume-mix, bank reconciliation, AR aging, forecasting, and driver-based scenarios.
Formulas run live in the browser; every result is checked against 277 independently
verified expected values.

![Build mode grading a SUMIFS formula](docs/media/build-mode.gif)

## Try it in 60 seconds
1. Open **Lab 7 — Bank Reconciliation** → *What-if* → drag **Amount tolerance** to 1.00.
   The reconciliation still ties to zero, but the items are now misclassified:
   *balanced isn't the same as correct.*
2. Open **Lab 9 — Forecasting** → *What-if* → **Inject outlier**. Watch the forecast error triple.
3. Open **Lab 2** → *Build* → type a hard-coded number instead of a formula. It is rejected.

## What's inside
| Area | Labs |
|---|---|
| Variance analysis | Budget vs. actual, flux (MoM/YoY), price-volume-mix, rate/volume, flexible budget, pivots |
| Reconciliations | Bank rec with automated matching, AR aging and subledger-to-GL tie-out |
| Forecasting | Moving average, trend × seasonality with backtest (MAPE/bias), driver-based scenarios |
| Model controls | Integrity-checks dashboard, reusable function library |

## How correctness is guaranteed
- 277 expected values and 12 what-if experiments, derived independently of this app
- CI blocks deployment unless every one is reproduced ([verification page](https://jiaxingxue.github.io/finance-excel-labs/#/verification))
- Grading checks *results*, not formula text, and re-tests answers under changed inputs to reject hard-coding

## Tech
React · TypeScript · Vite · HyperFormula (GPLv3) with custom Excel-compatible function plugins ·
ExcelJS export · Vitest · Playwright · GitHub Actions → GitHub Pages. Fully static, works offline.

## Screenshots
| Explore | What-if | Verification |
|---|---|---|
| ![](docs/media/explore.png) | ![](docs/media/whatif.png) | ![](docs/media/verification.png) |

## Run locally
    npm ci && npm run extract && npm run dev
    npm test            # unit + conformance
    npm run test:e2e    # end-to-end

## Docs
[Product requirements](docs/PRD.md) · [Architecture](docs/ARCHITECTURE.md) · [Demo script](docs/DEMO.md)

## Author
<Name from D5, or @jiaxingxue> — <one-line background, only if the owner provides it> <links from D6, if provided>

## License
GPLv3. See [LICENSE](LICENSE) and [third-party notices](THIRD_PARTY_NOTICES.md).
```
Also commit this PRD as `docs/PRD.md`. It demonstrates product thinking to technical reviewers.

### 17.7 Automated screenshots and GIF
- `npm run media` (Playwright) opens the **production build**, sets the viewport to 1440×900, light theme, and fixed data, then captures:
  - `explore.png`: Lab 2, cell `BvA!C6` selected, precedents highlighted
  - `whatif.png`: Lab 7 with tolerance = 1.00 and the "balanced ≠ correct" callout visible
  - `verification.png`: after "Run verification in this browser" completes
  - `build-mode.gif`: typing the `SUMIFS` formula, Fill Down, checks turning green (record a video with Playwright, then convert with ffmpeg; keep it under 3 MB and 15 s)
  - `og-image.png` (1200×630): app name, a grid screenshot, and "277/277 verified"
- Save the output to `docs/media/`. Re-run whenever the UI changes. CI does not regenerate media, to avoid noisy commits.

### 17.8 Demo scripts (`docs/DEMO.md`)

**3-minute version (interview default)**
| Time | Action | Say |
|---|---|---|
| 0:00 | Open the live URL on the home page | "Interactive Excel labs for FP&A. The formulas run live in the browser and every result is auto-verified." |
| 0:20 | Lab 2 → click `BvA!C6` | "Budget pulled from a long-format ledger with SUMIFS. The highlighted cells are its inputs." |
| 0:45 | Build mode → type `=500000` in `C6` → rejected; type the SUMIFS → accepted | "Grading checks results, and re-tests under changed inputs, so hard-coding can't pass. That's a control, not just a quiz." |
| 1:30 | Lab 7 → What-if → tolerance to 1.00 | "The rec still ties to zero, but items are misclassified. In banking and compliance work, a zero difference isn't evidence by itself; you review the nature of each reconciling item." |
| 2:15 | Lab 9 → Inject outlier | "One bad data point halves the growth estimate and triples the error, which is why you scan history before forecasting." |
| 2:40 | `/#/verification` → Run in browser | "277 expected values and 12 experiments reproduced right here. CI blocks deploys if any fail." |

**10-minute version (technical interview):** the 3-minute path, then:
1. Show the conformance test and one custom function plugin (`LOOKUP`, including the `LOOKUP(2,1/…)` trick) in the repo.
2. Show the GitHub Actions run: test → build → e2e → deploy gate.
3. Export `.xlsx` and open it in Excel to show the formulas are intact.
4. Walk through the PRD's risk table and the license decision (GPL vs. commercial).
5. The Modern Excel view: Tier A vs. `LET`/`LAMBDA`, and why the classic formulas are the graded, portable baseline.

### 17.9 Interview talking points (`docs/INTERVIEW_NOTES.md`)
Claude Code creates this file with these headings and fills them from the finished implementation:
- **Problem:** people learn finance modeling by reading, not doing; there was no safe, graded sandbox.
- **Domain judgment:** materiality rules (AND vs. OR), sign conventions, "balanced ≠ correct," backtesting against a naive benchmark.
- **Controls mindset:** the Checks sheet, conformance gate, anti-hard-coding, and verification page mirror internal-control design (preventive + detective).
- **Engineering decisions:** engine adapter, custom plugins instead of rewriting formulas, value-based grading, hash routing for static hosting, offline support.
- **Trade-offs and what's next:** real Excel embedding (Office.js), telemetry (opt-in), more labs (consolidation, cash flow).
- **Numbers to quote:** labs, formulas, assertions, experiments, test count, Lighthouse scores, recalculation time. Pull these from `verification.json` and the CI reports, never from memory.

### 17.10 Pre-interview checklist (`docs/DEMO.md`, bottom)
- [ ] Latest `main` run is green, and the site shows today's commit on `/about`
- [ ] Open the live URL once on the demo laptop **and** phone, so the service worker caches it for offline use
- [ ] Walk through the 3-minute script once, with Wi-Fi off
- [ ] Download a fresh `.xlsx` export and open it in Excel (for the 10-minute version)
- [ ] Browser zoom at 110–125% for screen sharing; close other tabs; light theme
- [ ] Backup: `docs/media/demo.mp4` (a 3-minute recording made with the Playwright video path) ready to play locally
- [ ] Have the README and the verification page open in tabs

### 17.11 Polish for a public audience
| Item | Requirement |
|---|---|
| Social preview | `<meta property="og:title|og:description|og:image">` and Twitter card tags. Absolute URLs built from the Pages URL. Upload `og-image.png` as the repo's social preview (Settings → General) |
| Favicon / title | `favicon.svg`; the page title changes per route ("Lab 7 — Bank Reconciliation · Excel Labs") |
| Landing clarity | Home hero states what the app is in one sentence, with two buttons: **Start Lab 1** and **Take the 60-second tour** (the README "Try it" steps as an in-app guided tour) |
| Footer | Author name, GitHub link, license, version + commit. No email or phone number (D6) |
| Robots | `robots.txt` allows indexing; add `sitemap.xml` listing the lab routes (optional) |
| Error page | Friendly error boundary with "Reset this lab" and a link to report an issue on GitHub |
| Performance | Lighthouse ≥ 90 for performance and ≥ 95 for accessibility on the **deployed** URL (checked in the smoke workflow) |

### 17.12 Licensing and attribution for the public repo
- Repository license: **GPLv3** (confirmed, D1), consistent with HyperFormula's `gpl-v3` license key. The in-app `/about` page states: "Formula engine: HyperFormula (GPLv3) by Handsontable."
- `THIRD_PARTY_NOTICES.md` lists every bundled dependency and its license (ExcelJS MIT, React MIT, etc.).
- Lesson content (the Labs document) is the owner's own material, included under the same repository license unless the owner chooses a separate content license (e.g., CC BY 4.0). Ask.
- Show the Labs document's verification statement unchanged, plus the note that the app calculates with an Excel-compatible engine, not Microsoft Excel. Make **no claim** of affiliation with Microsoft or Handsontable.
- *Note for the owner:* this is a practical reading of the licenses, not legal advice.

---

### References
- HyperFormula: [built-in functions](https://hyperformula.handsontable.com/guide/built-in-functions.html), [licensing](https://hyperformula.handsontable.com/docs/guide/licensing.html), [license key](https://hyperformula.handsontable.com/docs/guide/license-key.html), [custom functions](https://hyperformula.handsontable.com/guide/custom-functions.html), [arrays](https://hyperformula.handsontable.com/docs/guide/arrays.html), [GitHub](https://github.com/handsontable/hyperformula)
- IronCalc: [GitHub](https://github.com/ironcalc/IronCalc), [website](https://www.ironcalc.com/)
- GitHub Pages / Actions: [Pages documentation](https://docs.github.com/pages), [Using custom workflows with GitHub Pages](https://docs.github.com/en/pages/getting-started-with-github-pages/using-custom-workflows-with-github-pages)
- Vite: [Deploying a static site](https://vite.dev/guide/static-deploy)

---

## Change log

| Version | Date | Change |
|---|---|---|
| 1.1 | 2026-09-29 | Added Section 17 (public deployment and portfolio) |
| 1.2 | 2026-09-29 | Recorded the owner's decisions: GPLv3, repo `finance-excel-labs` |
| 1.3 | 2026-09-29 | Milestone 0 resolutions accepted by the owner (details in `docs/OPEN_ISSUES.md` #1–#12): hand-authored config moves to `src/config/labs.config.ts` and `src/data/` becomes generated-only, with `notice.json` and `manifest.json` added (§4.2, §4.5, §9.2); `workbook.json` keeps `{date}` values, converted to serials at engine load (§4.3); experiment `changes` convert only full `YYYY-MM-DD` strings, and `expect: null` means empty (§4.4); Appendix D/E JSON extraction is scoped to its own section (§4.2); local sub-path builds use `npm run build:pages` via `cross-env` (§17.3); the workflow uses Node 24 from `.nvmrc`, current action majors, job-scoped Pages permissions, and never cancels a `main` deploy (§17.4) |
| 1.4 | 2026-09-30 | Milestone 1: recorded which HyperFormula gaps are real, added the `INDEX` and `TEXT` overrides to the plugin table (§7.2, §7.4), and added the `verification:write` step to the workflow (§17.4) |
| 1.5 | 2026-09-30 | Milestone 2 kickoff, owner-approved (`docs/OPEN_ISSUES.md` #20–#24): the grid requirements are split across milestones. GR-7, GR-12, LS-2, and LS-3 move to M3; GR-8, GR-9, and GR-10 move to M4; Zustand is deferred to M5 (§9.1, §12). Also recorded: cell-link rules for bare and `$` references (#20), tiers hand-authored from Part 0.3 (#21), disabled "coming soon" mode tabs (#22), and an axe-core check in e2e (#24) |
| 1.6 | 2026-09-30 | Milestone 3 kickoff, owner-approved (`docs/OPEN_ISSUES.md` #28–#33): What-if control ranges settled, Lab 6 gets period and version dropdowns, Lab 7's add-a-row toggle is dropped, Lab 8's color scale is added and Lab 2's icon set skipped, charts are hand-rolled SVG. M3 ships as two PRs, M3a and M3b (§12) |
