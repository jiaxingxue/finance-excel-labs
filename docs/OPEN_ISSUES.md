# Open Issues

Known gaps, contradictions, and questions. Each has a status and, when resolved, where the resolution is recorded. Owner decisions live in [`DECISIONS.md`](DECISIONS.md).

**Status:** `RESOLVED` (decided and applied) · `DEFERRED` (resolution agreed; applied in a later milestone) · `OPEN` (needs a decision).

## Found during Milestone 0 review (2026-09-29)

All 12 resolutions were accepted by the owner on 2026-09-29.

| # | Issue | Resolution | Status | Where |
|---|---|---|---|---|
| 1 | **Node version.** The PRD workflow used Node 22 and older action majors; the owner develops on Node 24. | Node 24 everywhere: `.nvmrc` = `24`, `package.json` `engines.node: ">=24"` (no upper bound, so future upgrades don't break installs), and the workflow reads `.nvmrc`. Actions updated to the latest majors as of 2026-09-29 (checkout v7, setup-node v7, upload-artifact v7, upload-pages-artifact v5, configure-pages v6, deploy-pages v5). | RESOLVED | PRD §17.4, `.github/workflows/deploy.yml` |
| 2 | **Date format.** PRD §4.3 says dates become Excel serials, but `workbook.json` must deep-equal `reference/golden/spec.json`, which stores `{"date": "YYYY-MM-DD"}`. | `workbook.json` keeps `{date}` objects; the engine adapter converts them to 1900-system serials at load time. | RESOLVED (conversion applied in M1) | PRD §4.3, `src/engine/hyperformulaEngine.ts` |
| 3 | **`labs.config.ts` location.** PRD put the hand-authored config in `src/data/`, which CLAUDE.md says is entirely generated. | Moved to `src/config/labs.config.ts`. `src/data/` is generated-only; `npm run extract` deletes and rewrites it, and `--check` flags extra files. | RESOLVED | PRD §4.5, §9.2 |
| 4 | **POSIX-only env syntax.** `verify.md` and CLAUDE.md used `BASE_PATH=… npm run build`, which fails in PowerShell. In Git Bash it's worse: MSYS rewrites `/finance-excel-labs/` to `C:/Program Files/Git/finance-excel-labs/` (confirmed on the owner's machine). | `npm run build:pages` / `preview:pages` set the variable with `cross-env`. `verify.md` and CLAUDE.md updated. | RESOLVED | PRD §17.3, CLAUDE.md, `.claude/commands/verify.md` |
| 5 | **Workflow permissions and concurrency.** The PRD workflow granted `pages: write` and `id-token: write` to every job (including PR runs), and `cancel-in-progress: true` could cancel a Pages deploy mid-flight. | Workflow default `contents: read`; only `deploy` gets Pages permissions. Cancel in-progress runs only for pull requests. | RESOLVED | PRD §17.4, `.github/workflows/deploy.yml` |
| 6 | **`null` in experiment `expect`.** E3.1 expects `null` for `Flux!I8`, `I9`, `I10`, `I12`, `I13`; PRD §4.4 didn't define it. The lesson table shows those cells blank, and the matching assertions expect `""`. | `null` means empty: passes when the cell evaluates to an empty value or `""`. | RESOLVED (applied in M1) | PRD §4.4, `src/grading/match.ts` |
| 7 | **Text vs. date strings in experiment `changes`.** `AR!H2 = "2026-09-30"` is a date; `BvA!B1 = "2026-07"` (E2.2) is a text period. | Only a full `YYYY-MM-DD` string converts to a date serial; every other string stays text. | RESOLVED (applied in M1) | PRD §4.4, `src/grading/match.ts` |
| 8 | **JSON block scoping.** Appendix B.2 also contains JSON code blocks, so "the single fenced JSON array" is ambiguous document-wide. | The extractor looks only inside the Appendix D and E sections and fails unless exactly one `json` block is there (covered by a test). | RESOLVED | PRD §4.2, `scripts/extract/content.ts` |
| 9 | **What-if control ranges.** Appendix B.4 gives no min/max/step for most controls (Flux sliders `B4`/`B5`, Drivers sliders `B4:D6`, BankRec numeric inputs), but WI-1 requires them. | These are editorial values. Propose them to the owner in M3 when the controls are built. | DEFERRED (M3) | — |
| 10 | **Lab 11 `FAVFLAG` has an extra branch.** Its LAMBDA returns `"No budget"` when budget = 0; the workbook's `BvA!G6` has no such branch (only the hardened version shown after E2.2 does). | The simulated `FAVFLAG` follows the Lab 11 LAMBDA definition exactly. The Lab 11 test call (`BvA!G9`, budget ≠ 0) is unaffected. | DEFERRED (M5) | — |
| 11 | **Lab 1 helper column.** The lesson's `GL!F2` "UNMAPPED" completeness formula isn't in Appendix C (`GL` is exactly 37 rows × 5 columns = 185 cells). | No change: it's a lesson-only step. This is consistent with Lab 1 having no assertions and no graded Build mode (G2 covers Labs 2–10 and 12). | RESOLVED (no action) | — |
| 12 | **`verification:write` in M0.** The PRD workflow writes `public/verification.json` from conformance results, which don't exist until M1. | M0 ships every other workflow step; M1 adds `verification:write` (marked by a comment in the workflow). | RESOLVED (M1: runs after the tests, before the build) | PRD §17.4, `scripts/write-verification.ts` |

## Other notes from Milestone 0

| # | Issue | Resolution | Status |
|---|---|---|---|
| 13 | **Author email in git history.** The starter commit exposed a personal email. | The owner rewrote it to the GitHub noreply address and force-pushed (`4be5f05`). Local git config uses the noreply address. | RESOLVED |
| 14 | **`docs/DECISIONS.md` was gitignored**, so cloud sessions and CI couldn't see it. | Committed; removed from `.gitignore`. `START_HERE.md` stays ignored. | RESOLVED |
| 15 | **TypeScript 7.x can't be used yet.** `typescript-eslint` 8.71 supports TypeScript `>=4.8.4 <6.1.0`. | Pinned `typescript@~6.0.3`. Revisit when `typescript-eslint` supports 7.x. | OPEN (tooling; no owner action) |

## Found during Milestone 1 (2026-09-30)

| # | Issue | Resolution | Status | Where |
|---|---|---|---|---|
| 16 | **HyperFormula 3.4.0 function gaps.** The PRD §7.2 list was a hypothesis. Tested against the full workbook, stock HyperFormula passed 179/277 assertions. | Real gaps: `AVERAGEIFS`, `LOOKUP`, `INTERCEPT`, `TREND`, `FORECAST.LINEAR` (+ `FORECAST`) are missing. `TEXT` exists but ignores format codes. `INDEX(one-row range, n)` returns `#NUM!`. All seven get plugins. Two more gaps are fixed in configuration: bare `TRUE`/`FALSE` literals and `smartRounding`. Result: 277/277, 12/12. | RESOLVED | `docs/milestones/M1.md`, `src/engine/plugins/` |
| 17 | **HyperFormula crashes on formulas that reference a missing sheet.** If a cell's formula references a range on a sheet that doesn't exist (e.g. `=SUM(Ledger!E2:E37)`; sheet names are case-insensitive, so `Gl!` is not a missing sheet but `GL`, as in Excel; see #25), overwriting that cell throws an internal `Range does not exist` error. Reproduced in 3.4.0; references to existing sheets are unaffected. | The adapter rejects such formulas with a clear error (`unknown sheet "Ledger"`) in `load` and `setCell`, so they never reach HyperFormula. Build mode (M4) must show this message to the learner. Could be reported upstream. | RESOLVED (workaround) | `src/engine/hyperformulaEngine.ts`, `tests/engine.test.ts` |
| 18 | **Plugin scope limits.** No lab formula uses these: `LOOKUP` assumes ascending data (on unsorted data Excel's binary search can land on a different position); `TREND` supports only one `new_x` value and requires `known_x`; `INDEX` with row or column 0 returns `#VALUE!` (Excel returns the whole row or column). | Documented in each plugin's header. Revisit if Build mode (M4) or D8 lets learners type free-form formulas that hit them. | OPEN (no owner action) | `src/engine/plugins/` |
| 19 | **Floating-point noise in unformatted cells.** With `smartRounding: false` (#16), the engine returns full IEEE-754 doubles, so a cell with no `fmt` can hold e.g. `0.9999999999999999` (`Checks!C8`, the only such cell among the 58 unformatted numeric formula cells) or `0.30000000000000004`. Grading must keep the full value; only display is affected. | M2 display rule: a cell without `fmt` renders with SSF `"General"`, as Excel does (`0.30000000000000004` → `0.3`, `0.9999999999999999` → `1`, `1/3` → `0.333333333`). The formula bar and inspector show the value rounded to 15 significant digits (`Number(v.toPrecision(15))`), which is Excel's own precision limit. The engine and grader never round. M2 adds a unit test for both. | RESOLVED (M2) | `src/format/displayValue.ts`, `tests/display.test.ts`, `tests/e2e/lab.spec.ts` |

## Decided at the Milestone 2 kickoff (2026-09-30)

The owner approved the M2 component structure and layouts with these answers.

| # | Issue | Resolution | Status | Where |
|---|---|---|---|---|
| 20 | **Which lesson references become cell links (GR-11).** Lessons often cite cells without a sheet (`B1`, `A5:H5`), and many code spans are formula fragments (`$A6`, `Map!$B$2:$B$7`). | (1) Qualified references on a known sheet (`BvA!C6`, `Map!A2:A7`) are links in every lab, in code spans and in plain text. (2) Bare references are links only in labs with exactly one primary sheet (2, 3, 4, 5, 6, 9, 10, 12) and resolve to that sheet; in Labs 1, 7, 8, and 11 they stay text. (3) Anything containing `$` is never a link. A bare reference is linked only when a whole code span is the reference: in running text, strings such as `E7.1` (an experiment id) would be misread. | RESOLVED (M2) | `src/lesson/cellRefs.ts`, `tests/lesson-links.test.ts` |
| 21 | **Tier badge source (EX-2d).** No generated file carries tiers. | Hand-authored `src/config/tiers.ts`, transcribed from Part 0.3. A test parses the Part 0.3 table in the Labs document and requires an exact match. Functions the table doesn't name (`SUM`, `AVERAGE`, `COUNTIF`, `FORECAST`, `LINEST`, and the `FORECAST.ETS` helpers) are listed separately. A formula's tier is its newest function's tier; plain arithmetic is Tier A. | RESOLVED (M2) | `src/config/tiers.ts`, `tests/tiers.test.ts` |
| 22 | **Mode tabs before their modes exist.** | Build, What-if, Challenge, and Modern show as disabled tabs marked "coming soon", so the layout doesn't shift when they ship. | RESOLVED (M2) | `src/components/workspace/LabHeader.tsx` |
| 23 | **Milestone mapping of grid features.** PRD §12 put every grid requirement in M2. | M2 ships GR-1–GR-4, GR-6, GR-11, GR-13, undo/redo, EX-1–EX-4 (except EX-2c), LS-1, and LS-4. GR-7 (precedent outlines), GR-12 (conditional formatting), LS-2, and LS-3 move to **M3** with What-if. GR-8 (fill), GR-9 (copy/paste), and GR-10 (autocomplete) move to **M4** with Build mode. **Zustand** is deferred to **M5** with persistence; M2 uses one store read with `useSyncExternalStore`. | RESOLVED | PRD §9.1, §12 (v1.5) |
| 24 | **Automated accessibility check.** | `@axe-core/playwright` 4.13 runs on Lab 2 in the e2e suite and fails on any serious or critical WCAG 2.1 A/AA violation. It is dev-only and **MPL-2.0**, as is its `axe-core` dependency; neither is bundled in the site. | RESOLVED (M2) | `tests/e2e/lab.spec.ts`, `THIRD_PARTY_NOTICES.md` |

## Found during Milestone 2 (2026-09-30)

| # | Issue | Resolution | Status | Where |
|---|---|---|---|---|
| 25 | **Sheet names are case-insensitive.** HyperFormula, like Excel, resolves `Gl!E2` to `GL!E2`. The adapter's unknown-sheet check was already case-insensitive; only the example in #17 was wrong. | #17 now uses `Ledger!` as its example. The M2 e2e test types `=SUM(Ledger!E2:E37)` and checks that the learner sees the rejection message. | RESOLVED | #17, `tests/e2e/lab.spec.ts` |
| 26 | **EX-2(c): per-cell explanations in the inspector.** The lesson's "How the key formulas work" bullets name cells informally (`SUMIFS` (C6, D6)), so linking a bullet to a cell needs either a heuristic or new extracted data. | Deferred. The M2 inspector shows the formula, value, 15-digit value, tier, and the lesson's pattern-cell result. | OPEN (target M3 with LS-3; needs an owner choice between a heuristic and an extractor change) | — |
| 27 | **`verification.json` in local e2e runs.** The new e2e test fetches it from the build (the 404.html fallback would otherwise hide a missing file), so a local `/verify` must generate it the way CI does. | `/verify` step 4 now runs `npm test` with the JSON reporter, then `npm run verification:write`. The test's failure message explains how to generate the file. | RESOLVED | `.claude/commands/verify.md`, `tests/e2e/verification.spec.ts` |

## Decided at the Milestone 3 kickoff (2026-09-30)

The owner approved the M3 proposal (control ranges, challenge cards, charts) with these answers and adjustments. M3 ships as two PRs from one plan: **M3a** (`m3a-whatif`: What-if controls, GR-12, formula-bar commit on blur and label tooltips, home banner) and **M3b** (`m3b-challenge`, branched after M3a merges: challenge cards, charts CT-1–CT-4, GR-7, LS-2/LS-3, EX-2(c)).

| # | Issue | Resolution | Status | Where |
|---|---|---|---|---|
| 28 | **What-if control ranges (resolves #9).** | Approved as proposed, with two adjustments: `CostVar!B18`/`B19` use format `#,##0.00` (they step by 0.5), and Lab 6 gets dropdowns for `PivotLab!B2` (2025-08, 2026-07, 2026-08) and `PivotLab!B3` (ACT, BUD), which Appendix B.4 doesn't list. Rules: every range includes the base value and every Appendix E change value; sliders land exactly on them; Lab 10 churn starts at 0.5% because `Drivers!B21:D21` divide by churn. Sliders also have a typed value box and a tick at the base value. A unit test parses Appendix B.4 and requires every listed cell to have a control. | RESOLVED (M3a) | `src/config/labs.config.ts`, `tests/whatif.test.ts` |
| 29 | **Lab 7 "toggle to add a transaction row" (Appendix B.4).** The matching formulas use fixed ranges (`Bank!$D$2:$D$5`, `Book!$C$2:$C$6`), so a new row would be ignored, and widening the lab's formulas is not allowed. | Left out, by owner decision. Revisit only if the Labs document widens the ranges. | RESOLVED (not built) | `src/config/labs.config.ts` (comment) |
| 30 | **Conditional formats beyond formula and threshold rules.** | Lab 8's green→red color scale on `AR!E2:E7` is included. Lab 2's 3-arrow icon set on `E6:E12` is skipped (the lesson calls it "visual only"). Lab 7's "green when `=$B$19=0`, red otherwise" is two rules: `=$B$19=0` → green and `=$B$19<>0` → red. The Lab 12 suggestion to color the master check on every output sheet is not a lab rule and is not built. | RESOLVED (M3a) | `src/config/labs.config.ts` |
| 31 | **Charts.** | Hand-rolled SVG, no chart library. | DEFERRED (M3b) | — |
| 32 | **Challenge prompt text.** The Labs document has no question text for experiments. | Cards may show a generic UI prompt ("What do you expect to happen?") above the prediction step; no other new wording. | DEFERRED (M3b) | — |
| 33 | **Home banner.** The owner asked for "interactive in Explore mode, with What-if, Challenge, and Build coming next". | Because M3a ships What-if, the banner says Explore and What-if are live and Challenge and Build come next. The README status line matches. | RESOLVED (M3a) | `src/routes/Home.tsx`, `README.md` |

## Found during Milestone 3a (2026-09-30)

| # | Issue | Resolution | Status | Where |
|---|---|---|---|---|
| 34 | **Checks fail while an input is moved.** The checks grade against the Labs document's base inputs, so in What-if mode moving e.g. `BvA!B2` turns a check red. That is correct, but surprising. | The What-if panel shows a note while any input differs from its base: "Checks compare with the Labs document's base inputs…". "Reset inputs" restores them. | RESOLVED | `src/components/whatif/WhatIfPanel.tsx` |
| 35 | **Key outputs (WI-2) are not hand-picked.** | The delta list is every cell the lab's experiments watch (Appendix E `expect`), in order. Lab 6 has no experiments, so it shows a note instead. | RESOLVED | `src/whatif/controls.ts` |
| 36 | **Every slider step is an undo step.** Dragging writes the cell once per animation frame (WI-2's ≤ 50 ms), and each write is its own undo step, so Ctrl+Z after a drag walks back through the intermediate values. "Reset inputs" restores every input in one step. | Accepted for M3a. Revisit with persistence in M5 if it bothers testers (e.g. commit a drag as one undo step on release). | OPEN (no owner action) | `src/components/whatif/controls.tsx` |
