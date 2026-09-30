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
| 17 | **HyperFormula crashes on formulas that reference a missing sheet.** If a cell's formula references a range on a sheet that doesn't exist (e.g. a typo such as `=SUM(Gl!E2:E37)`), overwriting that cell throws an internal `Range does not exist` error. Reproduced in 3.4.0; references to existing sheets are unaffected. | The adapter rejects such formulas with a clear error (`unknown sheet "Gl"`) in `load` and `setCell`, so they never reach HyperFormula. Build mode (M4) must show this message to the learner. Could be reported upstream. | RESOLVED (workaround) | `src/engine/hyperformulaEngine.ts`, `tests/engine.test.ts` |
| 18 | **Plugin scope limits.** No lab formula uses these: `LOOKUP` assumes ascending data (on unsorted data Excel's binary search can land on a different position); `TREND` supports only one `new_x` value and requires `known_x`; `INDEX` with row or column 0 returns `#VALUE!` (Excel returns the whole row or column). | Documented in each plugin's header. Revisit if Build mode (M4) or D8 lets learners type free-form formulas that hit them. | OPEN (no owner action) | `src/engine/plugins/` |
