# CLAUDE.md — Excel Labs for Financial Analysis (`jiaxingxue/finance-excel-labs`)

Interactive, auto-graded Excel labs for FP&A and accounting, running an Excel-compatible formula engine in the browser. Static site deployed to GitHub Pages as a public portfolio project: https://jiaxingxue.github.io/finance-excel-labs/

## Read before any work
1. `docs/DECISIONS.md` — the owner's answers (license, repo, author). They override the PRD's defaults.
2. `docs/PRD.md` — full requirements. Section numbers below refer to it.
3. `docs/OPEN_ISSUES.md` — known gaps and questions (create it if missing).

## Non-negotiable rules
- **Source of truth:** `content/Excel_Implementation_Labs_Financial_Analysis.md`. Never invent formulas, numbers, or lesson text. Data files in `src/data/` are **generated** by `npm run extract`; never hand-edit them.
- **Golden cross-check:** `npm run extract` output must deep-equal `reference/golden/` (spec sheets, 277 assertions, 12 experiments). If they differ, the extractor is wrong. Do not "fix" the golden files.
- **Milestone order** (PRD §12). **M1 is a hard gate:** 277/277 assertions and 12/12 experiments must pass in the engine before any UI work beyond the M0 placeholder.
- **Never hard-code expected values** in app logic or tests. Tests load them from `src/data/assertions.json` / `experiments.json`.
- **Never rewrite lab formulas** to dodge an engine gap. Add a custom function plugin in `src/engine/plugins/` (PRD §7.4) with unit tests.
- **Grading is value-based**, with the anti-hard-coding check (PRD §6.3, BD-4).
- **License:** GPLv3 is confirmed (D1). Use `licenseKey: 'gpl-v3'`, ship the full GPLv3 `LICENSE`, and list dependency licenses in `THIRD_PARTY_NOTICES.md`. If D1 ever changes, stop public deploys until it's re-confirmed.
- **Privacy:** never commit an email address, phone number, API key, or license key. No analytics or network calls beyond static assets.
- Static hosting: hash routing; every data URL is prefixed with `import.meta.env.BASE_URL` (PRD §17.3).

## Workflow
- One branch per milestone (`m0-scaffold`, `m1-engine`, …). Small commits with clear messages.
- Before every commit: `npm run check` (= `npm run lint && npm run typecheck && npm test`). Before merging a milestone: run `/verify`.
- End each milestone by writing `docs/milestones/M<n>.md` (what works, what's left, known issues) and pausing for the owner's review.
- When blocked by an owner decision, add it to `docs/OPEN_ISSUES.md`, choose the documented default if one exists, and say so. Otherwise stop and ask.
- Prefer small, readable modules. TypeScript strict mode. No new runtime dependency without noting its license in `THIRD_PARTY_NOTICES.md`.

## Commands
<!-- Claude Code: fill these in during M0 and keep them current. -->
| Task | Command |
|---|---|
| Install | `npm ci` (Node 24, see `.nvmrc`; `.npmrc` sets `engine-strict`) |
| Extract content | `npm run extract` (`-- --check` fails if the output is stale) |
| Dev server | `npm run dev` |
| Unit + conformance tests | `npm test` |
| Pre-commit gate | `npm run check` (lint + typecheck + test; works in any shell) |
| E2E tests | `npm run build:pages` then `npm run test:e2e` (one-time: `npx playwright install chromium`) |
| Milestone screenshots | `npm run build:pages` then `npm run screenshots:m2` or `screenshots:m3` (writes `docs/milestones/m2/` or `m3/`) |
| Lint / types / format | `npm run lint` / `npm run typecheck` / `npm run format:check` |
| Production build | `npm run build` (base `/`) or `npm run build:pages` (base `/finance-excel-labs/`, mimics Pages) |
| Third-party notices | `npm run notices` (`-- --check` fails if stale); run after any dependency change |

The owner uses PowerShell on Windows; CI runs on Linux. Put any environment variable in an npm script with `cross-env`, never inline (`BASE_PATH=… npm run build`). The inline form fails in PowerShell, and Git Bash rewrites `/finance-excel-labs/` to `C:/Program Files/Git/finance-excel-labs/`.

## Project map
<!-- Claude Code: update as the structure is created (PRD §9.2). -->
- `content/` source Labs document · `reference/golden/` verified data (read-only) · `reference/generator/` scripts that produced it (provenance only; do not run in CI)
- `scripts/extract-content.ts` + `scripts/extract/` content pipeline · `scripts/generate-notices.ts`
- `src/data/` (generated only; never edit) · `src/config/labs.config.ts` (hand-authored lab config) · `src/content/types.ts` (data shapes)
- `src/engine/` formula engine adapter + plugins · `src/grading/` · `src/components/` · `src/routes/` · `src/styles/` (CSS tokens; components use CSS Modules)
- `tests/*.test.ts` (Vitest) · `tests/e2e/*.spec.ts` (Playwright, against the built site under `/finance-excel-labs/`)

## Lessons learned
- Sheet names are case-insensitive in HyperFormula and Excel (`Gl!E2` is `GL!E2`). Use a truly absent name such as `Ledger!` to test unknown-sheet handling.
- UI code imports `src/engine/address.ts` and `src/engine/types.ts`, never `src/engine/index.ts`, which would pull HyperFormula into an eagerly loaded chunk. Only `src/workspace/engineBundle.ts` imports the engine (PRD §8: the lesson renders first).
<!-- Add a one-line rule here whenever the same mistake happens twice. -->
