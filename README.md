# Excel Labs for Financial Analysis

> 🚧 **In development.** Live at **https://jiaxingxue.github.io/finance-excel-labs/**. Explore mode works (lesson, live grid, checks); Build, What-if, and Challenge modes arrive in later milestones.

Interactive, auto-graded Excel labs for FP&A and accounting: budget vs. actual, price–volume–mix, bank reconciliation, AR aging, forecasting, and driver-based scenarios. Formulas run live in the browser in an Excel-compatible engine, and every result is checked against 277 independently verified expected values.

## Planned labs
| Area | Labs |
|---|---|
| Variance analysis | Budget vs. actual, flux (MoM/YoY), price–volume–mix, rate/volume, flexible budget, pivots |
| Reconciliations | Bank rec with automated matching, AR aging and subledger-to-GL tie-out |
| Forecasting | Moving average, trend × seasonality with backtest (MAPE/bias), driver-based scenarios |
| Model controls | Integrity-checks dashboard, reusable function library |

## How correctness will be guaranteed
- **Independent expected values:** 277 expected values and 12 what-if experiments were computed outside this app (see [`reference/`](reference/README.md)).
- **A deploy gate:** the in-browser engine ([HyperFormula](https://github.com/handsontable/hyperformula), plus plugins for the functions it lacks) reproduces all 277 values and all 12 experiments in CI, and a deploy is blocked unless every one passes. The conformance suite is [`tests/conformance.test.ts`](tests/conformance.test.ts).
- **Grading checks results, not formula text,** and re-tests answers under changed inputs to reject hard-coding.

## Run locally
Requires **Node 24** (see `.nvmrc`). Commands work the same in PowerShell, cmd, and bash.

```
npm ci
npm run dev              # dev server at http://localhost:5173/
npm run check            # lint + typecheck + unit tests
```

| Task | Command |
|---|---|
| Regenerate data from the Labs document | `npm run extract` (`npm run extract -- --check` verifies it is current) |
| Unit tests + golden cross-check | `npm test` |
| Build as deployed on GitHub Pages | `npm run build:pages` (base path `/finance-excel-labs/`) |
| Preview that build | `npm run preview:pages` → http://localhost:4173/finance-excel-labs/ |
| End-to-end tests | `npx playwright install chromium` (once), then `npm run build:pages` and `npm run test:e2e` (it refuses a build made without the `/finance-excel-labs/` base). The `verification.json` test needs the file CI writes: first run `npm test -- --reporter=default --reporter=json --outputFile=reports/vitest.json` and `npm run verification:write` |
| Milestone screenshots | `npm run screenshots:m2` (after `npm run build:pages`) |
| Format | `npm run format` / `npm run format:check` |
| Third-party notices | `npm run notices` (after any dependency change) |

How content flows from the Labs document into the app: [docs/CONTENT_PIPELINE.md](docs/CONTENT_PIPELINE.md).

## Deploy
Every push to `main` runs [.github/workflows/deploy.yml](.github/workflows/deploy.yml): content sync check, lint, format, typecheck, unit and conformance tests, build, and end-to-end tests. The conformance results are published with the site as `verification.json`. The site is published to GitHub Pages **only if all of them pass**. Pull requests run the same checks without deploying.

One-time setup (owner): repository **Settings → Pages → Build and deployment → Source → GitHub Actions**.

## Project docs
- [Product requirements](docs/PRD.md)
- [Lab content (source of truth)](content/Excel_Implementation_Labs_Financial_Analysis.md)
- [Golden reference data](reference/README.md)
- [Owner decisions](docs/DECISIONS.md) · [Open issues](docs/OPEN_ISSUES.md) · [Milestone notes](docs/milestones/)

## License
GPLv3, for both the code and the lesson content. See [LICENSE](LICENSE) and [third-party notices](THIRD_PARTY_NOTICES.md).

<!-- Claude Code: replace this file with the full README template in PRD §17.6 during Milestone 7. -->
