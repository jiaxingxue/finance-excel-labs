# Excel Labs for Financial Analysis

> 🚧 **In development.** The live demo will be published at **https://jiaxingxue.github.io/finance-excel-labs/** once the first milestones are complete.

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
- **A deploy gate:** CI will block deployment unless every one of them is reproduced.
- **Grading checks results, not formula text,** and re-tests answers under changed inputs to reject hard-coding.

## Project docs
- [Product requirements](docs/PRD.md)
- [Lab content (source of truth)](content/Excel_Implementation_Labs_Financial_Analysis.md)
- [Golden reference data](reference/README.md)

## License
GPLv3 (license file added in the first milestone).

<!-- Claude Code: replace this file with the full README template in PRD §17.6 during Milestone 7. -->
