# Excel Implementation Labs for Financial Analysis
### Variance Analysis · Reconciliations · Forecasting Support — hands-on, cell-by-cell

*Companion to **Financial_Analysis_Study_Guide.md**. Prepared September 2026.*

> **Verification.** Every classic-Excel formula in this document was built into a single workbook (15 sheets, 668 formulas) and recalculated by a spreadsheet engine (LibreOffice Calc 24.2). All 277 expected-value checks and 12 what-if experiments shown here come from that recalculation, compared against independent Python calculations.
>
> **Microsoft 365–only formulas** (XLOOKUP, FILTER, UNIQUE, LET, LAMBDA, GROUPBY, PIVOTBY, SCAN, HSTACK) are shown as "modern equivalents." That engine cannot run them, so they were **not** machine-executed. Each one is designed to return the same result as its verified classic formula. Confirm them in Excel 365 before relying on them.

---

## Table of Contents

- [Part 0 — How This Workbook Is Organized](#part-0--how-this-workbook-is-organized)
- [Lab 1 — Data Foundation: Mapping Table and Long-Format Ledger](#lab-1--data-foundation-mapping-table-and-long-format-ledger)
- [Lab 2 — Budget vs. Actual Variance Report](#lab-2--budget-vs-actual-variance-report)
- [Lab 3 — Flux Analysis (Month-over-Month and Year-over-Year)](#lab-3--flux-analysis-month-over-month-and-year-over-year)
- [Lab 4 — Price–Volume–Mix Analysis and Variance Bridge](#lab-4--pricevolumemix-analysis-and-variance-bridge)
- [Lab 5 — Cost Rate/Volume Variances and the Flexible Budget](#lab-5--cost-ratevolume-variances-and-the-flexible-budget)
- [Lab 6 — Pivot-Based Analysis](#lab-6--pivot-based-analysis)
- [Lab 7 — Bank Reconciliation with Automated Matching](#lab-7--bank-reconciliation-with-automated-matching)
- [Lab 8 — AR Aging, Allowance, and Subledger-to-GL Reconciliation](#lab-8--ar-aging-allowance-and-subledger-to-gl-reconciliation)
- [Lab 9 — Forecasting: Moving Average, Trend, Seasonality, Backtest](#lab-9--forecasting-moving-average-trend-seasonality-backtest)
- [Lab 10 — Driver-Based Forecast, Scenarios, and Sensitivity](#lab-10--driver-based-forecast-scenarios-and-sensitivity)
- [Lab 11 — Reusable LAMBDA Function Library (Microsoft 365)](#lab-11--reusable-lambda-function-library-microsoft-365)
- [Lab 12 — Model Integrity Checks Dashboard](#lab-12--model-integrity-checks-dashboard)
- [Appendix A — Function Tiers and Fallback Formulas](#appendix-a--function-tiers-and-fallback-formulas)
- [Appendix B — Web App Blueprint](#appendix-b--web-app-blueprint)
- [Appendix C — Complete Cell-by-Cell Workbook Specification (JSON)](#appendix-c--complete-cell-by-cell-workbook-specification-json)
- [Appendix D — Assertions (Expected Values for Auto-Grading)](#appendix-d--assertions-expected-values-for-auto-grading)
- [Appendix E — What-If Experiments (JSON)](#appendix-e--what-if-experiments-json)

---

## Part 0 — How This Workbook Is Organized

### 0.1 Workbook map

All labs live in **one workbook** so that later labs can reference earlier sheets, just like a real finance model.

| Sheet | Lab | Purpose | Reads from |
|---|---|---|---|
| `Map` | 1 | Account → type / P&L line / sign | — |
| `GL` | 1 | Long-format ledger: actuals and budget by period, dept, account | — |
| `BvA` | 2 | Budget vs. actual report with favorable/unfavorable and materiality flags | `GL`, `Map` |
| `Flux` | 3 | Month-over-month and year-over-year fluctuation analysis | `GL` |
| `PVM` | 4 | Price–volume–mix decomposition and waterfall data | — |
| `CostVar` | 5 | Rate/volume cost variance; flexible budget | — |
| `PivotLab` | 6 | Formula-driven cross-tab (dept × account) | `GL`, `BvA` |
| `Book`, `Bank`, `BankRec` | 7 | Transaction matching and reconciliation proof | each other |
| `AR`, `ARSummary` | 8 | Aging buckets, allowance, subledger-to-GL tie-out | `AR` |
| `Forecast` | 9 | Moving average, linear trend, seasonal index, backtest | — |
| `Drivers` | 10 | Driver-based revenue model, scenario switch, sensitivity grid | — |
| `Checks` | 12 | One-page control dashboard | all |

### 0.2 Conventions used in every lab

1. **A1 addresses are exact.** Each formula is given for a specific cell (e.g., `BvA!C6`) and states how to fill it down or across. Build it in exactly that position and the expected results will match.
2. **Inputs vs. formulas.** Hard-coded inputs (thresholds, periods, rates) sit in labeled cells near the top of each sheet. Format them with **blue font** (the standard modeling convention); formulas stay black. Never type a number inside a formula.
3. **Bounded ranges, not whole columns.** Formulas use `GL!$E$2:$E$37`, not `GL!E:E`. Bounded ranges calculate faster, behave the same in every spreadsheet engine, and make the web-app version simpler. (In Excel you can later convert them to Tables; see Lab 1.)
4. **Periods as text keys.** Periods are stored as text like `"2026-08"`. That avoids date-serial and time-zone differences between engines and makes `SUMIFS` criteria unambiguous.
5. **The `$` rule.** Lock (`$`) whatever must **not** move when you fill. `$A6` locks the column (good when filling across), `B$1` locks the row (good when filling down), and `$B$1` locks both.

### 0.3 Function tiers

| Tier | Availability | Examples used here | Verified here? |
|---|---|---|---|
| **A — Classic** | Every Excel since 2010, LibreOffice, Google Sheets, most JavaScript formula engines | `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `INDEX`, `MATCH`, `LOOKUP`, `IF`, `AND`, `OR`, `ABS`, `ROUND`, `SUMPRODUCT`, `SLOPE`, `INTERCEPT`, `RSQ`, `TREND`, `EDATE`, `MONTH`, `TEXT`, `N`, `ISNUMBER`, `IFERROR` | ✅ Yes |
| **B — Excel 2016/2019** | Excel 2016+ | `FORECAST.LINEAR` ✅, `FORECAST.ETS` family, `IFS` | Partly (`FORECAST.LINEAR` yes) |
| **C — Dynamic arrays** | Excel 2021 and Microsoft 365 | `XLOOKUP`, `XMATCH`, `FILTER`, `UNIQUE`, `SORT`, `SEQUENCE`, `LET` | ❌ Shown as equivalents |
| **D — Microsoft 365 only** | Microsoft 365 (check your channel) | `LAMBDA`, `SCAN`, `MAP`, `HSTACK`, `VSTACK`, `GROUPBY`, `PIVOTBY` | ❌ Shown as equivalents |

**Recommendation for the future web app:** make the **Tier A formulas the canonical, auto-graded implementation**, because nearly any formula engine can execute them. Present Tier C/D formulas as an alternative "Modern Excel" view.

---

## Lab 1 — Data Foundation: Mapping Table and Long-Format Ledger

### Learning goals
- Structure data the way analysis needs it: one row per fact (**long format**).
- Separate **reference data** (account mapping) from **transaction data** (GL).
- Prepare ranges for `SUMIFS`, lookups, and PivotTables.

### Step 1 — Build the `Map` sheet

Type this into `Map!A1:D7`:

| | A | B | C | D |
|---|---|---|---|---|
| **1** | Account | Type | PL_Line | Sign |
| **2** | Revenue | Revenue | Net revenue | 1 |
| **3** | COGS | Expense | Cost of sales | -1 |
| **4** | Salaries | Expense | Operating expense | -1 |
| **5** | Marketing | Expense | Operating expense | -1 |
| **6** | Rent | Expense | Operating expense | -1 |
| **7** | Travel | Expense | Operating expense | -1 |

- **Type** drives favorable/unfavorable logic (revenue up = good, expense up = bad).
- **Sign** converts to a profit view (+1 revenue, −1 expense) if you ever need a signed total.

### Step 2 — Build the `GL` sheet

Enter the 36 data rows in `GL!A1:E37`. All amounts are positive, and the direction comes from `Map!Type`. The first rows look like this; the full data is in [Appendix C](#appendix-c--complete-cell-by-cell-workbook-specification-json).

| | A | B | C | D | E |
|---|---|---|---|---|---|
| **1** | Period | Dept | Account | Version | Amount |
| **2** | 2026-08 | Sales | Revenue | BUD | 300,000 |
| **3** | 2026-08 | Ops | Revenue | BUD | 200,000 |
| **4** | 2026-08 | Ops | COGS | BUD | 200,000 |
| **5** | 2026-08 | Sales | Salaries | BUD | 70,000 |
| **6** | 2026-08 | Ops | Salaries | BUD | 50,000 |
| **7** | 2026-08 | Sales | Marketing | BUD | 40,000 |
| **8** | 2026-08 | Ops | Rent | BUD | 25,000 |
| **9** | 2026-08 | Sales | Travel | BUD | 5,000 |
| **10** | 2026-08 | Ops | Travel | BUD | 3,000 |
| **11** | 2026-08 | Sales | Revenue | ACT | 270,000 |
| **12** | 2026-08 | Ops | Revenue | ACT | 198,000 |
| **13** | 2026-08 | Ops | COGS | ACT | 191,000 |

The table contains four blocks, each with the six accounts split across the Sales and Ops departments:

| Rows | Period | Version | What it is |
|---|---|---|---|
| 2–10 | 2026-08 | BUD | Budget for the month being analyzed |
| 11–19 | 2026-08 | ACT | Actuals for the month being analyzed |
| 20–28 | 2026-07 | ACT | Prior month actuals (for flux) |
| 29–37 | 2025-08 | ACT | Same month last year (for flux) |

### Step 3 — Make it robust (Excel UI)

| Action | How | Why |
|---|---|---|
| Convert to Tables | Click in the data → **Ctrl+T** → name them `tblMap`, `tblGL` (Table Design → Table Name) | Ranges auto-expand when you add rows |
| Data validation on Version | Select `GL!D2:D37` → Data → Data Validation → List → `ACT,BUD,FCST` | Prevents typos like "Act" that silently drop out of `SUMIFS` |
| Data validation on Account | List source `=Map!$A$2:$A$7` | Every GL account must exist in the mapping |
| Number format | `GL!E2:E37` → `#,##0` | Readability |
| Freeze header | View → Freeze Panes → Freeze Top Row | Navigation |

### Step 4 — Completeness check: every GL account is mapped

- **Classic (Tier A):** add a helper column in `GL!F2` and fill down: `=IF(COUNTIF(Map!$A$2:$A$7,C2)=0,"UNMAPPED","")`
- **Modern (Tier C):** one formula lists any unmapped accounts:

```excel
=LET(accts, UNIQUE(GL!C2:C37),
     missing, FILTER(accts, ISNA(XMATCH(accts, Map!A2:A7)), "All mapped"),
     missing)
```

With the lab data the answer is "All mapped." Lab 2 includes a second completeness check (`BvA!D14`) that proves the report captures 100% of GL actuals.

### Optional — Power Query instead of copy/paste

1. Data → Get Data → From File → From Workbook/CSV (your ERP export).
2. In the Power Query editor: set data types, trim text (Transform → Format → Trim), and filter out blank rows.
3. If the export is wide (months across columns), select the ID columns → Transform → **Unpivot Other Columns** to make it long.
4. Close & Load To → Table on the `GL` sheet. Next month: replace the file and click **Refresh All**.

---

## Lab 2 — Budget vs. Actual Variance Report

### Learning goals
- Pull budget and actual from a long table with `SUMIFS`.
- Compute $ and % variance correctly (sign-safe).
- Classify favorable/unfavorable for revenue vs. expense.
- Flag lines that need commentary using a two-part materiality rule.

### Inputs (`BvA!A1:B3`)

| Cell | Label | Value | Meaning |
|---|---|---|---|
| `B1` | Report period | `2026-08` (text) | Which period to report |
| `B2` | Pct threshold | 5% | Comment if \|var %\| ≥ this… |
| `B3` | Abs threshold | 5,000 | …**and** \|var $\| ≥ this |

### Build steps

1. Headers in `A5:H5`: Account, Type, Budget, Actual, Variance, Var %, Fav/Unfav, Needs comment.
2. Accounts in `A6:A11` (Revenue, COGS, Salaries, Marketing, Rent, Travel).
3. Enter the formulas in row 6 below and **fill down to row 11**.
4. Row 12 = Operating income. Enter `C12`/`D12`, then copy `E6:H6` into `E12:H12`.
5. Enter the completeness check in `D14`.

### Formulas (row 6 is the pattern; rows 7–11 are filled down)

| Cell | Formula | Result |
|---|---|---|
| `BvA!B6` | `=INDEX(Map!$B$2:$B$7,MATCH($A6,Map!$A$2:$A$7,0))` | Revenue |
| `BvA!C6` | `=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A6,GL!$D$2:$D$37,"BUD",GL!$A$2:$A$37,$B$1)` | 500,000 |
| `BvA!D6` | `=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A6,GL!$D$2:$D$37,"ACT",GL!$A$2:$A$37,$B$1)` | 468,000 |
| `BvA!E6` | `=D6-C6` | (32,000) |
| `BvA!F6` | `=IF(C6=0,"n/a",E6/ABS(C6))` | -6.4% |
| `BvA!G6` | `=IF(D6=C6,"On budget",IF((B6="Revenue")=(D6>C6),"Favorable","Unfavorable"))` | Unfavorable |
| `BvA!H6` | `=IF(ISNUMBER(F6),AND(ABS(E6)>=$B$3,ABS(F6)>=$B$2),FALSE)` | TRUE |
| `BvA!C12` | `=C6-SUM(C7:C11)` | 107,000 |
| `BvA!D12` | `=D6-SUM(D7:D11)` | 70,400 |
| `BvA!D14` | `=SUM(D6:D11)-SUMIFS(GL!$E$2:$E$37,GL!$D$2:$D$37,"ACT",GL!$A$2:$A$37,$B$1)` | 0 |

### How the key formulas work

- **`SUMIFS` (C6, D6):** `SUMIFS(sum_range, criteria_range1, criteria1, …)` adds the GL amounts where Account = `$A6`, Version = `"BUD"` (or `"ACT"`), and Period = `$B$1`. `$A6` keeps the column fixed but lets the row move when filling down; `$B$1` never moves.
- **`INDEX/MATCH` (B6):** `MATCH` finds the row of the account in `Map!A2:A7`, and `INDEX` returns the Type from the same row. This is the classic, universally supported lookup.
- **Sign-safe % (F6):** dividing by `ABS(C6)` keeps the sign meaningful when the budget is negative, and `IF(C6=0,"n/a",…)` prevents `#DIV/0!`.
- **Favorable logic (G6):** `(B6="Revenue")=(D6>C6)` compares two TRUE/FALSE values:

  | Type | Actual > Budget? | Comparison | Result |
  |---|---|---|---|
  | Revenue | TRUE | TRUE = TRUE | Favorable |
  | Revenue | FALSE | TRUE = FALSE | Unfavorable |
  | Expense | TRUE | FALSE = TRUE | Unfavorable |
  | Expense | FALSE | FALSE = FALSE | Favorable |

- **Materiality flag (H6):** both conditions must hold (`AND`). `ISNUMBER(F6)` guards against the `"n/a"` text, since `ABS("n/a")` would return `#VALUE!`.
- **Completeness check (D14):** the actuals pulled for the six mapped accounts minus **all** GL actuals for the period. Zero proves no account was missed. A non-zero value means an unmapped or misspelled account.

### Result (values after calculation)

| | A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|---|
| **5** | Account | Type | Budget | Actual | Variance | Var % | Fav/Unfav | Needs comment |
| **6** | Revenue | Revenue | 500,000 | 468,000 | (32,000) | -6.4% | Unfavorable | TRUE |
| **7** | COGS | Expense | 200,000 | 191,000 | (9,000) | -4.5% | Favorable | FALSE |
| **8** | Salaries | Expense | 120,000 | 124,500 | 4,500 | 3.8% | Unfavorable | FALSE |
| **9** | Marketing | Expense | 40,000 | 52,000 | 12,000 | 30.0% | Unfavorable | TRUE |
| **10** | Rent | Expense | 25,000 | 25,000 | 0 | 0.0% | On budget | FALSE |
| **11** | Travel | Expense | 8,000 | 5,100 | (2,900) | -36.3% | Favorable | FALSE |
| **12** | Operating income | Revenue | 107,000 | 70,400 | (36,600) | -34.2% | Unfavorable | TRUE |
| **13** |  |  |  |  |  |  |  |  |
| **14** | Check: mapped actuals minus all GL actuals (should be 0) |  |  | 0 |  |  |  |  |

### Modern Excel equivalents (Tier C/D, not machine-verified)

**Type lookup with XLOOKUP** (returns "Unmapped" instead of `#N/A`):
```excel
=XLOOKUP($A6, Map!$A$2:$A$7, Map!$B$2:$B$7, "Unmapped")
```

**The whole report as one spilling formula** (put in `J5`):
```excel
=LET(
  acc, Map!A2:A7,  typ, Map!B2:B7,  per, B1,
  bud, SUMIFS(GL!E2:E37, GL!C2:C37, acc, GL!D2:D37, "BUD", GL!A2:A37, per),
  act, SUMIFS(GL!E2:E37, GL!C2:C37, acc, GL!D2:D37, "ACT", GL!A2:A37, per),
  var, act - bud,
  pct, IF(bud = 0, "n/a", var / ABS(bud)),
  fav, IF(act = bud, "On budget", IF((typ = "Revenue") = (act > bud), "Favorable", "Unfavorable")),
  flag, IF(ISNUMBER(pct), (ABS(var) >= B3) * (ABS(pct) >= B2) = 1, FALSE),
  VSTACK({"Account","Type","Budget","Actual","Variance","Var %","Fav/Unfav","Needs comment"},
         HSTACK(acc, typ, bud, act, var, pct, fav, flag)))
```
Inside `LET`, `SUMIFS` accepts an **array** as its criteria (`acc`) and returns one result per account. `(cond1)*(cond2)=1` is the array-safe form of `AND`, because `AND` would collapse the whole array to a single TRUE/FALSE.

### Excel UI features to add

| Feature | Steps | Rule |
|---|---|---|
| Highlight unfavorable | Select `A6:H12` → Home → Conditional Formatting → New Rule → *Use a formula* | `=$G6="Unfavorable"` → light red fill |
| Bold lines needing comment | Same range, new rule | `=$H6=TRUE` → bold font |
| Icon for direction | Select `E6:E12` → Conditional Formatting → Icon Sets → 3 Arrows | Visual only |
| Period dropdown | Select `B1` → Data Validation → List | Source: `2025-08,2026-07,2026-08` |
| Commentary column | Add `I5` "Commentary" | Free text: *What / Why / So what* |

### Try it (verified what-if experiments)

**E2.1 — Lower the % threshold to 3%**  
Change: `BvA!B2` → `0.03`

| Watch cell | Before | After |
|---|---|---|
| `BvA!H6` | TRUE | TRUE |
| `BvA!H7` | FALSE | TRUE |
| `BvA!H8` | FALSE | FALSE |
| `BvA!H9` | TRUE | TRUE |
| `BvA!H10` | FALSE | FALSE |
| `BvA!H11` | FALSE | FALSE |
| `BvA!H12` | TRUE | TRUE |

**E2.2 — Switch the report period to 2026-07 (no budget loaded)**  
Change: `BvA!B1` → `2026-07`

| Watch cell | Before | After |
|---|---|---|
| `BvA!C6` | 500,000 | 0 |
| `BvA!D6` | 468,000 | 455,000 |
| `BvA!F6` | -6.4% | n/a |
| `BvA!G6` | Unfavorable | Favorable |
| `BvA!H6` | TRUE | FALSE |
| `BvA!D14` | 0 | 0 |

**Lesson from E2.2:** with no budget loaded for July, Revenue shows "Favorable" simply because actual > 0. Harden the formula so a missing plan can't masquerade as a result:
```excel
G6: =IF(C6=0,"No budget",IF(D6=C6,"On budget",IF((B6="Revenue")=(D6>C6),"Favorable","Unfavorable")))
```

### Common errors

| Symptom | Cause | Fix |
|---|---|---|
| Budget shows 0 | Period typed as a date in `B1` but stored as text in GL (or vice versa) | Make both text, or both dates |
| `#VALUE!` in H | `ABS()` applied to `"n/a"` | Use the `ISNUMBER` guard shown |
| Wrong rows after fill-down | Missing `$` on GL ranges | Use `GL!$E$2:$E$37` |
| `D14` ≠ 0 | GL has an account not in `Map` | Add it to `Map`, or fix the spelling |

---

## Lab 3 — Flux Analysis (Month-over-Month and Year-over-Year)

### Learning goals
- Compare the current period with the prior month and the same month last year.
- Apply an "either comparison breaches" rule to decide what needs explaining.

### Inputs (`Flux!A1:B5`)

| | A | B |
|---|---|---|
| **1** | Current period | 2026-08 |
| **2** | Prior month | 2026-07 |
| **3** | Prior year | 2025-08 |
| **4** | Pct threshold | 10.0% |
| **5** | Abs threshold | 5,000 |

### Formulas (row 8 pattern; fill down to row 13)

| Cell | Formula | Result |
|---|---|---|
| `Flux!B8` | `=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A8,GL!$D$2:$D$37,"ACT",GL!$A$2:$A$37,$B$1)` | 468,000 |
| `Flux!C8` | `=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A8,GL!$D$2:$D$37,"ACT",GL!$A$2:$A$37,$B$2)` | 455,000 |
| `Flux!D8` | `=B8-C8` | 13,000 |
| `Flux!E8` | `=IF(C8=0,"n/a",D8/ABS(C8))` | 2.9% |
| `Flux!F8` | `=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A8,GL!$D$2:$D$37,"ACT",GL!$A$2:$A$37,$B$3)` | 430,000 |
| `Flux!G8` | `=B8-F8` | 38,000 |
| `Flux!H8` | `=IF(F8=0,"n/a",G8/ABS(F8))` | 8.8% |
| `Flux!I8` | `=IF(OR(AND(ABS(D8)>=$B$5,ABS(N(E8))>=$B$4),AND(ABS(G8)>=$B$5,ABS(N(H8))>=$B$4)),"Explain","")` |  |

**Notes**
- The only difference between B, C, and F is the period cell they point to (`$B$1`, `$B$2`, `$B$3`). The account reference `$A8` stays the same.
- `N()` converts text such as `"n/a"` to 0, so `ABS(N(E8))` never errors. This is a compact alternative to the `ISNUMBER` guard used in Lab 2.
- The flag is `OR(AND(MoM breach), AND(YoY breach))`: either comparison crossing **both** thresholds requires an explanation.

### Result

| | A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|---|
| **7** | Account | Current | Prior month | MoM $ | MoM % | Prior year | YoY $ | YoY % | Flag |
| **8** | Revenue | 468,000 | 455,000 | 13,000 | 2.9% | 430,000 | 38,000 | 8.8% |  |
| **9** | COGS | 191,000 | 186,000 | 5,000 | 2.7% | 176,000 | 15,000 | 8.5% |  |
| **10** | Salaries | 124,500 | 121,000 | 3,500 | 2.9% | 110,000 | 14,500 | 13.2% | Explain |
| **11** | Marketing | 52,000 | 39,000 | 13,000 | 33.3% | 35,000 | 17,000 | 48.6% | Explain |
| **12** | Rent | 25,000 | 25,000 | 0 | 0.0% | 23,000 | 2,000 | 8.7% |  |
| **13** | Travel | 5,100 | 6,800 | (1,700) | -25.0% | 7,000 | (1,900) | -27.1% |  |

Salaries is flagged only on the year-over-year comparison (+13.2%, +14,500). Marketing breaches on both comparisons. Travel moved −25% month over month but stays under the $5,000 threshold, so it is not flagged.

### Modern equivalent (Tier C, not machine-verified)
```excel
=LET(acc, A8:A13,
     get, LAMBDA(p, SUMIFS(GL!E2:E37, GL!C2:C37, acc, GL!D2:D37, "ACT", GL!A2:A37, p)),
     cur, get(B1), pm, get(B2), py, get(B3),
     HSTACK(acc, cur, pm, cur - pm, (cur - pm) / ABS(pm), py, cur - py, (cur - py) / ABS(py)))
```

### Try it

**E3.1 — Raise the flux % threshold to 20%**  
Change: `Flux!B4` → `0.2`

| Watch cell | Before | After |
|---|---|---|
| `Flux!I8` |  |  |
| `Flux!I9` |  |  |
| `Flux!I10` | Explain |  |
| `Flux!I11` | Explain | Explain |
| `Flux!I12` |  |  |
| `Flux!I13` |  |  |

---

## Lab 4 — Price–Volume–Mix Analysis and Variance Bridge

### Learning goals
- Decompose a revenue variance into price, volume, and mix using the **budget-mix method**.
- Build a control check proving the three effects add up to the total.
- Chart the result as a waterfall.

### Inputs (`PVM!A4:E7`)

| | A | B | C | D | E |
|---|---|---|---|---|---|
| **4** | Product | Bud units | Bud price | Act units | Act price |
| **5** | Basic | 1,000 | 50 | 900 | 52 |
| **6** | Pro | 500 | 120 | 600 | 115 |
| **7** | Enterprise | 100 | 400 | 90 | 420 |

### Formulas (row 5 pattern; fill down to row 7; totals in row 8)

| Cell | Formula | Result |
|---|---|---|
| `PVM!F5` | `=B5*C5` | 50,000 |
| `PVM!G5` | `=D5*E5` | 46,800 |
| `PVM!H5` | `=B5/$B$8` | 62.50% |
| `PVM!I5` | `=$D$8*H5` | 993.75 |
| `PVM!J5` | `=(E5-C5)*D5` | 1,800.0 |
| `PVM!K5` | `=(I5-B5)*C5` | (312.5) |
| `PVM!L5` | `=(D5-I5)*C5` | (4,687.5) |
| `PVM!M5` | `=G5-F5` | (3,200) |
| `PVM!N5` | `=ROUND(J5+K5+L5-M5,6)` | 0 |

Totals row:

| Cell | Formula | Result |
|---|---|---|
| `PVM!B8` | `=SUM(B5:B7)` | 1,600 |
| `PVM!C8` | `=F8/B8` | 93.75 |
| `PVM!D8` | `=SUM(D5:D7)` | 1,590 |
| `PVM!E8` | `=G8/D8` | 96.60 |
| `PVM!F8` | `=SUM(F5:F7)` | 150,000 |
| `PVM!G8` | `=SUM(G5:G7)` | 153,600 |
| `PVM!H8` | `=SUM(H5:H7)` | 1 |
| `PVM!I8` | `=SUM(I5:I7)` | 1,590 |
| `PVM!J8` | `=SUM(J5:J7)` | 600 |
| `PVM!K8` | `=SUM(K5:K7)` | -937.5 |
| `PVM!L8` | `=SUM(L5:L7)` | 3,937.5 |
| `PVM!M8` | `=SUM(M5:M7)` | 3,600 |
| `PVM!N8` | `=SUM(N5:N7)` | 0 |

### How it works

| Effect | Formula (row 5) | Plain English |
|---|---|---|
| Budget mix % | `=B5/$B$8` | This product's share of budgeted units |
| Actual units at budget mix | `=$D$8*H5` | Actual total units, split in budget proportions |
| **Price** | `=(E5-C5)*D5` | Price change × what we actually sold |
| **Volume** | `=(I5-B5)*C5` | More/fewer total units at the budget mix, valued at budget price |
| **Mix** | `=(D5-I5)*C5` | Shift between products, valued at budget price |
| Check | `=ROUND(J5+K5+L5-M5,6)` | Must be 0; `ROUND` removes floating-point noise like 1E-12 |

### Result

| | A | B | C | D | E | F | G | H | I | J | K | L | M | N |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| **4** | Product | Bud units | Bud price | Act units | Act price | Bud rev | Act rev | Bud mix % | Act units @ bud mix | Price var | Volume var | Mix var | Total var | Check |
| **5** | Basic | 1,000 | 50 | 900 | 52 | 50,000 | 46,800 | 62.50% | 993.75 | 1,800.0 | (312.5) | (4,687.5) | (3,200) | 0 |
| **6** | Pro | 500 | 120 | 600 | 115 | 60,000 | 69,000 | 31.25% | 496.88 | (3,000.0) | (375.0) | 12,375.0 | 9,000 | 0 |
| **7** | Enterprise | 100 | 400 | 90 | 420 | 40,000 | 37,800 | 6.25% | 99.38 | 1,800.0 | (250.0) | (3,750.0) | (2,200) | 0 |
| **8** | Total | 1,600 | 93.75 | 1,590 | 96.60 | 150,000 | 153,600 | 1 | 1,590 | 600 | -937.5 | 3,937.5 | 3,600 | 0 |

**Interpretation:** revenue beat budget by 3,600. Mix was the main driver (3,937.5, as customers shifted to Pro), total volume was slightly low (-937.5), and net price was slightly positive (600). Note the Pro price cut: it hurt price but likely drove the favorable mix.

### Waterfall (variance bridge) chart

Waterfall data (`PVM!A10:B15`):

| | A | B |
|---|---|---|
| **10** | Waterfall step | Amount |
| **11** | Budget revenue | 150,000.0 |
| **12** | Price | 600.0 |
| **13** | Volume | (937.5) |
| **14** | Mix | 3,937.5 |
| **15** | Actual revenue | 153,600.0 |

1. Select `A10:B15` → Insert → Charts → **Waterfall**.
2. Click the "Budget revenue" bar → right-click → **Set as total**. Repeat for "Actual revenue".
3. Format → set the vertical axis minimum near 145,000 so the small steps are visible. State this truncation in the chart title.

### Modern equivalent (Tier C): the whole decomposition in one cell
```excel
=LET(bu,B5:B7, bp,C5:C7, au,D5:D7, ap,E5:E7,
     mix, bu/SUM(bu), aub, SUM(au)*mix,
     price, (ap-bp)*au, vol, (aub-bu)*bp, mx, (au-aub)*bp,
     HSTACK(A5:A7, price, vol, mx, price+vol+mx))
```

### Try it

**E4.1 — Pro sells at budget price (E6 = 120)**  
Change: `PVM!E6` → `120`

| Watch cell | Before | After |
|---|---|---|
| `PVM!J8` | 600 | 3,600 |
| `PVM!K8` | -937.5 | -937.5 |
| `PVM!L8` | 3,937.5 | 3,937.5 |
| `PVM!M8` | 3,600 | 6,600 |
| `PVM!N8` | 0 | 0 |

**E4.2 — Same total units, budget mix (D5=994, D6=497, D7=99)**  
Change: `PVM!D5` → `994`, `PVM!D6` → `497`, `PVM!D7` → `99`

| Watch cell | Before | After |
|---|---|---|
| `PVM!K8` | -937.5 | -937.5 |
| `PVM!L8` | 3,937.5 | -122.5 |
| `PVM!N8` | 0 | 0 |

In E4.2, mix is not exactly zero because whole units (994/497/99) can't hit the exact budget proportions (993.75/496.875/99.375). The check column stays at 0 in every case, which is the point of a control check.

---

## Lab 5 — Cost Rate/Volume Variances and the Flexible Budget

### Part A — Rate/volume (freight example)

Inputs `CostVar!B3:B6`: budget shipments 2,000 @ $12.00; actual 2,300 @ $12.50.

| Cell | Formula | Result |
|---|---|---|
| `CostVar!B8` | `=B3*B4` | 24,000 |
| `CostVar!B9` | `=B5*B6` | 28,750 |
| `CostVar!B10` | `=(B6-B4)*B5` | 1,150 |
| `CostVar!B11` | `=(B5-B3)*B4` | 3,600 |
| `CostVar!B12` | `=B9-B8` | 4,750 |
| `CostVar!B13` | `=B10+B11-B12` | 0 |
| `CostVar!C10` | `=IF(B10>0,"Unfavorable",IF(B10<0,"Favorable","None"))` | Unfavorable |
| `CostVar!C11` | `=IF(B11>0,"Unfavorable",IF(B11<0,"Favorable","None"))` | Unfavorable |
| `CostVar!C12` | `=IF(B12>0,"Unfavorable",IF(B12<0,"Favorable","None"))` | Unfavorable |

For **costs**, a positive variance means spending more than planned, which is **Unfavorable**. The label formula in column C encodes that.

| | A | B | C |
|---|---|---|---|
| **8** | Budget cost | 24,000 |  |
| **9** | Actual cost | 28,750 |  |
| **10** | Rate variance | 1,150 | Unfavorable |
| **11** | Volume variance | 3,600 | Unfavorable |
| **12** | Total variance | 4,750 | Unfavorable |
| **13** | Check (should be 0) | 0 |  |

### Part B — Flexible budget

A **flexible budget** restates the budget at the **actual** activity level. That splits the total variance into:
- **Sales-volume variance** (Flex − Static): the effect of selling a different number of units.
- **Flexible-budget variance** (Actual − Flex): the effect of prices and costs differing at the actual volume.

Inputs (`CostVar!A17:B24`):

| | A | B |
|---|---|---|
| **17** | Budget units | 10,000 |
| **18** | Budget price | 20 |
| **19** | Budget variable cost / unit | 12 |
| **20** | Budget fixed costs | 50,000 |
| **21** | Actual units | 11,000 |
| **22** | Actual revenue | 214,500 |
| **23** | Actual variable costs | 137,500 |
| **24** | Actual fixed costs | 52,000 |

Formulas:

| Cell | Formula | Result |
|---|---|---|
| `CostVar!B27` | `=$B$17` | 10,000 |
| `CostVar!C27` | `=$B$21` | 11,000 |
| `CostVar!D27` | `=$B$21` | 11,000 |
| `CostVar!E27` | `=D27-C27` | 0 |
| `CostVar!F27` | `=C27-B27` | 1,000 |
| `CostVar!B28` | `=B27*$B$18` | 200,000 |
| `CostVar!C28` | `=C27*$B$18` | 220,000 |
| `CostVar!D28` | `=$B$22` | 214,500 |
| `CostVar!E28` | `=D28-C28` | (5,500) |
| `CostVar!F28` | `=C28-B28` | 20,000 |
| `CostVar!B29` | `=B27*$B$19` | 120,000 |
| `CostVar!C29` | `=C27*$B$19` | 132,000 |
| `CostVar!D29` | `=$B$23` | 137,500 |
| `CostVar!E29` | `=D29-C29` | 5,500 |
| `CostVar!F29` | `=C29-B29` | 12,000 |
| `CostVar!B30` | `=B28-B29` | 80,000 |
| `CostVar!C30` | `=C28-C29` | 88,000 |
| `CostVar!D30` | `=D28-D29` | 77,000 |
| `CostVar!E30` | `=D30-C30` | (11,000) |
| `CostVar!F30` | `=C30-B30` | 8,000 |
| `CostVar!B31` | `=$B$20` | 50,000 |
| `CostVar!C31` | `=$B$20` | 50,000 |
| `CostVar!D31` | `=$B$24` | 52,000 |
| `CostVar!E31` | `=D31-C31` | 2,000 |
| `CostVar!F31` | `=C31-B31` | 0 |
| `CostVar!B32` | `=B30-B31` | 30,000 |
| `CostVar!C32` | `=C30-C31` | 38,000 |
| `CostVar!D32` | `=D30-D31` | 25,000 |
| `CostVar!E32` | `=D32-C32` | (13,000) |
| `CostVar!F32` | `=C32-B32` | 8,000 |

Result:

| | A | B | C | D | E | F |
|---|---|---|---|---|---|---|
| **26** | Line | Static budget | Flexible budget | Actual | Flex-budget variance (Actual - Flex) | Sales-volume variance (Flex - Static) |
| **27** | Units | 10,000 | 11,000 | 11,000 | 0 | 1,000 |
| **28** | Revenue | 200,000 | 220,000 | 214,500 | (5,500) | 20,000 |
| **29** | Variable costs | 120,000 | 132,000 | 137,500 | 5,500 | 12,000 |
| **30** | Contribution margin | 80,000 | 88,000 | 77,000 | (11,000) | 8,000 |
| **31** | Fixed costs | 50,000 | 50,000 | 52,000 | 2,000 | 0 |
| **32** | Operating income | 30,000 | 38,000 | 25,000 | (13,000) | 8,000 |

**Reading it:** static-budget operating income was 30,000 and actual was 25,000. Selling more units should have *added* 8,000 of profit (Flex − Static). Three things went the other way: revenue came in 5,500 below the flexible budget (lower price per unit), variable costs ran 5,500 above it (higher cost per unit), and fixed costs overran by 2,000. Together they produce a flexible-budget variance of (13,000).

> **Sign reading tip:** in column E, revenue and profit rows are "actual − flex" (negative = bad), while cost rows are also "actual − flex" (positive = bad). Label the direction for readers, as Part A's column C does.

### Try it

**E5.1 — Negotiate freight rate down to 11.50**  
Change: `CostVar!B6` → `11.5`

| Watch cell | Before | After |
|---|---|---|
| `CostVar!B10` | 1,150 | (1,150) |
| `CostVar!C10` | Unfavorable | Favorable |
| `CostVar!B12` | 4,750 | 2,450 |
| `CostVar!C12` | Unfavorable | Unfavorable |

---

## Lab 6 — Pivot-Based Analysis

Three ways to build the same dept × account summary for **2026-08 ACT**.

### Method 1 — Formula cross-tab with SUMIFS (Tier A, verified)

Inputs: `PivotLab!B2` = `2026-08`, `B3` = `ACT`. Accounts across `B5:G5`, departments down `A6:A7`.

The **one formula** in `B6`, filled across to `G6` and down to `G7`:

| Cell | Formula | Result |
|---|---|---|
| `PivotLab!B6` | `=SUMIFS(GL!$E$2:$E$37,GL!$B$2:$B$37,$A6,GL!$C$2:$C$37,B$5,GL!$A$2:$A$37,$B$2,GL!$D$2:$D$37,$B$3)` | 270,000 |
| `PivotLab!B8` | `=SUM(B6:B7)` | 468,000 |
| `PivotLab!B10` | `=SUM(B8:G8)-SUM(BvA!D6:D11)` | 0 |

The "mixed" references do the work: `$A6` (column locked, so filling across still reads the department) and `B$5` (row locked, so filling down still reads the account).

| | A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|---|
| **5** | Dept \ Account | Revenue | COGS | Salaries | Marketing | Rent | Travel |
| **6** | Sales | 270,000 | 0 | 72,000 | 52,000 | 0 | 3,600 |
| **7** | Ops | 198,000 | 191,000 | 52,500 | 0 | 25,000 | 1,500 |
| **8** | Total | 468,000 | 191,000 | 124,500 | 52,000 | 25,000 | 5,100 |
| **9** |  |  |  |  |  |  |  |
| **10** | Check: cross-tab total vs BvA actuals (should be 0) | 0 |  |  |  |  |  |

`B10` is a cross-check: the cross-tab total equals the BvA actuals (0 = agree).

### Method 2 — PivotTable (Excel UI)

1. Click inside `GL!A1:E37` (or `tblGL`) → Insert → **PivotTable** → New Worksheet.
2. Drag fields: **Rows** = Dept; **Columns** = Account; **Values** = Sum of Amount; **Filters** = Period, Version.
3. Set the filters to Period = 2026-08, Version = ACT. The numbers match Method 1.
4. PivotTable Analyze → **Insert Slicer** → Period, Version, Dept. Slicers are clickable filters.
5. **Drill-through:** double-click any value to see the underlying GL rows on a new sheet. This is the fastest way to investigate a variance.
6. **Budget vs. actual inside a pivot:** move Version to Columns, then PivotTable Analyze → Fields, Items & Sets → **Calculated Item** on Version: name `Var`, formula `=ACT-BUD`.
7. **Pull a pivot number into a report** (it stays correct even if the pivot layout moves):
   ```excel
   =GETPIVOTDATA("Amount", Pivot!$A$3, "Dept", "Sales", "Account", "Marketing")
   ```
8. After data changes: PivotTable Analyze → **Refresh** (or Data → Refresh All). PivotTables do *not* recalculate automatically.

### Method 3 — GROUPBY / PIVOTBY (Tier D, Microsoft 365; not machine-verified)

```excel
Dept × Account for the chosen period/version:
=PIVOTBY(GL!B2:B37, GL!C2:C37, GL!E2:E37, SUM,,,,,, (GL!A2:A37=B2)*(GL!D2:D37=B3)=1)

Totals by account only:
=GROUPBY(GL!C2:C37, GL!E2:E37, SUM,,,, (GL!A2:A37=B2)*(GL!D2:D37=B3)=1)

Tier C alternative (dynamic arrays, no PIVOTBY):
=LET(d, UNIQUE(GL!B2:B37), a, TRANSPOSE(Map!A2:A7),
     SUMIFS(GL!E2:E37, GL!B2:B37, d, GL!C2:C37, a, GL!A2:A37, B2, GL!D2:D37, B3))
```

**Argument positions:** the filter is the **10th** argument of `PIVOTBY` and the **7th** of `GROUPBY`; the blank commas skip the optional arguments in between. These functions recalculate automatically, unlike PivotTables.

### Which to use?

| Need | Best choice |
|---|---|
| Fast ad-hoc exploration, drill-through | PivotTable |
| A fixed report layout that must never move | SUMIFS cross-tab or GETPIVOTDATA |
| Auto-updating summary, Microsoft 365 users only | GROUPBY / PIVOTBY |
| Sharing with users on older Excel / web engines | SUMIFS cross-tab |

---

## Lab 7 — Bank Reconciliation with Automated Matching

### Learning goals
- Match book and bank transactions in two passes (exact, then tolerance).
- Classify unmatched items and build a reconciliation proof that must tie to zero.

### Inputs

`Book!A1:D6` (general ledger cash):

| | A | B | C | D |
|---|---|---|---|---|
| **1** | ID | Date | Amount | Ref |
| **2** | B1 | 2026-08-02 | 1,200.00 | INV1001 |
| **3** | B2 | 2026-08-05 | -350.00 | CHK2201 |
| **4** | B3 | 2026-08-10 | 4,999.50 | INV1002 |
| **5** | B4 | 2026-08-28 | -2,000.00 | CHK2202 |
| **6** | B5 | 2026-08-31 | 780.00 | DEP0831 |

`Bank!A1:D5` (bank statement):

| | A | B | C | D |
|---|---|---|---|---|
| **1** | ID | Date | Amount | Ref |
| **2** | K1 | 2026-08-03 | 1,200.00 | INV1001 |
| **3** | K2 | 2026-08-07 | -350.00 | CHK2201 |
| **4** | K3 | 2026-08-11 | 4,995.00 | INV1002 |
| **5** | K4 | 2026-08-15 | -25.00 | FEE |

`BankRec` parameters: amount tolerance `B2` = 5.00, day tolerance `B3` = 3, bank statement balance `B5` = 10,000.00, GL balance `B6` = 8,809.50.

### Matching formulas (row 2 pattern; fill down)

| Cell | Formula | Result |
|---|---|---|
| `Book!E2` | `=IF(COUNTIFS(Bank!$D$2:$D$5,$D2,Bank!$C$2:$C$5,$C2)>0,"Exact",IF(F2<>"","Tolerance","Unmatched"))` | Exact |
| `Book!F2` | `=IFERROR(LOOKUP(2,1/((Bank!$D$2:$D$5=$D2)*(ABS(Bank!$C$2:$C$5-$C2)<=BankRec!$B$2)*(ABS(Bank!$B$2:$B$5-$B2)<=BankRec!$B$3)),Bank!$A$2:$A$5),"")` | K1 |
| `Book!G2` | `=IF(COUNTIFS($D$2:$D$6,$D2,$C$2:$C$6,$C2)>1,"DUPLICATE","")` |  |

| Cell | Formula | Result |
|---|---|---|
| `Bank!E2` | `=IF(COUNTIFS(Book!$D$2:$D$6,$D2,Book!$C$2:$C$6,$C2)>0,"Exact",IF(F2<>"","Tolerance","Unmatched"))` | Exact |
| `Bank!F2` | `=IFERROR(LOOKUP(2,1/((Book!$D$2:$D$6=$D2)*(ABS(Book!$C$2:$C$6-$C2)<=BankRec!$B$2)*(ABS(Book!$B$2:$B$6-$B2)<=BankRec!$B$3)),Book!$A$2:$A$6),"")` | B1 |

### How the tolerance match works (the `LOOKUP(2,1/(…))` trick)

```
(Bank!$D$2:$D$5=$D2)                         → {TRUE,FALSE,FALSE,FALSE}  same reference?
*(ABS(Bank!$C$2:$C$5-$C2)<=BankRec!$B$2)      → amount within tolerance?
*(ABS(Bank!$B$2:$B$5-$B2)<=BankRec!$B$3)      → date within N days?
= {1,0,0,0}                                   (1 only where ALL conditions hold)
1/{1,0,0,0} = {1,#DIV/0!,#DIV/0!,#DIV/0!}
LOOKUP(2, that array, Bank!$A$2:$A$5)         → "K1"
```
`LOOKUP` searches for 2, can't find it, and returns the **last numeric** position, ignoring errors. That gives a multi-condition lookup that works in every Excel version with no Ctrl+Shift+Enter. `IFERROR(…,"")` returns blank when nothing matches.

**Status logic (E2):** exact reference + amount → "Exact"; otherwise a tolerance match → "Tolerance"; otherwise "Unmatched".

### Results

| | A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|---|
| **1** | ID | Date | Amount | Ref | Status | Matched bank ID | Duplicate? |
| **2** | B1 | 2026-08-02 | 1,200.00 | INV1001 | Exact | K1 |  |
| **3** | B2 | 2026-08-05 | -350.00 | CHK2201 | Exact | K2 |  |
| **4** | B3 | 2026-08-10 | 4,999.50 | INV1002 | Tolerance | K3 |  |
| **5** | B4 | 2026-08-28 | -2,000.00 | CHK2202 | Unmatched |  |  |
| **6** | B5 | 2026-08-31 | 780.00 | DEP0831 | Unmatched |  |  |

| | A | B | C | D | E | F |
|---|---|---|---|---|---|---|
| **1** | ID | Date | Amount | Ref | Status | Matched book ID |
| **2** | K1 | 2026-08-03 | 1,200.00 | INV1001 | Exact | B1 |
| **3** | K2 | 2026-08-07 | -350.00 | CHK2201 | Exact | B2 |
| **4** | K3 | 2026-08-11 | 4,995.00 | INV1002 | Tolerance | B3 |
| **5** | K4 | 2026-08-15 | -25.00 | FEE | Unmatched |  |

### The reconciliation proof (`BankRec`)

| Cell | Formula | Result |
|---|---|---|
| `BankRec!B9` | `=B5` | 10,000.00 |
| `BankRec!B10` | `=SUMIFS(Book!$C$2:$C$6,Book!$E$2:$E$6,"Unmatched",Book!$C$2:$C$6,">0")` | 780.00 |
| `BankRec!B11` | `=-SUMIFS(Book!$C$2:$C$6,Book!$E$2:$E$6,"Unmatched",Book!$C$2:$C$6,"<0")` | 2,000.00 |
| `BankRec!B12` | `=B9+B10-B11` | 8,780.00 |
| `BankRec!B14` | `=B6` | 8,809.50 |
| `BankRec!B15` | `=SUMIFS(Bank!$C$2:$C$5,Bank!$E$2:$E$5,"Unmatched")` | -25.00 |
| `BankRec!B16` | `=SUMIFS(Book!$C$2:$C$6,Book!$E$2:$E$6,"Tolerance")-SUMIFS(Bank!$C$2:$C$5,Bank!$E$2:$E$5,"Tolerance")` | 4.50 |
| `BankRec!B17` | `=B14+B15-B16` | 8,780.00 |
| `BankRec!B19` | `=ROUND(B12-B17,2)` | 0.00 |
| `BankRec!B20` | `=IF(B19=0,"Reconciled","Difference: "&TEXT(B19,"#,##0.00"))` | Reconciled |
| `BankRec!B22` | `=-(B15-B16)` | 29.50 |

| | A | B |
|---|---|---|
| **8** | BANK SIDE |  |
| **9** | Balance per bank | 10,000.00 |
| **10** | + Deposits in transit | 780.00 |
| **11** | - Outstanding checks | 2,000.00 |
| **12** | Adjusted bank balance | 8,780.00 |
| **13** | BOOK SIDE |  |
| **14** | Balance per GL | 8,809.50 |
| **15** | +/- Bank-only items (fees, interest) | -25.00 |
| **16** | - Book errors (book minus bank on tolerance matches) | 4.50 |
| **17** | Adjusted book balance | 8,780.00 |
| **18** |  |  |
| **19** | Difference | 0.00 |
| **20** | Status | Reconciled |
| **21** |  |  |
| **22** | Journal entry needed (bank fee + short-paid wire) | 29.50 |

The proof classifies each unmatched item by sign and source:
- Unmatched **book** deposits → deposits in transit
- Unmatched **book** payments → outstanding checks
- Unmatched **bank** items → bank-only items (need a journal entry)
- Tolerance matches → book errors (need a correcting entry)

`B22` is the journal entry: Dr Bank fees 29.50 / Cr Cash 29.50.

### Modern equivalents (Tier C, not machine-verified)
```excel
Tolerance match (first match):
=XLOOKUP(1, (Bank!$D$2:$D$5=D2)*(ABS(Bank!$C$2:$C$5-C2)<=BankRec!$B$2)
            *(ABS(Bank!$B$2:$B$5-B2)<=BankRec!$B$3), Bank!$A$2:$A$5, "")
   (add search_mode -1 as the 6th argument to return the LAST match, like LOOKUP)

All unmatched items on one sheet:
=VSTACK(FILTER(Book!A2:D6, Book!E2:E6="Unmatched", "none"),
        FILTER(Bank!A2:D5, Bank!E2:E5="Unmatched", "none"))
```

### Excel UI features

- Conditional formatting on `Book!A2:G6`: formula `=$E2="Unmatched"` → orange fill; `=$E2="Tolerance"` → yellow fill.
- Big status cell: format `BankRec!B20` with a green fill rule when `=$B$19=0` and red otherwise.
- Protect the sheet (Review → Protect Sheet) leaving only `B2:B6` unlocked. That stops accidental overwrites of the proof.

### Try it

**E7.1 — Tighten amount tolerance to 1.00**  
Change: `BankRec!B2` → `1`

| Watch cell | Before | After |
|---|---|---|
| `Book!E4` | Tolerance | Unmatched |
| `Bank!E4` | Tolerance | Unmatched |
| `BankRec!B10` | 780.00 | 5,779.50 |
| `BankRec!B15` | -25.00 | 4,970.00 |
| `BankRec!B16` | 4.50 | 0.00 |
| `BankRec!B12` | 8,780.00 | 13,779.50 |
| `BankRec!B17` | 8,780.00 | 13,779.50 |
| `BankRec!B19` | 0.00 | 0.00 |

**E7.2 — Bank statement balance keyed wrong (10,100)**  
Change: `BankRec!B5` → `10100`

| Watch cell | Before | After |
|---|---|---|
| `BankRec!B19` | 0.00 | 100.00 |
| `BankRec!B20` | Reconciled | Difference: 100.00 |

**Lesson from E7.1 — "balanced is not the same as correct":** with a $1 tolerance the short-paid wire no longer matches. The reconciliation *still* ties (difference 0), but the items are now misclassified: a real receipt appears as a "deposit in transit" and a "bank-only item." A reviewer must check each reconciling item's nature, not just the zero.

---

## Lab 8 — AR Aging, Allowance, and Subledger-to-GL Reconciliation

### Inputs

`AR!A1:D7` open invoices, and the as-of date in `AR!H2` = 2026-08-31:

| | A | B | C | D |
|---|---|---|---|---|
| **1** | Invoice | Customer | Due date | Open amount |
| **2** | A1 | Acme | 2026-05-15 | 3,000 |
| **3** | A2 | Beta | 2026-07-20 | 1,500 |
| **4** | A3 | Acme | 2026-08-25 | 2,200 |
| **5** | A4 | Cyan | 2026-09-10 | 900 |
| **6** | A5 | Beta | 2026-06-25 | 1,200 |
| **7** | A6 | Cyan | 2026-08-10 | 650 |

### Aging formulas (row 2 pattern; fill down)

| Cell | Formula | Result |
|---|---|---|
| `AR!E2` | `=$H$2-C2` | 108 |
| `AR!F2` | `=IF(E2<=0,"Current",IF(E2<=30,"1-30",IF(E2<=60,"31-60",IF(E2<=90,"61-90","90+"))))` | 90+ |

Dates are serial numbers, so `$H$2-C2` returns days. Negative days = not yet due.

**Three equivalent bucket formulas:**

| Tier | Formula for `F2` | Notes |
|---|---|---|
| A (verified) | `=IF(E2<=0,"Current",IF(E2<=30,"1-30",IF(E2<=60,"31-60",IF(E2<=90,"61-90","90+"))))` | Nested IF; works everywhere |
| A (verified separately) | `=LOOKUP(E2,{-99999,1,31,61,91},{"Current","1-30","31-60","61-90","90+"})` | Approximate-match lookup: finds the largest threshold ≤ days. Easy to extend |
| B/C | `=IFS(E2<=0,"Current",E2<=30,"1-30",E2<=60,"31-60",E2<=90,"61-90",TRUE,"90+")` | Flat and readable (Excel 2019+) |

The `LOOKUP` version was tested separately at the boundary values −10, 0, 1, 30, 31, 60, 61, 90, 91, and 108, and matched the nested IF at every one.

| | A | B | C | D | E | F |
|---|---|---|---|---|---|---|
| **1** | Invoice | Customer | Due date | Open amount | Days past due | Bucket |
| **2** | A1 | Acme | 2026-05-15 | 3,000 | 108 | 90+ |
| **3** | A2 | Beta | 2026-07-20 | 1,500 | 42 | 31-60 |
| **4** | A3 | Acme | 2026-08-25 | 2,200 | 6 | 1-30 |
| **5** | A4 | Cyan | 2026-09-10 | 900 | -10 | Current |
| **6** | A5 | Beta | 2026-06-25 | 1,200 | 67 | 61-90 |
| **7** | A6 | Cyan | 2026-08-10 | 650 | 21 | 1-30 |

### Summary, allowance, and customer matrix (`ARSummary`)

| Cell | Formula | Result |
|---|---|---|
| `ARSummary!B4` | `=SUMIFS(AR!$D$2:$D$7,AR!$F$2:$F$7,$A4)` | 900 |
| `ARSummary!C4` | `=B4/$B$9` | 9.5% |
| `ARSummary!E4` | `=B4*D4` | 4.50 |
| `ARSummary!B9` | `=SUM(B4:B8)` | 9,450 |
| `ARSummary!E9` | `=SUMPRODUCT(B4:B8,D4:D8)` | 1,816.50 |
| `ARSummary!B12` | `=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A12,AR!$F$2:$F$7,B$11)` | 0 |
| `ARSummary!G12` | `=SUM(B12:F12)` | 5,200 |
| `ARSummary!B18` | `=SUM(AR!$D$2:$D$7)` | 9,450 |
| `ARSummary!B19` | `=B17-B18` | 100 |

| | A | B | C | D | E |
|---|---|---|---|---|---|
| **3** | Bucket | Amount | % of total | Reserve % | Allowance |
| **4** | Current | 900 | 9.5% | 0.5% | 4.50 |
| **5** | 1-30 | 2,850 | 30.2% | 2.0% | 57.00 |
| **6** | 31-60 | 1,500 | 15.9% | 5.0% | 75.00 |
| **7** | 61-90 | 1,200 | 12.7% | 15.0% | 180.00 |
| **8** | 90+ | 3,000 | 31.7% | 50.0% | 1,500.00 |
| **9** | Total | 9,450 |  |  | 1,816.50 |

| | A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|---|
| **11** | Customer \ Bucket | Current | 1-30 | 31-60 | 61-90 | 90+ | Total |
| **12** | Acme | 0 | 2,200 | 0 | 0 | 3,000 | 5,200 |
| **13** | Beta | 0 | 0 | 1,500 | 1,200 | 0 | 2,700 |
| **14** | Cyan | 900 | 650 | 0 | 0 | 0 | 1,550 |

### Subledger-to-GL tie-out

| | A | B |
|---|---|---|
| **17** | GL control account (1200-AR) | 9,550 |
| **18** | Subledger total | 9,450 |
| **19** | Difference (investigate) | 100 |

The difference of 100 is a real reconciling item to investigate. Typical causes are a manual journal entry posted directly to the AR control account, or unapplied cash.

### Excel UI features
- Conditional formatting on `AR!E2:E7`: Color Scales (green → red) to visualize age.
- Sort `AR` by days past due, descending (Data → Sort) to create a collections worklist.
- PivotTable alternative: Rows = Customer, Columns = Bucket, Values = Sum of Open amount.

### Try it

**E8.1 — Roll the as-of date to 2026-09-30**  
Change: `AR!H2` → `2026-09-30`

| Watch cell | Before | After |
|---|---|---|
| `ARSummary!B4` | 900 | 0 |
| `ARSummary!B5` | 2,850 | 900 |
| `ARSummary!B6` | 1,500 | 2,850 |
| `ARSummary!B7` | 1,200 | 1,500 |
| `ARSummary!B8` | 3,000 | 4,200 |
| `ARSummary!E9` | 1,816.50 | 2,485.50 |

---

## Lab 9 — Forecasting: Moving Average, Trend, Seasonality, Backtest

### Data

`Forecast!A5:C34` holds 30 months of sales (Jan-2024 → Jun-2026) with a trend and a December peak. `B` is a period counter `t` = 1…30. The train/test split input is `B2` = 24 (train on the first 24 months, test on the last 6).

| | A | B | C |
|---|---|---|---|
| **4** | Month | t | Sales |
| **5** | Jan-24 | 1 | 90,609 |
| **6** | Feb-24 | 2 | 91,300 |
| **7** | Mar-24 | 3 | 104,501 |
| **8** | Apr-24 | 4 | 108,471 |
| **9** | May-24 | 5 | 107,398 |
| **10** | Jun-24 | 6 | 115,646 |

*(rows 11–34 continue the series; the full data is in Appendix C)*

### Step 1 — Helper columns (row 5 pattern; fill down to row 34)

| Cell | Formula | Result |
|---|---|---|
| `Forecast!D7` | `=AVERAGE(C5:C7)` | 95,470 |
| `Forecast!E5` | `=MONTH(A5)` | 1 |
| `Forecast!F5` | `=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B5` | 99,690 |
| `Forecast!G5` | `=C5/F5` | 0.9089 |
| `Forecast!H5` | `=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B5` | 98,073 |
| `Forecast!I5` | `=C5/H5` | 0.9239 |

- `D7` starts the 3-month moving average (it needs two prior months).
- **Trend (all)** uses all 30 months. **Trend (train)** uses only rows 5–28 (t ≤ 24). The backtest must never "peek" at the test months.
- **Ratio** = actual ÷ trend. A December ratio of about 1.2 means December runs ~20% above trend.

### Step 2 — Regression statistics

| Cell | Formula | Result |
|---|---|---|
| `Forecast!L19` | `=SLOPE($C$5:$C$34,$B$5:$B$34)` | 1,674.6 |
| `Forecast!L20` | `=INTERCEPT($C$5:$C$34,$B$5:$B$34)` | 98,015.0 |
| `Forecast!L21` | `=RSQ($C$5:$C$34,$B$5:$B$34)` | 0.694 |

Sales grow by about 1,674.6 per month. R² of 0.694 means the straight line alone explains about 69% of the variation; seasonality explains much of the rest.

### Step 3 — Seasonal index table (`K5:O16`)

| Cell | Formula | Result |
|---|---|---|
| `Forecast!L5` | `=AVERAGEIFS($G$5:$G$34,$E$5:$E$34,K5)` | 0.8885 |
| `Forecast!M5` | `=L5/AVERAGE($L$5:$L$16)` | 0.8843 |
| `Forecast!N5` | `=AVERAGEIFS($I$5:$I$34,$E$5:$E$34,K5,$B$5:$B$34,"<="&$B$2)` | 0.9017 |
| `Forecast!O5` | `=N5/AVERAGE($N$5:$N$16)` | 0.9017 |

Fill `L5:O5` down to row 16 (`K5:K16` = 1…12). Dividing by the average **normalizes** the index so the 12 values average exactly 1.0 (`M17`). Otherwise the seasonal forecast would be biased up or down overall.

| | K | L | M | N | O |
|---|---|---|---|---|---|
| **4** | Month # | Raw index (all) | Seasonal index (all) | Raw index (train) | Seasonal index (train) |
| **5** | 1 | 0.8885 | 0.8843 | 0.9017 | 0.9017 |
| **6** | 2 | 0.9044 | 0.9001 | 0.9134 | 0.9134 |
| **7** | 3 | 0.9941 | 0.9894 | 1.0017 | 1.0017 |
| **8** | 4 | 1.0058 | 1.0010 | 1.0092 | 1.0092 |
| **9** | 5 | 1.0225 | 1.0177 | 1.0175 | 1.0175 |
| **10** | 6 | 1.0676 | 1.0626 | 1.0593 | 1.0594 |
| **11** | 7 | 1.0725 | 1.0675 | 1.0636 | 1.0637 |
| **12** | 8 | 1.0403 | 1.0354 | 1.0299 | 1.0299 |
| **13** | 9 | 0.9817 | 0.9771 | 0.9702 | 0.9703 |
| **14** | 10 | 0.9499 | 0.9454 | 0.9372 | 0.9372 |
| **15** | 11 | 0.9484 | 0.9439 | 0.9341 | 0.9342 |
| **16** | 12 | 1.1812 | 1.1756 | 1.1617 | 1.1618 |
| **17** | Average |  | 1.0000 |  | 1.0000 |

### Step 4 — Forecast the next 6 months (rows 38–43)

| Cell | Formula | Result |
|---|---|---|
| `Forecast!A38` | `=EDATE($A$34,1)` | Jul-26 |
| `Forecast!B38` | `=$B$34+1` | 31 |
| `Forecast!C38` | `=FORECAST.LINEAR(B38,$C$5:$C$34,$B$5:$B$34)` | 149,928 |
| `Forecast!D38` | `=TREND($C$5:$C$34,$B$5:$B$34,B38)` | 149,928 |
| `Forecast!E38` | `=C38*INDEX($M$5:$M$16,MONTH(A38))` | 160,042 |
| `Forecast!A39` | `=EDATE(A38,1)` | Aug-26 |
| `Forecast!B39` | `=B38+1` | 32 |

| | A | B | C | D | E |
|---|---|---|---|---|---|
| **37** | Month | t | Linear (FORECAST.LINEAR) | Linear (TREND) | Trend x season |
| **38** | Jul-26 | 31 | 149,928 | 149,928 | 160,042 |
| **39** | Aug-26 | 32 | 151,603 | 151,603 | 156,971 |
| **40** | Sep-26 | 33 | 153,277 | 153,277 | 149,763 |
| **41** | Oct-26 | 34 | 154,952 | 154,952 | 146,499 |
| **42** | Nov-26 | 35 | 156,627 | 156,627 | 147,837 |
| **43** | Dec-26 | 36 | 158,301 | 158,301 | 186,102 |

- `FORECAST.LINEAR` and `TREND` give identical straight-line forecasts (C = D).
- **Trend × season** (E) layers in the seasonal pattern: December jumps to 186,102.

### Step 5 — Backtest against the last 6 months (rows 47–57)

| Cell | Formula | Result |
|---|---|---|
| `Forecast!A47` | `=A29` | Jan-26 |
| `Forecast!B47` | `=C29` | 121,543 |
| `Forecast!C47` | `=H29*INDEX($O$5:$O$16,E29)` | 129,621 |
| `Forecast!D47` | `=C17` | 106,332 |
| `Forecast!E47` | `=ABS(B47-C47)/ABS(B47)` | 6.65% |
| `Forecast!F47` | `=ABS(B47-D47)/ABS(B47)` | 12.51% |
| `Forecast!B54` | `=AVERAGE(E47:E52)` | 4.62% |
| `Forecast!B55` | `=(SUM(C47:C52)-SUM(B47:B52))/SUM(B47:B52)` | +4.48% |
| `Forecast!B56` | `=AVERAGE(F47:F52)` | 12.94% |
| `Forecast!B57` | `=B54<B56` | TRUE |

| | A | B | C | D | E | F |
|---|---|---|---|---|---|---|
| **46** | Month (test) | Actual | Model (train trend x train index) | Seasonal naive (t-12) | APE model | APE naive |
| **47** | Jan-26 | 121,543 | 129,621 | 106,332 | 6.65% | 12.51% |
| **48** | Feb-26 | 125,796 | 133,043 | 112,194 | 5.76% | 10.81% |
| **49** | Mar-26 | 140,065 | 147,817 | 121,935 | 5.53% | 12.94% |
| **50** | Apr-26 | 144,041 | 150,844 | 123,231 | 4.72% | 14.45% |
| **51** | May-26 | 149,925 | 154,019 | 130,938 | 2.73% | 12.66% |
| **52** | Jun-26 | 158,712 | 162,366 | 136,132 | 2.30% | 14.23% |
| **53** |  |  |  |  |  |  |
| **54** | MAPE (model) | 4.62% |  |  |  |  |
| **55** | Bias (model) | +4.48% |  |  |  |  |
| **56** | MAPE (seasonal naive) | 12.94% |  |  |  |  |
| **57** | Model beats naive? | TRUE |  |  |  |  |

**Reading it:** the model's MAPE is 4.62% against 12.94% for the seasonal-naive benchmark ("same month last year"), so the model adds value. However, the bias is +4.48%: the model **over-forecast consistently**. Growth may be slowing, and the next step is to investigate the drivers.

### Step 6 — Excel's built-in exponential smoothing (Tier B, Excel 2016+; not machine-verified)

Enter these next to the forecast block, e.g., in `G38` and `H38`, and fill down:
```excel
G38 (ETS forecast):     =FORECAST.ETS(A38, $C$5:$C$34, $A$5:$A$34, 1)
H38 (95% band radius):  =FORECAST.ETS.CONFINT(A38, $C$5:$C$34, $A$5:$A$34, 0.95, 1)
Upper bound:            =G38+H38
Lower bound:            =G38-H38
Detected season length: =FORECAST.ETS.SEASONALITY($C$5:$C$34, $A$5:$A$34)   → expect 12
```
- The timeline (`A5:A34`) must be evenly spaced (first of each month) with no duplicates.
- Seasonality argument: `1` = auto-detect, `0` = none, `12` = force 12 months.
- ETS results differ between spreadsheet programs (LibreOffice uses a different implementation), so **don't auto-grade ETS outputs with fixed expected values.** Grade that they are within a tolerance band, or only in genuine Excel.
- **Forecast Sheet shortcut:** select `A4:C34` → Data → **Forecast Sheet** → choose the end date and confidence interval → Create. Excel builds the ETS table and chart for you.

### Modern equivalents (Tier C/D, not machine-verified)
```excel
Seasonal index in one spill:
=LET(m, SEQUENCE(12), raw, AVERAGEIFS(G5:G34, E5:E34, m), raw / AVERAGE(raw))

Full regression stats (slope, intercept; with stats=TRUE also R², standard errors):
=LINEST(C5:C34, B5:B34, TRUE, TRUE)
```

### Charting
1. Select `A4:A34` and `C4:C34` → Insert → Line chart.
2. Right-click → Select Data → Add series "Trend × season" using `A38:A43` and `E38:E43`.
3. Format the forecast series with a dashed line, and label the axis units.

### Try it

**E9.1 — Insert an outlier: Apr-24 sales = 250,000**  
Change: `Forecast!C8` → `250000`

| Watch cell | Before | After |
|---|---|---|
| `Forecast!L19` | 1,674.6 | 950.4 |
| `Forecast!M8` | 1.0010 | 1.3696 |
| `Forecast!E41` | 146,499 | 134,252 |
| `Forecast!B54` | 4.62% | 16.18% |

**Lesson from E9.1:** one bad data point (a mis-keyed April) cut the estimated growth rate nearly in half, inflated April's seasonal index to about 1.37, and tripled the error. **Always scan history for outliers before forecasting** (a quick line chart is usually enough).

---

## Lab 10 — Driver-Based Forecast, Scenarios, and Sensitivity

### Inputs (`Drivers!A3:E10`)

| | A | B | C | D | E |
|---|---|---|---|---|---|
| **3** | Driver | Downside | Base | Upside | Active |
| **4** | New customers / month | 60 | 80 | 100 | 80 |
| **5** | Monthly churn | 3.0% | 2.0% | 1.5% | 2.0% |
| **6** | ARPU ($ / month) | 90.00 | 95.00 | 98.00 | 95.00 |
| **7** |  |  |  |  |  |
| **8** | Scenario # (1=Down, 2=Base, 3=Up) |  | 2 |  |  |
| **9** | Active scenario name |  | Base |  |  |
| **10** | Starting customers | 1,200 |  |  |  |

- `C8` is the **scenario switch** (1 = Downside, 2 = Base, 3 = Upside). Add a dropdown: Data Validation → List → `1,2,3`.
- Column E ("Active") pulls the selected scenario's value:

| Cell | Formula | Result |
|---|---|---|
| `Drivers!E4` | `=INDEX(B4:D4,$C$8)` | 80 |
| `Drivers!E5` | `=INDEX(B5:D5,$C$8)` | 2.0% |
| `Drivers!E6` | `=INDEX(B6:D6,$C$8)` | 95.00 |
| `Drivers!C9` | `=INDEX($B$3:$D$3,$C$8)` | Base |

Classic alternative: `=CHOOSE($C$8,B4,C4,D4)`.

### Customer roll-forward (columns B–G, rows 13–18)

| Cell | Formula | Result |
|---|---|---|
| `Drivers!B14` | `=$B$10` | 1,200 |
| `Drivers!B15` | `=-B14*$E$5` | (24) |
| `Drivers!B16` | `=$E$4` | 80 |
| `Drivers!B17` | `=B14+B15+B16` | 1,256 |
| `Drivers!B18` | `=B17*$E$6` | 119,320 |
| `Drivers!C13` | `=EDATE(B13,1)` | Nov-26 |
| `Drivers!C14` | `=B17` | 1,256 |
| `Drivers!H18` | `=SUM(B18:G18)` | 792,070 |

Fill `C13:C18` across to column G. Each month's opening customers equal the prior month's closing customers (`C14 = B17`). This "roll-forward" pattern is used everywhere in finance: customers, headcount, inventory, debt, cash.

| | A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|---|
| **13** | Month | Oct-26 | Nov-26 | Dec-26 | Jan-27 | Feb-27 | Mar-27 | 6-mo total |
| **14** | Opening customers | 1,200 | 1,256 | 1,311 | 1,365 | 1,417 | 1,469 |  |
| **15** | Churned | (24) | (25) | (26) | (27) | (28) | (29) |  |
| **16** | New | 80 | 80 | 80 | 80 | 80 | 80 |  |
| **17** | Closing customers | 1,256 | 1,311 | 1,365 | 1,417 | 1,469 | 1,520 |  |
| **18** | Revenue | 119,320 | 124,534 | 129,643 | 134,650 | 139,557 | 144,366 | 792,070 |

### All scenarios at once (closed form, row 21)

A roll-forward only shows the *active* scenario. To show every scenario side by side, use the closed-form formula for customers after k months. With churn *c*, new customers *n*, and starting customers *C₀*: Cₖ = C₀(1−c)ᵏ + n·(1−(1−c)ᵏ)/c. Summing revenue over k = 1…6 using the helper row `B25:G25` (= 1…6):

| Cell | Formula | Result |
|---|---|---|
| `Drivers!B21` | `=B6*SUMPRODUCT($B$10*(1-B5)^$B$25:$G$25+B4*(1-(1-B5)^$B$25:$G$25)/B5)` | 691,159 |
| `Drivers!C21` | `=C6*SUMPRODUCT($B$10*(1-C5)^$B$25:$G$25+C4*(1-(1-C5)^$B$25:$G$25)/C5)` | 792,070 |
| `Drivers!D21` | `=D6*SUMPRODUCT($B$10*(1-D5)^$B$25:$G$25+D4*(1-(1-D5)^$B$25:$G$25)/D5)` | 870,200 |

| | A | B | C | D |
|---|---|---|---|---|
| **21** | 6-mo revenue by scenario (closed form) | 691,159 | 792,070 | 870,200 |

`Checks!B9` confirms that the roll-forward total (`H18`) equals the closed form for the active scenario.

### Sensitivity grid: churn × ARPU (rows 28–33)

| Cell | Formula | Result |
|---|---|---|
| `Drivers!B29` | `=B$28*SUMPRODUCT($B$10*(1-$A29)^$B$25:$G$25+$E$4*(1-(1-$A29)^$B$25:$G$25)/$A29)` | 731,377 |

Fill `B29` across to `F29` and down to row 33 (mixed references `B$28` and `$A29`).

| | A | B | C | D | E | F |
|---|---|---|---|---|---|---|
| **28** | Churn \ ARPU | 85 | 90 | 95 | 100 | 105 |
| **29** | 1.0% | 731,377 | 774,399 | 817,421 | 860,444 | 903,466 |
| **30** | 1.5% | 719,944 | 762,294 | 804,644 | 846,993 | 889,343 |
| **31** | 2.0% | 708,694 | 750,382 | 792,070 | 833,757 | 875,445 |
| **32** | 2.5% | 697,623 | 738,659 | 779,696 | 820,733 | 861,769 |
| **33** | 3.0% | 686,729 | 727,125 | 767,520 | 807,916 | 848,312 |

The center cell (2.0% churn, $95 ARPU) equals the Base scenario. Each 0.5-point increase in churn costs about $12–13K of 6-month revenue at $95 ARPU.

### Excel-native alternatives

**What-If Data Table** (same grid, Excel's built-in way):
1. Set `C8` = 2 (Base). In an empty corner cell, e.g. `J28`, enter `=H18`.
2. Put ARPU values in `K28:O28` and churn values in `J29:J33`.
3. Select `J28:O33` → Data → What-If Analysis → **Data Table**. Row input cell = `C6` (Base ARPU); Column input cell = `C5` (Base churn).
4. Excel fills the grid with `{=TABLE(C6,C5)}`. Data Tables are Excel-only and slow in big models, which is why the formula grid above is the portable version. Use it in the web app.

**Goal Seek** ("what churn gives $800K?"):
1. Set `C8` = 2. Data → What-If Analysis → **Goal Seek**.
2. Set cell `H18`, To value `800000`, By changing cell `C5`.
3. Expected answer: churn ≈ **1.684%** (verified numerically: 0.016837). Equivalently, at 2% churn, ARPU would need to be ≈ **$95.95**.

**Scenario Manager** (Data → What-If Analysis → Scenario Manager) can store the three input sets. The switch cell in `C8` is more transparent and easier to audit.

### Modern equivalents (Tier D, not machine-verified)
```excel
Closing customers for all 6 months in one formula (spills across):
=TRANSPOSE(SCAN($B$10, SEQUENCE(6), LAMBDA(c, k, c*(1-$E$5)+$E$4)))

Six-month revenue:
=SUM(SCAN($B$10, SEQUENCE(6), LAMBDA(c, k, c*(1-$E$5)+$E$4))) * $E$6
```

### Try it

**E10.1 — Switch to the Upside scenario (C8 = 3)**  
Change: `Drivers!C8` → `3`

| Watch cell | Before | After |
|---|---|---|
| `Drivers!C9` | Base | Upside |
| `Drivers!G17` | 1,520 | 1,674 |
| `Drivers!H18` | 792,070 | 870,200 |
| `Checks!B9` | TRUE | TRUE |

**E10.2 — Base scenario with churn 3% (C5 = 0.03)**  
Change: `Drivers!C5` → `0.03`

| Watch cell | Before | After |
|---|---|---|
| `Drivers!H18` | 792,070 | 767,520 |
| `Drivers!B33` | 686,729 | 686,729 |
| `Drivers!D33` | 767,520 | 767,520 |

In E10.2 the grid cells (`B33`, `D33`) don't change, because the grid uses its own churn values in column A. Only the roll-forward `H18` changes, and it now equals the grid's 3%/$95 cell.

---

## Lab 11 — Reusable LAMBDA Function Library (Microsoft 365)

`LAMBDA` lets you define your own functions once and reuse them like built-ins. That standardizes logic across the team and removes copy-paste formula drift.

### How to create one
1. Formulas → **Name Manager** → New.
2. Name: `VARPCT`. Refers to: the LAMBDA formula below. OK.
3. Use it anywhere: `=VARPCT(D6, C6)`.
4. Test a LAMBDA before naming it by calling it inline: `=LAMBDA(a,b,(a-b)/ABS(b))(468000,500000)` → −6.4%.

### Library (not machine-verified; each should match the verified Tier A result shown)

| Name | Definition (paste into "Refers to") | Test call | Should equal |
|---|---|---|---|
| `VARPCT` | `=LAMBDA(actual,budget,IF(budget=0,"n/a",(actual-budget)/ABS(budget)))` | `=VARPCT(BvA!D6,BvA!C6)` | `BvA!F6` = -6.4% |
| `FAVFLAG` | `=LAMBDA(type,actual,budget,IF(budget=0,"No budget",IF(actual=budget,"On budget",IF((type="Revenue")=(actual>budget),"Favorable","Unfavorable"))))` | `=FAVFLAG(BvA!B9,BvA!D9,BvA!C9)` | `BvA!G9` = Unfavorable |
| `NEEDSCOMMENT` | `=LAMBDA(var,pct,absThr,pctThr,IF(ISNUMBER(pct),AND(ABS(var)>=absThr,ABS(pct)>=pctThr),FALSE))` | `=NEEDSCOMMENT(BvA!E9,BvA!F9,5000,0.05)` | `BvA!H9` = TRUE |
| `AGEBUCKET` | `=LAMBDA(days,IFS(days<=0,"Current",days<=30,"1-30",days<=60,"31-60",days<=90,"61-90",TRUE,"90+"))` | `=AGEBUCKET(AR!E2)` | `AR!F2` = 90+ |
| `MAPE` | `=LAMBDA(actual,forecast,AVERAGE(ABS(actual-forecast)/ABS(actual)))` | `=MAPE(Forecast!B47:B52,Forecast!C47:C52)` | `Forecast!B54` = 4.62% |
| `BIAS` | `=LAMBDA(actual,forecast,(SUM(forecast)-SUM(actual))/SUM(actual))` | `=BIAS(Forecast!B47:B52,Forecast!C47:C52)` | `Forecast!B55` = +4.48% |
| `CUSTROLL` | `=LAMBDA(start,churn,new,n,SCAN(start,SEQUENCE(n),LAMBDA(c,k,c*(1-churn)+new)))` | `=INDEX(CUSTROLL(1200,0.02,80,6),6)` | `Drivers!G17` = 1,520 |
| `PVMCHECK` | `=LAMBDA(price,vol,mix,total,ROUND(price+vol+mix-total,6)=0)` | `=PVMCHECK(PVM!J8,PVM!K8,PVM!L8,PVM!M8)` | TRUE |

**Tips**
- Document each LAMBDA in the Name Manager's *Comment* field (arguments and units).
- Share a library by copying one sheet that uses the names into a new workbook; the names travel with it.
- Keep a Tier A fallback column during rollout so anyone on older Excel can still verify results.

---

## Lab 12 — Model Integrity Checks Dashboard

Every serious model gets a **Checks** sheet: one place that proves the model is internally consistent.

| Cell | Formula | Result |
|---|---|---|
| `Checks!B3` | `=BvA!D14=0` | TRUE |
| `Checks!C3` | `=BvA!D14` | 0 |
| `Checks!B4` | `=PivotLab!B10=0` | TRUE |
| `Checks!C4` | `=PivotLab!B10` | 0 |
| `Checks!B5` | `=ROUND(PVM!N8,6)=0` | TRUE |
| `Checks!C5` | `=PVM!N8` | 0 |
| `Checks!B6` | `=CostVar!B13=0` | TRUE |
| `Checks!C6` | `=CostVar!B13` | 0 |
| `Checks!B7` | `=BankRec!B19=0` | TRUE |
| `Checks!C7` | `=BankRec!B19` | 0 |
| `Checks!B8` | `=ROUND(Forecast!M17,9)=1` | TRUE |
| `Checks!C8` | `=Forecast!M17` | 1 |
| `Checks!B9` | `=ROUND(Drivers!H18-INDEX(Drivers!B21:D21,Drivers!C8),2)=0` | TRUE |
| `Checks!C9` | `=Drivers!H18-INDEX(Drivers!B21:D21,Drivers!C8)` | 0 |
| `Checks!B10` | `=COUNTIF(B3:B9,FALSE)=0` | TRUE |
| `Checks!C12` | `=ARSummary!B19` | 100 |

| | A | B | C |
|---|---|---|---|
| **2** | Check | Pass? | Value |
| **3** | BvA: mapped actuals tie to GL | TRUE | 0 |
| **4** | PivotLab: cross-tab ties to BvA | TRUE | 0 |
| **5** | PVM: price+volume+mix = total | TRUE | 0 |
| **6** | CostVar: rate+volume = total | TRUE | 0 |
| **7** | BankRec: adjusted balances agree | TRUE | 0 |
| **8** | Forecast: seasonal index averages 1 | TRUE | 1 |
| **9** | Drivers: schedule = closed form | TRUE | 0 |
| **10** | ALL CONTROL CHECKS PASS | TRUE |  |
| **11** |  |  |  |
| **12** | Open item (informational): AR subledger vs GL difference |  | 100 |

- `B10` uses `COUNTIF(range,FALSE)=0` rather than `AND(range)`. It reads clearly and ignores blank rows.
- Row 12 is **informational**: a real reconciling item (the AR difference) that must be explained but isn't a model error, so it's excluded from the master check.
- Link the master check to the top of every output sheet, e.g., `="Model checks: "&IF(Checks!B10,"OK","FAILING — see Checks")`, and add conditional formatting (green/red).

**Checks worth adding to any finance model:** balance sheet balances (Assets − Liabilities − Equity = 0); cash flow ending cash = balance sheet cash; subtotals = sum of details; report totals = source totals; percentages sum to 100%; no `#REF!`/`#N/A` in outputs (`=SUMPRODUCT(--ISERROR(range))=0`); scenario switch within its allowed values.

---

## Appendix A — Function Tiers and Fallback Formulas

| Task | Tier D/C (modern) | Tier A fallback (verified pattern) |
|---|---|---|
| Lookup a value | `XLOOKUP(x, keys, vals, "none")` | `IFERROR(INDEX(vals, MATCH(x, keys, 0)), "none")` |
| Multi-condition lookup | `XLOOKUP(1, (c1)*(c2), vals, "")` | `IFERROR(LOOKUP(2, 1/((c1)*(c2)), vals), "")` (returns last match) |
| Count/sum with conditions | same | `COUNTIFS`, `SUMIFS` |
| Array AND inside formulas | `(c1)*(c2)=1` | `SUMPRODUCT((c1)*(c2))>0` |
| Filter rows | `FILTER(tbl, cond, "none")` | Helper column with a flag + AutoFilter, or `INDEX/SMALL(IF(...))` array formula |
| Distinct list | `UNIQUE(range)` | Remove Duplicates (Data tab) or a PivotTable |
| Grouped totals | `GROUPBY`, `PIVOTBY` | `SUMIFS` cross-tab (Lab 6) or PivotTable |
| Readable complex formula | `LET(name, value, …)` | Helper cells/columns |
| Reusable custom logic | `LAMBDA` + Name Manager | Copy the same formula pattern; document it |
| Range buckets | `IFS(...)` or `XLOOKUP(x, lows, labels, , -1)` | Nested `IF` or `LOOKUP(x, {lows}, {labels})` |
| Roll-forward sequence | `SCAN` | Row-by-row roll-forward (Lab 10) |
| Sequence of numbers | `SEQUENCE(n)` | Helper row typed 1…n (Lab 10 `B25:G25`) |
| Linear forecast | `FORECAST.LINEAR` | `FORECAST` (older name) or `TREND`, or `INTERCEPT+SLOPE*x` |
| Seasonal forecast | `FORECAST.ETS` | Trend × seasonal index (Lab 9) |

**Web engine note:** open-source JavaScript formula engines (for example HyperFormula or Formula.js) and embeddable spreadsheets each support their own function list, and support for dynamic-array functions varies. Check each engine's function list against the Tier A set above before choosing. If the app must run Tier C/D formulas exactly as Excel does, the alternative is to embed real Excel (Microsoft 365 / Office.js), which requires the user's Microsoft account.

---

## Appendix B — Web App Blueprint

This appendix turns the labs into a design you can build later.

### B.1 Architecture

```
┌────────────────────────────────────────────────────────────────┐
│ Browser                                                        │
│  ┌──────────────┐   ┌───────────────────┐   ┌───────────────┐  │
│  │ Lesson panel │   │ Spreadsheet grid  │   │ Checks panel  │  │
│  │ (markdown    │◄─►│ (editable cells,  │◄─►│ (assertions   │  │
│  │  from labs)  │   │  formula bar)     │   │  pass/fail)   │  │
│  └──────────────┘   └─────────┬─────────┘   └───────────────┘  │
│                               │ setCell / getValue             │
│                    ┌──────────▼──────────┐                     │
│                    │ Formula engine      │  loads spec JSON     │
│                    │ (Tier A canonical)  │  (Appendix C)        │
│                    └─────────────────────┘                     │
└────────────────────────────────────────────────────────────────┘
```

### B.2 Data contracts

**Workbook spec** (Appendix C), one object per sheet:
```json
{
  "workbook": "fin-analysis-excel-labs",
  "sheetOrder": ["Map", "GL", "..."],
  "sheets": {
    "BvA": {
      "B1": {"v": "2026-08"},
      "B2": {"v": 0.05, "fmt": "0.0%"},
      "C6": {"f": "=SUMIFS(GL!$E$2:$E$37, ...)", "fmt": "#,##0"}
    }
  }
}
```
- `v` = constant value; `{"date": "YYYY-MM-DD"}` = date (convert to an Excel serial or the engine's date type on load).
- `f` = formula (A1 syntax, English function names, comma separators).
- `fmt` = Excel number-format code for display.

**Assertion** (Appendix D), used for auto-grading:
```json
{"sheet": "BvA", "cell": "E6", "expected": -32000, "tolerance": 0.005}
```
Pass rule: for numbers, `|actual − expected| ≤ max(tolerance, |expected| × 1e-12)`; for text or booleans, exact equality.

**Experiment** (Appendix E), for guided what-if exercises:
```json
{"id": "E7.1", "lab": 7, "title": "...", "changes": {"BankRec!B2": 1},
 "expect": {"BankRec!B19": 0, "Book!E4": "Unmatched"}}
```

### B.3 Suggested lesson flow per lab (UI states)

| Mode | What the learner sees | Engine behavior |
|---|---|---|
| **Explore** | The completed lab sheet; clicking a cell shows its formula and an explanation | Load the full spec |
| **Build** | Inputs pre-filled, formula cells blank with hints ("SUMIFS: sum GL amounts where…") | Load inputs only; grade typed formulas against assertions |
| **What-if** | Sliders/dropdowns bound to input cells (thresholds, scenario #, tolerance) | `setCell` → recalc → show changed cells highlighted |
| **Challenge** | Experiment prompts ("Tighten tolerance to 1.00 — does the rec still tie? Why is it still wrong?") | Apply `changes`, compare with `expect` |
| **Modern view** | Tier C/D formula alongside the Tier A formula | Display only (or run if the engine supports it) |

### B.4 Inputs to expose as interactive controls

| Lab | Cell | Control |
|---|---|---|
| 2 | `BvA!B1` | Dropdown: 2025-08, 2026-07, 2026-08 |
| 2 | `BvA!B2`, `BvA!B3` | Sliders: 1–20%, 0–25,000 |
| 3 | `Flux!B4`, `Flux!B5` | Sliders |
| 4 | `PVM!B5:E7` | Editable table |
| 5 | `CostVar!B3:B6`, `B17:B24` | Numeric inputs |
| 7 | `BankRec!B2`, `B3`, `B5`, `B6` | Numeric inputs; toggle to add a transaction row |
| 8 | `AR!H2` | Date picker |
| 9 | `Forecast!C5:C34` | Editable series; "inject outlier" button |
| 10 | `Drivers!C8` | Segmented control: Down/Base/Up |
| 10 | `Drivers!B4:D6` | Sliders per driver |

### B.5 Grading formulas the learner types

Grade the **result**, not the text of the formula: many formulas are correct (e.g., `INDEX/MATCH` vs `XLOOKUP`). After the learner enters a formula, recalculate and check the relevant assertions. Optionally re-grade after applying an experiment's `changes`, to catch hard-coded numbers that only match by coincidence.

---

## Appendix C — Complete Cell-by-Cell Workbook Specification (JSON)

Every input value and formula in the verified workbook. Load these sheets in `sheetOrder` into a formula engine, or type them into Excel, and you reproduce every result in this guide. In the `.xlsx` file format, `FORECAST.LINEAR` is stored internally as `_xlfn.FORECAST.LINEAR`; type it normally in Excel.

#### Sheet `Map` (28 cells)

```json
{
  "A1": {"v": "Account"},
  "B1": {"v": "Type"},
  "C1": {"v": "PL_Line"},
  "D1": {"v": "Sign"},
  "A2": {"v": "Revenue"},
  "B2": {"v": "Revenue"},
  "C2": {"v": "Net revenue"},
  "D2": {"v": 1},
  "A3": {"v": "COGS"},
  "B3": {"v": "Expense"},
  "C3": {"v": "Cost of sales"},
  "D3": {"v": -1},
  "A4": {"v": "Salaries"},
  "B4": {"v": "Expense"},
  "C4": {"v": "Operating expense"},
  "D4": {"v": -1},
  "A5": {"v": "Marketing"},
  "B5": {"v": "Expense"},
  "C5": {"v": "Operating expense"},
  "D5": {"v": -1},
  "A6": {"v": "Rent"},
  "B6": {"v": "Expense"},
  "C6": {"v": "Operating expense"},
  "D6": {"v": -1},
  "A7": {"v": "Travel"},
  "B7": {"v": "Expense"},
  "C7": {"v": "Operating expense"},
  "D7": {"v": -1}
}
```

#### Sheet `GL` (185 cells)

```json
{
  "A1": {"v": "Period"},
  "B1": {"v": "Dept"},
  "C1": {"v": "Account"},
  "D1": {"v": "Version"},
  "E1": {"v": "Amount"},
  "A2": {"v": "2026-08"},
  "B2": {"v": "Sales"},
  "C2": {"v": "Revenue"},
  "D2": {"v": "BUD"},
  "E2": {"v": 300000},
  "A3": {"v": "2026-08"},
  "B3": {"v": "Ops"},
  "C3": {"v": "Revenue"},
  "D3": {"v": "BUD"},
  "E3": {"v": 200000},
  "A4": {"v": "2026-08"},
  "B4": {"v": "Ops"},
  "C4": {"v": "COGS"},
  "D4": {"v": "BUD"},
  "E4": {"v": 200000},
  "A5": {"v": "2026-08"},
  "B5": {"v": "Sales"},
  "C5": {"v": "Salaries"},
  "D5": {"v": "BUD"},
  "E5": {"v": 70000},
  "A6": {"v": "2026-08"},
  "B6": {"v": "Ops"},
  "C6": {"v": "Salaries"},
  "D6": {"v": "BUD"},
  "E6": {"v": 50000},
  "A7": {"v": "2026-08"},
  "B7": {"v": "Sales"},
  "C7": {"v": "Marketing"},
  "D7": {"v": "BUD"},
  "E7": {"v": 40000},
  "A8": {"v": "2026-08"},
  "B8": {"v": "Ops"},
  "C8": {"v": "Rent"},
  "D8": {"v": "BUD"},
  "E8": {"v": 25000},
  "A9": {"v": "2026-08"},
  "B9": {"v": "Sales"},
  "C9": {"v": "Travel"},
  "D9": {"v": "BUD"},
  "E9": {"v": 5000},
  "A10": {"v": "2026-08"},
  "B10": {"v": "Ops"},
  "C10": {"v": "Travel"},
  "D10": {"v": "BUD"},
  "E10": {"v": 3000},
  "A11": {"v": "2026-08"},
  "B11": {"v": "Sales"},
  "C11": {"v": "Revenue"},
  "D11": {"v": "ACT"},
  "E11": {"v": 270000},
  "A12": {"v": "2026-08"},
  "B12": {"v": "Ops"},
  "C12": {"v": "Revenue"},
  "D12": {"v": "ACT"},
  "E12": {"v": 198000},
  "A13": {"v": "2026-08"},
  "B13": {"v": "Ops"},
  "C13": {"v": "COGS"},
  "D13": {"v": "ACT"},
  "E13": {"v": 191000},
  "A14": {"v": "2026-08"},
  "B14": {"v": "Sales"},
  "C14": {"v": "Salaries"},
  "D14": {"v": "ACT"},
  "E14": {"v": 72000},
  "A15": {"v": "2026-08"},
  "B15": {"v": "Ops"},
  "C15": {"v": "Salaries"},
  "D15": {"v": "ACT"},
  "E15": {"v": 52500},
  "A16": {"v": "2026-08"},
  "B16": {"v": "Sales"},
  "C16": {"v": "Marketing"},
  "D16": {"v": "ACT"},
  "E16": {"v": 52000},
  "A17": {"v": "2026-08"},
  "B17": {"v": "Ops"},
  "C17": {"v": "Rent"},
  "D17": {"v": "ACT"},
  "E17": {"v": 25000},
  "A18": {"v": "2026-08"},
  "B18": {"v": "Sales"},
  "C18": {"v": "Travel"},
  "D18": {"v": "ACT"},
  "E18": {"v": 3600},
  "A19": {"v": "2026-08"},
  "B19": {"v": "Ops"},
  "C19": {"v": "Travel"},
  "D19": {"v": "ACT"},
  "E19": {"v": 1500},
  "A20": {"v": "2026-07"},
  "B20": {"v": "Sales"},
  "C20": {"v": "Revenue"},
  "D20": {"v": "ACT"},
  "E20": {"v": 262000},
  "A21": {"v": "2026-07"},
  "B21": {"v": "Ops"},
  "C21": {"v": "Revenue"},
  "D21": {"v": "ACT"},
  "E21": {"v": 193000},
  "A22": {"v": "2026-07"},
  "B22": {"v": "Ops"},
  "C22": {"v": "COGS"},
  "D22": {"v": "ACT"},
  "E22": {"v": 186000},
  "A23": {"v": "2026-07"},
  "B23": {"v": "Sales"},
  "C23": {"v": "Salaries"},
  "D23": {"v": "ACT"},
  "E23": {"v": 70000},
  "A24": {"v": "2026-07"},
  "B24": {"v": "Ops"},
  "C24": {"v": "Salaries"},
  "D24": {"v": "ACT"},
  "E24": {"v": 51000},
  "A25": {"v": "2026-07"},
  "B25": {"v": "Sales"},
  "C25": {"v": "Marketing"},
  "D25": {"v": "ACT"},
  "E25": {"v": 39000},
  "A26": {"v": "2026-07"},
  "B26": {"v": "Ops"},
  "C26": {"v": "Rent"},
  "D26": {"v": "ACT"},
  "E26": {"v": 25000},
  "A27": {"v": "2026-07"},
  "B27": {"v": "Sales"},
  "C27": {"v": "Travel"},
  "D27": {"v": "ACT"},
  "E27": {"v": 4200},
  "A28": {"v": "2026-07"},
  "B28": {"v": "Ops"},
  "C28": {"v": "Travel"},
  "D28": {"v": "ACT"},
  "E28": {"v": 2600},
  "A29": {"v": "2025-08"},
  "B29": {"v": "Sales"},
  "C29": {"v": "Revenue"},
  "D29": {"v": "ACT"},
  "E29": {"v": 250000},
  "A30": {"v": "2025-08"},
  "B30": {"v": "Ops"},
  "C30": {"v": "Revenue"},
  "D30": {"v": "ACT"},
  "E30": {"v": 180000},
  "A31": {"v": "2025-08"},
  "B31": {"v": "Ops"},
  "C31": {"v": "COGS"},
  "D31": {"v": "ACT"},
  "E31": {"v": 176000},
  "A32": {"v": "2025-08"},
  "B32": {"v": "Sales"},
  "C32": {"v": "Salaries"},
  "D32": {"v": "ACT"},
  "E32": {"v": 64000},
  "A33": {"v": "2025-08"},
  "B33": {"v": "Ops"},
  "C33": {"v": "Salaries"},
  "D33": {"v": "ACT"},
  "E33": {"v": 46000},
  "A34": {"v": "2025-08"},
  "B34": {"v": "Sales"},
  "C34": {"v": "Marketing"},
  "D34": {"v": "ACT"},
  "E34": {"v": 35000},
  "A35": {"v": "2025-08"},
  "B35": {"v": "Ops"},
  "C35": {"v": "Rent"},
  "D35": {"v": "ACT"},
  "E35": {"v": 23000},
  "A36": {"v": "2025-08"},
  "B36": {"v": "Sales"},
  "C36": {"v": "Travel"},
  "D36": {"v": "ACT"},
  "E36": {"v": 4000},
  "A37": {"v": "2025-08"},
  "B37": {"v": "Ops"},
  "C37": {"v": "Travel"},
  "D37": {"v": "ACT"},
  "E37": {"v": 3000}
}
```

#### Sheet `BvA` (72 cells)

```json
{
  "A1": {"v": "Report period"},
  "B1": {"v": "2026-08"},
  "A2": {"v": "Pct threshold"},
  "B2": {"v": 0.05, "fmt": "0.0%"},
  "A3": {"v": "Abs threshold"},
  "B3": {"v": 5000, "fmt": "#,##0"},
  "A5": {"v": "Account"},
  "B5": {"v": "Type"},
  "C5": {"v": "Budget"},
  "D5": {"v": "Actual"},
  "E5": {"v": "Variance"},
  "F5": {"v": "Var %"},
  "G5": {"v": "Fav/Unfav"},
  "H5": {"v": "Needs comment"},
  "A6": {"v": "Revenue"},
  "B6": {"f": "=INDEX(Map!$B$2:$B$7,MATCH($A6,Map!$A$2:$A$7,0))"},
  "C6": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A6,GL!$D$2:$D$37,\"BUD\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "D6": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A6,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "A7": {"v": "COGS"},
  "B7": {"f": "=INDEX(Map!$B$2:$B$7,MATCH($A7,Map!$A$2:$A$7,0))"},
  "C7": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A7,GL!$D$2:$D$37,\"BUD\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "D7": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A7,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "A8": {"v": "Salaries"},
  "B8": {"f": "=INDEX(Map!$B$2:$B$7,MATCH($A8,Map!$A$2:$A$7,0))"},
  "C8": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A8,GL!$D$2:$D$37,\"BUD\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "D8": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A8,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "A9": {"v": "Marketing"},
  "B9": {"f": "=INDEX(Map!$B$2:$B$7,MATCH($A9,Map!$A$2:$A$7,0))"},
  "C9": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A9,GL!$D$2:$D$37,\"BUD\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "D9": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A9,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "A10": {"v": "Rent"},
  "B10": {"f": "=INDEX(Map!$B$2:$B$7,MATCH($A10,Map!$A$2:$A$7,0))"},
  "C10": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A10,GL!$D$2:$D$37,\"BUD\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "D10": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A10,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "A11": {"v": "Travel"},
  "B11": {"f": "=INDEX(Map!$B$2:$B$7,MATCH($A11,Map!$A$2:$A$7,0))"},
  "C11": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A11,GL!$D$2:$D$37,\"BUD\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "D11": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A11,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "A12": {"v": "Operating income"},
  "B12": {"v": "Revenue"},
  "C12": {"f": "=C6-SUM(C7:C11)", "fmt": "#,##0"},
  "D12": {"f": "=D6-SUM(D7:D11)", "fmt": "#,##0"},
  "E6": {"f": "=D6-C6", "fmt": "#,##0;(#,##0)"},
  "F6": {"f": "=IF(C6=0,\"n/a\",E6/ABS(C6))", "fmt": "0.0%"},
  "G6": {"f": "=IF(D6=C6,\"On budget\",IF((B6=\"Revenue\")=(D6>C6),\"Favorable\",\"Unfavorable\"))"},
  "H6": {"f": "=IF(ISNUMBER(F6),AND(ABS(E6)>=$B$3,ABS(F6)>=$B$2),FALSE)"},
  "E7": {"f": "=D7-C7", "fmt": "#,##0;(#,##0)"},
  "F7": {"f": "=IF(C7=0,\"n/a\",E7/ABS(C7))", "fmt": "0.0%"},
  "G7": {"f": "=IF(D7=C7,\"On budget\",IF((B7=\"Revenue\")=(D7>C7),\"Favorable\",\"Unfavorable\"))"},
  "H7": {"f": "=IF(ISNUMBER(F7),AND(ABS(E7)>=$B$3,ABS(F7)>=$B$2),FALSE)"},
  "E8": {"f": "=D8-C8", "fmt": "#,##0;(#,##0)"},
  "F8": {"f": "=IF(C8=0,\"n/a\",E8/ABS(C8))", "fmt": "0.0%"},
  "G8": {"f": "=IF(D8=C8,\"On budget\",IF((B8=\"Revenue\")=(D8>C8),\"Favorable\",\"Unfavorable\"))"},
  "H8": {"f": "=IF(ISNUMBER(F8),AND(ABS(E8)>=$B$3,ABS(F8)>=$B$2),FALSE)"},
  "E9": {"f": "=D9-C9", "fmt": "#,##0;(#,##0)"},
  "F9": {"f": "=IF(C9=0,\"n/a\",E9/ABS(C9))", "fmt": "0.0%"},
  "G9": {"f": "=IF(D9=C9,\"On budget\",IF((B9=\"Revenue\")=(D9>C9),\"Favorable\",\"Unfavorable\"))"},
  "H9": {"f": "=IF(ISNUMBER(F9),AND(ABS(E9)>=$B$3,ABS(F9)>=$B$2),FALSE)"},
  "E10": {"f": "=D10-C10", "fmt": "#,##0;(#,##0)"},
  "F10": {"f": "=IF(C10=0,\"n/a\",E10/ABS(C10))", "fmt": "0.0%"},
  "G10": {"f": "=IF(D10=C10,\"On budget\",IF((B10=\"Revenue\")=(D10>C10),\"Favorable\",\"Unfavorable\"))"},
  "H10": {"f": "=IF(ISNUMBER(F10),AND(ABS(E10)>=$B$3,ABS(F10)>=$B$2),FALSE)"},
  "E11": {"f": "=D11-C11", "fmt": "#,##0;(#,##0)"},
  "F11": {"f": "=IF(C11=0,\"n/a\",E11/ABS(C11))", "fmt": "0.0%"},
  "G11": {"f": "=IF(D11=C11,\"On budget\",IF((B11=\"Revenue\")=(D11>C11),\"Favorable\",\"Unfavorable\"))"},
  "H11": {"f": "=IF(ISNUMBER(F11),AND(ABS(E11)>=$B$3,ABS(F11)>=$B$2),FALSE)"},
  "E12": {"f": "=D12-C12", "fmt": "#,##0;(#,##0)"},
  "F12": {"f": "=IF(C12=0,\"n/a\",E12/ABS(C12))", "fmt": "0.0%"},
  "G12": {"f": "=IF(D12=C12,\"On budget\",IF((B12=\"Revenue\")=(D12>C12),\"Favorable\",\"Unfavorable\"))"},
  "H12": {"f": "=IF(ISNUMBER(F12),AND(ABS(E12)>=$B$3,ABS(F12)>=$B$2),FALSE)"},
  "A14": {"v": "Check: mapped actuals minus all GL actuals (should be 0)"},
  "D14": {"f": "=SUM(D6:D11)-SUMIFS(GL!$E$2:$E$37,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"}
}
```

#### Sheet `Flux` (73 cells)

```json
{
  "A1": {"v": "Current period"},
  "B1": {"v": "2026-08"},
  "A2": {"v": "Prior month"},
  "B2": {"v": "2026-07"},
  "A3": {"v": "Prior year"},
  "B3": {"v": "2025-08"},
  "A4": {"v": "Pct threshold"},
  "B4": {"v": 0.1, "fmt": "0.0%"},
  "A5": {"v": "Abs threshold"},
  "B5": {"v": 5000, "fmt": "#,##0"},
  "A7": {"v": "Account"},
  "B7": {"v": "Current"},
  "C7": {"v": "Prior month"},
  "D7": {"v": "MoM $"},
  "E7": {"v": "MoM %"},
  "F7": {"v": "Prior year"},
  "G7": {"v": "YoY $"},
  "H7": {"v": "YoY %"},
  "I7": {"v": "Flag"},
  "A8": {"v": "Revenue"},
  "B8": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A8,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "C8": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A8,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$2)", "fmt": "#,##0"},
  "F8": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A8,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$3)", "fmt": "#,##0"},
  "D8": {"f": "=B8-C8", "fmt": "#,##0;(#,##0)"},
  "E8": {"f": "=IF(C8=0,\"n/a\",D8/ABS(C8))", "fmt": "0.0%"},
  "G8": {"f": "=B8-F8", "fmt": "#,##0;(#,##0)"},
  "H8": {"f": "=IF(F8=0,\"n/a\",G8/ABS(F8))", "fmt": "0.0%"},
  "I8": {"f": "=IF(OR(AND(ABS(D8)>=$B$5,ABS(N(E8))>=$B$4),AND(ABS(G8)>=$B$5,ABS(N(H8))>=$B$4)),\"Explain\",\"\")"},
  "A9": {"v": "COGS"},
  "B9": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A9,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "C9": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A9,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$2)", "fmt": "#,##0"},
  "F9": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A9,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$3)", "fmt": "#,##0"},
  "D9": {"f": "=B9-C9", "fmt": "#,##0;(#,##0)"},
  "E9": {"f": "=IF(C9=0,\"n/a\",D9/ABS(C9))", "fmt": "0.0%"},
  "G9": {"f": "=B9-F9", "fmt": "#,##0;(#,##0)"},
  "H9": {"f": "=IF(F9=0,\"n/a\",G9/ABS(F9))", "fmt": "0.0%"},
  "I9": {"f": "=IF(OR(AND(ABS(D9)>=$B$5,ABS(N(E9))>=$B$4),AND(ABS(G9)>=$B$5,ABS(N(H9))>=$B$4)),\"Explain\",\"\")"},
  "A10": {"v": "Salaries"},
  "B10": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A10,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "C10": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A10,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$2)", "fmt": "#,##0"},
  "F10": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A10,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$3)", "fmt": "#,##0"},
  "D10": {"f": "=B10-C10", "fmt": "#,##0;(#,##0)"},
  "E10": {"f": "=IF(C10=0,\"n/a\",D10/ABS(C10))", "fmt": "0.0%"},
  "G10": {"f": "=B10-F10", "fmt": "#,##0;(#,##0)"},
  "H10": {"f": "=IF(F10=0,\"n/a\",G10/ABS(F10))", "fmt": "0.0%"},
  "I10": {"f": "=IF(OR(AND(ABS(D10)>=$B$5,ABS(N(E10))>=$B$4),AND(ABS(G10)>=$B$5,ABS(N(H10))>=$B$4)),\"Explain\",\"\")"},
  "A11": {"v": "Marketing"},
  "B11": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A11,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "C11": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A11,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$2)", "fmt": "#,##0"},
  "F11": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A11,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$3)", "fmt": "#,##0"},
  "D11": {"f": "=B11-C11", "fmt": "#,##0;(#,##0)"},
  "E11": {"f": "=IF(C11=0,\"n/a\",D11/ABS(C11))", "fmt": "0.0%"},
  "G11": {"f": "=B11-F11", "fmt": "#,##0;(#,##0)"},
  "H11": {"f": "=IF(F11=0,\"n/a\",G11/ABS(F11))", "fmt": "0.0%"},
  "I11": {"f": "=IF(OR(AND(ABS(D11)>=$B$5,ABS(N(E11))>=$B$4),AND(ABS(G11)>=$B$5,ABS(N(H11))>=$B$4)),\"Explain\",\"\")"},
  "A12": {"v": "Rent"},
  "B12": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A12,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "C12": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A12,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$2)", "fmt": "#,##0"},
  "F12": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A12,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$3)", "fmt": "#,##0"},
  "D12": {"f": "=B12-C12", "fmt": "#,##0;(#,##0)"},
  "E12": {"f": "=IF(C12=0,\"n/a\",D12/ABS(C12))", "fmt": "0.0%"},
  "G12": {"f": "=B12-F12", "fmt": "#,##0;(#,##0)"},
  "H12": {"f": "=IF(F12=0,\"n/a\",G12/ABS(F12))", "fmt": "0.0%"},
  "I12": {"f": "=IF(OR(AND(ABS(D12)>=$B$5,ABS(N(E12))>=$B$4),AND(ABS(G12)>=$B$5,ABS(N(H12))>=$B$4)),\"Explain\",\"\")"},
  "A13": {"v": "Travel"},
  "B13": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A13,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$1)", "fmt": "#,##0"},
  "C13": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A13,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$2)", "fmt": "#,##0"},
  "F13": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$C$2:$C$37,$A13,GL!$D$2:$D$37,\"ACT\",GL!$A$2:$A$37,$B$3)", "fmt": "#,##0"},
  "D13": {"f": "=B13-C13", "fmt": "#,##0;(#,##0)"},
  "E13": {"f": "=IF(C13=0,\"n/a\",D13/ABS(C13))", "fmt": "0.0%"},
  "G13": {"f": "=B13-F13", "fmt": "#,##0;(#,##0)"},
  "H13": {"f": "=IF(F13=0,\"n/a\",G13/ABS(F13))", "fmt": "0.0%"},
  "I13": {"f": "=IF(OR(AND(ABS(D13)>=$B$5,ABS(N(E13))>=$B$4),AND(ABS(G13)>=$B$5,ABS(N(H13))>=$B$4)),\"Explain\",\"\")"}
}
```

#### Sheet `PVM` (82 cells)

```json
{
  "A4": {"v": "Product"},
  "B4": {"v": "Bud units"},
  "C4": {"v": "Bud price"},
  "D4": {"v": "Act units"},
  "E4": {"v": "Act price"},
  "F4": {"v": "Bud rev"},
  "G4": {"v": "Act rev"},
  "H4": {"v": "Bud mix %"},
  "I4": {"v": "Act units @ bud mix"},
  "J4": {"v": "Price var"},
  "K4": {"v": "Volume var"},
  "L4": {"v": "Mix var"},
  "M4": {"v": "Total var"},
  "N4": {"v": "Check"},
  "A5": {"v": "Basic"},
  "B5": {"v": 1000},
  "C5": {"v": 50},
  "D5": {"v": 900},
  "E5": {"v": 52},
  "F5": {"f": "=B5*C5", "fmt": "#,##0"},
  "G5": {"f": "=D5*E5", "fmt": "#,##0"},
  "H5": {"f": "=B5/$B$8", "fmt": "0.00%"},
  "I5": {"f": "=$D$8*H5", "fmt": "#,##0.00"},
  "J5": {"f": "=(E5-C5)*D5", "fmt": "#,##0.0;(#,##0.0)"},
  "K5": {"f": "=(I5-B5)*C5", "fmt": "#,##0.0;(#,##0.0)"},
  "L5": {"f": "=(D5-I5)*C5", "fmt": "#,##0.0;(#,##0.0)"},
  "M5": {"f": "=G5-F5", "fmt": "#,##0;(#,##0)"},
  "N5": {"f": "=ROUND(J5+K5+L5-M5,6)"},
  "A6": {"v": "Pro"},
  "B6": {"v": 500},
  "C6": {"v": 120},
  "D6": {"v": 600},
  "E6": {"v": 115},
  "F6": {"f": "=B6*C6", "fmt": "#,##0"},
  "G6": {"f": "=D6*E6", "fmt": "#,##0"},
  "H6": {"f": "=B6/$B$8", "fmt": "0.00%"},
  "I6": {"f": "=$D$8*H6", "fmt": "#,##0.00"},
  "J6": {"f": "=(E6-C6)*D6", "fmt": "#,##0.0;(#,##0.0)"},
  "K6": {"f": "=(I6-B6)*C6", "fmt": "#,##0.0;(#,##0.0)"},
  "L6": {"f": "=(D6-I6)*C6", "fmt": "#,##0.0;(#,##0.0)"},
  "M6": {"f": "=G6-F6", "fmt": "#,##0;(#,##0)"},
  "N6": {"f": "=ROUND(J6+K6+L6-M6,6)"},
  "A7": {"v": "Enterprise"},
  "B7": {"v": 100},
  "C7": {"v": 400},
  "D7": {"v": 90},
  "E7": {"v": 420},
  "F7": {"f": "=B7*C7", "fmt": "#,##0"},
  "G7": {"f": "=D7*E7", "fmt": "#,##0"},
  "H7": {"f": "=B7/$B$8", "fmt": "0.00%"},
  "I7": {"f": "=$D$8*H7", "fmt": "#,##0.00"},
  "J7": {"f": "=(E7-C7)*D7", "fmt": "#,##0.0;(#,##0.0)"},
  "K7": {"f": "=(I7-B7)*C7", "fmt": "#,##0.0;(#,##0.0)"},
  "L7": {"f": "=(D7-I7)*C7", "fmt": "#,##0.0;(#,##0.0)"},
  "M7": {"f": "=G7-F7", "fmt": "#,##0;(#,##0)"},
  "N7": {"f": "=ROUND(J7+K7+L7-M7,6)"},
  "A8": {"v": "Total"},
  "B8": {"f": "=SUM(B5:B7)"},
  "D8": {"f": "=SUM(D5:D7)"},
  "F8": {"f": "=SUM(F5:F7)"},
  "G8": {"f": "=SUM(G5:G7)"},
  "H8": {"f": "=SUM(H5:H7)"},
  "I8": {"f": "=SUM(I5:I7)"},
  "J8": {"f": "=SUM(J5:J7)"},
  "K8": {"f": "=SUM(K5:K7)"},
  "L8": {"f": "=SUM(L5:L7)"},
  "M8": {"f": "=SUM(M5:M7)"},
  "N8": {"f": "=SUM(N5:N7)"},
  "C8": {"f": "=F8/B8", "fmt": "0.00"},
  "E8": {"f": "=G8/D8", "fmt": "0.00"},
  "A10": {"v": "Waterfall step"},
  "B10": {"v": "Amount"},
  "A11": {"v": "Budget revenue"},
  "B11": {"f": "=F8", "fmt": "#,##0.0;(#,##0.0)"},
  "A12": {"v": "Price"},
  "B12": {"f": "=J8", "fmt": "#,##0.0;(#,##0.0)"},
  "A13": {"v": "Volume"},
  "B13": {"f": "=K8", "fmt": "#,##0.0;(#,##0.0)"},
  "A14": {"v": "Mix"},
  "B14": {"f": "=L8", "fmt": "#,##0.0;(#,##0.0)"},
  "A15": {"v": "Actual revenue"},
  "B15": {"f": "=G8", "fmt": "#,##0.0;(#,##0.0)"}
}
```

#### Sheet `CostVar` (83 cells)

```json
{
  "A1": {"v": "Part A - Rate / volume (freight)"},
  "A3": {"v": "Budget shipments"},
  "B3": {"v": 2000},
  "A4": {"v": "Budget rate / shipment"},
  "B4": {"v": 12},
  "A5": {"v": "Actual shipments"},
  "B5": {"v": 2300},
  "A6": {"v": "Actual rate / shipment"},
  "B6": {"v": 12.5},
  "A8": {"v": "Budget cost"},
  "B8": {"f": "=B3*B4", "fmt": "#,##0;(#,##0)"},
  "A9": {"v": "Actual cost"},
  "B9": {"f": "=B5*B6", "fmt": "#,##0;(#,##0)"},
  "A10": {"v": "Rate variance"},
  "B10": {"f": "=(B6-B4)*B5", "fmt": "#,##0;(#,##0)"},
  "A11": {"v": "Volume variance"},
  "B11": {"f": "=(B5-B3)*B4", "fmt": "#,##0;(#,##0)"},
  "A12": {"v": "Total variance"},
  "B12": {"f": "=B9-B8", "fmt": "#,##0;(#,##0)"},
  "A13": {"v": "Check (should be 0)"},
  "B13": {"f": "=B10+B11-B12", "fmt": "#,##0;(#,##0)"},
  "C10": {"f": "=IF(B10>0,\"Unfavorable\",IF(B10<0,\"Favorable\",\"None\"))"},
  "C11": {"f": "=IF(B11>0,\"Unfavorable\",IF(B11<0,\"Favorable\",\"None\"))"},
  "C12": {"f": "=IF(B12>0,\"Unfavorable\",IF(B12<0,\"Favorable\",\"None\"))"},
  "A15": {"v": "Part B - Flexible budget"},
  "A17": {"v": "Budget units"},
  "B17": {"v": 10000, "fmt": "#,##0"},
  "A18": {"v": "Budget price"},
  "B18": {"v": 20, "fmt": "#,##0"},
  "A19": {"v": "Budget variable cost / unit"},
  "B19": {"v": 12, "fmt": "#,##0"},
  "A20": {"v": "Budget fixed costs"},
  "B20": {"v": 50000, "fmt": "#,##0"},
  "A21": {"v": "Actual units"},
  "B21": {"v": 11000, "fmt": "#,##0"},
  "A22": {"v": "Actual revenue"},
  "B22": {"v": 214500, "fmt": "#,##0"},
  "A23": {"v": "Actual variable costs"},
  "B23": {"v": 137500, "fmt": "#,##0"},
  "A24": {"v": "Actual fixed costs"},
  "B24": {"v": 52000, "fmt": "#,##0"},
  "A26": {"v": "Line"},
  "B26": {"v": "Static budget"},
  "C26": {"v": "Flexible budget"},
  "D26": {"v": "Actual"},
  "E26": {"v": "Flex-budget variance (Actual - Flex)"},
  "F26": {"v": "Sales-volume variance (Flex - Static)"},
  "A27": {"v": "Units"},
  "B27": {"f": "=$B$17", "fmt": "#,##0"},
  "C27": {"f": "=$B$21", "fmt": "#,##0"},
  "D27": {"f": "=$B$21", "fmt": "#,##0"},
  "E27": {"f": "=D27-C27", "fmt": "#,##0;(#,##0)"},
  "F27": {"f": "=C27-B27", "fmt": "#,##0;(#,##0)"},
  "A28": {"v": "Revenue"},
  "B28": {"f": "=B27*$B$18", "fmt": "#,##0"},
  "C28": {"f": "=C27*$B$18", "fmt": "#,##0"},
  "D28": {"f": "=$B$22", "fmt": "#,##0"},
  "E28": {"f": "=D28-C28", "fmt": "#,##0;(#,##0)"},
  "F28": {"f": "=C28-B28", "fmt": "#,##0;(#,##0)"},
  "A29": {"v": "Variable costs"},
  "B29": {"f": "=B27*$B$19", "fmt": "#,##0"},
  "C29": {"f": "=C27*$B$19", "fmt": "#,##0"},
  "D29": {"f": "=$B$23", "fmt": "#,##0"},
  "E29": {"f": "=D29-C29", "fmt": "#,##0;(#,##0)"},
  "F29": {"f": "=C29-B29", "fmt": "#,##0;(#,##0)"},
  "A30": {"v": "Contribution margin"},
  "B30": {"f": "=B28-B29", "fmt": "#,##0"},
  "C30": {"f": "=C28-C29", "fmt": "#,##0"},
  "D30": {"f": "=D28-D29", "fmt": "#,##0"},
  "E30": {"f": "=D30-C30", "fmt": "#,##0;(#,##0)"},
  "F30": {"f": "=C30-B30", "fmt": "#,##0;(#,##0)"},
  "A31": {"v": "Fixed costs"},
  "B31": {"f": "=$B$20", "fmt": "#,##0"},
  "C31": {"f": "=$B$20", "fmt": "#,##0"},
  "D31": {"f": "=$B$24", "fmt": "#,##0"},
  "E31": {"f": "=D31-C31", "fmt": "#,##0;(#,##0)"},
  "F31": {"f": "=C31-B31", "fmt": "#,##0;(#,##0)"},
  "A32": {"v": "Operating income"},
  "B32": {"f": "=B30-B31", "fmt": "#,##0"},
  "C32": {"f": "=C30-C31", "fmt": "#,##0"},
  "D32": {"f": "=D30-D31", "fmt": "#,##0"},
  "E32": {"f": "=D32-C32", "fmt": "#,##0;(#,##0)"},
  "F32": {"f": "=C32-B32", "fmt": "#,##0;(#,##0)"}
}
```

#### Sheet `PivotLab` (34 cells)

```json
{
  "A2": {"v": "Period"},
  "B2": {"v": "2026-08"},
  "A3": {"v": "Version"},
  "B3": {"v": "ACT"},
  "A5": {"v": "Dept \\ Account"},
  "B5": {"v": "Revenue"},
  "C5": {"v": "COGS"},
  "D5": {"v": "Salaries"},
  "E5": {"v": "Marketing"},
  "F5": {"v": "Rent"},
  "G5": {"v": "Travel"},
  "A6": {"v": "Sales"},
  "B6": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$B$2:$B$37,$A6,GL!$C$2:$C$37,B$5,GL!$A$2:$A$37,$B$2,GL!$D$2:$D$37,$B$3)", "fmt": "#,##0"},
  "C6": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$B$2:$B$37,$A6,GL!$C$2:$C$37,C$5,GL!$A$2:$A$37,$B$2,GL!$D$2:$D$37,$B$3)", "fmt": "#,##0"},
  "D6": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$B$2:$B$37,$A6,GL!$C$2:$C$37,D$5,GL!$A$2:$A$37,$B$2,GL!$D$2:$D$37,$B$3)", "fmt": "#,##0"},
  "E6": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$B$2:$B$37,$A6,GL!$C$2:$C$37,E$5,GL!$A$2:$A$37,$B$2,GL!$D$2:$D$37,$B$3)", "fmt": "#,##0"},
  "F6": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$B$2:$B$37,$A6,GL!$C$2:$C$37,F$5,GL!$A$2:$A$37,$B$2,GL!$D$2:$D$37,$B$3)", "fmt": "#,##0"},
  "G6": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$B$2:$B$37,$A6,GL!$C$2:$C$37,G$5,GL!$A$2:$A$37,$B$2,GL!$D$2:$D$37,$B$3)", "fmt": "#,##0"},
  "A7": {"v": "Ops"},
  "B7": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$B$2:$B$37,$A7,GL!$C$2:$C$37,B$5,GL!$A$2:$A$37,$B$2,GL!$D$2:$D$37,$B$3)", "fmt": "#,##0"},
  "C7": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$B$2:$B$37,$A7,GL!$C$2:$C$37,C$5,GL!$A$2:$A$37,$B$2,GL!$D$2:$D$37,$B$3)", "fmt": "#,##0"},
  "D7": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$B$2:$B$37,$A7,GL!$C$2:$C$37,D$5,GL!$A$2:$A$37,$B$2,GL!$D$2:$D$37,$B$3)", "fmt": "#,##0"},
  "E7": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$B$2:$B$37,$A7,GL!$C$2:$C$37,E$5,GL!$A$2:$A$37,$B$2,GL!$D$2:$D$37,$B$3)", "fmt": "#,##0"},
  "F7": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$B$2:$B$37,$A7,GL!$C$2:$C$37,F$5,GL!$A$2:$A$37,$B$2,GL!$D$2:$D$37,$B$3)", "fmt": "#,##0"},
  "G7": {"f": "=SUMIFS(GL!$E$2:$E$37,GL!$B$2:$B$37,$A7,GL!$C$2:$C$37,G$5,GL!$A$2:$A$37,$B$2,GL!$D$2:$D$37,$B$3)", "fmt": "#,##0"},
  "A8": {"v": "Total"},
  "B8": {"f": "=SUM(B6:B7)", "fmt": "#,##0"},
  "C8": {"f": "=SUM(C6:C7)", "fmt": "#,##0"},
  "D8": {"f": "=SUM(D6:D7)", "fmt": "#,##0"},
  "E8": {"f": "=SUM(E6:E7)", "fmt": "#,##0"},
  "F8": {"f": "=SUM(F6:F7)", "fmt": "#,##0"},
  "G8": {"f": "=SUM(G6:G7)", "fmt": "#,##0"},
  "A10": {"v": "Check: cross-tab total vs BvA actuals (should be 0)"},
  "B10": {"f": "=SUM(B8:G8)-SUM(BvA!D6:D11)", "fmt": "#,##0"}
}
```

#### Sheet `Book` (42 cells)

```json
{
  "A1": {"v": "ID"},
  "B1": {"v": "Date"},
  "C1": {"v": "Amount"},
  "D1": {"v": "Ref"},
  "E1": {"v": "Status"},
  "F1": {"v": "Matched bank ID"},
  "G1": {"v": "Duplicate?"},
  "A2": {"v": "B1"},
  "B2": {"v": {"date": "2026-08-02"}, "fmt": "yyyy-mm-dd"},
  "C2": {"v": 1200.0, "fmt": "#,##0.00"},
  "D2": {"v": "INV1001"},
  "F2": {"f": "=IFERROR(LOOKUP(2,1/((Bank!$D$2:$D$5=$D2)*(ABS(Bank!$C$2:$C$5-$C2)<=BankRec!$B$2)*(ABS(Bank!$B$2:$B$5-$B2)<=BankRec!$B$3)),Bank!$A$2:$A$5),\"\")"},
  "E2": {"f": "=IF(COUNTIFS(Bank!$D$2:$D$5,$D2,Bank!$C$2:$C$5,$C2)>0,\"Exact\",IF(F2<>\"\",\"Tolerance\",\"Unmatched\"))"},
  "G2": {"f": "=IF(COUNTIFS($D$2:$D$6,$D2,$C$2:$C$6,$C2)>1,\"DUPLICATE\",\"\")"},
  "A3": {"v": "B2"},
  "B3": {"v": {"date": "2026-08-05"}, "fmt": "yyyy-mm-dd"},
  "C3": {"v": -350.0, "fmt": "#,##0.00"},
  "D3": {"v": "CHK2201"},
  "F3": {"f": "=IFERROR(LOOKUP(2,1/((Bank!$D$2:$D$5=$D3)*(ABS(Bank!$C$2:$C$5-$C3)<=BankRec!$B$2)*(ABS(Bank!$B$2:$B$5-$B3)<=BankRec!$B$3)),Bank!$A$2:$A$5),\"\")"},
  "E3": {"f": "=IF(COUNTIFS(Bank!$D$2:$D$5,$D3,Bank!$C$2:$C$5,$C3)>0,\"Exact\",IF(F3<>\"\",\"Tolerance\",\"Unmatched\"))"},
  "G3": {"f": "=IF(COUNTIFS($D$2:$D$6,$D3,$C$2:$C$6,$C3)>1,\"DUPLICATE\",\"\")"},
  "A4": {"v": "B3"},
  "B4": {"v": {"date": "2026-08-10"}, "fmt": "yyyy-mm-dd"},
  "C4": {"v": 4999.5, "fmt": "#,##0.00"},
  "D4": {"v": "INV1002"},
  "F4": {"f": "=IFERROR(LOOKUP(2,1/((Bank!$D$2:$D$5=$D4)*(ABS(Bank!$C$2:$C$5-$C4)<=BankRec!$B$2)*(ABS(Bank!$B$2:$B$5-$B4)<=BankRec!$B$3)),Bank!$A$2:$A$5),\"\")"},
  "E4": {"f": "=IF(COUNTIFS(Bank!$D$2:$D$5,$D4,Bank!$C$2:$C$5,$C4)>0,\"Exact\",IF(F4<>\"\",\"Tolerance\",\"Unmatched\"))"},
  "G4": {"f": "=IF(COUNTIFS($D$2:$D$6,$D4,$C$2:$C$6,$C4)>1,\"DUPLICATE\",\"\")"},
  "A5": {"v": "B4"},
  "B5": {"v": {"date": "2026-08-28"}, "fmt": "yyyy-mm-dd"},
  "C5": {"v": -2000.0, "fmt": "#,##0.00"},
  "D5": {"v": "CHK2202"},
  "F5": {"f": "=IFERROR(LOOKUP(2,1/((Bank!$D$2:$D$5=$D5)*(ABS(Bank!$C$2:$C$5-$C5)<=BankRec!$B$2)*(ABS(Bank!$B$2:$B$5-$B5)<=BankRec!$B$3)),Bank!$A$2:$A$5),\"\")"},
  "E5": {"f": "=IF(COUNTIFS(Bank!$D$2:$D$5,$D5,Bank!$C$2:$C$5,$C5)>0,\"Exact\",IF(F5<>\"\",\"Tolerance\",\"Unmatched\"))"},
  "G5": {"f": "=IF(COUNTIFS($D$2:$D$6,$D5,$C$2:$C$6,$C5)>1,\"DUPLICATE\",\"\")"},
  "A6": {"v": "B5"},
  "B6": {"v": {"date": "2026-08-31"}, "fmt": "yyyy-mm-dd"},
  "C6": {"v": 780.0, "fmt": "#,##0.00"},
  "D6": {"v": "DEP0831"},
  "F6": {"f": "=IFERROR(LOOKUP(2,1/((Bank!$D$2:$D$5=$D6)*(ABS(Bank!$C$2:$C$5-$C6)<=BankRec!$B$2)*(ABS(Bank!$B$2:$B$5-$B6)<=BankRec!$B$3)),Bank!$A$2:$A$5),\"\")"},
  "E6": {"f": "=IF(COUNTIFS(Bank!$D$2:$D$5,$D6,Bank!$C$2:$C$5,$C6)>0,\"Exact\",IF(F6<>\"\",\"Tolerance\",\"Unmatched\"))"},
  "G6": {"f": "=IF(COUNTIFS($D$2:$D$6,$D6,$C$2:$C$6,$C6)>1,\"DUPLICATE\",\"\")"}
}
```

#### Sheet `Bank` (30 cells)

```json
{
  "A1": {"v": "ID"},
  "B1": {"v": "Date"},
  "C1": {"v": "Amount"},
  "D1": {"v": "Ref"},
  "E1": {"v": "Status"},
  "F1": {"v": "Matched book ID"},
  "A2": {"v": "K1"},
  "B2": {"v": {"date": "2026-08-03"}, "fmt": "yyyy-mm-dd"},
  "C2": {"v": 1200.0, "fmt": "#,##0.00"},
  "D2": {"v": "INV1001"},
  "F2": {"f": "=IFERROR(LOOKUP(2,1/((Book!$D$2:$D$6=$D2)*(ABS(Book!$C$2:$C$6-$C2)<=BankRec!$B$2)*(ABS(Book!$B$2:$B$6-$B2)<=BankRec!$B$3)),Book!$A$2:$A$6),\"\")"},
  "E2": {"f": "=IF(COUNTIFS(Book!$D$2:$D$6,$D2,Book!$C$2:$C$6,$C2)>0,\"Exact\",IF(F2<>\"\",\"Tolerance\",\"Unmatched\"))"},
  "A3": {"v": "K2"},
  "B3": {"v": {"date": "2026-08-07"}, "fmt": "yyyy-mm-dd"},
  "C3": {"v": -350.0, "fmt": "#,##0.00"},
  "D3": {"v": "CHK2201"},
  "F3": {"f": "=IFERROR(LOOKUP(2,1/((Book!$D$2:$D$6=$D3)*(ABS(Book!$C$2:$C$6-$C3)<=BankRec!$B$2)*(ABS(Book!$B$2:$B$6-$B3)<=BankRec!$B$3)),Book!$A$2:$A$6),\"\")"},
  "E3": {"f": "=IF(COUNTIFS(Book!$D$2:$D$6,$D3,Book!$C$2:$C$6,$C3)>0,\"Exact\",IF(F3<>\"\",\"Tolerance\",\"Unmatched\"))"},
  "A4": {"v": "K3"},
  "B4": {"v": {"date": "2026-08-11"}, "fmt": "yyyy-mm-dd"},
  "C4": {"v": 4995.0, "fmt": "#,##0.00"},
  "D4": {"v": "INV1002"},
  "F4": {"f": "=IFERROR(LOOKUP(2,1/((Book!$D$2:$D$6=$D4)*(ABS(Book!$C$2:$C$6-$C4)<=BankRec!$B$2)*(ABS(Book!$B$2:$B$6-$B4)<=BankRec!$B$3)),Book!$A$2:$A$6),\"\")"},
  "E4": {"f": "=IF(COUNTIFS(Book!$D$2:$D$6,$D4,Book!$C$2:$C$6,$C4)>0,\"Exact\",IF(F4<>\"\",\"Tolerance\",\"Unmatched\"))"},
  "A5": {"v": "K4"},
  "B5": {"v": {"date": "2026-08-15"}, "fmt": "yyyy-mm-dd"},
  "C5": {"v": -25.0, "fmt": "#,##0.00"},
  "D5": {"v": "FEE"},
  "F5": {"f": "=IFERROR(LOOKUP(2,1/((Book!$D$2:$D$6=$D5)*(ABS(Book!$C$2:$C$6-$C5)<=BankRec!$B$2)*(ABS(Book!$B$2:$B$6-$B5)<=BankRec!$B$3)),Book!$A$2:$A$6),\"\")"},
  "E5": {"f": "=IF(COUNTIFS(Book!$D$2:$D$6,$D5,Book!$C$2:$C$6,$C5)>0,\"Exact\",IF(F5<>\"\",\"Tolerance\",\"Unmatched\"))"}
}
```

#### Sheet `BankRec` (33 cells)

```json
{
  "A1": {"v": "Bank reconciliation - August 2026"},
  "A2": {"v": "Amount tolerance"},
  "B2": {"v": 5, "fmt": "#,##0.00"},
  "A3": {"v": "Day tolerance"},
  "B3": {"v": 3},
  "A5": {"v": "Balance per bank statement"},
  "B5": {"v": 10000, "fmt": "#,##0.00"},
  "A6": {"v": "Balance per GL"},
  "B6": {"v": 8809.5, "fmt": "#,##0.00"},
  "A8": {"v": "BANK SIDE"},
  "A9": {"v": "Balance per bank"},
  "B9": {"f": "=B5", "fmt": "#,##0.00"},
  "A10": {"v": "+ Deposits in transit"},
  "B10": {"f": "=SUMIFS(Book!$C$2:$C$6,Book!$E$2:$E$6,\"Unmatched\",Book!$C$2:$C$6,\">0\")", "fmt": "#,##0.00"},
  "A11": {"v": "- Outstanding checks"},
  "B11": {"f": "=-SUMIFS(Book!$C$2:$C$6,Book!$E$2:$E$6,\"Unmatched\",Book!$C$2:$C$6,\"<0\")", "fmt": "#,##0.00"},
  "A12": {"v": "Adjusted bank balance"},
  "B12": {"f": "=B9+B10-B11", "fmt": "#,##0.00"},
  "A13": {"v": "BOOK SIDE"},
  "A14": {"v": "Balance per GL"},
  "B14": {"f": "=B6", "fmt": "#,##0.00"},
  "A15": {"v": "+/- Bank-only items (fees, interest)"},
  "B15": {"f": "=SUMIFS(Bank!$C$2:$C$5,Bank!$E$2:$E$5,\"Unmatched\")", "fmt": "#,##0.00"},
  "A16": {"v": "- Book errors (book minus bank on tolerance matches)"},
  "B16": {"f": "=SUMIFS(Book!$C$2:$C$6,Book!$E$2:$E$6,\"Tolerance\")-SUMIFS(Bank!$C$2:$C$5,Bank!$E$2:$E$5,\"Tolerance\")", "fmt": "#,##0.00"},
  "A17": {"v": "Adjusted book balance"},
  "B17": {"f": "=B14+B15-B16", "fmt": "#,##0.00"},
  "A19": {"v": "Difference"},
  "B19": {"f": "=ROUND(B12-B17,2)", "fmt": "#,##0.00"},
  "A20": {"v": "Status"},
  "B20": {"f": "=IF(B19=0,\"Reconciled\",\"Difference: \"&TEXT(B19,\"#,##0.00\"))", "fmt": "#,##0.00"},
  "A22": {"v": "Journal entry needed (bank fee + short-paid wire)"},
  "B22": {"f": "=-(B15-B16)", "fmt": "#,##0.00"}
}
```

#### Sheet `AR` (44 cells)

```json
{
  "A1": {"v": "Invoice"},
  "B1": {"v": "Customer"},
  "C1": {"v": "Due date"},
  "D1": {"v": "Open amount"},
  "E1": {"v": "Days past due"},
  "F1": {"v": "Bucket"},
  "H1": {"v": "As of"},
  "H2": {"v": {"date": "2026-08-31"}, "fmt": "yyyy-mm-dd"},
  "A2": {"v": "A1"},
  "B2": {"v": "Acme"},
  "C2": {"v": {"date": "2026-05-15"}, "fmt": "yyyy-mm-dd"},
  "D2": {"v": 3000},
  "E2": {"f": "=$H$2-C2", "fmt": "0"},
  "F2": {"f": "=IF(E2<=0,\"Current\",IF(E2<=30,\"1-30\",IF(E2<=60,\"31-60\",IF(E2<=90,\"61-90\",\"90+\"))))"},
  "A3": {"v": "A2"},
  "B3": {"v": "Beta"},
  "C3": {"v": {"date": "2026-07-20"}, "fmt": "yyyy-mm-dd"},
  "D3": {"v": 1500},
  "E3": {"f": "=$H$2-C3", "fmt": "0"},
  "F3": {"f": "=IF(E3<=0,\"Current\",IF(E3<=30,\"1-30\",IF(E3<=60,\"31-60\",IF(E3<=90,\"61-90\",\"90+\"))))"},
  "A4": {"v": "A3"},
  "B4": {"v": "Acme"},
  "C4": {"v": {"date": "2026-08-25"}, "fmt": "yyyy-mm-dd"},
  "D4": {"v": 2200},
  "E4": {"f": "=$H$2-C4", "fmt": "0"},
  "F4": {"f": "=IF(E4<=0,\"Current\",IF(E4<=30,\"1-30\",IF(E4<=60,\"31-60\",IF(E4<=90,\"61-90\",\"90+\"))))"},
  "A5": {"v": "A4"},
  "B5": {"v": "Cyan"},
  "C5": {"v": {"date": "2026-09-10"}, "fmt": "yyyy-mm-dd"},
  "D5": {"v": 900},
  "E5": {"f": "=$H$2-C5", "fmt": "0"},
  "F5": {"f": "=IF(E5<=0,\"Current\",IF(E5<=30,\"1-30\",IF(E5<=60,\"31-60\",IF(E5<=90,\"61-90\",\"90+\"))))"},
  "A6": {"v": "A5"},
  "B6": {"v": "Beta"},
  "C6": {"v": {"date": "2026-06-25"}, "fmt": "yyyy-mm-dd"},
  "D6": {"v": 1200},
  "E6": {"f": "=$H$2-C6", "fmt": "0"},
  "F6": {"f": "=IF(E6<=0,\"Current\",IF(E6<=30,\"1-30\",IF(E6<=60,\"31-60\",IF(E6<=90,\"61-90\",\"90+\"))))"},
  "A7": {"v": "A6"},
  "B7": {"v": "Cyan"},
  "C7": {"v": {"date": "2026-08-10"}, "fmt": "yyyy-mm-dd"},
  "D7": {"v": 650},
  "E7": {"f": "=$H$2-C7", "fmt": "0"},
  "F7": {"f": "=IF(E7<=0,\"Current\",IF(E7<=30,\"1-30\",IF(E7<=60,\"31-60\",IF(E7<=90,\"61-90\",\"90+\"))))"}
}
```

#### Sheet `ARSummary` (67 cells)

```json
{
  "A3": {"v": "Bucket"},
  "B3": {"v": "Amount"},
  "C3": {"v": "% of total"},
  "D3": {"v": "Reserve %"},
  "E3": {"v": "Allowance"},
  "A4": {"v": "Current"},
  "B4": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$F$2:$F$7,$A4)", "fmt": "#,##0"},
  "C4": {"f": "=B4/$B$9", "fmt": "0.0%"},
  "D4": {"v": 0.005, "fmt": "0.0%"},
  "E4": {"f": "=B4*D4", "fmt": "#,##0.00"},
  "A5": {"v": "1-30"},
  "B5": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$F$2:$F$7,$A5)", "fmt": "#,##0"},
  "C5": {"f": "=B5/$B$9", "fmt": "0.0%"},
  "D5": {"v": 0.02, "fmt": "0.0%"},
  "E5": {"f": "=B5*D5", "fmt": "#,##0.00"},
  "A6": {"v": "31-60"},
  "B6": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$F$2:$F$7,$A6)", "fmt": "#,##0"},
  "C6": {"f": "=B6/$B$9", "fmt": "0.0%"},
  "D6": {"v": 0.05, "fmt": "0.0%"},
  "E6": {"f": "=B6*D6", "fmt": "#,##0.00"},
  "A7": {"v": "61-90"},
  "B7": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$F$2:$F$7,$A7)", "fmt": "#,##0"},
  "C7": {"f": "=B7/$B$9", "fmt": "0.0%"},
  "D7": {"v": 0.15, "fmt": "0.0%"},
  "E7": {"f": "=B7*D7", "fmt": "#,##0.00"},
  "A8": {"v": "90+"},
  "B8": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$F$2:$F$7,$A8)", "fmt": "#,##0"},
  "C8": {"f": "=B8/$B$9", "fmt": "0.0%"},
  "D8": {"v": 0.5, "fmt": "0.0%"},
  "E8": {"f": "=B8*D8", "fmt": "#,##0.00"},
  "A9": {"v": "Total"},
  "B9": {"f": "=SUM(B4:B8)", "fmt": "#,##0"},
  "E9": {"f": "=SUMPRODUCT(B4:B8,D4:D8)", "fmt": "#,##0.00"},
  "A11": {"v": "Customer \\ Bucket"},
  "B11": {"v": "Current"},
  "C11": {"v": "1-30"},
  "D11": {"v": "31-60"},
  "E11": {"v": "61-90"},
  "F11": {"v": "90+"},
  "G11": {"v": "Total"},
  "A12": {"v": "Acme"},
  "B12": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A12,AR!$F$2:$F$7,B$11)", "fmt": "#,##0"},
  "C12": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A12,AR!$F$2:$F$7,C$11)", "fmt": "#,##0"},
  "D12": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A12,AR!$F$2:$F$7,D$11)", "fmt": "#,##0"},
  "E12": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A12,AR!$F$2:$F$7,E$11)", "fmt": "#,##0"},
  "F12": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A12,AR!$F$2:$F$7,F$11)", "fmt": "#,##0"},
  "G12": {"f": "=SUM(B12:F12)", "fmt": "#,##0"},
  "A13": {"v": "Beta"},
  "B13": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A13,AR!$F$2:$F$7,B$11)", "fmt": "#,##0"},
  "C13": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A13,AR!$F$2:$F$7,C$11)", "fmt": "#,##0"},
  "D13": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A13,AR!$F$2:$F$7,D$11)", "fmt": "#,##0"},
  "E13": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A13,AR!$F$2:$F$7,E$11)", "fmt": "#,##0"},
  "F13": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A13,AR!$F$2:$F$7,F$11)", "fmt": "#,##0"},
  "G13": {"f": "=SUM(B13:F13)", "fmt": "#,##0"},
  "A14": {"v": "Cyan"},
  "B14": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A14,AR!$F$2:$F$7,B$11)", "fmt": "#,##0"},
  "C14": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A14,AR!$F$2:$F$7,C$11)", "fmt": "#,##0"},
  "D14": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A14,AR!$F$2:$F$7,D$11)", "fmt": "#,##0"},
  "E14": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A14,AR!$F$2:$F$7,E$11)", "fmt": "#,##0"},
  "F14": {"f": "=SUMIFS(AR!$D$2:$D$7,AR!$B$2:$B$7,$A14,AR!$F$2:$F$7,F$11)", "fmt": "#,##0"},
  "G14": {"f": "=SUM(B14:F14)", "fmt": "#,##0"},
  "A17": {"v": "GL control account (1200-AR)"},
  "B17": {"v": 9550, "fmt": "#,##0"},
  "A18": {"v": "Subledger total"},
  "B18": {"f": "=SUM(AR!$D$2:$D$7)", "fmt": "#,##0"},
  "A19": {"v": "Difference (investigate)"},
  "B19": {"f": "=B17-B18", "fmt": "#,##0"}
}
```

#### Sheet `Forecast` (439 cells)

```json
{
  "A1": {"v": "Monthly sales history (Jan-2024 to Jun-2026) and forecasts"},
  "A2": {"v": "Train periods (t <=)"},
  "B2": {"v": 24},
  "A4": {"v": "Month"},
  "B4": {"v": "t"},
  "C4": {"v": "Sales"},
  "D4": {"v": "MA3"},
  "E4": {"v": "Month #"},
  "F4": {"v": "Trend (all)"},
  "G4": {"v": "Ratio (all)"},
  "H4": {"v": "Trend (train)"},
  "I4": {"v": "Ratio (train)"},
  "A5": {"v": {"date": "2024-01-01"}, "fmt": "mmm-yy"},
  "B5": {"v": 1},
  "C5": {"v": 90609.0, "fmt": "#,##0"},
  "E5": {"f": "=MONTH(A5)"},
  "F5": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B5", "fmt": "#,##0"},
  "G5": {"f": "=C5/F5", "fmt": "0.0000"},
  "H5": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B5", "fmt": "#,##0"},
  "I5": {"f": "=C5/H5", "fmt": "0.0000"},
  "A6": {"v": {"date": "2024-02-01"}, "fmt": "mmm-yy"},
  "B6": {"v": 2},
  "C6": {"v": 91300.0, "fmt": "#,##0"},
  "E6": {"f": "=MONTH(A6)"},
  "F6": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B6", "fmt": "#,##0"},
  "G6": {"f": "=C6/F6", "fmt": "0.0000"},
  "H6": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B6", "fmt": "#,##0"},
  "I6": {"f": "=C6/H6", "fmt": "0.0000"},
  "A7": {"v": {"date": "2024-03-01"}, "fmt": "mmm-yy"},
  "B7": {"v": 3},
  "C7": {"v": 104501.0, "fmt": "#,##0"},
  "D7": {"f": "=AVERAGE(C5:C7)", "fmt": "#,##0"},
  "E7": {"f": "=MONTH(A7)"},
  "F7": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B7", "fmt": "#,##0"},
  "G7": {"f": "=C7/F7", "fmt": "0.0000"},
  "H7": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B7", "fmt": "#,##0"},
  "I7": {"f": "=C7/H7", "fmt": "0.0000"},
  "A8": {"v": {"date": "2024-04-01"}, "fmt": "mmm-yy"},
  "B8": {"v": 4},
  "C8": {"v": 108471.0, "fmt": "#,##0"},
  "D8": {"f": "=AVERAGE(C6:C8)", "fmt": "#,##0"},
  "E8": {"f": "=MONTH(A8)"},
  "F8": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B8", "fmt": "#,##0"},
  "G8": {"f": "=C8/F8", "fmt": "0.0000"},
  "H8": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B8", "fmt": "#,##0"},
  "I8": {"f": "=C8/H8", "fmt": "0.0000"},
  "A9": {"v": {"date": "2024-05-01"}, "fmt": "mmm-yy"},
  "B9": {"v": 5},
  "C9": {"v": 107398.0, "fmt": "#,##0"},
  "D9": {"f": "=AVERAGE(C7:C9)", "fmt": "#,##0"},
  "E9": {"f": "=MONTH(A9)"},
  "F9": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B9", "fmt": "#,##0"},
  "G9": {"f": "=C9/F9", "fmt": "0.0000"},
  "H9": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B9", "fmt": "#,##0"},
  "I9": {"f": "=C9/H9", "fmt": "0.0000"},
  "A10": {"v": {"date": "2024-06-01"}, "fmt": "mmm-yy"},
  "B10": {"v": 6},
  "C10": {"v": 115646.0, "fmt": "#,##0"},
  "D10": {"f": "=AVERAGE(C8:C10)", "fmt": "#,##0"},
  "E10": {"f": "=MONTH(A10)"},
  "F10": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B10", "fmt": "#,##0"},
  "G10": {"f": "=C10/F10", "fmt": "0.0000"},
  "H10": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B10", "fmt": "#,##0"},
  "I10": {"f": "=C10/H10", "fmt": "0.0000"},
  "A11": {"v": {"date": "2024-07-01"}, "fmt": "mmm-yy"},
  "B11": {"v": 7},
  "C11": {"v": 117976.0, "fmt": "#,##0"},
  "D11": {"f": "=AVERAGE(C9:C11)", "fmt": "#,##0"},
  "E11": {"f": "=MONTH(A11)"},
  "F11": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B11", "fmt": "#,##0"},
  "G11": {"f": "=C11/F11", "fmt": "0.0000"},
  "H11": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B11", "fmt": "#,##0"},
  "I11": {"f": "=C11/H11", "fmt": "0.0000"},
  "A12": {"v": {"date": "2024-08-01"}, "fmt": "mmm-yy"},
  "B12": {"v": 8},
  "C12": {"v": 116498.0, "fmt": "#,##0"},
  "D12": {"f": "=AVERAGE(C10:C12)", "fmt": "#,##0"},
  "E12": {"f": "=MONTH(A12)"},
  "F12": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B12", "fmt": "#,##0"},
  "G12": {"f": "=C12/F12", "fmt": "0.0000"},
  "H12": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B12", "fmt": "#,##0"},
  "I12": {"f": "=C12/H12", "fmt": "0.0000"},
  "A13": {"v": {"date": "2024-09-01"}, "fmt": "mmm-yy"},
  "B13": {"v": 9},
  "C13": {"v": 111966.0, "fmt": "#,##0"},
  "D13": {"f": "=AVERAGE(C11:C13)", "fmt": "#,##0"},
  "E13": {"f": "=MONTH(A13)"},
  "F13": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B13", "fmt": "#,##0"},
  "G13": {"f": "=C13/F13", "fmt": "0.0000"},
  "H13": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B13", "fmt": "#,##0"},
  "I13": {"f": "=C13/H13", "fmt": "0.0000"},
  "A14": {"v": {"date": "2024-10-01"}, "fmt": "mmm-yy"},
  "B14": {"v": 10},
  "C14": {"v": 109524.0, "fmt": "#,##0"},
  "D14": {"f": "=AVERAGE(C12:C14)", "fmt": "#,##0"},
  "E14": {"f": "=MONTH(A14)"},
  "F14": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B14", "fmt": "#,##0"},
  "G14": {"f": "=C14/F14", "fmt": "0.0000"},
  "H14": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B14", "fmt": "#,##0"},
  "I14": {"f": "=C14/H14", "fmt": "0.0000"},
  "A15": {"v": {"date": "2024-11-01"}, "fmt": "mmm-yy"},
  "B15": {"v": 11},
  "C15": {"v": 111009.0, "fmt": "#,##0"},
  "D15": {"f": "=AVERAGE(C13:C15)", "fmt": "#,##0"},
  "E15": {"f": "=MONTH(A15)"},
  "F15": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B15", "fmt": "#,##0"},
  "G15": {"f": "=C15/F15", "fmt": "0.0000"},
  "H15": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B15", "fmt": "#,##0"},
  "I15": {"f": "=C15/H15", "fmt": "0.0000"},
  "A16": {"v": {"date": "2024-12-01"}, "fmt": "mmm-yy"},
  "B16": {"v": 12},
  "C16": {"v": 141356.0, "fmt": "#,##0"},
  "D16": {"f": "=AVERAGE(C14:C16)", "fmt": "#,##0"},
  "E16": {"f": "=MONTH(A16)"},
  "F16": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B16", "fmt": "#,##0"},
  "G16": {"f": "=C16/F16", "fmt": "0.0000"},
  "H16": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B16", "fmt": "#,##0"},
  "I16": {"f": "=C16/H16", "fmt": "0.0000"},
  "A17": {"v": {"date": "2025-01-01"}, "fmt": "mmm-yy"},
  "B17": {"v": 13},
  "C17": {"v": 106332.0, "fmt": "#,##0"},
  "D17": {"f": "=AVERAGE(C15:C17)", "fmt": "#,##0"},
  "E17": {"f": "=MONTH(A17)"},
  "F17": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B17", "fmt": "#,##0"},
  "G17": {"f": "=C17/F17", "fmt": "0.0000"},
  "H17": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B17", "fmt": "#,##0"},
  "I17": {"f": "=C17/H17", "fmt": "0.0000"},
  "A18": {"v": {"date": "2025-02-01"}, "fmt": "mmm-yy"},
  "B18": {"v": 14},
  "C18": {"v": 112194.0, "fmt": "#,##0"},
  "D18": {"f": "=AVERAGE(C16:C18)", "fmt": "#,##0"},
  "E18": {"f": "=MONTH(A18)"},
  "F18": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B18", "fmt": "#,##0"},
  "G18": {"f": "=C18/F18", "fmt": "0.0000"},
  "H18": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B18", "fmt": "#,##0"},
  "I18": {"f": "=C18/H18", "fmt": "0.0000"},
  "A19": {"v": {"date": "2025-03-01"}, "fmt": "mmm-yy"},
  "B19": {"v": 15},
  "C19": {"v": 121935.0, "fmt": "#,##0"},
  "D19": {"f": "=AVERAGE(C17:C19)", "fmt": "#,##0"},
  "E19": {"f": "=MONTH(A19)"},
  "F19": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B19", "fmt": "#,##0"},
  "G19": {"f": "=C19/F19", "fmt": "0.0000"},
  "H19": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B19", "fmt": "#,##0"},
  "I19": {"f": "=C19/H19", "fmt": "0.0000"},
  "A20": {"v": {"date": "2025-04-01"}, "fmt": "mmm-yy"},
  "B20": {"v": 16},
  "C20": {"v": 123231.0, "fmt": "#,##0"},
  "D20": {"f": "=AVERAGE(C18:C20)", "fmt": "#,##0"},
  "E20": {"f": "=MONTH(A20)"},
  "F20": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B20", "fmt": "#,##0"},
  "G20": {"f": "=C20/F20", "fmt": "0.0000"},
  "H20": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B20", "fmt": "#,##0"},
  "I20": {"f": "=C20/H20", "fmt": "0.0000"},
  "A21": {"v": {"date": "2025-05-01"}, "fmt": "mmm-yy"},
  "B21": {"v": 17},
  "C21": {"v": 130938.0, "fmt": "#,##0"},
  "D21": {"f": "=AVERAGE(C19:C21)", "fmt": "#,##0"},
  "E21": {"f": "=MONTH(A21)"},
  "F21": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B21", "fmt": "#,##0"},
  "G21": {"f": "=C21/F21", "fmt": "0.0000"},
  "H21": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B21", "fmt": "#,##0"},
  "I21": {"f": "=C21/H21", "fmt": "0.0000"},
  "A22": {"v": {"date": "2025-06-01"}, "fmt": "mmm-yy"},
  "B22": {"v": 18},
  "C22": {"v": 136132.0, "fmt": "#,##0"},
  "D22": {"f": "=AVERAGE(C20:C22)", "fmt": "#,##0"},
  "E22": {"f": "=MONTH(A22)"},
  "F22": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B22", "fmt": "#,##0"},
  "G22": {"f": "=C22/F22", "fmt": "0.0000"},
  "H22": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B22", "fmt": "#,##0"},
  "I22": {"f": "=C22/H22", "fmt": "0.0000"},
  "A23": {"v": {"date": "2025-07-01"}, "fmt": "mmm-yy"},
  "B23": {"v": 19},
  "C23": {"v": 138917.0, "fmt": "#,##0"},
  "D23": {"f": "=AVERAGE(C21:C23)", "fmt": "#,##0"},
  "E23": {"f": "=MONTH(A23)"},
  "F23": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B23", "fmt": "#,##0"},
  "G23": {"f": "=C23/F23", "fmt": "0.0000"},
  "H23": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B23", "fmt": "#,##0"},
  "I23": {"f": "=C23/H23", "fmt": "0.0000"},
  "A24": {"v": {"date": "2025-08-01"}, "fmt": "mmm-yy"},
  "B24": {"v": 20},
  "C24": {"v": 136110.0, "fmt": "#,##0"},
  "D24": {"f": "=AVERAGE(C22:C24)", "fmt": "#,##0"},
  "E24": {"f": "=MONTH(A24)"},
  "F24": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B24", "fmt": "#,##0"},
  "G24": {"f": "=C24/F24", "fmt": "0.0000"},
  "H24": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B24", "fmt": "#,##0"},
  "I24": {"f": "=C24/H24", "fmt": "0.0000"},
  "A25": {"v": {"date": "2025-09-01"}, "fmt": "mmm-yy"},
  "B25": {"v": 21},
  "C25": {"v": 129630.0, "fmt": "#,##0"},
  "D25": {"f": "=AVERAGE(C23:C25)", "fmt": "#,##0"},
  "E25": {"f": "=MONTH(A25)"},
  "F25": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B25", "fmt": "#,##0"},
  "G25": {"f": "=C25/F25", "fmt": "0.0000"},
  "H25": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B25", "fmt": "#,##0"},
  "I25": {"f": "=C25/H25", "fmt": "0.0000"},
  "A26": {"v": {"date": "2025-10-01"}, "fmt": "mmm-yy"},
  "B26": {"v": 22},
  "C26": {"v": 127508.0, "fmt": "#,##0"},
  "D26": {"f": "=AVERAGE(C24:C26)", "fmt": "#,##0"},
  "E26": {"f": "=MONTH(A26)"},
  "F26": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B26", "fmt": "#,##0"},
  "G26": {"f": "=C26/F26", "fmt": "0.0000"},
  "H26": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B26", "fmt": "#,##0"},
  "I26": {"f": "=C26/H26", "fmt": "0.0000"},
  "A27": {"v": {"date": "2025-11-01"}, "fmt": "mmm-yy"},
  "B27": {"v": 23},
  "C27": {"v": 128795.0, "fmt": "#,##0"},
  "D27": {"f": "=AVERAGE(C25:C27)", "fmt": "#,##0"},
  "E27": {"f": "=MONTH(A27)"},
  "F27": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B27", "fmt": "#,##0"},
  "G27": {"f": "=C27/F27", "fmt": "0.0000"},
  "H27": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B27", "fmt": "#,##0"},
  "I27": {"f": "=C27/H27", "fmt": "0.0000"},
  "A28": {"v": {"date": "2025-12-01"}, "fmt": "mmm-yy"},
  "B28": {"v": 24},
  "C28": {"v": 161091.0, "fmt": "#,##0"},
  "D28": {"f": "=AVERAGE(C26:C28)", "fmt": "#,##0"},
  "E28": {"f": "=MONTH(A28)"},
  "F28": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B28", "fmt": "#,##0"},
  "G28": {"f": "=C28/F28", "fmt": "0.0000"},
  "H28": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B28", "fmt": "#,##0"},
  "I28": {"f": "=C28/H28", "fmt": "0.0000"},
  "A29": {"v": {"date": "2026-01-01"}, "fmt": "mmm-yy"},
  "B29": {"v": 25},
  "C29": {"v": 121543.0, "fmt": "#,##0"},
  "D29": {"f": "=AVERAGE(C27:C29)", "fmt": "#,##0"},
  "E29": {"f": "=MONTH(A29)"},
  "F29": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B29", "fmt": "#,##0"},
  "G29": {"f": "=C29/F29", "fmt": "0.0000"},
  "H29": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B29", "fmt": "#,##0"},
  "I29": {"f": "=C29/H29", "fmt": "0.0000"},
  "A30": {"v": {"date": "2026-02-01"}, "fmt": "mmm-yy"},
  "B30": {"v": 26},
  "C30": {"v": 125796.0, "fmt": "#,##0"},
  "D30": {"f": "=AVERAGE(C28:C30)", "fmt": "#,##0"},
  "E30": {"f": "=MONTH(A30)"},
  "F30": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B30", "fmt": "#,##0"},
  "G30": {"f": "=C30/F30", "fmt": "0.0000"},
  "H30": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B30", "fmt": "#,##0"},
  "I30": {"f": "=C30/H30", "fmt": "0.0000"},
  "A31": {"v": {"date": "2026-03-01"}, "fmt": "mmm-yy"},
  "B31": {"v": 27},
  "C31": {"v": 140065.0, "fmt": "#,##0"},
  "D31": {"f": "=AVERAGE(C29:C31)", "fmt": "#,##0"},
  "E31": {"f": "=MONTH(A31)"},
  "F31": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B31", "fmt": "#,##0"},
  "G31": {"f": "=C31/F31", "fmt": "0.0000"},
  "H31": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B31", "fmt": "#,##0"},
  "I31": {"f": "=C31/H31", "fmt": "0.0000"},
  "A32": {"v": {"date": "2026-04-01"}, "fmt": "mmm-yy"},
  "B32": {"v": 28},
  "C32": {"v": 144041.0, "fmt": "#,##0"},
  "D32": {"f": "=AVERAGE(C30:C32)", "fmt": "#,##0"},
  "E32": {"f": "=MONTH(A32)"},
  "F32": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B32", "fmt": "#,##0"},
  "G32": {"f": "=C32/F32", "fmt": "0.0000"},
  "H32": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B32", "fmt": "#,##0"},
  "I32": {"f": "=C32/H32", "fmt": "0.0000"},
  "A33": {"v": {"date": "2026-05-01"}, "fmt": "mmm-yy"},
  "B33": {"v": 29},
  "C33": {"v": 149925.0, "fmt": "#,##0"},
  "D33": {"f": "=AVERAGE(C31:C33)", "fmt": "#,##0"},
  "E33": {"f": "=MONTH(A33)"},
  "F33": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B33", "fmt": "#,##0"},
  "G33": {"f": "=C33/F33", "fmt": "0.0000"},
  "H33": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B33", "fmt": "#,##0"},
  "I33": {"f": "=C33/H33", "fmt": "0.0000"},
  "A34": {"v": {"date": "2026-06-01"}, "fmt": "mmm-yy"},
  "B34": {"v": 30},
  "C34": {"v": 158712.0, "fmt": "#,##0"},
  "D34": {"f": "=AVERAGE(C32:C34)", "fmt": "#,##0"},
  "E34": {"f": "=MONTH(A34)"},
  "F34": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)+SLOPE($C$5:$C$34,$B$5:$B$34)*B34", "fmt": "#,##0"},
  "G34": {"f": "=C34/F34", "fmt": "0.0000"},
  "H34": {"f": "=INTERCEPT($C$5:$C$28,$B$5:$B$28)+SLOPE($C$5:$C$28,$B$5:$B$28)*B34", "fmt": "#,##0"},
  "I34": {"f": "=C34/H34", "fmt": "0.0000"},
  "K4": {"v": "Month #"},
  "L4": {"v": "Raw index (all)"},
  "M4": {"v": "Seasonal index (all)"},
  "N4": {"v": "Raw index (train)"},
  "O4": {"v": "Seasonal index (train)"},
  "K5": {"v": 1},
  "L5": {"f": "=AVERAGEIFS($G$5:$G$34,$E$5:$E$34,K5)", "fmt": "0.0000"},
  "M5": {"f": "=L5/AVERAGE($L$5:$L$16)", "fmt": "0.0000"},
  "N5": {"f": "=AVERAGEIFS($I$5:$I$34,$E$5:$E$34,K5,$B$5:$B$34,\"<=\"&$B$2)", "fmt": "0.0000"},
  "O5": {"f": "=N5/AVERAGE($N$5:$N$16)", "fmt": "0.0000"},
  "K6": {"v": 2},
  "L6": {"f": "=AVERAGEIFS($G$5:$G$34,$E$5:$E$34,K6)", "fmt": "0.0000"},
  "M6": {"f": "=L6/AVERAGE($L$5:$L$16)", "fmt": "0.0000"},
  "N6": {"f": "=AVERAGEIFS($I$5:$I$34,$E$5:$E$34,K6,$B$5:$B$34,\"<=\"&$B$2)", "fmt": "0.0000"},
  "O6": {"f": "=N6/AVERAGE($N$5:$N$16)", "fmt": "0.0000"},
  "K7": {"v": 3},
  "L7": {"f": "=AVERAGEIFS($G$5:$G$34,$E$5:$E$34,K7)", "fmt": "0.0000"},
  "M7": {"f": "=L7/AVERAGE($L$5:$L$16)", "fmt": "0.0000"},
  "N7": {"f": "=AVERAGEIFS($I$5:$I$34,$E$5:$E$34,K7,$B$5:$B$34,\"<=\"&$B$2)", "fmt": "0.0000"},
  "O7": {"f": "=N7/AVERAGE($N$5:$N$16)", "fmt": "0.0000"},
  "K8": {"v": 4},
  "L8": {"f": "=AVERAGEIFS($G$5:$G$34,$E$5:$E$34,K8)", "fmt": "0.0000"},
  "M8": {"f": "=L8/AVERAGE($L$5:$L$16)", "fmt": "0.0000"},
  "N8": {"f": "=AVERAGEIFS($I$5:$I$34,$E$5:$E$34,K8,$B$5:$B$34,\"<=\"&$B$2)", "fmt": "0.0000"},
  "O8": {"f": "=N8/AVERAGE($N$5:$N$16)", "fmt": "0.0000"},
  "K9": {"v": 5},
  "L9": {"f": "=AVERAGEIFS($G$5:$G$34,$E$5:$E$34,K9)", "fmt": "0.0000"},
  "M9": {"f": "=L9/AVERAGE($L$5:$L$16)", "fmt": "0.0000"},
  "N9": {"f": "=AVERAGEIFS($I$5:$I$34,$E$5:$E$34,K9,$B$5:$B$34,\"<=\"&$B$2)", "fmt": "0.0000"},
  "O9": {"f": "=N9/AVERAGE($N$5:$N$16)", "fmt": "0.0000"},
  "K10": {"v": 6},
  "L10": {"f": "=AVERAGEIFS($G$5:$G$34,$E$5:$E$34,K10)", "fmt": "0.0000"},
  "M10": {"f": "=L10/AVERAGE($L$5:$L$16)", "fmt": "0.0000"},
  "N10": {"f": "=AVERAGEIFS($I$5:$I$34,$E$5:$E$34,K10,$B$5:$B$34,\"<=\"&$B$2)", "fmt": "0.0000"},
  "O10": {"f": "=N10/AVERAGE($N$5:$N$16)", "fmt": "0.0000"},
  "K11": {"v": 7},
  "L11": {"f": "=AVERAGEIFS($G$5:$G$34,$E$5:$E$34,K11)", "fmt": "0.0000"},
  "M11": {"f": "=L11/AVERAGE($L$5:$L$16)", "fmt": "0.0000"},
  "N11": {"f": "=AVERAGEIFS($I$5:$I$34,$E$5:$E$34,K11,$B$5:$B$34,\"<=\"&$B$2)", "fmt": "0.0000"},
  "O11": {"f": "=N11/AVERAGE($N$5:$N$16)", "fmt": "0.0000"},
  "K12": {"v": 8},
  "L12": {"f": "=AVERAGEIFS($G$5:$G$34,$E$5:$E$34,K12)", "fmt": "0.0000"},
  "M12": {"f": "=L12/AVERAGE($L$5:$L$16)", "fmt": "0.0000"},
  "N12": {"f": "=AVERAGEIFS($I$5:$I$34,$E$5:$E$34,K12,$B$5:$B$34,\"<=\"&$B$2)", "fmt": "0.0000"},
  "O12": {"f": "=N12/AVERAGE($N$5:$N$16)", "fmt": "0.0000"},
  "K13": {"v": 9},
  "L13": {"f": "=AVERAGEIFS($G$5:$G$34,$E$5:$E$34,K13)", "fmt": "0.0000"},
  "M13": {"f": "=L13/AVERAGE($L$5:$L$16)", "fmt": "0.0000"},
  "N13": {"f": "=AVERAGEIFS($I$5:$I$34,$E$5:$E$34,K13,$B$5:$B$34,\"<=\"&$B$2)", "fmt": "0.0000"},
  "O13": {"f": "=N13/AVERAGE($N$5:$N$16)", "fmt": "0.0000"},
  "K14": {"v": 10},
  "L14": {"f": "=AVERAGEIFS($G$5:$G$34,$E$5:$E$34,K14)", "fmt": "0.0000"},
  "M14": {"f": "=L14/AVERAGE($L$5:$L$16)", "fmt": "0.0000"},
  "N14": {"f": "=AVERAGEIFS($I$5:$I$34,$E$5:$E$34,K14,$B$5:$B$34,\"<=\"&$B$2)", "fmt": "0.0000"},
  "O14": {"f": "=N14/AVERAGE($N$5:$N$16)", "fmt": "0.0000"},
  "K15": {"v": 11},
  "L15": {"f": "=AVERAGEIFS($G$5:$G$34,$E$5:$E$34,K15)", "fmt": "0.0000"},
  "M15": {"f": "=L15/AVERAGE($L$5:$L$16)", "fmt": "0.0000"},
  "N15": {"f": "=AVERAGEIFS($I$5:$I$34,$E$5:$E$34,K15,$B$5:$B$34,\"<=\"&$B$2)", "fmt": "0.0000"},
  "O15": {"f": "=N15/AVERAGE($N$5:$N$16)", "fmt": "0.0000"},
  "K16": {"v": 12},
  "L16": {"f": "=AVERAGEIFS($G$5:$G$34,$E$5:$E$34,K16)", "fmt": "0.0000"},
  "M16": {"f": "=L16/AVERAGE($L$5:$L$16)", "fmt": "0.0000"},
  "N16": {"f": "=AVERAGEIFS($I$5:$I$34,$E$5:$E$34,K16,$B$5:$B$34,\"<=\"&$B$2)", "fmt": "0.0000"},
  "O16": {"f": "=N16/AVERAGE($N$5:$N$16)", "fmt": "0.0000"},
  "K17": {"v": "Average"},
  "M17": {"f": "=AVERAGE(M5:M16)", "fmt": "0.0000"},
  "O17": {"f": "=AVERAGE(O5:O16)", "fmt": "0.0000"},
  "K19": {"v": "Slope / month"},
  "L19": {"f": "=SLOPE($C$5:$C$34,$B$5:$B$34)", "fmt": "#,##0.0"},
  "K20": {"v": "Intercept"},
  "L20": {"f": "=INTERCEPT($C$5:$C$34,$B$5:$B$34)", "fmt": "#,##0.0"},
  "K21": {"v": "R-squared"},
  "L21": {"f": "=RSQ($C$5:$C$34,$B$5:$B$34)", "fmt": "0.000"},
  "A37": {"v": "Month"},
  "B37": {"v": "t"},
  "C37": {"v": "Linear (FORECAST.LINEAR)"},
  "D37": {"v": "Linear (TREND)"},
  "E37": {"v": "Trend x season"},
  "A38": {"f": "=EDATE($A$34,1)", "fmt": "mmm-yy"},
  "B38": {"f": "=$B$34+1"},
  "C38": {"f": "=FORECAST.LINEAR(B38,$C$5:$C$34,$B$5:$B$34)", "fmt": "#,##0"},
  "D38": {"f": "=TREND($C$5:$C$34,$B$5:$B$34,B38)", "fmt": "#,##0"},
  "E38": {"f": "=C38*INDEX($M$5:$M$16,MONTH(A38))", "fmt": "#,##0"},
  "A39": {"f": "=EDATE(A38,1)", "fmt": "mmm-yy"},
  "B39": {"f": "=B38+1"},
  "C39": {"f": "=FORECAST.LINEAR(B39,$C$5:$C$34,$B$5:$B$34)", "fmt": "#,##0"},
  "D39": {"f": "=TREND($C$5:$C$34,$B$5:$B$34,B39)", "fmt": "#,##0"},
  "E39": {"f": "=C39*INDEX($M$5:$M$16,MONTH(A39))", "fmt": "#,##0"},
  "A40": {"f": "=EDATE(A39,1)", "fmt": "mmm-yy"},
  "B40": {"f": "=B39+1"},
  "C40": {"f": "=FORECAST.LINEAR(B40,$C$5:$C$34,$B$5:$B$34)", "fmt": "#,##0"},
  "D40": {"f": "=TREND($C$5:$C$34,$B$5:$B$34,B40)", "fmt": "#,##0"},
  "E40": {"f": "=C40*INDEX($M$5:$M$16,MONTH(A40))", "fmt": "#,##0"},
  "A41": {"f": "=EDATE(A40,1)", "fmt": "mmm-yy"},
  "B41": {"f": "=B40+1"},
  "C41": {"f": "=FORECAST.LINEAR(B41,$C$5:$C$34,$B$5:$B$34)", "fmt": "#,##0"},
  "D41": {"f": "=TREND($C$5:$C$34,$B$5:$B$34,B41)", "fmt": "#,##0"},
  "E41": {"f": "=C41*INDEX($M$5:$M$16,MONTH(A41))", "fmt": "#,##0"},
  "A42": {"f": "=EDATE(A41,1)", "fmt": "mmm-yy"},
  "B42": {"f": "=B41+1"},
  "C42": {"f": "=FORECAST.LINEAR(B42,$C$5:$C$34,$B$5:$B$34)", "fmt": "#,##0"},
  "D42": {"f": "=TREND($C$5:$C$34,$B$5:$B$34,B42)", "fmt": "#,##0"},
  "E42": {"f": "=C42*INDEX($M$5:$M$16,MONTH(A42))", "fmt": "#,##0"},
  "A43": {"f": "=EDATE(A42,1)", "fmt": "mmm-yy"},
  "B43": {"f": "=B42+1"},
  "C43": {"f": "=FORECAST.LINEAR(B43,$C$5:$C$34,$B$5:$B$34)", "fmt": "#,##0"},
  "D43": {"f": "=TREND($C$5:$C$34,$B$5:$B$34,B43)", "fmt": "#,##0"},
  "E43": {"f": "=C43*INDEX($M$5:$M$16,MONTH(A43))", "fmt": "#,##0"},
  "A46": {"v": "Month (test)"},
  "B46": {"v": "Actual"},
  "C46": {"v": "Model (train trend x train index)"},
  "D46": {"v": "Seasonal naive (t-12)"},
  "E46": {"v": "APE model"},
  "F46": {"v": "APE naive"},
  "A47": {"f": "=A29", "fmt": "mmm-yy"},
  "B47": {"f": "=C29", "fmt": "#,##0"},
  "C47": {"f": "=H29*INDEX($O$5:$O$16,E29)", "fmt": "#,##0"},
  "D47": {"f": "=C17", "fmt": "#,##0"},
  "E47": {"f": "=ABS(B47-C47)/ABS(B47)", "fmt": "0.00%"},
  "F47": {"f": "=ABS(B47-D47)/ABS(B47)", "fmt": "0.00%"},
  "A48": {"f": "=A30", "fmt": "mmm-yy"},
  "B48": {"f": "=C30", "fmt": "#,##0"},
  "C48": {"f": "=H30*INDEX($O$5:$O$16,E30)", "fmt": "#,##0"},
  "D48": {"f": "=C18", "fmt": "#,##0"},
  "E48": {"f": "=ABS(B48-C48)/ABS(B48)", "fmt": "0.00%"},
  "F48": {"f": "=ABS(B48-D48)/ABS(B48)", "fmt": "0.00%"},
  "A49": {"f": "=A31", "fmt": "mmm-yy"},
  "B49": {"f": "=C31", "fmt": "#,##0"},
  "C49": {"f": "=H31*INDEX($O$5:$O$16,E31)", "fmt": "#,##0"},
  "D49": {"f": "=C19", "fmt": "#,##0"},
  "E49": {"f": "=ABS(B49-C49)/ABS(B49)", "fmt": "0.00%"},
  "F49": {"f": "=ABS(B49-D49)/ABS(B49)", "fmt": "0.00%"},
  "A50": {"f": "=A32", "fmt": "mmm-yy"},
  "B50": {"f": "=C32", "fmt": "#,##0"},
  "C50": {"f": "=H32*INDEX($O$5:$O$16,E32)", "fmt": "#,##0"},
  "D50": {"f": "=C20", "fmt": "#,##0"},
  "E50": {"f": "=ABS(B50-C50)/ABS(B50)", "fmt": "0.00%"},
  "F50": {"f": "=ABS(B50-D50)/ABS(B50)", "fmt": "0.00%"},
  "A51": {"f": "=A33", "fmt": "mmm-yy"},
  "B51": {"f": "=C33", "fmt": "#,##0"},
  "C51": {"f": "=H33*INDEX($O$5:$O$16,E33)", "fmt": "#,##0"},
  "D51": {"f": "=C21", "fmt": "#,##0"},
  "E51": {"f": "=ABS(B51-C51)/ABS(B51)", "fmt": "0.00%"},
  "F51": {"f": "=ABS(B51-D51)/ABS(B51)", "fmt": "0.00%"},
  "A52": {"f": "=A34", "fmt": "mmm-yy"},
  "B52": {"f": "=C34", "fmt": "#,##0"},
  "C52": {"f": "=H34*INDEX($O$5:$O$16,E34)", "fmt": "#,##0"},
  "D52": {"f": "=C22", "fmt": "#,##0"},
  "E52": {"f": "=ABS(B52-C52)/ABS(B52)", "fmt": "0.00%"},
  "F52": {"f": "=ABS(B52-D52)/ABS(B52)", "fmt": "0.00%"},
  "A54": {"v": "MAPE (model)"},
  "B54": {"f": "=AVERAGE(E47:E52)", "fmt": "0.00%"},
  "A55": {"v": "Bias (model)"},
  "B55": {"f": "=(SUM(C47:C52)-SUM(B47:B52))/SUM(B47:B52)", "fmt": "+0.00%;-0.00%"},
  "A56": {"v": "MAPE (seasonal naive)"},
  "B56": {"f": "=AVERAGE(F47:F52)", "fmt": "0.00%"},
  "A57": {"v": "Model beats naive?"},
  "B57": {"f": "=B54<B56"}
}
```

#### Sheet `Drivers` (118 cells)

```json
{
  "A1": {"v": "Driver-based subscription revenue forecast"},
  "A3": {"v": "Driver"},
  "B3": {"v": "Downside"},
  "C3": {"v": "Base"},
  "D3": {"v": "Upside"},
  "E3": {"v": "Active"},
  "A4": {"v": "New customers / month"},
  "B4": {"v": 60, "fmt": "#,##0"},
  "C4": {"v": 80, "fmt": "#,##0"},
  "D4": {"v": 100, "fmt": "#,##0"},
  "E4": {"f": "=INDEX(B4:D4,$C$8)", "fmt": "#,##0"},
  "A5": {"v": "Monthly churn"},
  "B5": {"v": 0.03, "fmt": "0.0%"},
  "C5": {"v": 0.02, "fmt": "0.0%"},
  "D5": {"v": 0.015, "fmt": "0.0%"},
  "E5": {"f": "=INDEX(B5:D5,$C$8)", "fmt": "0.0%"},
  "A6": {"v": "ARPU ($ / month)"},
  "B6": {"v": 90, "fmt": "#,##0.00"},
  "C6": {"v": 95, "fmt": "#,##0.00"},
  "D6": {"v": 98, "fmt": "#,##0.00"},
  "E6": {"f": "=INDEX(B6:D6,$C$8)", "fmt": "#,##0.00"},
  "A8": {"v": "Scenario # (1=Down, 2=Base, 3=Up)"},
  "C8": {"v": 2},
  "A9": {"v": "Active scenario name"},
  "C9": {"f": "=INDEX($B$3:$D$3,$C$8)"},
  "A10": {"v": "Starting customers"},
  "B10": {"v": 1200, "fmt": "#,##0"},
  "A13": {"v": "Month"},
  "B13": {"v": {"date": "2026-10-01"}, "fmt": "mmm-yy"},
  "B14": {"f": "=$B$10", "fmt": "#,##0"},
  "B15": {"f": "=-B14*$E$5", "fmt": "#,##0;(#,##0)"},
  "B16": {"f": "=$E$4", "fmt": "#,##0"},
  "B17": {"f": "=B14+B15+B16", "fmt": "#,##0"},
  "B18": {"f": "=B17*$E$6", "fmt": "#,##0"},
  "C13": {"f": "=EDATE(B13,1)", "fmt": "mmm-yy"},
  "C14": {"f": "=B17", "fmt": "#,##0"},
  "C15": {"f": "=-C14*$E$5", "fmt": "#,##0;(#,##0)"},
  "C16": {"f": "=$E$4", "fmt": "#,##0"},
  "C17": {"f": "=C14+C15+C16", "fmt": "#,##0"},
  "C18": {"f": "=C17*$E$6", "fmt": "#,##0"},
  "D13": {"f": "=EDATE(C13,1)", "fmt": "mmm-yy"},
  "D14": {"f": "=C17", "fmt": "#,##0"},
  "D15": {"f": "=-D14*$E$5", "fmt": "#,##0;(#,##0)"},
  "D16": {"f": "=$E$4", "fmt": "#,##0"},
  "D17": {"f": "=D14+D15+D16", "fmt": "#,##0"},
  "D18": {"f": "=D17*$E$6", "fmt": "#,##0"},
  "E13": {"f": "=EDATE(D13,1)", "fmt": "mmm-yy"},
  "E14": {"f": "=D17", "fmt": "#,##0"},
  "E15": {"f": "=-E14*$E$5", "fmt": "#,##0;(#,##0)"},
  "E16": {"f": "=$E$4", "fmt": "#,##0"},
  "E17": {"f": "=E14+E15+E16", "fmt": "#,##0"},
  "E18": {"f": "=E17*$E$6", "fmt": "#,##0"},
  "F13": {"f": "=EDATE(E13,1)", "fmt": "mmm-yy"},
  "F14": {"f": "=E17", "fmt": "#,##0"},
  "F15": {"f": "=-F14*$E$5", "fmt": "#,##0;(#,##0)"},
  "F16": {"f": "=$E$4", "fmt": "#,##0"},
  "F17": {"f": "=F14+F15+F16", "fmt": "#,##0"},
  "F18": {"f": "=F17*$E$6", "fmt": "#,##0"},
  "G13": {"f": "=EDATE(F13,1)", "fmt": "mmm-yy"},
  "G14": {"f": "=F17", "fmt": "#,##0"},
  "G15": {"f": "=-G14*$E$5", "fmt": "#,##0;(#,##0)"},
  "G16": {"f": "=$E$4", "fmt": "#,##0"},
  "G17": {"f": "=G14+G15+G16", "fmt": "#,##0"},
  "G18": {"f": "=G17*$E$6", "fmt": "#,##0"},
  "A14": {"v": "Opening customers"},
  "A15": {"v": "Churned"},
  "A16": {"v": "New"},
  "A17": {"v": "Closing customers"},
  "A18": {"v": "Revenue"},
  "H13": {"v": "6-mo total"},
  "H18": {"f": "=SUM(B18:G18)", "fmt": "#,##0"},
  "A21": {"v": "6-mo revenue by scenario (closed form)"},
  "B21": {"f": "=B6*SUMPRODUCT($B$10*(1-B5)^$B$25:$G$25+B4*(1-(1-B5)^$B$25:$G$25)/B5)", "fmt": "#,##0"},
  "C21": {"f": "=C6*SUMPRODUCT($B$10*(1-C5)^$B$25:$G$25+C4*(1-(1-C5)^$B$25:$G$25)/C5)", "fmt": "#,##0"},
  "D21": {"f": "=D6*SUMPRODUCT($B$10*(1-D5)^$B$25:$G$25+D4*(1-(1-D5)^$B$25:$G$25)/D5)", "fmt": "#,##0"},
  "A25": {"v": "Helper: month index k"},
  "B25": {"v": 1},
  "C25": {"v": 2},
  "D25": {"v": 3},
  "E25": {"v": 4},
  "F25": {"v": 5},
  "G25": {"v": 6},
  "A28": {"v": "Churn \\ ARPU"},
  "B28": {"v": 85},
  "C28": {"v": 90},
  "D28": {"v": 95},
  "E28": {"v": 100},
  "F28": {"v": 105},
  "A29": {"v": 0.01, "fmt": "0.0%"},
  "B29": {"f": "=B$28*SUMPRODUCT($B$10*(1-$A29)^$B$25:$G$25+$E$4*(1-(1-$A29)^$B$25:$G$25)/$A29)", "fmt": "#,##0"},
  "C29": {"f": "=C$28*SUMPRODUCT($B$10*(1-$A29)^$B$25:$G$25+$E$4*(1-(1-$A29)^$B$25:$G$25)/$A29)", "fmt": "#,##0"},
  "D29": {"f": "=D$28*SUMPRODUCT($B$10*(1-$A29)^$B$25:$G$25+$E$4*(1-(1-$A29)^$B$25:$G$25)/$A29)", "fmt": "#,##0"},
  "E29": {"f": "=E$28*SUMPRODUCT($B$10*(1-$A29)^$B$25:$G$25+$E$4*(1-(1-$A29)^$B$25:$G$25)/$A29)", "fmt": "#,##0"},
  "F29": {"f": "=F$28*SUMPRODUCT($B$10*(1-$A29)^$B$25:$G$25+$E$4*(1-(1-$A29)^$B$25:$G$25)/$A29)", "fmt": "#,##0"},
  "A30": {"v": 0.015, "fmt": "0.0%"},
  "B30": {"f": "=B$28*SUMPRODUCT($B$10*(1-$A30)^$B$25:$G$25+$E$4*(1-(1-$A30)^$B$25:$G$25)/$A30)", "fmt": "#,##0"},
  "C30": {"f": "=C$28*SUMPRODUCT($B$10*(1-$A30)^$B$25:$G$25+$E$4*(1-(1-$A30)^$B$25:$G$25)/$A30)", "fmt": "#,##0"},
  "D30": {"f": "=D$28*SUMPRODUCT($B$10*(1-$A30)^$B$25:$G$25+$E$4*(1-(1-$A30)^$B$25:$G$25)/$A30)", "fmt": "#,##0"},
  "E30": {"f": "=E$28*SUMPRODUCT($B$10*(1-$A30)^$B$25:$G$25+$E$4*(1-(1-$A30)^$B$25:$G$25)/$A30)", "fmt": "#,##0"},
  "F30": {"f": "=F$28*SUMPRODUCT($B$10*(1-$A30)^$B$25:$G$25+$E$4*(1-(1-$A30)^$B$25:$G$25)/$A30)", "fmt": "#,##0"},
  "A31": {"v": 0.02, "fmt": "0.0%"},
  "B31": {"f": "=B$28*SUMPRODUCT($B$10*(1-$A31)^$B$25:$G$25+$E$4*(1-(1-$A31)^$B$25:$G$25)/$A31)", "fmt": "#,##0"},
  "C31": {"f": "=C$28*SUMPRODUCT($B$10*(1-$A31)^$B$25:$G$25+$E$4*(1-(1-$A31)^$B$25:$G$25)/$A31)", "fmt": "#,##0"},
  "D31": {"f": "=D$28*SUMPRODUCT($B$10*(1-$A31)^$B$25:$G$25+$E$4*(1-(1-$A31)^$B$25:$G$25)/$A31)", "fmt": "#,##0"},
  "E31": {"f": "=E$28*SUMPRODUCT($B$10*(1-$A31)^$B$25:$G$25+$E$4*(1-(1-$A31)^$B$25:$G$25)/$A31)", "fmt": "#,##0"},
  "F31": {"f": "=F$28*SUMPRODUCT($B$10*(1-$A31)^$B$25:$G$25+$E$4*(1-(1-$A31)^$B$25:$G$25)/$A31)", "fmt": "#,##0"},
  "A32": {"v": 0.025, "fmt": "0.0%"},
  "B32": {"f": "=B$28*SUMPRODUCT($B$10*(1-$A32)^$B$25:$G$25+$E$4*(1-(1-$A32)^$B$25:$G$25)/$A32)", "fmt": "#,##0"},
  "C32": {"f": "=C$28*SUMPRODUCT($B$10*(1-$A32)^$B$25:$G$25+$E$4*(1-(1-$A32)^$B$25:$G$25)/$A32)", "fmt": "#,##0"},
  "D32": {"f": "=D$28*SUMPRODUCT($B$10*(1-$A32)^$B$25:$G$25+$E$4*(1-(1-$A32)^$B$25:$G$25)/$A32)", "fmt": "#,##0"},
  "E32": {"f": "=E$28*SUMPRODUCT($B$10*(1-$A32)^$B$25:$G$25+$E$4*(1-(1-$A32)^$B$25:$G$25)/$A32)", "fmt": "#,##0"},
  "F32": {"f": "=F$28*SUMPRODUCT($B$10*(1-$A32)^$B$25:$G$25+$E$4*(1-(1-$A32)^$B$25:$G$25)/$A32)", "fmt": "#,##0"},
  "A33": {"v": 0.03, "fmt": "0.0%"},
  "B33": {"f": "=B$28*SUMPRODUCT($B$10*(1-$A33)^$B$25:$G$25+$E$4*(1-(1-$A33)^$B$25:$G$25)/$A33)", "fmt": "#,##0"},
  "C33": {"f": "=C$28*SUMPRODUCT($B$10*(1-$A33)^$B$25:$G$25+$E$4*(1-(1-$A33)^$B$25:$G$25)/$A33)", "fmt": "#,##0"},
  "D33": {"f": "=D$28*SUMPRODUCT($B$10*(1-$A33)^$B$25:$G$25+$E$4*(1-(1-$A33)^$B$25:$G$25)/$A33)", "fmt": "#,##0"},
  "E33": {"f": "=E$28*SUMPRODUCT($B$10*(1-$A33)^$B$25:$G$25+$E$4*(1-(1-$A33)^$B$25:$G$25)/$A33)", "fmt": "#,##0"},
  "F33": {"f": "=F$28*SUMPRODUCT($B$10*(1-$A33)^$B$25:$G$25+$E$4*(1-(1-$A33)^$B$25:$G$25)/$A33)", "fmt": "#,##0"}
}
```

#### Sheet `Checks` (28 cells)

```json
{
  "A2": {"v": "Check"},
  "B2": {"v": "Pass?"},
  "C2": {"v": "Value"},
  "A3": {"v": "BvA: mapped actuals tie to GL"},
  "B3": {"f": "=BvA!D14=0"},
  "C3": {"f": "=BvA!D14"},
  "A4": {"v": "PivotLab: cross-tab ties to BvA"},
  "B4": {"f": "=PivotLab!B10=0"},
  "C4": {"f": "=PivotLab!B10"},
  "A5": {"v": "PVM: price+volume+mix = total"},
  "B5": {"f": "=ROUND(PVM!N8,6)=0"},
  "C5": {"f": "=PVM!N8"},
  "A6": {"v": "CostVar: rate+volume = total"},
  "B6": {"f": "=CostVar!B13=0"},
  "C6": {"f": "=CostVar!B13"},
  "A7": {"v": "BankRec: adjusted balances agree"},
  "B7": {"f": "=BankRec!B19=0"},
  "C7": {"f": "=BankRec!B19"},
  "A8": {"v": "Forecast: seasonal index averages 1"},
  "B8": {"f": "=ROUND(Forecast!M17,9)=1"},
  "C8": {"f": "=Forecast!M17"},
  "A9": {"v": "Drivers: schedule = closed form"},
  "B9": {"f": "=ROUND(Drivers!H18-INDEX(Drivers!B21:D21,Drivers!C8),2)=0"},
  "C9": {"f": "=Drivers!H18-INDEX(Drivers!B21:D21,Drivers!C8)"},
  "A10": {"v": "ALL CONTROL CHECKS PASS"},
  "B10": {"f": "=COUNTIF(B3:B9,FALSE)=0"},
  "A12": {"v": "Open item (informational): AR subledger vs GL difference"},
  "C12": {"f": "=ARSummary!B19"}
}
```


---

## Appendix D — Assertions (Expected Values for Auto-Grading)

Each expected value was calculated independently in Python and matched the spreadsheet engine's recalculation of Appendix C.

```json
[
  {"sheet": "BvA", "cell": "C6", "expected": 500000, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "D6", "expected": 468000, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "E6", "expected": -32000, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "C7", "expected": 200000, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "D7", "expected": 191000, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "E7", "expected": -9000, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "C8", "expected": 120000, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "D8", "expected": 124500, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "E8", "expected": 4500, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "C9", "expected": 40000, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "D9", "expected": 52000, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "E9", "expected": 12000, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "C10", "expected": 25000, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "D10", "expected": 25000, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "E10", "expected": 0, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "C11", "expected": 8000, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "D11", "expected": 5100, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "E11", "expected": -2900, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "F6", "expected": -0.064, "tolerance": 1e-09},
  {"sheet": "BvA", "cell": "F9", "expected": 0.3, "tolerance": 1e-09},
  {"sheet": "BvA", "cell": "F11", "expected": -0.3625, "tolerance": 1e-09},
  {"sheet": "BvA", "cell": "G6", "expected": "Unfavorable", "tolerance": 0.005},
  {"sheet": "BvA", "cell": "G7", "expected": "Favorable", "tolerance": 0.005},
  {"sheet": "BvA", "cell": "G10", "expected": "On budget", "tolerance": 0.005},
  {"sheet": "BvA", "cell": "G11", "expected": "Favorable", "tolerance": 0.005},
  {"sheet": "BvA", "cell": "H6", "expected": true, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "H7", "expected": false, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "H8", "expected": false, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "H9", "expected": true, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "H11", "expected": false, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "C12", "expected": 107000, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "D12", "expected": 70400, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "E12", "expected": -36600, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "G12", "expected": "Unfavorable", "tolerance": 0.005},
  {"sheet": "BvA", "cell": "H12", "expected": true, "tolerance": 0.005},
  {"sheet": "BvA", "cell": "D14", "expected": 0, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "B8", "expected": 468000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "C8", "expected": 455000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "F8", "expected": 430000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "D8", "expected": 13000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "G8", "expected": 38000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "I8", "expected": "", "tolerance": 0.005},
  {"sheet": "Flux", "cell": "B9", "expected": 191000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "C9", "expected": 186000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "F9", "expected": 176000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "D9", "expected": 5000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "G9", "expected": 15000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "I9", "expected": "", "tolerance": 0.005},
  {"sheet": "Flux", "cell": "B10", "expected": 124500, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "C10", "expected": 121000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "F10", "expected": 110000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "D10", "expected": 3500, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "G10", "expected": 14500, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "I10", "expected": "Explain", "tolerance": 0.005},
  {"sheet": "Flux", "cell": "B11", "expected": 52000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "C11", "expected": 39000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "F11", "expected": 35000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "D11", "expected": 13000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "G11", "expected": 17000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "I11", "expected": "Explain", "tolerance": 0.005},
  {"sheet": "Flux", "cell": "B12", "expected": 25000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "C12", "expected": 25000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "F12", "expected": 23000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "D12", "expected": 0, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "G12", "expected": 2000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "I12", "expected": "", "tolerance": 0.005},
  {"sheet": "Flux", "cell": "B13", "expected": 5100, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "C13", "expected": 6800, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "F13", "expected": 7000, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "D13", "expected": -1700, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "G13", "expected": -1900, "tolerance": 0.005},
  {"sheet": "Flux", "cell": "I13", "expected": "", "tolerance": 0.005},
  {"sheet": "Flux", "cell": "E11", "expected": 0.3333333333333333, "tolerance": 1e-09},
  {"sheet": "Flux", "cell": "H10", "expected": 0.1318181818181818, "tolerance": 1e-09},
  {"sheet": "PVM", "cell": "J5", "expected": 1800, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "K5", "expected": -312.5, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "L5", "expected": -4687.5, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "M5", "expected": -3200, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "N5", "expected": 0, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "J6", "expected": -3000, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "K6", "expected": -375, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "L6", "expected": 12375, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "M6", "expected": 9000, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "N6", "expected": 0, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "J7", "expected": 1800, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "K7", "expected": -250, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "L7", "expected": -3750, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "M7", "expected": -2200, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "N7", "expected": 0, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "F8", "expected": 150000, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "G8", "expected": 153600, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "J8", "expected": 600, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "K8", "expected": -937.5, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "L8", "expected": 3937.5, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "M8", "expected": 3600, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "N8", "expected": 0, "tolerance": 0.005},
  {"sheet": "PVM", "cell": "I5", "expected": 993.75, "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "B8", "expected": 24000, "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "B9", "expected": 28750, "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "B10", "expected": 1150, "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "B11", "expected": 3600, "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "B12", "expected": 4750, "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "B13", "expected": 0, "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "C10", "expected": "Unfavorable", "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "B32", "expected": 30000, "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "C32", "expected": 38000, "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "D32", "expected": 25000, "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "E28", "expected": -5500, "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "E29", "expected": 5500, "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "E31", "expected": 2000, "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "E32", "expected": -13000, "tolerance": 0.005},
  {"sheet": "CostVar", "cell": "F32", "expected": 8000, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "B6", "expected": 270000, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "C6", "expected": 0, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "D6", "expected": 72000, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "E6", "expected": 52000, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "F6", "expected": 0, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "G6", "expected": 3600, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "B7", "expected": 198000, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "C7", "expected": 191000, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "D7", "expected": 52500, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "E7", "expected": 0, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "F7", "expected": 25000, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "G7", "expected": 1500, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "B8", "expected": 468000, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "G8", "expected": 5100, "tolerance": 0.005},
  {"sheet": "PivotLab", "cell": "B10", "expected": 0, "tolerance": 0.005},
  {"sheet": "Book", "cell": "E2", "expected": "Exact", "tolerance": 0.005},
  {"sheet": "Book", "cell": "E3", "expected": "Exact", "tolerance": 0.005},
  {"sheet": "Book", "cell": "E4", "expected": "Tolerance", "tolerance": 0.005},
  {"sheet": "Book", "cell": "E5", "expected": "Unmatched", "tolerance": 0.005},
  {"sheet": "Book", "cell": "E6", "expected": "Unmatched", "tolerance": 0.005},
  {"sheet": "Book", "cell": "F2", "expected": "K1", "tolerance": 0.005},
  {"sheet": "Book", "cell": "F3", "expected": "K2", "tolerance": 0.005},
  {"sheet": "Book", "cell": "F4", "expected": "K3", "tolerance": 0.005},
  {"sheet": "Book", "cell": "F5", "expected": "", "tolerance": 0.005},
  {"sheet": "Book", "cell": "F6", "expected": "", "tolerance": 0.005},
  {"sheet": "Book", "cell": "G2", "expected": "", "tolerance": 0.005},
  {"sheet": "Bank", "cell": "E2", "expected": "Exact", "tolerance": 0.005},
  {"sheet": "Bank", "cell": "E4", "expected": "Tolerance", "tolerance": 0.005},
  {"sheet": "Bank", "cell": "E5", "expected": "Unmatched", "tolerance": 0.005},
  {"sheet": "Bank", "cell": "F4", "expected": "B3", "tolerance": 0.005},
  {"sheet": "BankRec", "cell": "B10", "expected": 780, "tolerance": 0.005},
  {"sheet": "BankRec", "cell": "B11", "expected": 2000, "tolerance": 0.005},
  {"sheet": "BankRec", "cell": "B12", "expected": 8780, "tolerance": 0.005},
  {"sheet": "BankRec", "cell": "B15", "expected": -25, "tolerance": 0.005},
  {"sheet": "BankRec", "cell": "B16", "expected": 4.5, "tolerance": 0.005},
  {"sheet": "BankRec", "cell": "B17", "expected": 8780, "tolerance": 0.005},
  {"sheet": "BankRec", "cell": "B19", "expected": 0, "tolerance": 0.005},
  {"sheet": "BankRec", "cell": "B20", "expected": "Reconciled", "tolerance": 0.005},
  {"sheet": "BankRec", "cell": "B22", "expected": 29.5, "tolerance": 0.005},
  {"sheet": "AR", "cell": "E2", "expected": 108, "tolerance": 0.005},
  {"sheet": "AR", "cell": "F2", "expected": "90+", "tolerance": 0.005},
  {"sheet": "AR", "cell": "E3", "expected": 42, "tolerance": 0.005},
  {"sheet": "AR", "cell": "F3", "expected": "31-60", "tolerance": 0.005},
  {"sheet": "AR", "cell": "E4", "expected": 6, "tolerance": 0.005},
  {"sheet": "AR", "cell": "F4", "expected": "1-30", "tolerance": 0.005},
  {"sheet": "AR", "cell": "E5", "expected": -10, "tolerance": 0.005},
  {"sheet": "AR", "cell": "F5", "expected": "Current", "tolerance": 0.005},
  {"sheet": "AR", "cell": "E6", "expected": 67, "tolerance": 0.005},
  {"sheet": "AR", "cell": "F6", "expected": "61-90", "tolerance": 0.005},
  {"sheet": "AR", "cell": "E7", "expected": 21, "tolerance": 0.005},
  {"sheet": "AR", "cell": "F7", "expected": "1-30", "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "B4", "expected": 900, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "B5", "expected": 2850, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "B6", "expected": 1500, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "B7", "expected": 1200, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "B8", "expected": 3000, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "B9", "expected": 9450, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "E9", "expected": 1816.5, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "G12", "expected": 5200, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "G13", "expected": 2700, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "G14", "expected": 1550, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "B12", "expected": 0, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "F12", "expected": 3000, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "C12", "expected": 2200, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "D13", "expected": 1500, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "E13", "expected": 1200, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "B14", "expected": 900, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "C14", "expected": 650, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "B18", "expected": 9450, "tolerance": 0.005},
  {"sheet": "ARSummary", "cell": "B19", "expected": 100, "tolerance": 0.005},
  {"sheet": "Forecast", "cell": "L19", "expected": 1674.618687430478, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "L20", "expected": 98015.04367816096, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "M5", "expected": 0.8843045388549597, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "O5", "expected": 0.9016950420810913, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "M6", "expected": 0.9000931136321721, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "O6", "expected": 0.9134076913053546, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "M7", "expected": 0.9894215161578155, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "O7", "expected": 1.0017499427226255, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "M8", "expected": 1.0010092372483812, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "O8", "expected": 1.0092407424247487, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "M9", "expected": 1.017681686648264, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "O9", "expected": 1.0175296160967362, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "M10", "expected": 1.06260173514462, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "O10", "expected": 1.0593517772301515, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "M11", "expected": 1.0674551866763862, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "O11", "expected": 1.0636664596610945, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "M12", "expected": 1.0354096967498465, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "O12", "expected": 1.029928316009592, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "M13", "expected": 0.9770721341450566, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "O13", "expected": 0.9702541360511192, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "M14", "expected": 0.9454480616844231, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "O14", "expected": 0.9372400048746478, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "M15", "expected": 0.9438840616508716, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "O15", "expected": 0.934171333667518, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "M16", "expected": 1.1756190314072057, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "O16", "expected": 1.16176493787532, "tolerance": 1e-09},
  {"sheet": "Forecast", "cell": "C38", "expected": 149928.2229885058, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "D38", "expected": 149928.2229885058, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "E38", "expected": 160041.6592582543, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "C47", "expected": 129621.0341491588, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "C39", "expected": 151602.84167593624, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "D39", "expected": 151602.84167593624, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "E39", "expected": 156971.05232609613, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "C48", "expected": 133043.26580865565, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "C40", "expected": 153277.46036336673, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "D40", "expected": 153277.46036336673, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "E40", "expected": 149763.13531356904, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "C49", "expected": 147817.49157732268, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "C41", "expected": 154952.0790507972, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "D41", "expected": 154952.0790507972, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "E41", "expected": 146499.14279254773, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "C50", "expected": 150843.73705152702, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "C42", "expected": 156626.6977382277, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "D42", "expected": 156626.6977382277, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "E42", "expected": 147837.44362412175, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "C51", "expected": 154019.29850445394, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "C43", "expected": 158301.31642565818, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "D43", "expected": 158301.31642565818, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "E43", "expected": 186102.04028681785, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "C52", "expected": 162366.03395872313, "tolerance": 1e-06},
  {"sheet": "Forecast", "cell": "B54", "expected": 0.04616377740364807, "tolerance": 1e-12},
  {"sheet": "Forecast", "cell": "B55", "expected": 0.044791890612870124, "tolerance": 1e-12},
  {"sheet": "Forecast", "cell": "B56", "expected": 0.12935047052999243, "tolerance": 1e-12},
  {"sheet": "Forecast", "cell": "B57", "expected": true, "tolerance": 0.005},
  {"sheet": "Forecast", "cell": "M17", "expected": 1.0, "tolerance": 1e-12},
  {"sheet": "Drivers", "cell": "C9", "expected": "Base", "tolerance": 0.005},
  {"sheet": "Drivers", "cell": "H18", "expected": 792069.592181376, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "B21", "expected": 691158.827474712, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "C21", "expected": 792069.592181376, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "D21", "expected": 870199.8167901551, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "G17", "expected": 1519.6413335808002, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "B29", "expected": 731377.109024022, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "C29", "expected": 774399.291907788, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "D29", "expected": 817421.4747915539, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "E29", "expected": 860443.65767532, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "F29", "expected": 903465.840559086, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "B30", "expected": 719944.3796023043, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "C30", "expected": 762294.0489906752, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "D30", "expected": 804643.718379046, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "E30", "expected": 846993.3877674169, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "F30", "expected": 889343.0571557876, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "B31", "expected": 708693.845635968, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "C31", "expected": 750381.718908672, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "D31", "expected": 792069.592181376, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "E31", "expected": 833757.4654540799, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "F31", "expected": 875445.338726784, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "B32", "expected": 697622.8357983399, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "C32", "expected": 738659.4731982421, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "D32", "expected": 779696.1105981445, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "E32", "expected": 820732.7479980469, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "F32", "expected": 861769.3853979493, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "B33", "expected": 686728.710534918, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "C33", "expected": 727124.517036972, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "D33", "expected": 767520.3235390261, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "E33", "expected": 807916.13004108, "tolerance": 1e-06},
  {"sheet": "Drivers", "cell": "F33", "expected": 848311.936543134, "tolerance": 1e-06},
  {"sheet": "Checks", "cell": "B3", "expected": true, "tolerance": 0.005},
  {"sheet": "Checks", "cell": "B4", "expected": true, "tolerance": 0.005},
  {"sheet": "Checks", "cell": "B5", "expected": true, "tolerance": 0.005},
  {"sheet": "Checks", "cell": "B6", "expected": true, "tolerance": 0.005},
  {"sheet": "Checks", "cell": "B7", "expected": true, "tolerance": 0.005},
  {"sheet": "Checks", "cell": "B8", "expected": true, "tolerance": 0.005},
  {"sheet": "Checks", "cell": "B9", "expected": true, "tolerance": 0.005},
  {"sheet": "Checks", "cell": "B10", "expected": true, "tolerance": 0.005},
  {"sheet": "Checks", "cell": "C12", "expected": 100, "tolerance": 0.005}
]
```

---

## Appendix E — What-If Experiments (JSON)

Each experiment applies `changes` to the base workbook. `expect` lists the recalculated values of the watched cells.

```json
[
 {
  "id": "E2.1",
  "lab": 2,
  "title": "Lower the % threshold to 3%",
  "changes": {
   "BvA!B2": 0.03
  },
  "expect": {
   "BvA!H6": true,
   "BvA!H7": true,
   "BvA!H8": false,
   "BvA!H9": true,
   "BvA!H10": false,
   "BvA!H11": false,
   "BvA!H12": true
  }
 },
 {
  "id": "E2.2",
  "lab": 2,
  "title": "Switch the report period to 2026-07 (no budget loaded)",
  "changes": {
   "BvA!B1": "2026-07"
  },
  "expect": {
   "BvA!C6": 0,
   "BvA!D6": 455000,
   "BvA!F6": "n/a",
   "BvA!G6": "Favorable",
   "BvA!H6": false,
   "BvA!D14": 0
  }
 },
 {
  "id": "E3.1",
  "lab": 3,
  "title": "Raise the flux % threshold to 20%",
  "changes": {
   "Flux!B4": 0.2
  },
  "expect": {
   "Flux!I8": null,
   "Flux!I9": null,
   "Flux!I10": null,
   "Flux!I11": "Explain",
   "Flux!I12": null,
   "Flux!I13": null
  }
 },
 {
  "id": "E4.1",
  "lab": 4,
  "title": "Pro sells at budget price (E6 = 120)",
  "changes": {
   "PVM!E6": 120
  },
  "expect": {
   "PVM!J8": 3600,
   "PVM!K8": -937.5,
   "PVM!L8": 3937.5,
   "PVM!M8": 6600,
   "PVM!N8": 0
  }
 },
 {
  "id": "E4.2",
  "lab": 4,
  "title": "Same total units, budget mix (D5=994, D6=497, D7=99)",
  "changes": {
   "PVM!D5": 994,
   "PVM!D6": 497,
   "PVM!D7": 99
  },
  "expect": {
   "PVM!K8": -937.5,
   "PVM!L8": -122.5,
   "PVM!N8": 0
  }
 },
 {
  "id": "E5.1",
  "lab": 5,
  "title": "Negotiate freight rate down to 11.50",
  "changes": {
   "CostVar!B6": 11.5
  },
  "expect": {
   "CostVar!B10": -1150,
   "CostVar!C10": "Favorable",
   "CostVar!B12": 2450,
   "CostVar!C12": "Unfavorable"
  }
 },
 {
  "id": "E7.1",
  "lab": 7,
  "title": "Tighten amount tolerance to 1.00",
  "changes": {
   "BankRec!B2": 1
  },
  "expect": {
   "Book!E4": "Unmatched",
   "Bank!E4": "Unmatched",
   "BankRec!B10": 5779.5,
   "BankRec!B15": 4970,
   "BankRec!B16": 0,
   "BankRec!B12": 13779.5,
   "BankRec!B17": 13779.5,
   "BankRec!B19": 0
  }
 },
 {
  "id": "E7.2",
  "lab": 7,
  "title": "Bank statement balance keyed wrong (10,100)",
  "changes": {
   "BankRec!B5": 10100
  },
  "expect": {
   "BankRec!B19": 100,
   "BankRec!B20": "Difference: 100.00"
  }
 },
 {
  "id": "E8.1",
  "lab": 8,
  "title": "Roll the as-of date to 2026-09-30",
  "changes": {
   "AR!H2": "2026-09-30"
  },
  "expect": {
   "ARSummary!B4": 0,
   "ARSummary!B5": 900,
   "ARSummary!B6": 2850,
   "ARSummary!B7": 1500,
   "ARSummary!B8": 4200,
   "ARSummary!E9": 2485.5
  }
 },
 {
  "id": "E9.1",
  "lab": 9,
  "title": "Insert an outlier: Apr-24 sales = 250,000",
  "changes": {
   "Forecast!C8": 250000
  },
  "expect": {
   "Forecast!L19": 950.443604004449,
   "Forecast!M8": 1.36956149864869,
   "Forecast!E41": 134252.380595987,
   "Forecast!B54": 0.161791702130333
  }
 },
 {
  "id": "E10.1",
  "lab": 10,
  "title": "Switch to the Upside scenario (C8 = 3)",
  "changes": {
   "Drivers!C8": 3
  },
  "expect": {
   "Drivers!C9": "Upside",
   "Drivers!G17": 1673.91487359523,
   "Drivers!H18": 870199.816790155,
   "Checks!B9": true
  }
 },
 {
  "id": "E10.2",
  "lab": 10,
  "title": "Base scenario with churn 3% (C5 = 0.03)",
  "changes": {
   "Drivers!C5": 0.03
  },
  "expect": {
   "Drivers!H18": 767520.323539026,
   "Drivers!B33": 686728.710534918,
   "Drivers!D33": 767520.323539026
  }
 }
]
```

---
*End of Excel Implementation Labs.*
