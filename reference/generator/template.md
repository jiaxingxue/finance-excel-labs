# Excel Implementation Labs for Financial Analysis
### Variance Analysis · Reconciliations · Forecasting Support — hands-on, cell-by-cell

*Companion to **Financial_Analysis_Study_Guide.md**. Prepared September 2026.*

> **Verification.** Every classic-Excel formula in this document was built into a single workbook (15 sheets, {{N_FORMULAS}} formulas) and recalculated by a spreadsheet engine (LibreOffice Calc 24.2). All {{N_ASSERTIONS}} expected-value checks and {{N_EXPERIMENTS}} what-if experiments shown here come from that recalculation, compared against independent Python calculations.
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

{{VALUES:Map!A1:D7}}

- **Type** drives favorable/unfavorable logic (revenue up = good, expense up = bad).
- **Sign** converts to a profit view (+1 revenue, −1 expense) if you ever need a signed total.

### Step 2 — Build the `GL` sheet

Enter the 36 data rows in `GL!A1:E37`. All amounts are positive, and the direction comes from `Map!Type`. The first rows look like this; the full data is in [Appendix C](#appendix-c--complete-cell-by-cell-workbook-specification-json).

{{VALUES:GL!A1:E13}}

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

{{FORMULAS:BvA!B6:H6,C12,D12,D14}}

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

{{VALUES:BvA!A5:H14}}

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

{{EXPERIMENTS:2}}

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

{{VALUES:Flux!A1:B5}}

### Formulas (row 8 pattern; fill down to row 13)

{{FORMULAS:Flux!B8:I8}}

**Notes**
- The only difference between B, C, and F is the period cell they point to (`$B$1`, `$B$2`, `$B$3`). The account reference `$A8` stays the same.
- `N()` converts text such as `"n/a"` to 0, so `ABS(N(E8))` never errors. This is a compact alternative to the `ISNUMBER` guard used in Lab 2.
- The flag is `OR(AND(MoM breach), AND(YoY breach))`: either comparison crossing **both** thresholds requires an explanation.

### Result

{{VALUES:Flux!A7:I13}}

Salaries is flagged only on the year-over-year comparison (+13.2%, +14,500). Marketing breaches on both comparisons. Travel moved −25% month over month but stays under the $5,000 threshold, so it is not flagged.

### Modern equivalent (Tier C, not machine-verified)
```excel
=LET(acc, A8:A13,
     get, LAMBDA(p, SUMIFS(GL!E2:E37, GL!C2:C37, acc, GL!D2:D37, "ACT", GL!A2:A37, p)),
     cur, get(B1), pm, get(B2), py, get(B3),
     HSTACK(acc, cur, pm, cur - pm, (cur - pm) / ABS(pm), py, cur - py, (cur - py) / ABS(py)))
```

### Try it

{{EXPERIMENTS:3}}

---

## Lab 4 — Price–Volume–Mix Analysis and Variance Bridge

### Learning goals
- Decompose a revenue variance into price, volume, and mix using the **budget-mix method**.
- Build a control check proving the three effects add up to the total.
- Chart the result as a waterfall.

### Inputs (`PVM!A4:E7`)

{{VALUES:PVM!A4:E7}}

### Formulas (row 5 pattern; fill down to row 7; totals in row 8)

{{FORMULAS:PVM!F5:N5}}

Totals row:

{{FORMULAS:PVM!B8:N8}}

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

{{VALUES:PVM!A4:N8}}

**Interpretation:** revenue beat budget by {{VAL:PVM!M8}}. Mix was the main driver ({{VAL:PVM!L8}}, as customers shifted to Pro), total volume was slightly low ({{VAL:PVM!K8}}), and net price was slightly positive ({{VAL:PVM!J8}}). Note the Pro price cut: it hurt price but likely drove the favorable mix.

### Waterfall (variance bridge) chart

Waterfall data (`PVM!A10:B15`):

{{VALUES:PVM!A10:B15}}

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

{{EXPERIMENTS:4}}

In E4.2, mix is not exactly zero because whole units (994/497/99) can't hit the exact budget proportions (993.75/496.875/99.375). The check column stays at 0 in every case, which is the point of a control check.

---

## Lab 5 — Cost Rate/Volume Variances and the Flexible Budget

### Part A — Rate/volume (freight example)

Inputs `CostVar!B3:B6`: budget shipments 2,000 @ $12.00; actual 2,300 @ $12.50.

{{FORMULAS:CostVar!B8:B13,C10:C12}}

For **costs**, a positive variance means spending more than planned, which is **Unfavorable**. The label formula in column C encodes that.

{{VALUES:CostVar!A8:C13}}

### Part B — Flexible budget

A **flexible budget** restates the budget at the **actual** activity level. That splits the total variance into:
- **Sales-volume variance** (Flex − Static): the effect of selling a different number of units.
- **Flexible-budget variance** (Actual − Flex): the effect of prices and costs differing at the actual volume.

Inputs (`CostVar!A17:B24`):

{{VALUES:CostVar!A17:B24}}

Formulas:

{{FORMULAS:CostVar!B27:F32}}

Result:

{{VALUES:CostVar!A26:F32}}

**Reading it:** static-budget operating income was {{VAL:CostVar!B32}} and actual was {{VAL:CostVar!D32}}. Selling more units should have *added* {{VAL:CostVar!F32}} of profit (Flex − Static). Three things went the other way: revenue came in 5,500 below the flexible budget (lower price per unit), variable costs ran 5,500 above it (higher cost per unit), and fixed costs overran by 2,000. Together they produce a flexible-budget variance of {{VAL:CostVar!E32}}.

> **Sign reading tip:** in column E, revenue and profit rows are "actual − flex" (negative = bad), while cost rows are also "actual − flex" (positive = bad). Label the direction for readers, as Part A's column C does.

### Try it

{{EXPERIMENTS:5}}

---

## Lab 6 — Pivot-Based Analysis

Three ways to build the same dept × account summary for **2026-08 ACT**.

### Method 1 — Formula cross-tab with SUMIFS (Tier A, verified)

Inputs: `PivotLab!B2` = `2026-08`, `B3` = `ACT`. Accounts across `B5:G5`, departments down `A6:A7`.

The **one formula** in `B6`, filled across to `G6` and down to `G7`:

{{FORMULAS:PivotLab!B6,B8,B10}}

The "mixed" references do the work: `$A6` (column locked, so filling across still reads the department) and `B$5` (row locked, so filling down still reads the account).

{{VALUES:PivotLab!A5:G10}}

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

{{VALUES:Book!A1:D6}}

`Bank!A1:D5` (bank statement):

{{VALUES:Bank!A1:D5}}

`BankRec` parameters: amount tolerance `B2` = {{VAL:BankRec!B2}}, day tolerance `B3` = {{VAL:BankRec!B3}}, bank statement balance `B5` = {{VAL:BankRec!B5}}, GL balance `B6` = {{VAL:BankRec!B6}}.

### Matching formulas (row 2 pattern; fill down)

{{FORMULAS:Book!E2:G2}}

{{FORMULAS:Bank!E2:F2}}

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

{{VALUES:Book!A1:G6}}

{{VALUES:Bank!A1:F5}}

### The reconciliation proof (`BankRec`)

{{FORMULAS:BankRec!B9:B12,B14:B17,B19,B20,B22}}

{{VALUES:BankRec!A8:B22}}

The proof classifies each unmatched item by sign and source:
- Unmatched **book** deposits → deposits in transit
- Unmatched **book** payments → outstanding checks
- Unmatched **bank** items → bank-only items (need a journal entry)
- Tolerance matches → book errors (need a correcting entry)

`B22` is the journal entry: Dr Bank fees {{VAL:BankRec!B22}} / Cr Cash {{VAL:BankRec!B22}}.

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

{{EXPERIMENTS:7}}

**Lesson from E7.1 — "balanced is not the same as correct":** with a $1 tolerance the short-paid wire no longer matches. The reconciliation *still* ties (difference 0), but the items are now misclassified: a real receipt appears as a "deposit in transit" and a "bank-only item." A reviewer must check each reconciling item's nature, not just the zero.

---

## Lab 8 — AR Aging, Allowance, and Subledger-to-GL Reconciliation

### Inputs

`AR!A1:D7` open invoices, and the as-of date in `AR!H2` = {{VAL:AR!H2}}:

{{VALUES:AR!A1:D7}}

### Aging formulas (row 2 pattern; fill down)

{{FORMULAS:AR!E2:F2}}

Dates are serial numbers, so `$H$2-C2` returns days. Negative days = not yet due.

**Three equivalent bucket formulas:**

| Tier | Formula for `F2` | Notes |
|---|---|---|
| A (verified) | `=IF(E2<=0,"Current",IF(E2<=30,"1-30",IF(E2<=60,"31-60",IF(E2<=90,"61-90","90+"))))` | Nested IF; works everywhere |
| A (verified separately) | `=LOOKUP(E2,{-99999,1,31,61,91},{"Current","1-30","31-60","61-90","90+"})` | Approximate-match lookup: finds the largest threshold ≤ days. Easy to extend |
| B/C | `=IFS(E2<=0,"Current",E2<=30,"1-30",E2<=60,"31-60",E2<=90,"61-90",TRUE,"90+")` | Flat and readable (Excel 2019+) |

The `LOOKUP` version was tested separately at the boundary values −10, 0, 1, 30, 31, 60, 61, 90, 91, and 108, and matched the nested IF at every one.

{{VALUES:AR!A1:F7}}

### Summary, allowance, and customer matrix (`ARSummary`)

{{FORMULAS:ARSummary!B4:E4,B9,E9,B12,G12,B18,B19}}

{{VALUES:ARSummary!A3:E9}}

{{VALUES:ARSummary!A11:G14}}

### Subledger-to-GL tie-out

{{VALUES:ARSummary!A17:B19}}

The difference of {{VAL:ARSummary!B19}} is a real reconciling item to investigate. Typical causes are a manual journal entry posted directly to the AR control account, or unapplied cash.

### Excel UI features
- Conditional formatting on `AR!E2:E7`: Color Scales (green → red) to visualize age.
- Sort `AR` by days past due, descending (Data → Sort) to create a collections worklist.
- PivotTable alternative: Rows = Customer, Columns = Bucket, Values = Sum of Open amount.

### Try it

{{EXPERIMENTS:8}}

---

## Lab 9 — Forecasting: Moving Average, Trend, Seasonality, Backtest

### Data

`Forecast!A5:C34` holds 30 months of sales (Jan-2024 → Jun-2026) with a trend and a December peak. `B` is a period counter `t` = 1…30. The train/test split input is `B2` = 24 (train on the first 24 months, test on the last 6).

{{VALUES:Forecast!A4:C10}}

*(rows 11–34 continue the series; the full data is in Appendix C)*

### Step 1 — Helper columns (row 5 pattern; fill down to row 34)

{{FORMULAS:Forecast!D7,E5:I5}}

- `D7` starts the 3-month moving average (it needs two prior months).
- **Trend (all)** uses all 30 months. **Trend (train)** uses only rows 5–28 (t ≤ 24). The backtest must never "peek" at the test months.
- **Ratio** = actual ÷ trend. A December ratio of about 1.2 means December runs ~20% above trend.

### Step 2 — Regression statistics

{{FORMULAS:Forecast!L19:L21}}

Sales grow by about {{VAL:Forecast!L19}} per month. R² of {{VAL:Forecast!L21}} means the straight line alone explains about 69% of the variation; seasonality explains much of the rest.

### Step 3 — Seasonal index table (`K5:O16`)

{{FORMULAS:Forecast!L5:O5}}

Fill `L5:O5` down to row 16 (`K5:K16` = 1…12). Dividing by the average **normalizes** the index so the 12 values average exactly 1.0 (`M17`). Otherwise the seasonal forecast would be biased up or down overall.

{{VALUES:Forecast!K4:O17}}

### Step 4 — Forecast the next 6 months (rows 38–43)

{{FORMULAS:Forecast!A38:E38,A39:B39}}

{{VALUES:Forecast!A37:E43}}

- `FORECAST.LINEAR` and `TREND` give identical straight-line forecasts (C = D).
- **Trend × season** (E) layers in the seasonal pattern: December jumps to {{VAL:Forecast!E43}}.

### Step 5 — Backtest against the last 6 months (rows 47–57)

{{FORMULAS:Forecast!A47:F47,B54:B57}}

{{VALUES:Forecast!A46:F57}}

**Reading it:** the model's MAPE is {{VAL:Forecast!B54}} against {{VAL:Forecast!B56}} for the seasonal-naive benchmark ("same month last year"), so the model adds value. However, the bias is {{VAL:Forecast!B55}}: the model **over-forecast consistently**. Growth may be slowing, and the next step is to investigate the drivers.

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

{{EXPERIMENTS:9}}

**Lesson from E9.1:** one bad data point (a mis-keyed April) cut the estimated growth rate nearly in half, inflated April's seasonal index to about 1.37, and tripled the error. **Always scan history for outliers before forecasting** (a quick line chart is usually enough).

---

## Lab 10 — Driver-Based Forecast, Scenarios, and Sensitivity

### Inputs (`Drivers!A3:E10`)

{{VALUES:Drivers!A3:E10}}

- `C8` is the **scenario switch** (1 = Downside, 2 = Base, 3 = Upside). Add a dropdown: Data Validation → List → `1,2,3`.
- Column E ("Active") pulls the selected scenario's value:

{{FORMULAS:Drivers!E4:E6,C9}}

Classic alternative: `=CHOOSE($C$8,B4,C4,D4)`.

### Customer roll-forward (columns B–G, rows 13–18)

{{FORMULAS:Drivers!B13:B18,C13:C14,H18}}

Fill `C13:C18` across to column G. Each month's opening customers equal the prior month's closing customers (`C14 = B17`). This "roll-forward" pattern is used everywhere in finance: customers, headcount, inventory, debt, cash.

{{VALUES:Drivers!A13:H18}}

### All scenarios at once (closed form, row 21)

A roll-forward only shows the *active* scenario. To show every scenario side by side, use the closed-form formula for customers after k months. With churn *c*, new customers *n*, and starting customers *C₀*: Cₖ = C₀(1−c)ᵏ + n·(1−(1−c)ᵏ)/c. Summing revenue over k = 1…6 using the helper row `B25:G25` (= 1…6):

{{FORMULAS:Drivers!B21:D21}}

{{VALUES:Drivers!A21:D21}}

`Checks!B9` confirms that the roll-forward total (`H18`) equals the closed form for the active scenario.

### Sensitivity grid: churn × ARPU (rows 28–33)

{{FORMULAS:Drivers!B29}}

Fill `B29` across to `F29` and down to row 33 (mixed references `B$28` and `$A29`).

{{VALUES:Drivers!A28:F33}}

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

{{EXPERIMENTS:10}}

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
| `VARPCT` | `=LAMBDA(actual,budget,IF(budget=0,"n/a",(actual-budget)/ABS(budget)))` | `=VARPCT(BvA!D6,BvA!C6)` | `BvA!F6` = {{VAL:BvA!F6}} |
| `FAVFLAG` | `=LAMBDA(type,actual,budget,IF(budget=0,"No budget",IF(actual=budget,"On budget",IF((type="Revenue")=(actual>budget),"Favorable","Unfavorable"))))` | `=FAVFLAG(BvA!B9,BvA!D9,BvA!C9)` | `BvA!G9` = {{VAL:BvA!G9}} |
| `NEEDSCOMMENT` | `=LAMBDA(var,pct,absThr,pctThr,IF(ISNUMBER(pct),AND(ABS(var)>=absThr,ABS(pct)>=pctThr),FALSE))` | `=NEEDSCOMMENT(BvA!E9,BvA!F9,5000,0.05)` | `BvA!H9` = {{VAL:BvA!H9}} |
| `AGEBUCKET` | `=LAMBDA(days,IFS(days<=0,"Current",days<=30,"1-30",days<=60,"31-60",days<=90,"61-90",TRUE,"90+"))` | `=AGEBUCKET(AR!E2)` | `AR!F2` = {{VAL:AR!F2}} |
| `MAPE` | `=LAMBDA(actual,forecast,AVERAGE(ABS(actual-forecast)/ABS(actual)))` | `=MAPE(Forecast!B47:B52,Forecast!C47:C52)` | `Forecast!B54` = {{VAL:Forecast!B54}} |
| `BIAS` | `=LAMBDA(actual,forecast,(SUM(forecast)-SUM(actual))/SUM(actual))` | `=BIAS(Forecast!B47:B52,Forecast!C47:C52)` | `Forecast!B55` = {{VAL:Forecast!B55}} |
| `CUSTROLL` | `=LAMBDA(start,churn,new,n,SCAN(start,SEQUENCE(n),LAMBDA(c,k,c*(1-churn)+new)))` | `=INDEX(CUSTROLL(1200,0.02,80,6),6)` | `Drivers!G17` = {{VAL:Drivers!G17}} |
| `PVMCHECK` | `=LAMBDA(price,vol,mix,total,ROUND(price+vol+mix-total,6)=0)` | `=PVMCHECK(PVM!J8,PVM!K8,PVM!L8,PVM!M8)` | TRUE |

**Tips**
- Document each LAMBDA in the Name Manager's *Comment* field (arguments and units).
- Share a library by copying one sheet that uses the names into a new workbook; the names travel with it.
- Keep a Tier A fallback column during rollout so anyone on older Excel can still verify results.

---

## Lab 12 — Model Integrity Checks Dashboard

Every serious model gets a **Checks** sheet: one place that proves the model is internally consistent.

{{FORMULAS:Checks!B3:C9,B10,C12}}

{{VALUES:Checks!A2:C12}}

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

{{SPEC}}

---

## Appendix D — Assertions (Expected Values for Auto-Grading)

Each expected value was calculated independently in Python and matched the spreadsheet engine's recalculation of Appendix C.

{{ASSERTIONS}}

---

## Appendix E — What-If Experiments (JSON)

Each experiment applies `changes` to the base workbook. `expect` lists the recalculated values of the watched cells.

{{EXPERIMENTS_JSON}}

---
*End of Excel Implementation Labs.*
