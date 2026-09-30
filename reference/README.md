# Golden Reference Data (read-only)

These files were produced **independently of the app**, before any app code existed. They exist so that the app's extraction script and formula engine can be checked against an outside source, not only against themselves.

## How they were made
1. `generator/build.py` defines the lab workbook (15 sheets, 668 formulas), writes it to `.xlsx`, and has **LibreOffice Calc 24.2** recalculate it headlessly.
2. It compares 277 expected values, computed separately in Python/NumPy, against the recalculated workbook: **277/277 match**.
3. `generator/experiments.py` applies 12 what-if changes, recalculates each in LibreOffice, and records the results.
4. `generator/render.py` fills `generator/template.md` with these verified values to produce `content/Excel_Implementation_Labs_Financial_Analysis.md`.

The generator scripts are included **for provenance only**. They need Python, openpyxl, NumPy, and LibreOffice, and should not be run in CI.

## Files
| File | Contents | App counterpart |
|---|---|---|
| `golden/spec.json` | `sheetOrder`, `sheets` (every cell: `v` value, `f` formula, `fmt` number format), `assertions` | `src/data/workbook.json` (sheets must deep-equal) |
| `golden/assertions.json` | 277 assertions `{sheet, cell, expected, tolerance}` | `src/data/assertions.json` (must deep-equal) |
| `golden/experiments.json` | 12 experiments `{id, lab, title, changes, expect}` (same format as the Labs document's Appendix E) | `src/data/experiments.json` (must deep-equal) |
| `golden/experiments.full.json` | Same experiments with before/after values for every watched cell | Reference only |
| `golden/excel_labs_verified.xlsx` | The recalculated workbook, with cached values. Open it in Excel to inspect any lab | Reference only; useful for export round-trip comparisons |

## Notes
- Dates in `spec.json` are `{"date": "YYYY-MM-DD"}`. Convert them to Excel 1900-system serials when loading.
- In the `.xlsx`, `FORECAST.LINEAR` is stored as `_xlfn.FORECAST.LINEAR`, as the file format requires.
- Microsoft 365–only formulas (LET, LAMBDA, XLOOKUP, …) are **not** in this workbook. They appear only as text in the Labs document.
- `.claude/settings.json` denies edits to `reference/golden/`. If you believe a golden value is wrong, record it in `docs/OPEN_ISSUES.md` with evidence, and let the owner decide.
