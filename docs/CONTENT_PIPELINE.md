# Content Pipeline

All lesson text, workbook cells, expected values, and experiments come from one file:
[`content/Excel_Implementation_Labs_Financial_Analysis.md`](../content/Excel_Implementation_Labs_Financial_Analysis.md) (the "Labs document"). The app never hard-codes any of it.

```
Labs document ──npm run extract──► src/data/*.json ──► app + tests
                                        │
                     tests/golden.test.ts: deep-equals reference/golden/
```

## Commands

| Command | What it does |
|---|---|
| `npm run extract` | Deletes `src/data/` and regenerates every file from the Labs document |
| `npm run extract -- --check` | Regenerates in memory and fails if any file in `src/data/` is missing, stale, or extra (CI runs this) |

## Generated files (`src/data/`, never edit by hand)

| File | Source in the Labs document | Validation |
|---|---|---|
| `workbook.json` | Appendix C: each `#### Sheet \`NAME\` (N cells)` heading and its `json` block | 15 sheets; each sheet has exactly N cells; each cell has exactly one of `v`/`f`; formulas start with `=`; dates are `{"date": "YYYY-MM-DD"}` |
| `assertions.json` | The single `json` block inside Appendix D | 277 entries; every `sheet!cell` exists in the workbook |
| `experiments.json` | The single `json` block inside Appendix E | 12 entries; every `changes`/`expect` reference exists in the workbook |
| `exercises.json` | Lab formula-table rows `` | `SHEET!ADDR` | `=FORMULA` | RESULT | `` | Each formula equals the Appendix C formula for that cell; no duplicates |
| `lessons/<id>.json` | Level-2 sections: Part 0, Labs 1–12, Appendices A–B | Labs 1–12 all present and in order; markdown is verbatim (only the trailing `---` separator is trimmed) |
| `notice.json` | The verification blockquote before the Table of Contents | Present |
| `manifest.json` | All of the above | Counts, lesson index (id, kind, n, title, slug), SHA-256 of the Labs document |

Dates stay as `{date}` objects in `workbook.json`. The engine adapter converts them to 1900-system serials when it loads the workbook (OPEN_ISSUES #2).

## Golden cross-check

`reference/golden/` was produced independently of this app (see [`reference/README.md`](../reference/README.md)). `tests/golden.test.ts` requires:

- `workbook.json` sheet order and every sheet's cells deep-equal `spec.json`
- `assertions.json` deep-equals `assertions.json` (and `spec.json`'s `assertions`)
- `experiments.json` deep-equals `experiments.json`

If this test fails, **the extractor is wrong**. Do not edit the golden files.

## Updating content

1. Edit the Labs document.
2. Run `npm run extract`. It stops with a specific message if a count, reference, or formula no longer lines up.
3. Run `npm test`. If the golden cross-check fails because the content itself changed on purpose, the golden files must be regenerated with `reference/generator/` (Python + LibreOffice, outside CI) and reviewed by the owner.
4. Commit the Labs document and `src/data/` together.

## Code

| Module | Role |
|---|---|
| `scripts/extract-content.ts` | CLI: write or `--check` |
| `scripts/extract/content.ts` | Pure extraction: markdown → file contents |
| `scripts/extract/markdown.ts` | Fence-aware section and code-block scanning |
| `scripts/extract/validate.ts` | Shape checks for cells, assertions, and experiments |
| `src/content/types.ts` | Data types shared by the extractor and the app |

The scripts run directly on Node 24 (native TypeScript type stripping). No build step or `tsx` is needed, which is why the code uses only erasable TypeScript syntax (`erasableSyntaxOnly`).
