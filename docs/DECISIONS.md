# Owner Decisions

Claude Code reads this file before starting work. Rows marked **CONFIRMED** are final. Rows marked **PENDING** can be decided later: until then Claude Code uses the listed default, and the "Needed by" column says when each answer starts to matter. To decide a row, fill in the Answer, change Status to `CONFIRMED`, and add a line to the change log.

| ID | Decision | Answer | Status | Default while pending | Needed by |
|---|---|---|---|---|---|
| **D1** | License for the public repo | **GPLv3.** The app uses HyperFormula's free `gpl-v3` license key; the repo ships the full GPLv3 `LICENSE` | **CONFIRMED** | — | — |
| **D2** | GitHub username | **`jiaxingxue`** | **CONFIRMED** | — | — |
| **D3** | Repository name | **`finance-excel-labs`** → repo `https://github.com/jiaxingxue/finance-excel-labs`, site `https://jiaxingxue.github.io/finance-excel-labs/` | **CONFIRMED** | — | — |
| **D4** | Custom domain | — | PENDING | None (use the github.io URL) | Any time; optional |
| **D5** | Author name shown in the footer and README | — | PENDING | Show the GitHub handle `@jiaxingxue` only | Before M7 (portfolio release) |
| **D6** | Public links (LinkedIn, portfolio site). **Never an email or phone number.** | — | PENDING | None; the footer links only to the GitHub repo | Before M7 |
| **D7** | License for the lesson content (the Labs document) | **GPLv3, same as the code** | **CONFIRMED** | — | — |
| **D8** | Allow Tier C functions (e.g., XLOOKUP, FILTER) in Build mode if the engine computes them correctly? (PRD Q5) | — | PENDING | Yes | Before M4 (Build mode) |
| **D9** | Genuine Microsoft Excel embedding (Office.js) | **Not in v1** | **CONFIRMED** | — | — |
| **D10** | Include the earlier study guides (financial analysis concepts; CFO controls/COSO) as reading content? (PRD Q6) | **Not in v1: placeholder link on the About page** | **CONFIRMED** | — | — |
| **D11** | Branding: app name, colors, logo (PRD Q4) | — | PENDING | Name "Excel Labs for Financial Analysis"; neutral palette | Before M6 (polish) |
| **D12** | Pause for owner review after each milestone? | **Yes** | **CONFIRMED** | — | — |

## Change log
| Date | Change |
|---|---|
| 2026-09-29 | Created with recommended defaults |
| 2026-09-29 | Confirmed D1 (GPLv3), D2 (`jiaxingxue`), D3 (`finance-excel-labs`), D10 (placeholder link on About page) |
| 2026-09-29 | Confirmed D7 (lesson content under GPLv3, same as code; the repo is public). This file is now committed (removed from `.gitignore`) |
