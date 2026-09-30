---
description: Run the full verification pipeline (content sync, lint, types, conformance, build, e2e) and report results
---

Run the project's verification pipeline and report the results. Stop at the first failing step, explain the failure, and propose a fix. Do not apply the fix until I say so.

Steps, in order:
1. `npm run extract -- --check`: generated data must match the Labs document.
2. Cross-check the generated data against `reference/golden/`: 15 sheets whose cells deep-equal `spec.json`, 277 assertions, 12 experiments. Report any difference.
3. `npm run lint` and `npm run typecheck`.
4. `npm test`: report the conformance numbers explicitly as "assertions X/277, experiments Y/12", plus the total unit-test count.
5. `npm run build:pages` (builds with `BASE_PATH=/finance-excel-labs/` via cross-env; works in PowerShell and Git Bash): report bundle sizes and flag the initial JS if it exceeds 600 KB gzipped (PRD §8).
6. `npm run test:e2e` (only if steps 1–5 pass).

Finish with a summary table (step, status, key numbers). Then state which acceptance criteria from PRD §11 are now met for the current milestone and which remain.

$ARGUMENTS
