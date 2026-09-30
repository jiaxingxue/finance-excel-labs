// npm run verification:write → reports/vitest.json (conformance results) → public/verification.json
// (PRD §17.5). CI runs it after the tests and before the build, so the file ships with the site.
// Counts come from the "assertions" and "experiments" blocks of tests/conformance.test.ts; totals
// come from src/data/, so a missing test counts as a failure.

import { execSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { HyperFormula } from 'hyperformula';
import type { Assertion, Experiment } from '../src/content/types.ts';
import { ENGINE_NAME, PLUGIN_FUNCTIONS } from '../src/engine/index.ts';

interface VitestReport {
  testResults: {
    assertionResults: { ancestorTitles: string[]; title: string; status: string }[];
  }[];
}

const root = join(import.meta.dirname, '..');
const reportFile = join(root, 'reports', 'vitest.json');
const outFile = join(root, 'public', 'verification.json');
const readJson = <T>(file: string): T => JSON.parse(readFileSync(file, 'utf8')) as T;

if (!existsSync(reportFile)) {
  console.error(
    'verification:write: reports/vitest.json not found. Run ' +
      '`npm test -- --reporter=default --reporter=json --outputFile=reports/vitest.json` first.',
  );
  process.exit(1);
}

const results = readJson<VitestReport>(reportFile).testResults.flatMap((f) => f.assertionResults);

/** Passing tests in a top-level describe block whose titles match `pattern`. */
function passed(block: string, pattern: RegExp): number {
  const tests = results.filter((t) => t.ancestorTitles[0] === block && pattern.test(t.title));
  return tests.filter((t) => t.status === 'passed').length;
}

const totals = {
  assertions: readJson<Assertion[]>(join(root, 'src/data/assertions.json')).length,
  experiments: readJson<Experiment[]>(join(root, 'src/data/experiments.json')).length,
};
// Per-item test titles: "BvA!C6" and "E2.1 <title>".
const assertions = { passed: passed('assertions', /^[^!]+![A-Z]+\d+$/), total: totals.assertions };
const experiments = { passed: passed('experiments', /^E\d+\.\d+ /), total: totals.experiments };

function commit(): string {
  const sha = process.env.VITE_COMMIT_SHA ?? process.env.GITHUB_SHA;
  if (sha) return sha.slice(0, 7);
  try {
    return execSync('git rev-parse --short HEAD', { stdio: ['ignore', 'pipe', 'ignore'] })
      .toString()
      .trim();
  } catch {
    return 'unknown';
  }
}

const verification = {
  assertions,
  experiments,
  engine: { name: ENGINE_NAME, version: HyperFormula.version, plugins: PLUGIN_FUNCTIONS },
  commit: commit(),
  builtAt: new Date().toISOString(),
  label: `${assertions.passed}/${assertions.total} passing`,
};

mkdirSync(join(root, 'public'), { recursive: true });
writeFileSync(outFile, `${JSON.stringify(verification, null, 2)}\n`);
console.log(
  `verification:write: assertions ${assertions.passed}/${assertions.total}, ` +
    `experiments ${experiments.passed}/${experiments.total} → public/verification.json`,
);
