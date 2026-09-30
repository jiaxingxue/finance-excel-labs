// Conformance suite (PRD §7.3) — the M1 gate. Expected values come only from src/data/.
// scripts/write-verification.ts counts the results of the "assertions" and "experiments" blocks
// below by name, so keep those describe titles stable.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';
import type { Assertion, Experiment, WorkbookSpec } from '../src/content/types.ts';
import {
  createEngine,
  formatRef,
  isErrorValue,
  isoDateToSerial,
  type CellValue,
  type FormulaEngine,
} from '../src/engine/index.ts';
import { checkAssertion, runExperiment } from '../src/grading/conformance.ts';

const read = <T>(file: string): T =>
  JSON.parse(readFileSync(join(import.meta.dirname, '../src/data', file), 'utf8')) as T;

const workbook = read<WorkbookSpec>('workbook.json');
const assertions = read<Assertion[]>('assertions.json');
const experiments = read<Experiment[]>('experiments.json');

let engine: FormulaEngine;
beforeAll(() => {
  engine = createEngine(workbook);
});

const show = (v: CellValue) => JSON.stringify(v);

describe('assertions', () => {
  it('covers all 277', () => expect(assertions).toHaveLength(277));

  it.each(assertions.map((a) => [`${a.sheet}!${a.cell}`, a] as const))('%s', (_, assertion) => {
    const result = checkAssertion(engine, assertion);
    expect(result.pass, `got ${show(result.actual)}, expected ${show(assertion.expected)}`).toBe(
      true,
    );
  });
});

describe('experiments', () => {
  it('covers all 12', () => expect(experiments).toHaveLength(12));

  it.each(experiments.map((e) => [`${e.id} ${e.title}`, e] as const))('%s', (_, experiment) => {
    const result = runExperiment(engine, experiment);
    const failures = result.checks
      .filter((c) => !c.pass)
      .map((c) => `${formatRef(c.ref)}: got ${show(c.actual)}, expected ${show(c.expected)}`);
    expect(failures).toEqual([]);
    expect(result.checks).toHaveLength(Object.keys(experiment.expect).length);
  });

  it('leave the base workbook untouched', () => {
    for (const assertion of assertions) expect(checkAssertion(engine, assertion).pass).toBe(true);
  });
});

describe('error-free base workbook', () => {
  it('no formula cell evaluates to an error', () => {
    const errors: string[] = [];
    for (const sheet of workbook.sheetOrder) {
      for (const [cell, spec] of Object.entries(workbook.sheets[sheet]!)) {
        if (!('f' in spec)) continue;
        const value = engine.getValue({ sheet, cell });
        if (isErrorValue(value)) errors.push(`${sheet}!${cell} = ${show(value)}`);
      }
    }
    expect(errors).toEqual([]);
  });

  it('cells expected to be "" are empty strings, not errors', () => {
    const blanks = assertions.filter((a) => a.expected === '');
    expect(blanks.length).toBeGreaterThan(0);
    for (const a of blanks) expect(engine.getValue(a)).toBe('');
  });
});

/** A one-sheet workbook with the given formulas in column A. */
function evaluate(...formulas: string[]): CellValue[] {
  const cells = Object.fromEntries(formulas.map((f, i) => [`A${i + 1}`, { f }]));
  const scratch = createEngine({ sheetOrder: ['T'], sheets: { T: cells } });
  try {
    return formulas.map((_, i) => scratch.getValue({ sheet: 'T', cell: `A${i + 1}` }));
  } finally {
    scratch.destroy();
  }
}

/** Runs a lab formula from workbook.json and compares with its assertion. */
function labFormulaMatchesAssertion(sheet: string, cell: string): void {
  const spec = workbook.sheets[sheet]?.[cell];
  const assertion = assertions.find((a) => a.sheet === sheet && a.cell === cell);
  expect(spec && 'f' in spec, `${sheet}!${cell} is a formula`).toBe(true);
  expect(assertion, `${sheet}!${cell} has an assertion`).toBeDefined();
  expect(checkAssertion(engine, assertion!).pass).toBe(true);
}

describe('targeted semantics (PRD §7.3 table)', () => {
  it('LOOKUP last-match trick', () => {
    expect(evaluate('=LOOKUP(2,1/({0,1,0,1}),{"a","b","c","d"})')).toEqual(['d']);
  });

  it('LOOKUP with no match is an error (caught by IFERROR)', () => {
    expect(evaluate('=IFERROR(LOOKUP(2,1/({0,0}),{"a","b"}),"")')).toEqual(['']);
  });

  it('approximate LOOKUP buckets', () => {
    const xs = [-10, 0, 1, 30, 31, 60, 61, 90, 91, 108];
    const buckets = '{-99999,1,31,61,91},{"Current","1-30","31-60","61-90","90+"}';
    expect(evaluate(...xs.map((x) => `=LOOKUP(${x},${buckets})`))).toEqual([
      'Current',
      'Current',
      '1-30',
      '1-30',
      '31-60',
      '31-60',
      '61-90',
      '61-90',
      '90+',
      '90+',
    ]);
  });

  it('boolean equality with bare TRUE/FALSE literals', () => {
    expect(evaluate('=(TRUE)=(FALSE)', '=TRUE', '=FALSE=FALSE')).toEqual([false, true, true]);
  });

  it('N() on text', () => {
    expect(evaluate('=N("n/a")')).toEqual([0]);
  });

  it('SUMIFS with a text-period criterion (BvA!C6)', () => labFormulaMatchesAssertion('BvA', 'C6'));

  it('AVERAGEIFS with a concatenated criterion (Forecast!N5)', () => {
    // N5 has no assertion of its own; only the ratios O5:O16 (= N/AVERAGE(N)) are asserted.
    // So check N5 against a plain average of its inputs, and check the dependent assertions.
    const formula = '=AVERAGEIFS($I$5:$I$34,$E$5:$E$34,K5,$B$5:$B$34,"<="&$B$2)';
    expect(engine.getFormula({ sheet: 'Forecast', cell: 'N5' })).toBe(formula);
    const at = (cell: string) => engine.getValue({ sheet: 'Forecast', cell });
    const [month, cutoff] = [at('K5'), at('B2')] as number[];
    const matched: number[] = [];
    for (let row = 5; row <= 34; row++) {
      if (at(`E${row}`) === month && (at(`B${row}`) as number) <= cutoff!) {
        matched.push(at(`I${row}`) as number);
      }
    }
    expect(matched.length).toBeGreaterThan(0);
    expect(at('N5')).toBeCloseTo(matched.reduce((s, v) => s + v, 0) / matched.length, 9);
    for (let row = 5; row <= 16; row++) labFormulaMatchesAssertion('Forecast', `O${row}`);
  });

  it('SUMPRODUCT with array power (Drivers!C21)', () =>
    labFormulaMatchesAssertion('Drivers', 'C21'));

  it('TEXT format', () => {
    expect(evaluate('=TEXT(100,"#,##0.00")')).toEqual(['100.00']);
  });

  it('EDATE on a date serial', () => {
    expect(evaluate('=EDATE(DATE(2026,6,1),1)')).toEqual([isoDateToSerial('2026-07-01')]);
  });

  it('divide-by-zero guard', () => {
    expect(evaluate('=IF(0=0,"n/a",1/0)')).toEqual(['n/a']);
  });
});
