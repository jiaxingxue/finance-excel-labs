// src/config/tiers.ts is hand-transcribed; these tests hold it to the Labs document's Part 0.3 table.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import workbook from '../src/data/workbook.json' with { type: 'json' };
import type { WorkbookSpec } from '../src/content/types.ts';
import {
  TIER_EXTRAS,
  TIER_TABLE,
  formulaFunctions,
  formulaTier,
  tierOf,
  type Tier,
} from '../src/config/tiers.ts';

const labsDocument = readFileSync(
  join(import.meta.dirname, '../content/Excel_Implementation_Labs_Financial_Analysis.md'),
  'utf8',
);

/** Tier → the backticked names in the "Examples used here" column of the Part 0.3 table. */
function part03Table(): Record<string, string[]> {
  const start = labsDocument.indexOf('### 0.3 Function tiers');
  expect(start, 'Part 0.3 heading').toBeGreaterThan(-1);
  const section = labsDocument.slice(start, labsDocument.indexOf('\n#', start + 1));
  const table: Record<string, string[]> = {};
  for (const line of section.split('\n')) {
    const row = /^\|\s*\*\*([A-D]) —[^|]*\|[^|]*\|([^|]*)\|/.exec(line);
    if (row) table[row[1]!] = [...row[2]!.matchAll(/`([^`]+)`/g)].map((m) => m[1]!);
  }
  return table;
}

describe('src/config/tiers.ts', () => {
  it('matches the Part 0.3 table of the Labs document exactly', () => {
    expect(part03Table()).toEqual(TIER_TABLE);
  });

  it('keeps extras out of the table, so each function has one tier', () => {
    const tabled = new Set(Object.values(TIER_TABLE).flat());
    expect(Object.keys(TIER_EXTRAS).filter((f) => tabled.has(f))).toEqual([]);
  });

  it('knows the tier of every function the workbook uses, and all are graded tiers (A/B)', () => {
    const spec = workbook as WorkbookSpec;
    const used = new Set<string>();
    for (const cells of Object.values(spec.sheets)) {
      for (const c of Object.values(cells))
        if ('f' in c) formulaFunctions(c.f).forEach((f) => used.add(f));
    }
    const tiers = Object.fromEntries([...used].sort().map((f) => [f, tierOf(f)]));
    expect(Object.values(tiers).filter((t) => t !== 'A' && t !== 'B')).toEqual([]);
    expect(tiers['FORECAST.LINEAR']).toBe<Tier>('B');
  });

  it('rates a formula by its newest function', () => {
    expect(formulaTier('=D6-C6')).toBe('A');
    expect(formulaTier('=SUMIFS(GL!E2:E37,GL!C2:C37,"IF(")')).toBe('A');
    expect(formulaTier('=ROUND(FORECAST.LINEAR(A38,C5:C34,A5:A34),0)')).toBe('B');
    expect(formulaTier('=IFERROR(xlookup(A1,B:B,C:C),"")')).toBe('C');
    expect(formulaTier('=NOTAFUNCTION(1)')).toBeNull();
  });
});
