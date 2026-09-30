// Golden cross-check (CLAUDE.md): the committed generated data must deep-equal the
// independently produced reference data. If this fails, the extractor is wrong.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import type { Assertion, Experiment, WorkbookSpec } from '../src/content/types.ts';

const root = join(import.meta.dirname, '..');
const read = <T>(...path: string[]): T =>
  JSON.parse(readFileSync(join(root, ...path), 'utf8')) as T;

const golden = {
  spec: read<WorkbookSpec & { assertions: Assertion[] }>('reference/golden/spec.json'),
  assertions: read<Assertion[]>('reference/golden/assertions.json'),
  experiments: read<Experiment[]>('reference/golden/experiments.json'),
};
const data = {
  workbook: read<WorkbookSpec>('src/data/workbook.json'),
  assertions: read<Assertion[]>('src/data/assertions.json'),
  experiments: read<Experiment[]>('src/data/experiments.json'),
};

describe('generated data matches reference/golden', () => {
  it('has the same 15 sheets in the same order', () => {
    expect(data.workbook.sheetOrder).toHaveLength(15);
    expect(data.workbook.sheetOrder).toEqual(golden.spec.sheetOrder);
  });

  it.each(golden.spec.sheetOrder)('sheet %s deep-equals spec.json', (name) => {
    expect(data.workbook.sheets[name]).toEqual(golden.spec.sheets[name]);
  });

  it('has 277 assertions that deep-equal assertions.json', () => {
    expect(data.assertions).toHaveLength(277);
    expect(data.assertions).toEqual(golden.assertions);
    expect(data.assertions).toEqual(golden.spec.assertions);
  });

  it('has 12 experiments that deep-equal experiments.json', () => {
    expect(data.experiments).toHaveLength(12);
    expect(data.experiments).toEqual(golden.experiments);
  });
});
