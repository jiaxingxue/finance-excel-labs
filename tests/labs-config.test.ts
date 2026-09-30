// Build-time validation of the hand-authored lab config against the generated data (PRD §4.5).

import { describe, expect, it } from 'vitest';
import assertions from '../src/data/assertions.json' with { type: 'json' };
import manifest from '../src/data/manifest.json' with { type: 'json' };
import workbook from '../src/data/workbook.json' with { type: 'json' };
import { labs } from '../src/config/labs.config.ts';

describe('src/config/labs.config.ts', () => {
  it('defines Labs 1–12 in order, each with a lesson', () => {
    expect(labs.map((l) => l.n)).toEqual(Array.from({ length: 12 }, (_, i) => i + 1));
    const lessonLabs = manifest.lessons.filter((l) => l.kind === 'lab').map((l) => l.n);
    expect(lessonLabs).toEqual(labs.map((l) => l.n));
  });

  it('shows every workbook sheet in at least one lab', () => {
    const shown = new Set(labs.flatMap((l) => l.sheets));
    expect(workbook.sheetOrder.filter((s) => !shown.has(s))).toEqual([]);
  });

  it('only references sheets that exist', () => {
    const known = new Set(workbook.sheetOrder);
    const unknown = labs
      .flatMap((l) => [...l.sheets, ...l.primarySheets])
      .filter((s) => !known.has(s));
    expect(unknown).toEqual([]);
  });

  it('lists each primary sheet among the lab’s shown sheets', () => {
    for (const lab of labs) {
      for (const s of lab.primarySheets) expect(lab.sheets, `Lab ${lab.n}`).toContain(s);
    }
  });

  it('gives every sheet at most one owning lab', () => {
    const owners = labs.flatMap((l) => l.primarySheets);
    expect(owners.length).toBe(new Set(owners).size);
  });

  it('assigns every assertion to exactly one lab through its sheet', () => {
    const owned = new Set(labs.flatMap((l) => l.primarySheets));
    const orphans = assertions
      .filter((a) => !owned.has(a.sheet))
      .map((a) => `${a.sheet}!${a.cell}`);
    expect(orphans).toEqual([]);
  });
});
