// What-if controls (WI-1, OPEN_ISSUES #28) and conditional formats (GR-12), checked against the
// Labs document (Appendix B.4), the generated workbook, and the experiments. No expected value is
// written here: bases come from workbook.json, experiment values from experiments.json, and
// conditional-format results from the engine's own values.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import experimentsJson from '../src/data/experiments.json' with { type: 'json' };
import workbookJson from '../src/data/workbook.json' with { type: 'json' };
import type { CellSpec, Experiment, WorkbookSpec } from '../src/content/types.ts';
import { labs, type Control } from '../src/config/labs.config.ts';
import { isoDateToSerial, parseRef, rangeAddresses } from '../src/engine/address.ts';
import { createEngine, type CellRef } from '../src/engine/index.ts';
import {
  controlCells,
  controlLabelCell,
  groupTitleCell,
  keyOutputs,
  labControlCells,
  tableLabels,
} from '../src/whatif/controls.ts';
import { shiftFormula, sheetFormats } from '../src/workspace/conditionalFormat.ts';
import { WorkbookStore } from '../src/workspace/workbookStore.ts';

const experiments = experimentsJson as unknown as Experiment[];
const workbook = workbookJson as WorkbookSpec;
const doc = readFileSync(
  join(import.meta.dirname, '../content/Excel_Implementation_Labs_Financial_Analysis.md'),
  'utf8',
);

const key = ({ sheet, cell }: CellRef) => `${sheet}!${cell}`;
const specOf = ({ sheet, cell }: CellRef): CellSpec | undefined => workbook.sheets[sheet]?.[cell];

/** A constant's value as the engine stores it (dates as serials). */
function baseOf(ref: CellRef): string | number | boolean {
  const spec = specOf(ref);
  if (!spec || 'f' in spec) throw new Error(`${key(ref)} is not an input cell`);
  return typeof spec.v === 'object' ? isoDateToSerial(spec.v.date)! : spec.v;
}

/** Experiment values written to a cell, converted the way experiments are applied. */
function experimentValues(ref: CellRef): (string | number)[] {
  return experiments.flatMap((e) => {
    const v = e.changes[key(ref)];
    if (v === undefined) return [];
    return [typeof v === 'string' ? (isoDateToSerial(v) ?? v) : v];
  });
}

/** Cells listed in Appendix B.4's table. A bare range (`B17:B24`) is on the row's last sheet. */
function appendixB4Cells(): { lab: number; cells: string[] }[] {
  const section = doc.slice(doc.indexOf('### B.4'), doc.indexOf('### B.5'));
  const rows = section.split('\n').filter((l) => /^\| \d+ \|/.test(l));
  expect(rows.length).toBeGreaterThan(0);
  return rows.map((row) => {
    const [, lab, cellsColumn] = row.split('|').map((c) => c.trim());
    let sheet = '';
    const cells: string[] = [];
    for (const [, token] of cellsColumn!.matchAll(/`([^`]+)`/g)) {
      const [qualifier, range] = token!.includes('!') ? token!.split('!') : [sheet, token!];
      sheet = qualifier!;
      const [start, end = start] = range!.split(':') as [string, string?];
      cells.push(
        ...rangeAddresses(start, end)
          .flat()
          .map((c) => `${sheet}!${c}`),
      );
    }
    return { lab: Number(lab), cells };
  });
}

const allControls = labs.flatMap((lab) =>
  lab.controls.flatMap((group) => group.controls.map((control) => ({ lab, control }))),
);

describe('What-if controls (WI-1)', () => {
  it('cover every input cell in Appendix B.4, in the lab that lists it', () => {
    const missing: string[] = [];
    for (const { lab, cells } of appendixB4Cells()) {
      const config = labs.find((l) => l.n === lab)!;
      const covered = new Set(labControlCells(config, experiments).map(key));
      missing.push(...cells.filter((c) => !covered.has(c)).map((c) => `Lab ${lab}: ${c}`));
    }
    expect(missing).toEqual([]);
  });

  it('write only input cells (constants), never formulas', () => {
    for (const { lab, control } of allControls) {
      for (const ref of controlCells(control, experiments)) {
        const spec = specOf(ref);
        expect(spec, `Lab ${lab.n}: ${key(ref)}`).toBeDefined();
        expect('f' in spec!, `Lab ${lab.n}: ${key(ref)} is a formula`).toBe(false);
      }
    }
  });

  it('reach every experiment change of their lab, so experiments can be driven by controls', () => {
    for (const lab of labs.filter((l) => l.controls.length > 0)) {
      const covered = new Set(labControlCells(lab, experiments).map(key));
      for (const e of experiments.filter((x) => x.lab === lab.n)) {
        for (const cell of Object.keys(e.changes)) expect(covered, `${e.id}`).toContain(cell);
      }
    }
  });

  const onGrid = (v: number, min: number, step: number) => {
    const steps = (v - min) / step;
    return Math.abs(steps - Math.round(steps)) < 1e-9;
  };

  it('include the base value and every experiment value in their ranges', () => {
    for (const { lab, control } of allControls) {
      const cells = controlCells(control, experiments);
      const values = cells.flatMap((ref) => [baseOf(ref), ...experimentValues(ref)]);
      const where = `Lab ${lab.n}: ${cells.map(key).join(', ')}`;
      const c = control as Control;
      switch (c.type) {
        case 'slider':
          for (const v of values) {
            expect(typeof v, where).toBe('number');
            expect(v as number, where).toBeGreaterThanOrEqual(c.min);
            expect(v as number, where).toBeLessThanOrEqual(c.max);
            expect(onGrid(v as number, c.min, c.step), `${where}: ${v} is between steps`).toBe(
              true,
            );
          }
          break;
        case 'number':
          for (const v of values) {
            expect(typeof v, where).toBe('number');
            if (c.min !== undefined) expect(v as number, where).toBeGreaterThanOrEqual(c.min);
            if (c.max !== undefined) expect(v as number, where).toBeLessThanOrEqual(c.max);
          }
          break;
        case 'select':
          for (const v of values) expect(c.options, where).toContain(v);
          break;
        case 'segmented':
          for (const v of values)
            expect(
              c.options.map((o) => o.value),
              where,
            ).toContain(v);
          break;
        case 'date':
          for (const v of values) {
            expect(v as number, where).toBeGreaterThanOrEqual(isoDateToSerial(c.min)!);
            expect(v as number, where).toBeLessThanOrEqual(isoDateToSerial(c.max)!);
          }
          break;
        case 'table':
          expect(c.columns.length, where).toBe(
            rangeAddresses(...(c.range.split('!')[1]!.split(':') as [string, string]))[0]!.length,
          );
          for (const v of values) expect(typeof v, where).toBe('number');
          break;
        case 'experiment':
          break;
      }
    }
  });

  it('never let churn reach 0 (the closed form in Drivers!B21:D21 divides by it)', () => {
    const churn = allControls.filter(({ control }) =>
      controlCells(control, experiments).some((r) => r.sheet === 'Drivers' && r.cell.endsWith('5')),
    );
    expect(churn).toHaveLength(3);
    for (const { control } of churn) {
      expect(control.type).toBe('slider');
      if (control.type === 'slider') expect(control.min).toBeGreaterThan(0);
    }
  });

  it('take every label from a workbook cell that holds text', () => {
    const labelCells: CellRef[] = [];
    for (const lab of labs) {
      for (const group of lab.controls) {
        const title = groupTitleCell(group);
        if (title) labelCells.push(title);
        for (const c of group.controls) {
          if (c.type === 'table') {
            const { rows, cols } = tableLabels(c);
            labelCells.push(...rows, ...cols);
          } else if (c.type !== 'experiment') {
            labelCells.push(controlLabelCell(c));
          }
        }
      }
    }
    for (const ref of labelCells) {
      const spec = specOf(ref);
      expect(spec && 'v' in spec && spec.v !== '', `${key(ref)} has no label`).toBe(true);
    }
  });

  it('name existing experiments for action buttons (WI-3: Lab 9 injects E9.1)', () => {
    const actions = allControls.filter(({ control }) => control.type === 'experiment');
    expect(actions.map(({ lab }) => lab.n)).toEqual([9]);
    for (const { control } of actions) {
      expect(controlCells(control, experiments).map(key)).toEqual(
        Object.keys(
          experiments.find((e) => e.id === (control as { experiment: string }).experiment)!.changes,
        ),
      );
    }
  });

  it('list key outputs from the lab’s experiments, each cell once (WI-2)', () => {
    const lab2 = keyOutputs(2, experiments).map(key);
    const expected = [
      ...new Set(experiments.filter((e) => e.lab === 2).flatMap((e) => Object.keys(e.expect))),
    ];
    expect(lab2).toEqual(expected);
    expect(keyOutputs(6, experiments)).toEqual([]);
  });
});

describe('conditional formats (GR-12)', () => {
  it('shifts relative references and leaves $-locked parts and strings alone', () => {
    expect(shiftFormula('=$G6="Unfavorable"', 3, 2)).toBe('=$G9="Unfavorable"');
    expect(shiftFormula('=$B$19=0', 5, 5)).toBe('=$B$19=0');
    expect(shiftFormula('=A1+B$2+$C3+GL!D4', 1, 1)).toBe('=B2+C$2+$C4+GL!E5');
    expect(shiftFormula('=E2="G6 A1"', 1, 0)).toBe('=E3="G6 A1"');
    expect(shiftFormula('=LOG10(A1)', 1, 0)).toBe('=LOG10(A2)');
    expect(shiftFormula('=A1', -1, 0)).toBe('=#REF!');
  });

  it('only reference sheets in the workbook', () => {
    for (const lab of labs) {
      for (const rule of lab.conditionalFormats) {
        expect(workbook.sheetOrder).toContain(rule.range.split('!')[0]);
      }
    }
  });

  const store = new WorkbookStore(createEngine(workbook), workbook);
  const rules = labs.flatMap((l) => l.conditionalFormats);

  it('Lab 2: rows with "Unfavorable" in column G are red; rows with TRUE in H are bold', () => {
    const formats = sheetFormats(rules, 'BvA', store);
    for (let row = 6; row <= 12; row++) {
      const unfavorable = store.getValue({ sheet: 'BvA', cell: `G${row}` }) === 'Unfavorable';
      const flagged = store.getValue({ sheet: 'BvA', cell: `H${row}` }) === true;
      for (const col of 'ABCDEFGH') {
        const f = formats.get(`${col}${row}`) ?? {};
        expect(f.fill, `BvA!${col}${row}`).toBe(unfavorable ? 'redFill' : undefined);
        expect(f.bold, `BvA!${col}${row}`).toBe(flagged ? true : undefined);
      }
    }
    expect(formats.has('A5')).toBe(false);
  });

  it('Lab 7: Book rows follow their status; BankRec!B20 is green only while B19 is 0', () => {
    const book = sheetFormats(rules, 'Book', store);
    for (let row = 2; row <= 6; row++) {
      const status = store.getValue({ sheet: 'Book', cell: `E${row}` });
      const fill =
        status === 'Unmatched' ? 'orangeFill' : status === 'Tolerance' ? 'yellowFill' : undefined;
      expect(book.get(`A${row}`)?.fill, `Book!A${row}`).toBe(fill);
    }
    expect(sheetFormats(rules, 'BankRec', store).get('B20')?.fill).toBe('greenFill');
    const e72 = experiments.find((e) => e.id === 'E7.2')!;
    const changed = new WorkbookStore(createEngine(workbook), workbook);
    for (const [ref, v] of Object.entries(e72.changes)) changed.enter(parseRef(ref), v);
    expect(sheetFormats(rules, 'BankRec', changed).get('B20')?.fill).toBe('redFill');
  });

  it('Lab 8: the color scale runs from the fewest to the most days past due', () => {
    const formats = sheetFormats(rules, 'AR', store);
    const days = rangeAddresses('E2', 'E7')
      .flat()
      .map((cell) => ({
        cell,
        v: store.getValue({ sheet: 'AR', cell }) as number,
      }));
    const lo = days.reduce((a, b) => (b.v < a.v ? b : a));
    const hi = days.reduce((a, b) => (b.v > a.v ? b : a));
    expect(formats.get(lo.cell)?.scale).toBe(0);
    expect(formats.get(hi.cell)?.scale).toBe(1);
    for (const { cell } of days) expect(formats.get(cell)?.scale).toBeGreaterThanOrEqual(0);
  });

  it('is case-insensitive about sheet names and ignores other sheets', () => {
    expect(sheetFormats(rules, 'bva', store).size).toBeGreaterThan(0);
    expect(sheetFormats(rules, 'GL', store).size).toBe(0);
  });
});
