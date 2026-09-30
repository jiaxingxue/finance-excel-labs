// Engine adapter (PRD §7.3 interface), address/date helpers, and grading comparison rules.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import type { Assertion, WorkbookSpec } from '../src/content/types.ts';
import { parseA1, referencedSheets, toA1 } from '../src/engine/address.ts';
import {
  createEngine,
  isoDateToSerial,
  parseRef,
  serialToIsoDate,
  type FormulaEngine,
} from '../src/engine/index.ts';
import {
  assertionPasses,
  experimentChangeContent,
  experimentValuePasses,
} from '../src/grading/match.ts';

const workbook = JSON.parse(
  readFileSync(join(import.meta.dirname, '../src/data/workbook.json'), 'utf8'),
) as WorkbookSpec;

/** A1:A3 inputs, B1 =A1*2 (B2:B3 blank until filled), C1 =SUM(B1:B3), T!A1 =S!C1+1. */
function small(): FormulaEngine {
  return createEngine({
    sheetOrder: ['S', 'T'],
    sheets: {
      S: {
        A1: { v: 1 },
        A2: { v: 2 },
        A3: { v: 3 },
        B1: { f: '=A1*2', fmt: '0.00' },
        C1: { f: '=SUM(B1:B3)' },
        D1: { v: 'label' },
      },
      T: { A1: { f: '=S!C1+1' } },
    },
  });
}

describe('loading and exporting', () => {
  const engine = createEngine(workbook);

  it('round-trips workbook.json exactly (formulas, text, dates, formats)', () => {
    expect(engine.sheetNames()).toEqual(workbook.sheetOrder);
    expect(engine.exportCells()).toEqual(workbook);
  });

  it('loads {date} constants as 1900-system serials', () => {
    for (const [sheet, cells] of Object.entries(workbook.sheets)) {
      for (const [cell, spec] of Object.entries(cells)) {
        if ('v' in spec && typeof spec.v === 'object') {
          expect(engine.getValue({ sheet, cell })).toBe(isoDateToSerial(spec.v.date));
        }
      }
    }
  });

  it('keeps text constants as text, even when they look like dates or numbers', () => {
    for (const [sheet, cells] of Object.entries(workbook.sheets)) {
      for (const [cell, spec] of Object.entries(cells)) {
        if ('v' in spec && typeof spec.v === 'string') {
          expect(engine.getValue({ sheet, cell })).toBe(spec.v);
        }
      }
    }
  });

  it('returns formulas verbatim and null for constants', () => {
    const [sheet, cells] = Object.entries(workbook.sheets)[0]!;
    for (const [cell, spec] of Object.entries(cells)) {
      expect(engine.getFormula({ sheet, cell })).toBe('f' in spec ? spec.f : null);
    }
  });

  it('rejects unknown sheets and invalid addresses', () => {
    expect(() => engine.getValue({ sheet: 'Nope', cell: 'A1' })).toThrow(/Unknown sheet/);
    expect(() => engine.getValue({ sheet: 'BvA', cell: 'A0' })).toThrow(/Invalid cell address/);
  });
});

describe('editing', () => {
  it('setCell recalculates and reports every changed cell', () => {
    const engine = small();
    const changes = engine.setCell({ sheet: 'S', cell: 'A1' }, 10);
    const byRef = Object.fromEntries(changes.map((c) => [`${c.ref.sheet}!${c.ref.cell}`, c.value]));
    expect(byRef).toMatchObject({ 'S!A1': 10, 'S!B1': 20, 'S!C1': 20, 'T!A1': 21 });
  });

  it('setCell: "=" starts a formula, other strings are text, "" and null clear', () => {
    const engine = small();
    const at = { sheet: 'S', cell: 'E1' };
    engine.setCell(at, '=A2+A3');
    expect([engine.getValue(at), engine.getFormula(at)]).toEqual([5, '=A2+A3']);
    engine.setCell(at, '2026-07');
    expect(engine.getValue(at)).toBe('2026-07');
    engine.setCell(at, "'quoted");
    expect(engine.exportCells().sheets.S!.E1).toEqual({ v: "'quoted" });
    engine.setCell(at, '');
    expect(engine.getValue(at)).toBeNull();
    engine.setCell(at, true);
    expect(engine.getValue(at)).toBe(true);
    engine.setCell(at, null);
    expect(engine.exportCells().sheets.S!.E1).toBeUndefined();
  });

  it('rejects formulas that refer to unknown sheets (HyperFormula would crash on overwrite)', () => {
    const engine = small();
    const at = { sheet: 'S', cell: 'E1' };
    expect(() => engine.setCell(at, '=SUMIFS(GL!E2:E37,GL!C2:C37,1)')).toThrow(
      /unknown sheet "GL"/,
    );
    // Sheet names are case-insensitive; "!" inside text is not a reference.
    engine.setCell(at, '=s!A1+t!A1+LEN("GL!A1")');
    expect(engine.getValue(at)).toBe(1 + 3 + 5);
    engine.setCell(at, '=S!A2');
    expect(engine.getValue(at)).toBe(2);
    expect(() =>
      createEngine({ sheetOrder: ['S'], sheets: { S: { A1: { f: '=X!A1' } } } }),
    ).toThrow(/S!A1: the formula refers to an unknown sheet "X"/);
  });

  it('errors come back as ErrorValue objects', () => {
    const engine = small();
    engine.setCell({ sheet: 'S', cell: 'E1' }, '=1/0');
    expect(engine.getValue({ sheet: 'S', cell: 'E1' })).toMatchObject({ error: '#DIV/0!' });
  });

  it('undo and redo; both are no-ops with nothing to undo', () => {
    const engine = small();
    engine.undo();
    engine.redo();
    engine.setCell({ sheet: 'S', cell: 'A1' }, 10);
    engine.undo();
    expect(engine.getValue({ sheet: 'S', cell: 'C1' })).toBe(2);
    engine.redo();
    expect(engine.getValue({ sheet: 'S', cell: 'C1' })).toBe(20);
  });

  it('fill copies a formula down with relative references, as one undo step', () => {
    const engine = small();
    engine.fill({ sheet: 'S', start: 'B1', end: 'B1' }, { sheet: 'S', start: 'B1', end: 'B3' });
    expect(['B2', 'B3'].map((cell) => engine.getFormula({ sheet: 'S', cell }))).toEqual([
      '=A2*2',
      '=A3*2',
    ]);
    expect(engine.getValue({ sheet: 'S', cell: 'C1' })).toBe(12);
    engine.undo();
    expect(engine.getFormula({ sheet: 'S', cell: 'B2' })).toBeNull();
  });

  it('fill repeats a multi-cell source across a larger target', () => {
    const engine = small();
    engine.fill({ sheet: 'S', start: 'A1', end: 'A2' }, { sheet: 'S', start: 'A1', end: 'A5' });
    expect(['A3', 'A4', 'A5'].map((cell) => engine.getValue({ sheet: 'S', cell }))).toEqual([
      1, 2, 1,
    ]);
  });

  it('keeps number formats with the cell position', () => {
    const engine = small();
    engine.setCell({ sheet: 'S', cell: 'B1' }, 5);
    expect(engine.exportCells().sheets.S!.B1).toEqual({ v: 5, fmt: '0.00' });
  });
});

describe('dependency graph (EN-6)', () => {
  it('precedents lists the cells and ranges a formula reads', () => {
    const engine = small();
    expect(engine.precedents({ sheet: 'S', cell: 'C1' })).toEqual([
      { sheet: 'S', start: 'B1', end: 'B3' },
    ]);
    expect(engine.precedents({ sheet: 'T', cell: 'A1' })).toEqual([
      { sheet: 'S', start: 'C1', end: 'C1' },
    ]);
  });

  it('dependents follows references through ranges and across sheets', () => {
    const engine = small();
    expect(engine.dependents({ sheet: 'S', cell: 'B1' })).toEqual([{ sheet: 'S', cell: 'C1' }]);
    expect(engine.dependents({ sheet: 'S', cell: 'C1' })).toEqual([{ sheet: 'T', cell: 'A1' }]);
    expect(engine.dependents({ sheet: 'S', cell: 'D1' })).toEqual([]);
  });

  it('works on the lab workbook: GL amounts feed the BvA SUMIFS', () => {
    const engine = createEngine(workbook);
    const deps = engine.dependents({ sheet: 'GL', cell: 'E2' });
    expect(deps).toContainEqual({ sheet: 'BvA', cell: 'C6' });
    expect(engine.precedents({ sheet: 'BvA', cell: 'C6' })).toContainEqual({
      sheet: 'GL',
      start: 'E2',
      end: 'E37',
    });
  });
});

describe('clone', () => {
  it('copies the current state and is independent of the original', () => {
    const engine = small();
    engine.setCell({ sheet: 'S', cell: 'A1' }, 10);
    const copy = engine.clone();
    expect(copy.getValue({ sheet: 'T', cell: 'A1' })).toBe(21);
    copy.setCell({ sheet: 'S', cell: 'A2' }, 100);
    expect(copy.getValue({ sheet: 'S', cell: 'C1' })).toBe(20);
    expect(engine.getValue({ sheet: 'S', cell: 'C1' })).toBe(20);
    expect(engine.getValue({ sheet: 'S', cell: 'A2' })).toBe(2);
    copy.destroy();
    expect(() => copy.getValue({ sheet: 'S', cell: 'A1' })).toThrow(/No workbook loaded/);
  });
});

describe('addresses and dates', () => {
  it('converts A1 addresses both ways', () => {
    for (const address of ['A1', 'Z9', 'AA10', 'AZ1', 'BA2', 'ZZ100', 'AAA1']) {
      expect(toA1(parseA1(address))).toBe(address);
    }
    expect(parseA1('C6')).toEqual({ row: 5, col: 2 });
    expect(() => parseA1('a1')).toThrow();
  });

  it('finds the sheets a formula refers to', () => {
    expect(referencedSheets('=SUM(GL!$E$2:$E$37)+Map!A1+GL!B2')).toEqual(['GL', 'Map']);
    expect(referencedSheets("='My ''Q3'' Sheet'!A1+'Other'!B2")).toEqual([
      "My 'Q3' Sheet",
      'Other',
    ]);
    expect(referencedSheets('=IF(A1="x!y","Bank!A1",B1)')).toEqual([]);
  });

  it('parses "Sheet!A1" references', () => {
    expect(parseRef('Flux!I8')).toEqual({ sheet: 'Flux', cell: 'I8' });
    expect(() => parseRef('I8')).toThrow();
  });

  it('converts dates with the 1900 system and rejects anything but a full valid date', () => {
    expect(isoDateToSerial('1900-03-01')).toBe(61);
    expect(isoDateToSerial('1899-12-31')).toBe(1);
    expect(serialToIsoDate(isoDateToSerial('2026-08-31')!)).toBe('2026-08-31');
    for (const bad of ['2026-07', '2026-02-30', '2026-13-01', '26-08-31', 'x']) {
      expect(isoDateToSerial(bad)).toBeUndefined();
    }
  });
});

describe('grading comparison rules (PRD §4.4)', () => {
  const assertion = (expected: Assertion['expected'], tolerance = 0.005): Assertion => ({
    sheet: 'S',
    cell: 'A1',
    expected,
    tolerance,
  });

  it('assertion numbers use max(tolerance, |expected| * 1e-12)', () => {
    expect(assertionPasses(100.005, assertion(100))).toBe(true);
    expect(assertionPasses(100.006, assertion(100))).toBe(false);
    expect(assertionPasses(1e12 + 0.5, assertion(1e12, 1e-9))).toBe(true);
    expect(assertionPasses('100', assertion(100))).toBe(false);
  });

  it('assertion strings and booleans need strict equality; "" also accepts empty', () => {
    expect(assertionPasses('Favorable', assertion('Favorable'))).toBe(true);
    expect(assertionPasses('favorable', assertion('Favorable'))).toBe(false);
    expect(assertionPasses(1, assertion(true))).toBe(false);
    expect(assertionPasses(null, assertion(''))).toBe(true);
    expect(assertionPasses({ error: '#N/A' }, assertion(''))).toBe(false);
  });

  it('experiment numbers use 1e-6 relative or 0.005 absolute, whichever is larger', () => {
    expect(experimentValuePasses(1_000_000.9, 1_000_000)).toBe(true);
    expect(experimentValuePasses(1_000_001.1, 1_000_000)).toBe(false);
    expect(experimentValuePasses(0.004, 0)).toBe(true);
  });

  it('experiment null means empty', () => {
    expect(experimentValuePasses(null, null)).toBe(true);
    expect(experimentValuePasses('', null)).toBe(true);
    expect(experimentValuePasses(0, null)).toBe(false);
  });

  it('only a full YYYY-MM-DD change becomes a date serial', () => {
    expect(experimentChangeContent('2026-09-30')).toBe(isoDateToSerial('2026-09-30'));
    expect(experimentChangeContent('2026-07')).toBe('2026-07');
    expect(experimentChangeContent(3)).toBe(3);
  });
});
