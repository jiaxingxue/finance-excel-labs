// The M2 display rule (OPEN_ISSUES #19): cells without `fmt` render with SSF "General"; the
// formula bar and inspector show 15 significant digits; the engine and grader never round.

import { describe, expect, it } from 'vitest';
import assertions from '../src/data/assertions.json' with { type: 'json' };
import workbook from '../src/data/workbook.json' with { type: 'json' };
import type { Assertion, WorkbookSpec } from '../src/content/types.ts';
import { createEngine, isoDateToSerial, type CellValue } from '../src/engine/index.ts';
import { formatCell, formulaBarText, toSignificant15 } from '../src/format/displayValue.ts';
import { assertionPasses } from '../src/grading/match.ts';

const spec = workbook as WorkbookSpec;

describe('formatCell: General for cells without fmt', () => {
  it.each([
    [0.30000000000000004, '0.3'],
    [0.9999999999999999, '1'],
    [1 / 3, '0.333333333'],
    [46265, '46265'],
    [-32000, '-32000'],
  ])('%s → %s', (value, text) => {
    expect(formatCell(value)).toBe(text);
  });

  it('uses the cell format code through SSF', () => {
    expect(formatCell(-32000, '#,##0;(#,##0)')).toBe('(32,000)');
    expect(formatCell(-0.064, '0.0%')).toBe('-6.4%');
    expect(formatCell(isoDateToSerial('2026-08-31')!, 'yyyy-mm-dd')).toBe('2026-08-31');
  });

  it('shows booleans, text, blanks, and errors as Excel does (GR-2)', () => {
    expect(formatCell(true)).toBe('TRUE');
    expect(formatCell(false, '0.0%')).toBe('FALSE');
    expect(formatCell('n/a', '0.0%')).toBe('n/a');
    expect(formatCell(null)).toBe('');
    expect(formatCell({ error: '#DIV/0!' })).toBe('#DIV/0!');
  });
});

describe('15 significant digits in the formula bar and inspector', () => {
  it.each([
    [0.30000000000000004, '0.3'],
    [0.9999999999999999, '1'],
    [691158.827474712, '691158.827474712'],
    [1 / 3, '0.333333333333333'],
    [-32000, '-32000'],
  ])('%s → %s', (value, text) => {
    expect(toSignificant15(value)).toBe(text);
  });

  it('shows formulas as written and constants as entered', () => {
    expect(formulaBarText(500000, '=SUM(A1:A2)')).toBe('=SUM(A1:A2)');
    expect(formulaBarText(0.1 + 0.2, null)).toBe('0.3');
    expect(formulaBarText(isoDateToSerial('2026-09-30')!, null, 'yyyy-mm-dd')).toBe('2026-09-30');
    expect(formulaBarText('2026-08', null)).toBe('2026-08');
    expect(formulaBarText(null, null)).toBe('');
  });
});

describe('the display rule on the real workbook', () => {
  const engine = createEngine(spec);

  it('Checks!C8 holds floating-point noise that display hides but grading keeps', () => {
    const raw = engine.getValue({ sheet: 'Checks', cell: 'C8' });
    expect(spec.sheets.Checks!.C8).not.toHaveProperty('fmt');
    expect(raw).not.toBe(1); // the engine does not round…
    expect(formatCell(raw)).toBe('1'); // …the grid shows General…
    expect(toSignificant15(raw as number)).toBe('1'); // …and the inspector 15 digits
  });

  it('grades the full-precision value, not the displayed text', () => {
    // Every assertion still passes on the raw engine value; the display layer is never consulted.
    const failures = (assertions as Assertion[]).filter(
      (a) => !assertionPasses(engine.getValue(a), a),
    );
    expect(failures).toEqual([]);
  });

  it('formats every workbook value without throwing', () => {
    for (const [sheet, cells] of Object.entries(spec.sheets)) {
      for (const [cell, c] of Object.entries(cells)) {
        const value: CellValue = engine.getValue({ sheet, cell });
        expect(() => formatCell(value, c.fmt), `${sheet}!${cell}`).not.toThrow();
      }
    }
  });
});
