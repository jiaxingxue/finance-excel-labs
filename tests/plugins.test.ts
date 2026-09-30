// Unit tests for the gap-filling plugins (PRD §7.4), on small standalone sheets. The lab formulas
// that use them are covered by tests/conformance.test.ts; these pin down Excel's edge cases.

import { describe, expect, it } from 'vitest';
import type { SheetSpec } from '../src/content/types.ts';
import { PLUGIN_FUNCTIONS, createEngine, type CellValue } from '../src/engine/index.ts';

/**
 * Loads `data` (rows of constants starting at A1) and evaluates each formula in column Z.
 * Returns values, with errors reduced to their code (e.g. "#N/A").
 */
function run(data: (number | string | boolean | null)[][], ...formulas: string[]): unknown[] {
  const cells: SheetSpec = {};
  data.forEach((row, r) =>
    row.forEach((v, c) => {
      if (v !== null) cells[`${String.fromCharCode(65 + c)}${r + 1}`] = { v };
    }),
  );
  formulas.forEach((f, i) => (cells[`Z${i + 1}`] = { f }));
  const engine = createEngine({ sheetOrder: ['S'], sheets: { S: cells } });
  try {
    return formulas.map((_, i) => simplify(engine.getValue({ sheet: 'S', cell: `Z${i + 1}` })));
  } finally {
    engine.destroy();
  }
}

const simplify = (v: CellValue) => (typeof v === 'object' && v !== null ? v.error : v);

describe('plugin registry', () => {
  it('lists every function a plugin provides', () => {
    expect(PLUGIN_FUNCTIONS).toEqual([
      'AVERAGEIFS',
      'FORECAST',
      'FORECAST.LINEAR',
      'INDEX',
      'INTERCEPT',
      'LOOKUP',
      'TEXT',
      'TREND',
    ]);
  });
});

describe('AVERAGEIFS', () => {
  // A: region, B: amount, C: month
  const data = [
    ['East', 10, 1],
    ['West', 20, 1],
    ['East', 30, 2],
    ['east', 'n/a', 2],
    ['Eastern', 50, 3],
  ];

  it('averages numeric cells where every criterion matches', () => {
    expect(run(data, '=AVERAGEIFS(B1:B5,A1:A5,"East")')).toEqual([20]);
    expect(run(data, '=AVERAGEIFS(B1:B5,A1:A5,"East",C1:C5,">1")')).toEqual([30]);
  });

  it('supports comparison operators and concatenated criteria', () => {
    expect(
      run(
        data,
        '=AVERAGEIFS(B1:B5,C1:C5,"<=2")',
        '=AVERAGEIFS(B1:B5,C1:C5,"<>1")',
        '=AVERAGEIFS(B1:B5,C1:C5,"<="&1)',
      ),
    ).toEqual([20, 40, 15]);
  });

  it('supports wildcards and matches text case-insensitively', () => {
    expect(
      run(data, '=AVERAGEIFS(B1:B5,A1:A5,"East*")', '=AVERAGEIFS(B1:B5,A1:A5,"?ast")'),
    ).toEqual([30, 20]);
  });

  it('returns #DIV/0! when nothing numeric matches', () => {
    expect(run(data, '=AVERAGEIFS(B1:B5,A1:A5,"North")')).toEqual(['#DIV/0!']);
  });

  it('returns #VALUE! when range sizes differ', () => {
    expect(run(data, '=AVERAGEIFS(B1:B5,A1:A4,"East")')).toEqual(['#VALUE!']);
  });
});

describe('LOOKUP', () => {
  it('vector form: last value <= lookup in an ascending vector', () => {
    const data = [
      [1, 'one'],
      [5, 'five'],
      [5, 'five-b'],
      [9, 'nine'],
    ];
    expect(
      run(
        data,
        '=LOOKUP(5,A1:A4,B1:B4)',
        '=LOOKUP(7,A1:A4,B1:B4)',
        '=LOOKUP(100,A1:A4,B1:B4)',
        '=LOOKUP(0,A1:A4,B1:B4)',
      ),
    ).toEqual(['five-b', 'five-b', 'nine', '#N/A']);
  });

  it('skips errors in an array expression (last-match idiom)', () => {
    const data = [
      ['x', 'r1'],
      ['y', 'r2'],
      ['x', 'r3'],
      ['z', 'r4'],
    ];
    expect(run(data, '=LOOKUP(2,1/(A1:A4="x"),B1:B4)')).toEqual(['r3']);
  });

  it('compares text case-insensitively and ignores other types', () => {
    const data = [['apple'], [3], ['Cherry'], [true]];
    expect(run(data, '=LOOKUP("banana",A1:A4)', '=LOOKUP("zzz",A1:A4)')).toEqual([
      'apple',
      'Cherry',
    ]);
  });

  it('array form searches the first row or column of a 2-D array', () => {
    expect(run([], '=LOOKUP(2,{1,2,3;"a","b","c"})', '=LOOKUP(2.5,{1,"a";2,"b";3,"c"})')).toEqual([
      'b',
      'b',
    ]);
  });

  it('propagates an error lookup value', () => {
    expect(run([], '=LOOKUP(1/0,{1,2},{"a","b"})')).toEqual(['#DIV/0!']);
  });
});

describe('INTERCEPT, FORECAST.LINEAR, FORECAST, TREND', () => {
  // y = 3x + 2, exactly; one row has text in y (the pair is skipped).
  const data = [
    [1, 5],
    [2, 8],
    [3, 11],
    [4, 'n/a'],
    [5, 17],
  ];

  it('computes the least-squares line', () => {
    expect(
      run(
        data,
        '=INTERCEPT(B1:B5,A1:A5)',
        '=FORECAST.LINEAR(10,B1:B5,A1:A5)',
        '=FORECAST(10,B1:B5,A1:A5)',
        '=TREND(B1:B5,A1:A5,10)',
        '=TREND(B1:B5,A1:A5,A2)',
      ),
    ).toEqual([2, 32, 32, 32, 8]);
  });

  it('agrees with the built-in SLOPE on noisy data', () => {
    const noisy = [
      [1, 2.1],
      [2, 3.9],
      [3, 6.2],
      [4, 7.8],
      [5, 10.1],
    ];
    const [intercept, forecast, slope] = run(
      noisy,
      '=INTERCEPT(B1:B5,A1:A5)',
      '=FORECAST.LINEAR(6,B1:B5,A1:A5)',
      '=SLOPE(B1:B5,A1:A5)',
    ) as number[];
    expect(forecast).toBeCloseTo(intercept! + slope! * 6, 12);
    // Mean point lies on the line: mean(y) = intercept + slope * mean(x).
    expect(intercept! + slope! * 3).toBeCloseTo((2.1 + 3.9 + 6.2 + 7.8 + 10.1) / 5, 12);
  });

  it('TREND can fit through the origin (const = FALSE)', () => {
    expect(run([], '=TREND({2;4.5;6},{1;2;3},5,FALSE)', '=TREND({3;5;7},{1;2;3},0,TRUE)')).toEqual([
      (5 * (2 + 9 + 18)) / 14,
      1,
    ]);
  });

  it('returns Excel errors for bad input', () => {
    expect(
      run(
        data,
        '=INTERCEPT(B1:B5,A1:A4)',
        '=INTERCEPT({1;1},{2;2})',
        '=INTERCEPT(B1,A1)',
        '=TREND(B1:B5,A1:A4,1)',
        '=TREND(B1:B5,A1:A5,A1:A2)',
        '=INTERCEPT({1;2},{1;1/0})',
      ),
    ).toEqual(['#N/A', '#DIV/0!', '#DIV/0!', '#REF!', '#VALUE!', '#DIV/0!']);
  });
});

describe('TEXT', () => {
  it('applies Excel number formats', () => {
    expect(
      run(
        [],
        '=TEXT(1234567.891,"#,##0.00")',
        '=TEXT(-1234.5,"#,##0;(#,##0)")',
        '=TEXT(0.256,"0.0%")',
        '=TEXT(DATE(2026,8,31),"yyyy-mm-dd")',
        '=TEXT(DATE(2026,8,31),"mmm-yy")',
        '=TEXT(5,"000")',
      ),
    ).toEqual(['1,234,567.89', '(1,235)', '25.6%', '2026-08-31', 'Aug-26', '005']);
  });

  it('handles text, blanks and booleans like Excel', () => {
    expect(
      run([['12.5', null, true]], '=TEXT(A1,"0.00")', '=TEXT(B1,"0.00")', '=TEXT(C1,"0")'),
    ).toEqual(['12.50', '0.00', 'TRUE']);
  });
});

describe('INDEX', () => {
  const data = [
    ['a', 'b', 'c'],
    ['d', 'e', 'f'],
  ];

  it('indexes a one-row range by column when column_num is omitted', () => {
    expect(run(data, '=INDEX(A1:C1,2)', '=INDEX(A1:C1,1,3)')).toEqual(['b', 'c']);
  });

  it('keeps row/column indexing for columns and 2-D ranges', () => {
    expect(run(data, '=INDEX(A1:A2,2)', '=INDEX(A1:C2,2,3)', '=INDEX(A1:C2,2)')).toEqual([
      'd',
      'f',
      'd',
    ]);
  });

  it('returns #REF! outside the range and #VALUE! below 1', () => {
    expect(run(data, '=INDEX(A1:C1,4)', '=INDEX(A1:C2,3,1)', '=INDEX(A1:C2,0,1)')).toEqual([
      '#REF!',
      '#REF!',
      '#VALUE!',
    ]);
  });

  it('works with MATCH and inline arrays', () => {
    expect(run(data, '=INDEX(A1:C2,MATCH("d",A1:A2,0),2)', '=INDEX({10,20,30},3)')).toEqual([
      'e',
      30,
    ]);
  });
});

describe('engine configuration', () => {
  it('treats bare TRUE and FALSE as literals', () => {
    expect(
      run([[true], [false], [false]], '=COUNTIF(A1:A3,FALSE)', '=AND(TRUE,NOT(FALSE))'),
    ).toEqual([2, true]);
  });

  it('keeps full floating-point precision (no smart rounding)', () => {
    expect(run([], '=1/3', '=0.1+0.2')).toEqual([1 / 3, 0.1 + 0.2]);
  });

  it('shows 0 for a reference to a blank cell, as Excel does', () => {
    expect(run([], '=A1', '=A1&""')).toEqual([0, '']);
  });
});
