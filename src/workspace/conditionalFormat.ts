// Display-only conditional formatting (GR-12). A rule's formula is written for the top-left cell
// of its range; for every other cell its relative references move by the same offset, as in
// Excel. The engine evaluates each shifted formula (PRD §9.3); nothing here interprets formulas.

import type { ConditionalFormat, FormatStyle } from '../config/labs.config.ts';
import { parseA1, rangeAddresses, toA1 } from '../engine/address.ts';
import type { CellValue } from '../engine/types.ts';
import { parseRange } from '../lesson/cellRefs.ts';

export interface CellFormat {
  fill?: Exclude<FormatStyle, 'bold'>;
  bold?: boolean;
  /** Position on a two-color scale, 0 (lowest) to 1 (highest). */
  scale?: number;
}

const STRING_LITERAL = /"(?:[^"]|"")*"/g;
// A cell reference not glued to a name or function call: `$G6`, `B$19`, `GL!A2`.
const CELL_REF = /(?<![\w.$])(\$?)([A-Z]{1,3})(\$?)([1-9][0-9]*)(?![\w(!])/g;

/** Moves the relative parts of every cell reference, leaving `$`-locked parts and strings alone. */
export function shiftFormula(formula: string, dRow: number, dCol: number): string {
  let out = '';
  let last = 0;
  for (const m of formula.matchAll(STRING_LITERAL)) {
    out += shiftRefs(formula.slice(last, m.index), dRow, dCol) + m[0];
    last = m.index + m[0].length;
  }
  return out + shiftRefs(formula.slice(last), dRow, dCol);
}

function shiftRefs(code: string, dRow: number, dCol: number): string {
  return code.replace(
    CELL_REF,
    (_whole, colLock: string, letters: string, rowLock: string, digits: string) => {
      const pos = parseA1(`${letters}${digits}`);
      const moved = {
        row: rowLock ? pos.row : pos.row + dRow,
        col: colLock ? pos.col : pos.col + dCol,
      };
      if (moved.row < 0 || moved.col < 0) return '#REF!';
      const [newLetters, newDigits] = /^([A-Z]+)(\d+)$/.exec(toA1(moved))!.slice(1) as [
        string,
        string,
      ];
      return `${colLock}${newLetters}${rowLock}${newDigits}`;
    },
  );
}

/** Excel treats TRUE and any non-zero number as a match; text, blanks, and errors don't match. */
function matches(value: CellValue): boolean {
  return value === true || (typeof value === 'number' && value !== 0);
}

export interface FormatSource {
  evaluate(formula: string, sheet: string): CellValue;
  getValue(ref: { sheet: string; cell: string }): CellValue;
}

/**
 * Formats for one sheet, keyed by address. Rules apply in order; for each property the first
 * matching rule wins (Excel's rule priority), so a fill and a bold rule combine.
 */
export function sheetFormats(
  rules: readonly ConditionalFormat[],
  sheet: string,
  source: FormatSource,
): Map<string, CellFormat> {
  const formats = new Map<string, CellFormat>();
  const format = (cell: string) => {
    let f = formats.get(cell);
    if (!f) formats.set(cell, (f = {}));
    return f;
  };
  const results = new Map<string, boolean>(); // many cells in a row share one shifted formula

  for (const rule of rules) {
    const range = parseRange(rule.range);
    if (!range || range.sheet.toLowerCase() !== sheet.toLowerCase()) continue;
    const rows = rangeAddresses(range.start, range.end);

    if (rule.type === 'colorScale') {
      const numbers = rows.flat().flatMap((cell) => {
        const v = source.getValue({ sheet: range.sheet, cell });
        return typeof v === 'number' ? [[cell, v] as const] : [];
      });
      const values = numbers.map(([, v]) => v);
      const [lo, hi] = [Math.min(...values), Math.max(...values)];
      for (const [cell, v] of numbers) {
        const f = format(cell);
        f.scale ??= hi === lo ? 0 : (v - lo) / (hi - lo);
      }
      continue;
    }

    const origin = parseA1(range.start);
    for (const cell of rows.flat()) {
      const pos = parseA1(cell);
      const formula = shiftFormula(rule.formula, pos.row - origin.row, pos.col - origin.col);
      let hit = results.get(formula);
      if (hit === undefined) {
        hit = matches(source.evaluate(formula, range.sheet));
        results.set(formula, hit);
      }
      if (!hit) continue;
      const f = format(cell);
      if (rule.style === 'bold') f.bold ??= true;
      else f.fill ??= rule.style;
    }
  }
  return formats;
}
