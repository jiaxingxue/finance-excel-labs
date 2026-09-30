// How cell values are shown (PRD §4.3, GR-2, OPEN_ISSUES #19). Display only: the engine and the
// grader always work with the full-precision value.

import SSF from 'ssf';
// Imports address.ts and types.ts directly, not engine/index.ts, so HyperFormula stays in the
// lazily loaded engine chunk.
import { serialToIsoDate } from '../engine/address.ts';
import { isErrorValue, type CellValue } from '../engine/types.ts';

/**
 * The grid text for a value. Numbers use the cell's Excel format code through SSF; a cell with no
 * `fmt` uses "General", as Excel does, so floating-point noise such as 0.30000000000000004 shows
 * as 0.3. Booleans show as TRUE/FALSE and errors as their Excel error text.
 */
export function formatCell(value: CellValue, fmt?: string): string {
  if (value === null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (isErrorValue(value)) return value.error;
  if (typeof value === 'string') return value;
  try {
    return SSF.format(fmt ?? 'General', value);
  } catch {
    return SSF.format('General', value);
  }
}

/**
 * A number rounded to 15 significant digits, Excel's own precision limit: 0.9999999999999999 → "1".
 * Used by the formula bar and the inspector.
 */
export function toSignificant15(n: number): string {
  return String(Number(n.toPrecision(15)));
}

/** True when an Excel format code displays a date (e.g. "yyyy-mm-dd", "mmm-yy"). */
export function isDateFormat(fmt: string | undefined): boolean {
  if (!fmt) return false;
  try {
    return SSF.is_date(fmt);
  } catch {
    return false;
  }
}

/**
 * What the formula bar shows for a cell (GR-3): the formula if there is one, otherwise the constant.
 * Numbers use 15 significant digits; date-formatted numbers show as YYYY-MM-DD, the form they
 * were entered in and the form the input parser accepts back.
 */
export function formulaBarText(value: CellValue, formula: string | null, fmt?: string): string {
  if (formula !== null) return formula;
  if (typeof value === 'number') {
    return isDateFormat(fmt) ? serialToIsoDate(value) : toSignificant15(value);
  }
  return formatCell(value);
}
