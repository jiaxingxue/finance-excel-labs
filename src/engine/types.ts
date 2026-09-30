// The engine adapter contract (PRD §7.3). The app and the grader talk to this interface only,
// so HyperFormula can be swapped for another engine that passes the same conformance suite.

import type { WorkbookSpec } from '../content/types.ts';

/** One cell, e.g. `{ sheet: 'BvA', cell: 'C6' }` (same field names as an Assertion). */
export interface CellRef {
  sheet: string;
  cell: string;
}

/** A rectangular range on one sheet; `start === end` for a single cell. */
export interface RangeRef {
  sheet: string;
  start: string;
  end: string;
}

/** An Excel error such as `#DIV/0!` or `#N/A`. */
export interface ErrorValue {
  error: string;
  message?: string;
}

/** `null` is an empty cell. */
export type CellValue = number | string | boolean | ErrorValue | null;

export interface CellEntry {
  ref: CellRef;
  content: string | number | boolean | null;
}

export interface ChangedCell {
  ref: CellRef;
  value: CellValue;
}

export interface FormulaEngine {
  load(workbook: WorkbookSpec): void;
  sheetNames(): string[];
  /**
   * `=…` is a formula; any other string is literal text; `null` or `""` clears the cell.
   * Parsing typed input ("1,000", "5%", dates) is the caller's job.
   * Throws if a formula refers to a sheet that doesn't exist.
   */
  setCell(ref: CellRef, content: string | number | boolean | null): ChangedCell[];
  /** Several `setCell`s with one recalculation and one undo step. All or nothing. */
  setCells(entries: CellEntry[]): ChangedCell[];
  getValue(ref: CellRef): CellValue;
  getFormula(ref: CellRef): string | null;
  /**
   * Evaluates a formula that isn't stored in any cell, on the given sheet, without changing the
   * workbook (used for conditional-format rules, PRD §9.3). Relative references are read as-is.
   * A formula that can't be parsed returns an error value instead of throwing.
   */
  evaluate(formula: string, sheet: string): CellValue;
  /** Excel "fill": repeats `source` across `target`, adjusting relative references. One undo step. */
  fill(source: RangeRef, target: RangeRef): ChangedCell[];
  /** Cells and ranges referenced directly by the cell's formula. */
  precedents(ref: CellRef): RangeRef[];
  /** Formula cells that reference this cell directly (including through a range). */
  dependents(ref: CellRef): CellRef[];
  /** An independent copy of the current state (no undo history), for hidden re-grading (BD-4). */
  clone(): FormulaEngine;
  undo(): void;
  redo(): void;
  /** Current state in the `workbook.json` shape. Formats and date constants are preserved. */
  exportCells(): WorkbookSpec;
  /** Releases the engine's memory. The instance can't be used afterwards. */
  destroy(): void;
}

export function isErrorValue(value: CellValue): value is ErrorValue {
  return typeof value === 'object' && value !== null;
}
