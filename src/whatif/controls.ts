// What-if controls (WI-1–WI-4) as data: which cells each control writes, where its label comes
// from, and the lab's key output cells. Pure functions, so tests can check the config against
// the Labs document without a browser.

import type { Experiment } from '../content/types.ts';
import type { Control, ControlGroup, LabConfig, TableControl } from '../config/labs.config.ts';
import { parseA1, parseRef, rangeAddresses, toA1 } from '../engine/address.ts';
import type { CellRef, CellValue } from '../engine/types.ts';
import { parseRange } from '../lesson/cellRefs.ts';

/** `"PVM!B5:E7"` → the table's cells, row by row. */
export function tableCells(table: TableControl): CellRef[][] {
  const range = mustParseRange(table.range);
  return rangeAddresses(range.start, range.end).map((row) =>
    row.map((cell) => ({ sheet: range.sheet, cell })),
  );
}

/** Label cells of a table's rows (and columns), on the table's sheet. */
export function tableLabels(table: TableControl): { rows: CellRef[]; cols: CellRef[] } {
  const { sheet } = mustParseRange(table.range);
  const flat = (text: string) => {
    const [start, end = start] = text.split(':') as [string, string?];
    return rangeAddresses(start, end)
      .flat()
      .map((cell) => ({ sheet, cell }));
  };
  return { rows: flat(table.rowLabels), cols: table.colLabels ? flat(table.colLabels) : [] };
}

/** Every cell a control can write. An experiment button writes its experiment's change cells. */
export function controlCells(control: Control, experiments: readonly Experiment[]): CellRef[] {
  switch (control.type) {
    case 'table':
      return tableCells(control).flat();
    case 'experiment':
      return Object.keys(findExperiment(control.experiment, experiments).changes).map(parseRef);
    default:
      return [parseRef(control.cell)];
  }
}

/** The input cells a lab's controls can change; "Reset inputs" restores exactly these (WI-4). */
export function labControlCells(lab: LabConfig, experiments: readonly Experiment[]): CellRef[] {
  const seen = new Map<string, CellRef>();
  for (const group of lab.controls) {
    for (const control of group.controls) {
      for (const ref of controlCells(control, experiments))
        seen.set(`${ref.sheet}!${ref.cell}`, ref);
    }
  }
  return [...seen.values()];
}

/** The cell whose text labels a single-cell control: `labelCell`, or column A of the same row. */
export function controlLabelCell(
  control: Exclude<Control, TableControl | { type: 'experiment' }>,
): CellRef {
  const { sheet, cell } = parseRef(control.cell);
  return { sheet, cell: control.labelCell ?? toA1({ row: parseA1(cell).row, col: 0 }) };
}

export function groupTitleCell(group: ControlGroup): CellRef | null {
  return group.titleCell ? parseRef(group.titleCell) : null;
}

/**
 * Key output cells for the delta list (WI-2): every cell the lab's experiments watch, in the
 * order the experiments list them. Taken from Appendix E, not chosen by hand.
 */
export function keyOutputs(lab: number, experiments: readonly Experiment[]): CellRef[] {
  const seen = new Map<string, CellRef>();
  for (const e of experiments) {
    if (e.lab !== lab) continue;
    for (const key of Object.keys(e.expect)) if (!seen.has(key)) seen.set(key, parseRef(key));
  }
  return [...seen.values()];
}

export function findExperiment(id: string, experiments: readonly Experiment[]): Experiment {
  const found = experiments.find((e) => e.id === id);
  if (!found) throw new Error(`Unknown experiment "${id}"`);
  return found;
}

function mustParseRange(text: string) {
  const range = parseRange(text);
  if (!range) throw new Error(`Invalid range "${text}"`);
  return range;
}

/** Same displayed value: equal primitives, or the same error. */
export function sameValue(a: CellValue, b: CellValue): boolean {
  if (typeof a === 'object' && a !== null && typeof b === 'object' && b !== null) {
    return a.error === b.error;
  }
  return a === b;
}
