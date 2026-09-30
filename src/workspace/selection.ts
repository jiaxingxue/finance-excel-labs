// Grid selection and keyboard movement (GR-13). Pure functions, so the Excel jump rules are testable.

import { parseA1, toA1, type GridPosition } from '../engine/address.ts';
import type { RangeRef } from '../engine/types.ts';

/** `anchor` is the active cell; `focus` is the far corner of the selected range. */
export interface Selection {
  sheet: string;
  anchor: GridPosition;
  focus: GridPosition;
}

export interface Bounds {
  rows: number;
  cols: number;
}

export type Direction = 'up' | 'down' | 'left' | 'right';

const STEP: Record<Direction, GridPosition> = {
  up: { row: -1, col: 0 },
  down: { row: 1, col: 0 },
  left: { row: 0, col: -1 },
  right: { row: 0, col: 1 },
};

export function cellSelection(sheet: string, cell: string): Selection {
  const pos = parseA1(cell);
  return { sheet, anchor: pos, focus: pos };
}

export function rangeSelection({ sheet, start, end }: RangeRef): Selection {
  return { sheet, anchor: parseA1(start), focus: parseA1(end) };
}

export function activeCell(sel: Selection): string {
  return toA1(sel.anchor);
}

/** Top-left and bottom-right corners, whichever way the range was drawn. */
export function normalize(sel: Selection): {
  top: number;
  left: number;
  bottom: number;
  right: number;
} {
  return {
    top: Math.min(sel.anchor.row, sel.focus.row),
    left: Math.min(sel.anchor.col, sel.focus.col),
    bottom: Math.max(sel.anchor.row, sel.focus.row),
    right: Math.max(sel.anchor.col, sel.focus.col),
  };
}

export function selectionLabel(sel: Selection): string {
  const { top, left, bottom, right } = normalize(sel);
  const start = toA1({ row: top, col: left });
  const end = toA1({ row: bottom, col: right });
  return start === end ? `${sel.sheet}!${start}` : `${sel.sheet}!${start}:${end}`;
}

function clamp(pos: GridPosition, bounds: Bounds): GridPosition {
  return {
    row: Math.min(Math.max(pos.row, 0), bounds.rows - 1),
    col: Math.min(Math.max(pos.col, 0), bounds.cols - 1),
  };
}

/** One step, or `by` steps (Page Up/Down), stopping at the edge. */
export function step(pos: GridPosition, dir: Direction, bounds: Bounds, by = 1): GridPosition {
  const d = STEP[dir];
  return clamp({ row: pos.row + d.row * by, col: pos.col + d.col * by }, bounds);
}

/**
 * Ctrl+arrow, as in Excel: inside a block of filled cells, go to the block's last cell; otherwise
 * go to the next filled cell; with none ahead, go to the edge.
 */
export function jump(
  pos: GridPosition,
  dir: Direction,
  bounds: Bounds,
  isFilled: (pos: GridPosition) => boolean,
): GridPosition {
  const d = STEP[dir];
  const inside = (p: GridPosition) =>
    p.row >= 0 && p.col >= 0 && p.row < bounds.rows && p.col < bounds.cols;
  const next = (p: GridPosition) => ({ row: p.row + d.row, col: p.col + d.col });

  let cur = pos;
  let ahead = next(cur);
  if (!inside(ahead)) return cur;
  if (isFilled(cur) && isFilled(ahead)) {
    while (inside(next(ahead)) && isFilled(next(ahead))) ahead = next(ahead);
    return ahead;
  }
  cur = ahead;
  while (!isFilled(cur)) {
    const n = next(cur);
    if (!inside(n)) return cur;
    cur = n;
  }
  return cur;
}
