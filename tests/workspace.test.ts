// Grid selection rules (GR-13) and the workspace store around the engine.

import { describe, expect, it } from 'vitest';
import assertions from '../src/data/assertions.json' with { type: 'json' };
import workbook from '../src/data/workbook.json' with { type: 'json' };
import type { Assertion, WorkbookSpec } from '../src/content/types.ts';
import { parseA1, toA1, type GridPosition } from '../src/engine/address.ts';
import { createEngine } from '../src/engine/index.ts';
import { assertionPasses } from '../src/grading/match.ts';
import {
  cellSelection,
  jump,
  normalize,
  rangeSelection,
  selectionLabel,
  step,
} from '../src/workspace/selection.ts';
import { WorkbookStore } from '../src/workspace/workbookStore.ts';

const spec = workbook as WorkbookSpec;

describe('selection', () => {
  const bounds = { rows: 20, cols: 8 };

  it('steps and clamps at the edges', () => {
    expect(step({ row: 0, col: 0 }, 'up', bounds)).toEqual({ row: 0, col: 0 });
    expect(step({ row: 5, col: 2 }, 'down', bounds, 10)).toEqual({ row: 15, col: 2 });
    expect(step({ row: 15, col: 7 }, 'right', bounds)).toEqual({ row: 15, col: 7 });
  });

  it('Ctrl+arrow jumps like Excel', () => {
    // Filled: A1:A3 and A6:A8 in column A.
    const filled = new Set(['A1', 'A2', 'A3', 'A6', 'A7', 'A8']);
    const isFilled = (p: GridPosition) => filled.has(toA1(p));
    const down = (cell: string) => toA1(jump(parseA1(cell), 'down', bounds, isFilled));
    expect(down('A1')).toBe('A3'); // to the end of the block
    expect(down('A3')).toBe('A6'); // to the next block
    expect(down('A4')).toBe('A6'); // from a blank cell to the next filled one
    expect(down('A8')).toBe('A20'); // nothing ahead: to the edge
    expect(toA1(jump(parseA1('A6'), 'up', bounds, isFilled))).toBe('A3');
  });

  it('labels and normalizes ranges drawn in any direction', () => {
    const sel = rangeSelection({ sheet: 'BvA', start: 'H5', end: 'A5' });
    expect(normalize(sel)).toEqual({ top: 4, left: 0, bottom: 4, right: 7 });
    expect(selectionLabel(sel)).toBe('BvA!A5:H5');
    expect(selectionLabel(cellSelection('BvA', 'C6'))).toBe('BvA!C6');
  });
});

describe('WorkbookStore', () => {
  const make = () => new WorkbookStore(createEngine(spec), spec);
  const bva = (cell: string) => ({ sheet: 'BvA', cell });

  it('recalculates, reports changed cells, and notifies subscribers', () => {
    const store = make();
    let calls = 0;
    store.subscribe(() => calls++);
    expect(store.enter(bva('B3'), 1)).toBeNull();
    expect(calls).toBe(1);
    expect(store.getVersion()).toBe(1);
    expect(store.lastChange.keys.has('BvA!B3')).toBe(true);
    expect([...store.lastChange.keys].some((k) => k.startsWith('BvA!H'))).toBe(true);
    store.destroy();
  });

  it('returns an error message instead of throwing for an unknown sheet', () => {
    const store = make();
    expect(store.enter(bva('J1'), '=SUM(Ledger!E2:E37)')).toMatch(/Ledger/);
    expect(store.getVersion()).toBe(0);
    store.destroy();
  });

  it('undo and reset restore every assertion', () => {
    const store = make();
    const bvaChecks = (assertions as Assertion[]).filter((a) => a.sheet === 'BvA');
    const failing = () => bvaChecks.filter((a) => !assertionPasses(store.getValue(a), a)).length;
    store.enter(bva('B3'), 1);
    expect(failing()).toBeGreaterThan(0);
    store.undo();
    expect(failing()).toBe(0);
    store.enter(bva('B3'), 1);
    store.enter(bva('Z40'), 'note');
    expect(store.extent('BvA')).toEqual({ rows: 40, cols: 26 });
    store.reset();
    expect(failing()).toBe(0);
    expect(store.extent('BvA')).toEqual({ rows: 14, cols: 8 });
    expect(store.lastChange.keys.size).toBe(0);
    store.destroy();
  });

  it('reads number formats from the Labs document', () => {
    const store = make();
    expect(store.fmt(bva('F6'))).toBe(spec.sheets.BvA!.F6!.fmt);
    expect(store.fmt(bva('Z99'))).toBeUndefined();
    store.destroy();
  });
});
