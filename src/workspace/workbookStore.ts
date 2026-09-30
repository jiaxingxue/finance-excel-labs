// The lab workspace's view of the formula engine. The engine lives outside React state (PRD §9.1);
// components subscribe to a version number and read values through this store.

import type { WorkbookSpec } from '../content/types.ts';
import { parseA1 } from '../engine/address.ts';
import type { CellRef, CellValue, FormulaEngine } from '../engine/types.ts';
import type { CellInput } from './parseInput.ts';

export interface Extent {
  /** Number of rows and columns that hold content (1-based counts). */
  rows: number;
  cols: number;
}

export interface LastChange {
  /** Increases with every edit, so a cell changed twice in a row flashes twice. */
  seq: number;
  /** "Sheet!A1" keys of the cells whose value changed in the last recalculation (GR-6). */
  keys: ReadonlySet<string>;
}

const keyOf = ({ sheet, cell }: CellRef) => `${sheet}!${cell}`;

export class WorkbookStore {
  private version = 0;
  private readonly listeners = new Set<() => void>();
  private extents = new Map<string, Extent>();
  lastChange: LastChange = { seq: 0, keys: new Set() };

  private readonly engine: FormulaEngine;
  private readonly original: WorkbookSpec;

  constructor(engine: FormulaEngine, original: WorkbookSpec) {
    this.engine = engine;
    this.original = original;
    this.resetExtents();
  }

  get sheetOrder(): readonly string[] {
    return this.original.sheetOrder;
  }

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  getVersion = (): number => this.version;

  getValue(ref: CellRef): CellValue {
    return this.engine.getValue(ref);
  }

  getFormula(ref: CellRef): string | null {
    return this.engine.getFormula(ref);
  }

  /** The cell's Excel number format from the Labs document. Edits don't change formats. */
  fmt({ sheet, cell }: CellRef): string | undefined {
    return this.original.sheets[sheet]?.[cell]?.fmt;
  }

  extent(sheet: string): Extent {
    return this.extents.get(sheet) ?? { rows: 0, cols: 0 };
  }

  /** Commits typed content. Returns an error message instead of throwing (PRD §8 robustness). */
  enter(ref: CellRef, content: CellInput): string | null {
    let changed;
    try {
      changed = this.engine.setCell(ref, content);
    } catch (error) {
      return error instanceof Error ? error.message : String(error);
    }
    this.grow(ref);
    this.lastChange = {
      seq: this.lastChange.seq + 1,
      keys: new Set(changed.map((c) => keyOf(c.ref))),
    };
    this.emit();
    return null;
  }

  undo(): void {
    this.engine.undo();
    this.quietEmit();
  }

  redo(): void {
    this.engine.redo();
    this.quietEmit();
  }

  /** Restores every sheet to the Labs document's workbook (EX-3). */
  reset(): void {
    this.engine.load(this.original);
    this.resetExtents();
    this.quietEmit();
  }

  destroy(): void {
    this.listeners.clear();
    this.engine.destroy();
  }

  private grow({ sheet, cell }: CellRef): void {
    const { row, col } = parseA1(cell);
    const e = this.extent(sheet);
    this.extents.set(sheet, { rows: Math.max(e.rows, row + 1), cols: Math.max(e.cols, col + 1) });
  }

  private resetExtents(): void {
    this.extents = new Map();
    for (const [sheet, cells] of Object.entries(this.original.sheets)) {
      this.extents.set(sheet, { rows: 0, cols: 0 });
      for (const cell of Object.keys(cells)) this.grow({ sheet, cell });
    }
  }

  /** Recalculated without flashing (undo, redo, reset). */
  private quietEmit(): void {
    this.lastChange = { seq: this.lastChange.seq + 1, keys: new Set() };
    this.emit();
  }

  private emit(): void {
    this.version++;
    for (const listener of this.listeners) listener();
  }
}
