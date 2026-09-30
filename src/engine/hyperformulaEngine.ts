// FormulaEngine implemented on HyperFormula (PRD §7.2–7.4). All HyperFormula-specific details —
// configuration, plugin registration, addresses, value and date conversion — stay in this file.

import {
  DetailedCellError,
  HyperFormula,
  type CellValue as HfCellValue,
  type ConfigParams,
  type ExportedChange,
  type RawCellContent,
  type SerializedNamedExpression,
  type SimpleCellAddress,
  type SimpleCellRange,
} from 'hyperformula';
import type { CellSpec, SheetSpec, WorkbookSpec } from '../content/types.ts';
import { isoDateToSerial, parseA1, referencedSheets, serialToIsoDate, toA1 } from './address.ts';
import { registerPlugins } from './plugins/index.ts';
import type { CellRef, CellValue, ChangedCell, FormulaEngine, RangeRef } from './types.ts';

// HyperFormula is licensed under GPLv3 for this project (docs/DECISIONS.md D1).
const CONFIG: Partial<ConfigParams> = {
  licenseKey: 'gpl-v3',
  // Array arithmetic inside functions: SUMPRODUCT(…^range), LOOKUP(2,1/(…),…) (EN-3).
  useArrayArithmetic: true,
  // The default rounds results to ~10 significant digits (Drivers!B21 lost precision).
  smartRounding: false,
  // Excel shows 0 for a formula that references a blank cell (`=A1`).
  evaluateNullToZero: true,
  // Excel's 1900 date system (serial 1 = 1900-01-01, exact from 1900-03-01 on).
  nullDate: { year: 1899, month: 12, day: 30 },
};

// HyperFormula parses TRUE and FALSE only as functions — TRUE() — so `=(TRUE)=(FALSE)` and
// `COUNTIF(range,FALSE)` give #NAME?. Defining them as names makes the bare literals work.
const NAMED_EXPRESSIONS: SerializedNamedExpression[] = [
  { name: 'TRUE', expression: '=TRUE()' },
  { name: 'FALSE', expression: '=FALSE()' },
];

/** HyperFormula treats a leading apostrophe as "this is text", as Excel does. */
const TEXT_PREFIX = "'";

export class HyperFormulaEngine implements FormulaEngine {
  private hf: HyperFormula | undefined;
  /** Number formats by sheet and address. They stay with the position, as in Excel. */
  private formats = new Map<string, Map<string, string>>();
  /** "Sheet!A1" of cells loaded as date constants, so exportCells can write them back as dates. */
  private dateCells = new Set<string>();

  constructor(workbook?: WorkbookSpec) {
    registerPlugins();
    if (workbook) this.load(workbook);
  }

  load(workbook: WorkbookSpec): void {
    this.hf?.destroy();
    this.formats = new Map();
    this.dateCells = new Set();
    const sheets: Record<string, RawCellContent[][]> = {};
    for (const name of workbook.sheetOrder) {
      const spec = workbook.sheets[name];
      if (!spec) throw new Error(`Sheet "${name}" is in sheetOrder but has no cells`);
      const grid: RawCellContent[][] = [];
      const formats = new Map<string, string>();
      for (const [address, cell] of Object.entries(spec)) {
        const { row, col } = parseA1(address);
        if ('f' in cell) assertKnownSheets(cell.f, workbook.sheetOrder, `${name}!${address}`);
        (grid[row] ??= [])[col] = this.toRawContent(name, address, cell);
        if (cell.fmt !== undefined) formats.set(address, cell.fmt);
      }
      sheets[name] = Array.from(grid, (row) => row ?? []);
      this.formats.set(name, formats);
    }
    this.hf = HyperFormula.buildFromSheets(sheets, CONFIG, NAMED_EXPRESSIONS);
  }

  sheetNames(): string[] {
    return this.engine.getSheetNames();
  }

  setCell(ref: CellRef, content: string | number | boolean | null): ChangedCell[] {
    let raw: RawCellContent;
    if (content === null || content === '') {
      raw = null;
    } else if (typeof content === 'string' && content.startsWith('=')) {
      assertKnownSheets(content, this.sheetNames(), `${ref.sheet}!${ref.cell}`);
      raw = content;
    } else {
      raw = typeof content === 'string' ? TEXT_PREFIX + content : content;
    }
    return this.mapChanges(this.engine.setCellContents(this.address(ref), raw));
  }

  getValue(ref: CellRef): CellValue {
    return toCellValue(this.engine.getCellValue(this.address(ref)));
  }

  getFormula(ref: CellRef): string | null {
    return this.engine.getCellFormula(this.address(ref)) ?? null;
  }

  fill(source: RangeRef, target: RangeRef): ChangedCell[] {
    const [src, dst] = [this.range(source), this.range(target)];
    const data = this.engine.getFillRangeData(src, dst);
    return this.mapChanges(this.engine.setCellContents(dst.start, data));
  }

  precedents(ref: CellRef): RangeRef[] {
    return this.engine.getCellPrecedents(this.address(ref)).map((p) => this.toRangeRef(p));
  }

  dependents(ref: CellRef): CellRef[] {
    // HyperFormula reports a range as the dependent of each cell in it, and the formulas that use
    // the range as the range's dependents, so walk through ranges to reach formula cells.
    const result = new Map<string, CellRef>();
    const seenRanges = new Set<string>();
    const queue: (SimpleCellAddress | SimpleCellRange)[] = [this.address(ref)];
    while (queue.length > 0) {
      for (const dep of this.engine.getCellDependents(queue.shift()!)) {
        if ('start' in dep) {
          const key = JSON.stringify(dep);
          if (!seenRanges.has(key)) {
            seenRanges.add(key);
            queue.push(dep);
          }
        } else {
          const cell = this.toCellRef(dep);
          result.set(`${cell.sheet}!${cell.cell}`, cell);
        }
      }
    }
    return [...result.values()];
  }

  clone(): FormulaEngine {
    return new HyperFormulaEngine(this.exportCells());
  }

  undo(): void {
    if (this.engine.isThereSomethingToUndo()) this.engine.undo();
  }

  redo(): void {
    if (this.engine.isThereSomethingToRedo()) this.engine.redo();
  }

  exportCells(): WorkbookSpec {
    const workbook: WorkbookSpec = { sheetOrder: this.sheetNames(), sheets: {} };
    for (const name of workbook.sheetOrder) {
      const sheet: SheetSpec = {};
      const formats = this.formats.get(name);
      this.engine.getSheetSerialized(this.sheetId(name)).forEach((cells, row) =>
        cells.forEach((raw, col) => {
          const address = toA1({ row, col });
          const cell = this.toCellSpec(name, address, raw);
          if (!cell) return;
          const fmt = formats?.get(address);
          sheet[address] = fmt === undefined ? cell : { ...cell, fmt };
        }),
      );
      workbook.sheets[name] = sheet;
    }
    return workbook;
  }

  destroy(): void {
    this.hf?.destroy();
    this.hf = undefined;
  }

  // --- conversions ------------------------------------------------------------------------------

  private get engine(): HyperFormula {
    if (!this.hf) throw new Error('No workbook loaded');
    return this.hf;
  }

  private toRawContent(sheet: string, address: string, cell: CellSpec): RawCellContent {
    if ('f' in cell) return cell.f;
    const { v } = cell;
    if (typeof v === 'object') {
      const serial = isoDateToSerial(v.date);
      if (serial === undefined) throw new Error(`${sheet}!${address}: invalid date "${v.date}"`);
      this.dateCells.add(`${sheet}!${address}`);
      return serial;
    }
    return typeof v === 'string' ? TEXT_PREFIX + v : v;
  }

  private toCellSpec(sheet: string, address: string, raw: RawCellContent): CellSpec | undefined {
    if (raw === null || raw === undefined || raw === '') return undefined;
    if (typeof raw === 'string') {
      if (raw.startsWith('=')) return { f: raw };
      return { v: raw.startsWith(TEXT_PREFIX) ? raw.slice(TEXT_PREFIX.length) : raw };
    }
    if (typeof raw === 'number' && this.dateCells.has(`${sheet}!${address}`)) {
      return { v: { date: serialToIsoDate(raw) } };
    }
    if (typeof raw === 'number' || typeof raw === 'boolean') return { v: raw };
    throw new Error(`${sheet}!${address}: unexpected cell content`);
  }

  private sheetId(name: string): number {
    const id = this.engine.getSheetId(name);
    if (id === undefined) throw new Error(`Unknown sheet "${name}"`);
    return id;
  }

  private address({ sheet, cell }: CellRef): SimpleCellAddress {
    return { sheet: this.sheetId(sheet), ...parseA1(cell) };
  }

  private range({ sheet, start, end }: RangeRef): SimpleCellRange {
    return {
      start: this.address({ sheet, cell: start }),
      end: this.address({ sheet, cell: end }),
    };
  }

  private toCellRef(address: SimpleCellAddress): CellRef {
    const sheet = this.engine.getSheetName(address.sheet);
    if (sheet === undefined) throw new Error(`Unknown sheet id ${address.sheet}`);
    return { sheet, cell: toA1(address) };
  }

  private toRangeRef(target: SimpleCellAddress | SimpleCellRange): RangeRef {
    const [start, end] = 'start' in target ? [target.start, target.end] : [target, target];
    const { sheet, cell } = this.toCellRef(start);
    return { sheet, start: cell, end: toA1(end) };
  }

  private mapChanges(changes: ExportedChange[]): ChangedCell[] {
    const result: ChangedCell[] = [];
    for (const change of changes) {
      if ('address' in change) {
        result.push({ ref: this.toCellRef(change.address), value: toCellValue(change.newValue) });
      }
    }
    return result;
  }
}

/**
 * Throws if the formula refers to a sheet that doesn't exist. Excel would ask for an external
 * workbook; HyperFormula 3.4.0 accepts it but later throws an internal "Range does not exist"
 * error when that cell is overwritten (docs/OPEN_ISSUES.md #17), so such formulas never reach it.
 */
function assertKnownSheets(formula: string, sheets: string[], where: string): void {
  const known = new Set(sheets.map((s) => s.toLowerCase()));
  const unknown = referencedSheets(formula).filter((s) => !known.has(s.toLowerCase()));
  if (unknown.length > 0) {
    throw new Error(`${where}: the formula refers to an unknown sheet "${unknown[0]}"`);
  }
}

function toCellValue(value: HfCellValue): CellValue {
  if (value instanceof DetailedCellError) {
    return value.message ? { error: value.value, message: value.message } : { error: value.value };
  }
  return value;
}
