// A1-style addresses and 1900-system date serials (PRD §4.3).

import type { CellRef } from './types.ts';

export interface GridPosition {
  /** 0-based */
  row: number;
  /** 0-based */
  col: number;
}

const A1 = /^([A-Z]{1,3})([1-9][0-9]*)$/;

export function parseA1(address: string): GridPosition {
  const match = A1.exec(address);
  if (!match) throw new Error(`Invalid cell address "${address}"`);
  let col = 0;
  for (const ch of match[1]!) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(match[2]) - 1, col: col - 1 };
}

export function toA1({ row, col }: GridPosition): string {
  let letters = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  }
  return `${letters}${row + 1}`;
}

/** `"Flux!I8"` → `{ sheet: 'Flux', cell: 'I8' }` (the key format of experiment changes/expect). */
export function parseRef(ref: string): CellRef {
  const bang = ref.lastIndexOf('!');
  if (bang <= 0) throw new Error(`Invalid cell reference "${ref}"`);
  const cell = ref.slice(bang + 1);
  parseA1(cell);
  return { sheet: ref.slice(0, bang), cell };
}

export function formatRef({ sheet, cell }: CellRef): string {
  return `${sheet}!${cell}`;
}

const STRING_LITERAL = /"(?:[^"]|"")*"/g;
const SHEET_PREFIX = /'((?:[^']|'')+)'!|([A-Za-z_À-￿][\w.À-￿]*)!/g;

/** Sheet names a formula refers to (`GL!A1`, `'My Sheet'!A1`), ignoring text in string literals. */
export function referencedSheets(formula: string): string[] {
  const code = formula.replace(STRING_LITERAL, '""');
  const names = new Set<string>();
  for (const [, quoted, bare] of code.matchAll(SHEET_PREFIX)) {
    names.add(quoted !== undefined ? quoted.replaceAll("''", "'") : bare!);
  }
  return [...names];
}

// Excel's 1900 system: serial = days since 1899-12-30, exact for dates from 1900-03-01 on
// (Excel's fictitious 1900-02-29 only shifts earlier dates).
const EPOCH_MS = Date.UTC(1899, 11, 30);
const DAY_MS = 86_400_000;
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/;

/** `"2026-08-31"` → 46265. Returns `undefined` unless the string is a full, valid YYYY-MM-DD date. */
export function isoDateToSerial(iso: string): number | undefined {
  const match = ISO_DATE.exec(iso);
  if (!match) return undefined;
  const [y, m, d] = [Number(match[1]), Number(match[2]), Number(match[3])];
  const ms = Date.UTC(y, m - 1, d);
  const date = new Date(ms);
  if (date.getUTCFullYear() !== y || date.getUTCMonth() !== m - 1 || date.getUTCDate() !== d) {
    return undefined;
  }
  return Math.round((ms - EPOCH_MS) / DAY_MS);
}

/** 46265 → `"2026-08-31"`. The fractional (time) part is dropped. */
export function serialToIsoDate(serial: number): string {
  return new Date(EPOCH_MS + Math.floor(serial) * DAY_MS).toISOString().slice(0, 10);
}

/** The addresses in a range, row by row: `B5:C6` → `[['B5', 'C5'], ['B6', 'C6']]`. */
export function rangeAddresses(start: string, end: string): string[][] {
  const [a, b] = [parseA1(start), parseA1(end)];
  const rows: string[][] = [];
  for (let row = Math.min(a.row, b.row); row <= Math.max(a.row, b.row); row++) {
    const cells: string[] = [];
    for (let col = Math.min(a.col, b.col); col <= Math.max(a.col, b.col); col++) {
      cells.push(toA1({ row, col }));
    }
    rows.push(cells);
  }
  return rows;
}
