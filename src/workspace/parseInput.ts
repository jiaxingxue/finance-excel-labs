// Turns what a learner types into cell content (GR-4). The engine adapter leaves this to the
// caller: `setCell` treats every string that doesn't start with "=" as literal text.

import { isoDateToSerial } from '../engine/address.ts';

export type CellInput = string | number | boolean | null;

const NUMBER = /^[+-]?(\d{1,3}(,\d{3})+|\d+)?(\.\d+)?([eE][+-]?\d+)?$/;

/**
 * - `""` clears the cell; `=…` is a formula.
 * - `TRUE` / `FALSE` (any case) are booleans.
 * - Numbers may use thousands separators and a trailing `%` (`1,000`, `5%`, `-0.25`).
 * - A full `YYYY-MM-DD` date becomes a 1900-system serial (the same rule as experiment changes).
 * - Anything else is text, including text periods such as `2026-08`.
 */
export function parseInput(typed: string): CellInput {
  const text = typed.trim();
  if (text === '') return null;
  if (text.startsWith('=')) return text;
  const upper = text.toUpperCase();
  if (upper === 'TRUE') return true;
  if (upper === 'FALSE') return false;
  const percent = text.endsWith('%');
  const body = percent ? text.slice(0, -1).trim() : text;
  if (/\d/.test(body) && NUMBER.test(body)) {
    const n = Number(body.replaceAll(',', ''));
    if (Number.isFinite(n)) return percent ? n / 100 : n;
  }
  return isoDateToSerial(text) ?? typed;
}
