// Runtime shape checks for JSON parsed out of the Labs document.
// The document is trusted content, but a typo there must fail loudly, not load quietly.

import type { Assertion, CellSpec, Experiment } from '../../src/content/types.ts';

export class ContentError extends Error {}

export function fail(msg: string): never {
  throw new ContentError(msg);
}

const isObject = (x: unknown): x is Record<string, unknown> =>
  typeof x === 'object' && x !== null && !Array.isArray(x);

const A1 = /^[A-Z]{1,3}[1-9]\d*$/;
const SHEET_REF = /^[A-Za-z][A-Za-z0-9_]*![A-Z]{1,3}[1-9]\d*$/;
const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

export function isA1(s: string): boolean {
  return A1.test(s);
}

export function checkCell(where: string, addr: string, raw: unknown): CellSpec {
  if (!isA1(addr)) fail(`${where}: "${addr}" is not an A1 address`);
  if (!isObject(raw)) fail(`${where}!${addr}: cell must be an object`);
  const keys = Object.keys(raw);
  for (const k of keys) {
    if (k !== 'v' && k !== 'f' && k !== 'fmt') fail(`${where}!${addr}: unknown key "${k}"`);
  }
  if ('fmt' in raw && typeof raw.fmt !== 'string') fail(`${where}!${addr}: fmt must be a string`);
  if ('f' in raw === 'v' in raw) fail(`${where}!${addr}: needs exactly one of "v" or "f"`);
  if ('f' in raw) {
    if (typeof raw.f !== 'string' || !raw.f.startsWith('=')) {
      fail(`${where}!${addr}: formula must be a string starting with "="`);
    }
  } else {
    const v = raw.v;
    const scalar = typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean';
    const date =
      isObject(v) &&
      Object.keys(v).length === 1 &&
      typeof v.date === 'string' &&
      ISO_DATE.test(v.date);
    if (!scalar && !date) fail(`${where}!${addr}: v must be a scalar or {"date": "YYYY-MM-DD"}`);
  }
  return raw as CellSpec;
}

export function checkAssertion(i: number, raw: unknown): Assertion {
  const where = `Assertion #${i + 1}`;
  if (!isObject(raw)) fail(`${where}: must be an object`);
  const { sheet, cell, expected, tolerance } = raw;
  if (Object.keys(raw).length !== 4)
    fail(`${where}: expected keys sheet, cell, expected, tolerance`);
  if (typeof sheet !== 'string') fail(`${where}: sheet must be a string`);
  if (typeof cell !== 'string' || !isA1(cell)) fail(`${where}: cell must be an A1 address`);
  if (!['number', 'string', 'boolean'].includes(typeof expected)) {
    fail(`${where}: expected must be a number, string, or boolean`);
  }
  if (typeof tolerance !== 'number' || !(tolerance >= 0)) {
    fail(`${where}: tolerance must be a non-negative number`);
  }
  return raw as unknown as Assertion;
}

export function checkExperiment(i: number, raw: unknown): Experiment {
  const where = `Experiment #${i + 1}`;
  if (!isObject(raw)) fail(`${where}: must be an object`);
  const { id, lab, title, changes, expect } = raw;
  if (typeof id !== 'string') fail(`${where}: id must be a string`);
  if (typeof lab !== 'number' || !Number.isInteger(lab)) fail(`${id}: lab must be an integer`);
  if (typeof title !== 'string') fail(`${id}: title must be a string`);
  if (!isObject(changes) || Object.keys(changes).length === 0) {
    fail(`${id}: changes must be a non-empty object`);
  }
  if (!isObject(expect) || Object.keys(expect).length === 0) {
    fail(`${id}: expect must be a non-empty object`);
  }
  for (const [ref, v] of Object.entries(changes)) {
    if (!SHEET_REF.test(ref)) fail(`${id}: "${ref}" is not a Sheet!A1 reference`);
    if (typeof v !== 'number' && typeof v !== 'string') {
      fail(`${id}: change ${ref} must be a number or string`);
    }
  }
  for (const [ref, v] of Object.entries(expect)) {
    if (!SHEET_REF.test(ref)) fail(`${id}: "${ref}" is not a Sheet!A1 reference`);
    if (v !== null && !['number', 'string', 'boolean'].includes(typeof v)) {
      fail(`${id}: expect ${ref} must be a number, string, boolean, or null`);
    }
  }
  return raw as unknown as Experiment;
}

export function parseJson(where: string, body: string): unknown {
  try {
    return JSON.parse(body);
  } catch (e) {
    return fail(`${where}: invalid JSON (${(e as Error).message})`);
  }
}
