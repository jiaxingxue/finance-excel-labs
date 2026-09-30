// Value comparison rules for grading (PRD §4.4).

import type { Assertion } from '../content/types.ts';
import { isoDateToSerial, type CellValue } from '../engine/index.ts';

/** Assertion numbers pass if |actual − expected| ≤ max(tolerance, |expected| × 1e-12). */
export function assertionPasses(actual: CellValue, assertion: Assertion): boolean {
  const { expected, tolerance } = assertion;
  if (typeof expected === 'number') {
    return (
      typeof actual === 'number' &&
      Math.abs(actual - expected) <= Math.max(tolerance, Math.abs(expected) * 1e-12)
    );
  }
  if (expected === '') return isEmpty(actual);
  return actual === expected;
}

const EXPERIMENT_RELATIVE = 1e-6;
const EXPERIMENT_ABSOLUTE = 0.005;

/** Experiment numbers use 1e-6 relative or 0.005 absolute, whichever is larger; `null` means empty. */
export function experimentValuePasses(
  actual: CellValue,
  expected: number | string | boolean | null,
): boolean {
  if (expected === null || expected === '') return isEmpty(actual);
  if (typeof expected === 'number') {
    return (
      typeof actual === 'number' &&
      Math.abs(actual - expected) <=
        Math.max(Math.abs(expected) * EXPERIMENT_RELATIVE, EXPERIMENT_ABSOLUTE)
    );
  }
  return actual === expected;
}

/** An experiment change as cell content: only a full YYYY-MM-DD string becomes a date serial. */
export function experimentChangeContent(value: number | string): number | string {
  return typeof value === 'string' ? (isoDateToSerial(value) ?? value) : value;
}

function isEmpty(value: CellValue): boolean {
  return value === null || value === '';
}
