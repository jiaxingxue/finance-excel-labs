// Evaluating assertions and experiments against an engine (PRD §7.3, §9.4). Used by the
// conformance suite now and by the in-browser verification page later (PRD §17.5).

import type { Assertion, Experiment } from '../content/types.ts';
import { parseRef, type CellRef, type CellValue, type FormulaEngine } from '../engine/index.ts';
import { assertionPasses, experimentChangeContent, experimentValuePasses } from './match.ts';

export interface AssertionResult {
  assertion: Assertion;
  actual: CellValue;
  pass: boolean;
}

export function checkAssertion(engine: FormulaEngine, assertion: Assertion): AssertionResult {
  const actual = engine.getValue(assertion);
  return { assertion, actual, pass: assertionPasses(actual, assertion) };
}

export interface ExperimentCheck {
  ref: CellRef;
  expected: number | string | boolean | null;
  actual: CellValue;
  pass: boolean;
}

export interface ExperimentResult {
  experiment: Experiment;
  checks: ExperimentCheck[];
  pass: boolean;
}

/** Applies the experiment's changes to `engine` (mutating it) and checks every expected cell. */
export function applyExperiment(engine: FormulaEngine, experiment: Experiment): ExperimentResult {
  for (const [ref, value] of Object.entries(experiment.changes)) {
    engine.setCell(parseRef(ref), experimentChangeContent(value));
  }
  const checks = Object.entries(experiment.expect).map(([key, expected]) => {
    const ref = parseRef(key);
    const actual = engine.getValue(ref);
    return { ref, expected, actual, pass: experimentValuePasses(actual, expected) };
  });
  return { experiment, checks, pass: checks.every((c) => c.pass) };
}

/** Runs the experiment on a throwaway clone, leaving `base` untouched. */
export function runExperiment(base: FormulaEngine, experiment: Experiment): ExperimentResult {
  const clone = base.clone();
  try {
    return applyExperiment(clone, experiment);
  } finally {
    clone.destroy();
  }
}
