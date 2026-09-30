// Helpers shared by the gap-filling plugins (PRD §7.4).

import { CellError, EmptyValue, ErrorType, FunctionPlugin, SimpleRangeValue } from 'hyperformula';

/** What a plugin callback may return to HyperFormula. */
export type PluginResult =
  number | string | boolean | CellError | SimpleRangeValue | typeof EmptyValue;

/** HyperFormula's plugin method signature. Its AST and state types aren't exported, so borrow them. */
type RunFunction = LabPlugin['runFunction'];
export type PluginAst = { args: Parameters<RunFunction>[0] };
export type PluginState = Parameters<RunFunction>[1];
type PluginImpl = Parameters<RunFunction>[3];

export abstract class LabPlugin extends FunctionPlugin {
  /** Evaluates the arguments per the function's metadata, then calls `impl` (HyperFormula's contract). */
  protected run(name: string, ast: PluginAst, state: PluginState, impl: PluginImpl) {
    return this.runFunction(ast.args, state, this.metadata(name), impl);
  }
}

/**
 * The plain number inside a cell value, or `undefined` for anything that isn't a number.
 * HyperFormula wraps some numbers (dates, percents, currency) in a RichNumber `{ val }`.
 */
export function toNumber(value: unknown): number | undefined {
  if (typeof value === 'number') return value;
  if (typeof value === 'object' && value !== null && !(value instanceof CellError)) {
    const val = (value as { val?: unknown }).val;
    if (typeof val === 'number') return val;
  }
  return undefined;
}

/** The value as HyperFormula's criterion lambdas expect it (RichNumbers unwrapped). */
export function toRaw<T>(value: T): T | number {
  return toNumber(value) ?? value;
}

export function cellError(type: ErrorType, message?: string): CellError {
  return new CellError(type, message);
}

/** All values of a range or array, row by row. */
export function flatten(range: SimpleRangeValue): unknown[] {
  return range.valuesFromTopLeftCorner();
}
