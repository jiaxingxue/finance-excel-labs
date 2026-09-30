// Ordinary least squares: INTERCEPT, FORECAST.LINEAR (alias FORECAST) and TREND — none are built
// into HyperFormula 3.4 (its SLOPE and RSQ are, and pass conformance).
//
// Excel semantics: known_y and known_x are paired by position; a pair is used only when both values
// are numbers (text, booleans and blanks inside ranges are skipped); an error in either propagates.
// Different sizes → #N/A (TREND: #REF!); fewer than two points or zero variance in x → #DIV/0!.

import { CellError, ErrorType, FunctionArgumentType, SimpleRangeValue } from 'hyperformula';
import {
  LabPlugin,
  cellError,
  flatten,
  toNumber,
  type PluginAst,
  type PluginState,
} from './shared.ts';

interface Fit {
  intercept: number;
  slope: number;
}

export function leastSquares(
  knownY: SimpleRangeValue,
  knownX: SimpleRangeValue,
  sizeError: ErrorType,
  throughOrigin = false,
): Fit | CellError {
  const ys = flatten(knownY);
  const xs = flatten(knownX);
  if (ys.length !== xs.length) {
    return cellError(sizeError, 'known_y and known_x must have the same number of values.');
  }
  const pairs: [number, number][] = [];
  for (let i = 0; i < ys.length; i++) {
    const [y, x] = [ys[i], xs[i]];
    if (y instanceof CellError) return y;
    if (x instanceof CellError) return x;
    const [ny, nx] = [toNumber(y), toNumber(x)];
    if (ny !== undefined && nx !== undefined) pairs.push([nx, ny]);
  }
  if (pairs.length < (throughOrigin ? 1 : 2)) {
    return cellError(ErrorType.DIV_BY_ZERO, 'Not enough numeric data points.');
  }

  if (throughOrigin) {
    let sxy = 0;
    let sxx = 0;
    for (const [x, y] of pairs) {
      sxy += x * y;
      sxx += x * x;
    }
    if (sxx === 0) return cellError(ErrorType.DIV_BY_ZERO, 'All known_x values are zero.');
    return { intercept: 0, slope: sxy / sxx };
  }

  // Mean-centred sums (numerically stable, as in Excel).
  const n = pairs.length;
  const meanX = pairs.reduce((s, [x]) => s + x, 0) / n;
  const meanY = pairs.reduce((s, [, y]) => s + y, 0) / n;
  let sxy = 0;
  let sxx = 0;
  for (const [x, y] of pairs) {
    sxy += (x - meanX) * (y - meanY);
    sxx += (x - meanX) ** 2;
  }
  if (sxx === 0) return cellError(ErrorType.DIV_BY_ZERO, 'The known_x values have zero variance.');
  const slope = sxy / sxx;
  return { intercept: meanY - slope * meanX, slope };
}

export class RegressionPlugin extends LabPlugin {
  static implementedFunctions = {
    INTERCEPT: {
      method: 'intercept',
      parameters: [
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.RANGE },
      ],
    },
    'FORECAST.LINEAR': {
      method: 'forecastLinear',
      parameters: [
        { argumentType: FunctionArgumentType.NUMBER },
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.RANGE },
      ],
    },
    TREND: {
      method: 'trend',
      parameters: [
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true },
      ],
    },
  };

  static aliases = { FORECAST: 'FORECAST.LINEAR' };

  intercept(ast: PluginAst, state: PluginState) {
    return this.run('INTERCEPT', ast, state, (ys: SimpleRangeValue, xs: SimpleRangeValue) => {
      const fit = leastSquares(ys, xs, ErrorType.NA);
      return fit instanceof CellError ? fit : fit.intercept;
    });
  }

  forecastLinear(ast: PluginAst, state: PluginState) {
    return this.run(
      'FORECAST.LINEAR',
      ast,
      state,
      (x: number, ys: SimpleRangeValue, xs: SimpleRangeValue) => {
        const fit = leastSquares(ys, xs, ErrorType.NA);
        return fit instanceof CellError ? fit : fit.intercept + fit.slope * x;
      },
    );
  }

  /**
   * Single-value form only (PRD §7.4): known_x and new_x are required and new_x must be one value.
   * Not supported: omitted known_x (Excel's default 1..n) and array results.
   */
  trend(ast: PluginAst, state: PluginState) {
    return this.run(
      'TREND',
      ast,
      state,
      (ys: SimpleRangeValue, xs: SimpleRangeValue, newX: SimpleRangeValue, useConst: boolean) => {
        const newValues = flatten(newX);
        if (newValues.length !== 1) {
          return cellError(ErrorType.VALUE, 'TREND is supported only with a single new_x value.');
        }
        const [value] = newValues;
        if (value instanceof CellError) return value;
        const x = toNumber(value);
        if (x === undefined) return cellError(ErrorType.VALUE, 'new_x must be a number.');
        const fit = leastSquares(ys, xs, ErrorType.REF, !useConst);
        return fit instanceof CellError ? fit : fit.intercept + fit.slope * x;
      },
    );
  }
}
