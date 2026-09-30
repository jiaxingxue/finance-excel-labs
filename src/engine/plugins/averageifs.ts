// AVERAGEIFS(average_range, criteria_range1, criterion1, …) — not built into HyperFormula 3.4.
// Criteria are parsed by HyperFormula's own criterion builder, so operators (<, <=, >, >=, <>, =),
// wildcards and text/number coercion behave exactly as they do in its SUMIFS and COUNTIFS.

import { CellError, ErrorType, FunctionArgumentType, SimpleRangeValue } from 'hyperformula';
import {
  LabPlugin,
  cellError,
  flatten,
  toNumber,
  toRaw,
  type PluginAst,
  type PluginState,
} from './shared.ts';

export class AverageIfsPlugin extends LabPlugin {
  static implementedFunctions = {
    AVERAGEIFS: {
      method: 'averageifs',
      parameters: [
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.NOERROR },
      ],
      repeatLastArgs: 2,
    },
  };

  averageifs(ast: PluginAst, state: PluginState) {
    return this.run(
      'AVERAGEIFS',
      ast,
      state,
      (values: SimpleRangeValue, ...pairs: (SimpleRangeValue | string | number | boolean)[]) => {
        const averaged = flatten(values);
        const tests: { cells: unknown[]; matches: (v: never) => boolean }[] = [];
        for (let i = 0; i < pairs.length; i += 2) {
          const range = pairs[i] as SimpleRangeValue;
          if (!range.sameDimensionsAs(values)) {
            return cellError(ErrorType.VALUE, 'AVERAGEIFS ranges must have the same size.');
          }
          const criterion = this.interpreter.criterionBuilder.fromCellValue(
            toRaw(pairs[i + 1]) as string | number | boolean,
            this.arithmeticHelper,
          );
          if (criterion === undefined) {
            return cellError(ErrorType.VALUE, 'Invalid AVERAGEIFS criterion.');
          }
          tests.push({ cells: flatten(range), matches: criterion.lambda });
        }

        let sum = 0;
        let count = 0;
        for (let k = 0; k < averaged.length; k++) {
          if (!tests.every((t) => t.matches(toRaw(t.cells[k]) as never))) continue;
          const cell = averaged[k];
          if (cell instanceof CellError) return cell;
          const n = toNumber(cell);
          if (n !== undefined) {
            sum += n;
            count++;
          }
        }
        return count > 0 ? sum / count : cellError(ErrorType.DIV_BY_ZERO, 'No cells match.');
      },
    );
  }
}
