// INDEX(array, row_num, [column_num]) — overrides HyperFormula 3.4's built-in, which always reads the
// second argument as a row number, so INDEX(B3:D3, 2) on a one-row range returns #NUM! instead of C3
// (Drivers!C9 and E4:E6 use this form).
//
// Excel semantics implemented here: with column_num omitted, a one-row array is indexed by column
// and anything else by row (column 1). Out-of-range indexes return #REF! as in Excel.
// Not supported (kept as in HyperFormula): row_num or column_num of 0, which in Excel returns a
// whole column or row, returns #VALUE!.

import { ErrorType, FunctionArgumentType, SimpleRangeValue } from 'hyperformula';
import {
  LabPlugin,
  cellError,
  type PluginAst,
  type PluginResult,
  type PluginState,
} from './shared.ts';

export class IndexPlugin extends LabPlugin {
  static implementedFunctions = {
    INDEX: {
      method: 'index',
      parameters: [
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.NUMBER },
        { argumentType: FunctionArgumentType.NUMBER, optionalArg: true },
      ],
    },
  };

  index(ast: PluginAst, state: PluginState) {
    return this.run(
      'INDEX',
      ast,
      state,
      (array: SimpleRangeValue, rowArg: number, colArg?: number) => {
        let row = Math.trunc(rowArg);
        let col = colArg === undefined ? 1 : Math.trunc(colArg);
        if (colArg === undefined && array.height() === 1) [row, col] = [1, row];
        if (row < 1 || col < 1) {
          return cellError(ErrorType.VALUE, 'INDEX row and column numbers must be at least 1.');
        }
        if (row > array.height() || col > array.width()) {
          return cellError(ErrorType.REF, 'INDEX is outside the range.');
        }
        return array.data[row - 1]![col - 1] as PluginResult;
      },
    );
  }
}
