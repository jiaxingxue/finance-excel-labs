// LOOKUP(lookup_value, lookup_vector, [result_vector]) — not built into HyperFormula 3.4.
//
// Excel semantics (PRD §7.4): approximate match that assumes lookup_vector is ascending and returns
// the result for the LAST position whose value is <= lookup_value. Error values, blanks and values of
// a different type are skipped, which is what makes the `LOOKUP(2, 1/(cond), results)` last-match
// idiom work. Text compares case-insensitively. No match returns #N/A.
// Array form (no result_vector, 2-D lookup array): searches the first row if the array is wider than
// tall, otherwise the first column, and returns from the last row/column.

import {
  CellError,
  EmptyValue,
  ErrorType,
  FunctionArgumentType,
  SimpleRangeValue,
} from 'hyperformula';
import {
  LabPlugin,
  cellError,
  flatten,
  toNumber,
  type PluginAst,
  type PluginResult,
  type PluginState,
} from './shared.ts';

type Comparable = number | string | boolean;

export class LookupPlugin extends LabPlugin {
  static implementedFunctions = {
    LOOKUP: {
      method: 'lookup',
      parameters: [
        { argumentType: FunctionArgumentType.NOERROR },
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.RANGE, optionalArg: true },
      ],
    },
  };

  lookup(ast: PluginAst, state: PluginState) {
    return this.run(
      'LOOKUP',
      ast,
      state,
      (key: unknown, lookupArray: SimpleRangeValue, resultArray?: SimpleRangeValue) => {
        const [searched, results] = resultArray
          ? [flatten(lookupArray), flatten(resultArray)]
          : arrayForm(lookupArray);
        const index = lastMatchAtOrBelow(searched, normalize(key));
        if (index < 0) {
          return cellError(ErrorType.NA, 'No value is less than or equal to the lookup value.');
        }
        if (index >= results.length) {
          return cellError(ErrorType.NA, 'The result vector is too short.');
        }
        return results[index] as PluginResult;
      },
    );
  }
}

/** For LOOKUP's two-argument form: which values to search and which to return. */
function arrayForm(array: SimpleRangeValue): [unknown[], unknown[]] {
  const rows = array.data;
  const height = array.height();
  const width = array.width();
  if (height === 1 || width === 1) {
    const values = flatten(array);
    return [values, values];
  }
  if (width > height) return [rows[0]!, rows[height - 1]!];
  return [rows.map((r) => r[0]), rows.map((r) => r[width - 1])];
}

/** Numbers (incl. dates), text and booleans are comparable; errors and blanks are not. */
function normalize(value: unknown): Comparable | undefined {
  if (value instanceof CellError || value === EmptyValue) return undefined;
  const n = toNumber(value);
  if (n !== undefined) return n;
  if (typeof value === 'string') return value.toLowerCase();
  if (typeof value === 'boolean') return value;
  return undefined;
}

/** Binary search, over the entries of the key's type, for the last one <= key. */
function lastMatchAtOrBelow(values: unknown[], key: Comparable | undefined): number {
  if (key === undefined) return -1;
  const candidates: { index: number; value: Comparable }[] = [];
  values.forEach((raw, index) => {
    const value = normalize(raw);
    if (value !== undefined && typeof value === typeof key) candidates.push({ index, value });
  });
  let lo = 0;
  let hi = candidates.length - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (candidates[mid]!.value <= key) {
      found = candidates[mid]!.index;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return found;
}
