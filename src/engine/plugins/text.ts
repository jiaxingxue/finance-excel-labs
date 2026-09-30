// TEXT(value, format_text) — overrides HyperFormula 3.4's built-in, which only understands a few
// date/time codes: TEXT(100,"#,##0.00") returns "100,##0.00" there. Formatting is delegated to
// SheetJS SSF (PRD §7.4), the same library the grid will use to display number formats.

import { CellError, EmptyValue, ErrorType, FunctionArgumentType } from 'hyperformula';
import SSF from 'ssf';
import { LabPlugin, cellError, toNumber, type PluginAst, type PluginState } from './shared.ts';

export function formatValue(value: unknown, formatText: string): string | CellError {
  // Excel: blanks format as 0, numeric text is converted, booleans come back as TRUE/FALSE.
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  let input: number | string;
  if (value === EmptyValue) input = 0;
  else if (typeof value === 'string') input = numericText(value) ?? value;
  else {
    const n = toNumber(value);
    if (n === undefined) return cellError(ErrorType.VALUE, 'TEXT needs a number or text value.');
    input = n;
  }
  try {
    return SSF.format(formatText, input);
  } catch {
    return cellError(ErrorType.VALUE, `Unsupported number format "${formatText}".`);
  }
}

function numericText(text: string): number | undefined {
  const trimmed = text.trim();
  if (trimmed === '') return undefined;
  const n = Number(trimmed);
  return Number.isFinite(n) ? n : undefined;
}

export class TextPlugin extends LabPlugin {
  static implementedFunctions = {
    TEXT: {
      method: 'text',
      parameters: [
        { argumentType: FunctionArgumentType.NOERROR },
        { argumentType: FunctionArgumentType.STRING },
      ],
    },
  };

  text(ast: PluginAst, state: PluginState) {
    return this.run('TEXT', ast, state, (value: unknown, formatText: string) =>
      formatValue(value, formatText),
    );
  }
}
