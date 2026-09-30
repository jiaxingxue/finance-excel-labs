// Public entry point of the formula engine. Everything outside src/engine/ imports from here.

import type { WorkbookSpec } from '../content/types.ts';
import { HyperFormulaEngine } from './hyperformulaEngine.ts';
import type { FormulaEngine } from './types.ts';

export * from './types.ts';
export { formatRef, isoDateToSerial, parseRef, serialToIsoDate } from './address.ts';
export { PLUGIN_FUNCTIONS } from './plugins/index.ts';

export const ENGINE_NAME = 'HyperFormula';

export function createEngine(workbook?: WorkbookSpec): FormulaEngine {
  return new HyperFormulaEngine(workbook);
}
