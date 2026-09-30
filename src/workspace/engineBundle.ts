// Everything the lab workspace needs from the formula engine, in one lazily loaded chunk (PRD §8:
// the lesson renders first). Import it only through loadEngine.ts; a static import would pull
// HyperFormula into the lab page's own chunk.

import workbookJson from '../data/workbook.json' with { type: 'json' };
import type { WorkbookSpec } from '../content/types.ts';
import { createEngine } from '../engine/index.ts';
import { WorkbookStore } from './workbookStore.ts';

export const workbook = workbookJson as WorkbookSpec;

export function createStore(): WorkbookStore {
  return new WorkbookStore(createEngine(workbook), workbook);
}
