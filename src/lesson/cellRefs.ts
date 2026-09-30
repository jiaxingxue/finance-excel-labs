// Which cell references in lesson text become clickable links (GR-11, OPEN_ISSUES #20).
//
// 1. Qualified references (`BvA!C6`, `BvA!A5:H5`) on a known sheet are always links, in code spans
//    and in plain text.
// 2. Bare references (`C6`, `A5:H5`) are links only in a lab with exactly one primary sheet, and
//    they resolve to that sheet. Only whole code spans count: in plain text, strings such as
//    "E7.1" (an experiment id) would be misread as cells.
// 3. Anything containing `$` is a formula fragment (`$A6`, `Map!$B$2:$B$7`) and is never a link.

import type { LabConfig } from '../config/labs.config.ts';
import type { RangeRef } from '../engine/types.ts';

export interface LinkScope {
  knownSheets: ReadonlySet<string>;
  /** Sheet for bare references, or null when bare references aren't linked. */
  defaultSheet: string | null;
}

export function linkScopeForLab(lab: LabConfig, sheetOrder: readonly string[]): LinkScope {
  return {
    knownSheets: new Set(sheetOrder),
    defaultSheet: lab.primarySheets.length === 1 ? lab.primarySheets[0]! : null,
  };
}

const ADDR = '[A-Z]{1,3}[1-9][0-9]*';
const QUALIFIED = new RegExp(`^([A-Za-z][A-Za-z0-9_]*)!(${ADDR})(?::(${ADDR}))?$`);
const BARE = new RegExp(`^(${ADDR})(?::(${ADDR}))?$`);
// In running text: not glued to other reference characters on either side.
const QUALIFIED_IN_TEXT = new RegExp(
  `(?<![\\w$!'.])([A-Za-z][A-Za-z0-9_]*)!(${ADDR})(?::(${ADDR}))?(?![\\w$:!(])`,
  'g',
);

/** The range a whole code span refers to, or null if the span isn't a linkable reference. */
export function resolveCodeRef(text: string, scope: LinkScope): RangeRef | null {
  const trimmed = text.trim();
  if (trimmed.includes('$')) return null;
  const qualified = QUALIFIED.exec(trimmed);
  if (qualified) {
    const [, sheet, start, end] = qualified;
    return scope.knownSheets.has(sheet!)
      ? { sheet: sheet!, start: start!, end: end ?? start! }
      : null;
  }
  const bare = BARE.exec(trimmed);
  if (bare && scope.defaultSheet) {
    return { sheet: scope.defaultSheet, start: bare[1]!, end: bare[2] ?? bare[1]! };
  }
  return null;
}

export interface TextRef {
  /** Offsets into the text: `text.slice(index, index + length)` is the reference. */
  index: number;
  length: number;
  range: RangeRef;
}

/** Qualified references on known sheets inside plain text. */
export function findTextRefs(text: string, scope: LinkScope): TextRef[] {
  const refs: TextRef[] = [];
  for (const m of text.matchAll(QUALIFIED_IN_TEXT)) {
    const [whole, sheet, start, end] = m;
    if (!scope.knownSheets.has(sheet!)) continue;
    refs.push({
      index: m.index,
      length: whole.length,
      range: { sheet: sheet!, start: start!, end: end ?? start! },
    });
  }
  return refs;
}

/** `{ sheet: 'BvA', start: 'A5', end: 'H5' }` → `"BvA!A5:H5"`. */
export function formatRange({ sheet, start, end }: RangeRef): string {
  return start === end ? `${sheet}!${start}` : `${sheet}!${start}:${end}`;
}

/** Inverse of `formatRange`, or null if the text isn't a qualified cell or range. */
export function parseRange(text: string): RangeRef | null {
  const m = QUALIFIED.exec(text);
  return m ? { sheet: m[1]!, start: m[2]!, end: m[3] ?? m[2]! } : null;
}
