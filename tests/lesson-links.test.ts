// Cell-link rules for lesson text (GR-11, OPEN_ISSUES #20).

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import workbook from '../src/data/workbook.json' with { type: 'json' };
import type { Lesson } from '../src/content/types.ts';
import { labs } from '../src/config/labs.config.ts';
import {
  findTextRefs,
  formatRange,
  linkScopeForLab,
  resolveCodeRef,
  type LinkScope,
} from '../src/lesson/cellRefs.ts';
import { highlightFormula } from '../src/lesson/highlightFormula.ts';
import { rehypeCellLinks } from '../src/lesson/rehypeCellLinks.ts';

const scopeOf = (n: number): LinkScope => linkScopeForLab(labs[n - 1]!, workbook.sheetOrder);
const code = (text: string, n: number) => {
  const range = resolveCodeRef(text, scopeOf(n));
  return range && formatRange(range);
};

describe('rule 1: qualified references are links in every lab', () => {
  it.each([2, 7, 8, 11])('Lab %i', (n) => {
    expect(code('BvA!C6', n)).toBe('BvA!C6');
    expect(code('Bank!A2:D5', n)).toBe('Bank!A2:D5');
  });

  it('ignores unknown sheets and non-references', () => {
    expect(code('Nope!C6', 2)).toBeNull();
    expect(code('=D6-C6', 2)).toBeNull();
    expect(code('SUMIFS', 2)).toBeNull();
    expect(code('2026-08', 2)).toBeNull();
  });

  it('finds qualified references in plain text', () => {
    const text = 'Totals tie to GL!E2:E37 and BvA!D14. See E7.1, Map!$A$2, or Nope!A1.';
    expect(findTextRefs(text, scopeOf(7)).map((r) => formatRange(r.range))).toEqual([
      'GL!E2:E37',
      'BvA!D14',
    ]);
  });
});

describe('rule 2: bare references resolve only in single-primary-sheet labs', () => {
  it.each([
    [2, 'BvA'],
    [3, 'Flux'],
    [4, 'PVM'],
    [5, 'CostVar'],
    [6, 'PivotLab'],
    [9, 'Forecast'],
    [10, 'Drivers'],
    [12, 'Checks'],
  ])('Lab %i → %s', (n, sheet) => {
    expect(code('C6', n)).toBe(`${sheet}!C6`);
    expect(code('A5:H5', n)).toBe(`${sheet}!A5:H5`);
  });

  it.each([1, 7, 8, 11])('Lab %i (not exactly one primary sheet) leaves them as text', (n) => {
    expect(code('C6', n)).toBeNull();
    expect(code('A5:H5', n)).toBeNull();
  });

  it('never links bare references in plain text', () => {
    expect(findTextRefs('Enter C12/D12, then E6:H6', scopeOf(2))).toEqual([]);
  });
});

describe('rule 3: references containing $ are formula fragments', () => {
  it.each(['$A6', '$B$1', 'C$6', 'Map!$B$2:$B$7', 'GL!$E$2:$E$37', 'Bank!A2:$D$5'])(
    '%s is not a link',
    (text) => {
      for (const n of [2, 7]) expect(code(text, n)).toBeNull();
      expect(findTextRefs(text, scopeOf(2))).toEqual([]);
    },
  );
});

describe('every link in the real lessons', () => {
  const lessonsDir = join(import.meta.dirname, '../src/data/lessons');
  const known = new Set(workbook.sheetOrder);

  it.each(labs.map((l) => l.n))('Lab %i links only to known sheets, never with $', (n) => {
    const lesson = JSON.parse(
      readFileSync(join(lessonsDir, `lab-${String(n).padStart(2, '0')}.json`), 'utf8'),
    ) as Lesson;
    const prose = lesson.markdown.replace(/```[\s\S]*?```/g, '');
    const scope = scopeOf(n);
    const links = [...prose.matchAll(/`([^`\n]+)`/g)]
      .map((m) => resolveCodeRef(m[1]!, scope))
      .filter((r) => r !== null);
    for (const r of links) {
      expect(known.has(r.sheet), formatRange(r)).toBe(true);
      expect(formatRange(r)).not.toContain('$');
      if (scope.defaultSheet === null) expect(prose).toContain(formatRange(r));
    }
  });

  it('Lab 2 links its inputs table and formula table', () => {
    const lesson = JSON.parse(readFileSync(join(lessonsDir, 'lab-02.json'), 'utf8')) as Lesson;
    const links = [...lesson.markdown.matchAll(/`([^`\n]+)`/g)]
      .map((m) => resolveCodeRef(m[1]!, scopeOf(2)))
      .filter((r) => r !== null)
      .map(formatRange);
    expect(links).toEqual(expect.arrayContaining(['BvA!A1:B3', 'BvA!B1', 'BvA!C6', 'BvA!D14']));
  });
});

describe('rehypeCellLinks', () => {
  const text = (value: string) => ({ type: 'text', value });
  const el = (tagName: string, children: unknown[]) => ({
    type: 'element',
    tagName,
    properties: {},
    children,
  });

  it('wraps code-span and plain-text references, and leaves code blocks alone', () => {
    const tree = {
      type: 'root',
      children: [
        el('p', [text('See '), el('code', [text('C6')]), text(' and GL!E2.')]),
        el('pre', [el('code', [text('=SUM(BvA!C6)')])]),
      ],
    };
    rehypeCellLinks(scopeOf(2))(tree as never);
    const p = tree.children[0]!.children as { tagName?: string; properties?: object }[];
    expect(p.map((n) => n.tagName ?? 'text')).toEqual(['text', 'a', 'text', 'a', 'text']);
    expect(p[1]!.properties).toEqual({ dataCellRef: 'BvA!C6' });
    expect(p[3]!.properties).toEqual({ dataCellRef: 'GL!E2' });
    expect(tree.children[1]).toEqual(el('pre', [el('code', [text('=SUM(BvA!C6)')])]));
  });
});

describe('highlightFormula (LS-1)', () => {
  it('marks strings, functions, and references', () => {
    const tokens = highlightFormula('=XLOOKUP($A6, Map!$A$2:$A$7, "Unmapped")');
    expect(tokens.filter((t) => t.kind !== 'text')).toEqual([
      { kind: 'function', text: 'XLOOKUP' },
      { kind: 'ref', text: '$A6' },
      { kind: 'ref', text: 'Map!$A$2:$A$7' },
      { kind: 'string', text: '"Unmapped"' },
    ]);
    expect(tokens.map((t) => t.text).join('')).toBe('=XLOOKUP($A6, Map!$A$2:$A$7, "Unmapped")');
  });

  it('handles dotted function names and leaves prose alone', () => {
    const tokens = highlightFormula('Upper bound: =FORECAST.ETS(A38, "a""b")');
    expect(tokens.filter((t) => t.kind !== 'text').map((t) => t.text)).toEqual([
      'FORECAST.ETS',
      'A38',
      '"a""b"',
    ]);
  });
});
