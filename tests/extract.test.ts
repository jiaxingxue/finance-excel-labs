import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { extractContent, renderFiles, slugify } from '../scripts/extract/content.ts';
import { fences, sections, toLines } from '../scripts/extract/markdown.ts';
import { checkCell, checkExperiment, ContentError } from '../scripts/extract/validate.ts';

const root = join(import.meta.dirname, '..');
const doc = readFileSync(
  join(root, 'content', 'Excel_Implementation_Labs_Financial_Analysis.md'),
  'utf8',
);

/** Replace exactly one occurrence, so a mutation can't silently miss. */
function mutate(find: string, replace: string): string {
  const parts = doc.split(find);
  if (parts.length !== 2) throw new Error(`"${find}" occurs ${parts.length - 1} times`);
  return parts.join(replace);
}

describe('markdown scanning', () => {
  const md = [
    '## One',
    '```json',
    '## not a heading',
    '```',
    '   ```excel',
    '=A1',
    '   ```',
    '### Sub',
    '## Two',
    'text',
  ].join('\n');
  const lines = toLines(md);

  it('ignores headings inside fences, including indented fences', () => {
    expect(sections(lines, 2).map((s) => s.heading)).toEqual(['One', 'Two']);
    expect(sections(lines, 2)[0]).toMatchObject({ start: 0, end: 8 });
  });

  it('finds fences with their language and body', () => {
    expect(fences(lines).map((f) => [f.lang, f.body])).toEqual([
      ['json', '## not a heading'],
      ['excel', '=A1'],
    ]);
  });

  it('normalizes CRLF line endings', () => {
    expect(toLines('a\r\nb\rc')).toEqual(['a', 'b', 'c']);
  });

  it('rejects an unclosed fence', () => {
    expect(() => fences(toLines('```json\n{}'))).toThrow(/Unclosed/);
  });
});

describe('cell and experiment validation', () => {
  it('accepts constants, dates, and formulas', () => {
    expect(() => checkCell('S', 'A1', { v: 1, fmt: '0' })).not.toThrow();
    expect(() => checkCell('S', 'A1', { v: { date: '2026-06-30' } })).not.toThrow();
    expect(() => checkCell('S', 'A1', { f: '=1+1' })).not.toThrow();
  });

  it('rejects malformed cells', () => {
    expect(() => checkCell('S', 'a1', { v: 1 })).toThrow(ContentError);
    expect(() => checkCell('S', 'A1', { v: 1, f: '=1' })).toThrow(/exactly one/);
    expect(() => checkCell('S', 'A1', { f: '1+1' })).toThrow(/starting with "="/);
    expect(() => checkCell('S', 'A1', { v: { date: '2026-06' } })).toThrow(ContentError);
    expect(() => checkCell('S', 'A1', { v: 1, style: 'x' })).toThrow(/unknown key/);
  });

  it('allows null in expect (empty cell) but not in changes', () => {
    const base = { id: 'E', lab: 1, title: 't', changes: { 'S!A1': 1 }, expect: { 'S!B1': null } };
    expect(() => checkExperiment(0, base)).not.toThrow();
    expect(() => checkExperiment(0, { ...base, changes: { 'S!A1': null } })).toThrow(ContentError);
  });
});

describe('extractContent on the Labs document', () => {
  const content = extractContent(doc);

  it('produces the counts PRD §11 #1 requires', () => {
    expect(content.manifest.counts).toMatchObject({ sheets: 15, assertions: 277, experiments: 12 });
    expect(content.manifest.counts.formulas).toBe(668);
  });

  it('keeps dates as {date} objects (converted to serials at engine load, OPEN_ISSUES #2)', () => {
    const dates = Object.values(content.workbook.sheets)
      .flatMap((s) => Object.values(s))
      .filter((c) => 'v' in c && typeof c.v === 'object');
    expect(dates.length).toBeGreaterThan(0);
    for (const c of dates)
      expect(c).toMatchObject({ v: { date: expect.stringMatching(/^\d{4}-\d{2}-\d{2}$/) } });
  });

  it('extracts Part 0, Labs 1–12, and Appendices A–B as lessons, not C–E', () => {
    expect(content.lessons.map((l) => l.id)).toEqual([
      'part-0',
      ...Array.from({ length: 12 }, (_, i) => `lab-${String(i + 1).padStart(2, '0')}`),
      'appendix-a',
      'appendix-b',
    ]);
    const lab2 = content.lessons.find((l) => l.id === 'lab-02')!;
    expect(lab2.title).toBe('Budget vs. Actual Variance Report');
    expect(lab2.markdown.startsWith('## Lab 2 — Budget vs. Actual Variance Report\n')).toBe(true);
    expect(lab2.markdown).not.toContain('## Lab 3');
  });

  it('keeps lesson markdown verbatim', () => {
    for (const lesson of content.lessons) expect(doc).toContain(lesson.markdown.trimEnd());
  });

  it('marks pattern cells from lab formula tables', () => {
    const c6 = content.exercises.find((e) => e.sheet === 'BvA' && e.cell === 'C6');
    expect(c6).toMatchObject({ lab: 2, result: '500,000' });
    expect(content.workbook.sheets.BvA?.C6).toMatchObject({ f: c6?.formula });
    expect(new Set(content.exercises.map((e) => e.lab)).has(12)).toBe(true);
  });

  it('slugifies titles, including en dashes', () => {
    expect(slugify('Price–Volume–Mix Analysis and Variance Bridge')).toBe(
      'price-volume-mix-analysis-and-variance-bridge',
    );
  });
});

describe('extractContent fails loudly on content drift', () => {
  it('when a sheet heading miscounts its cells', () => {
    expect(() =>
      extractContent(mutate('#### Sheet `Map` (28 cells)', '#### Sheet `Map` (29 cells)')),
    ).toThrow(/heading says 29 cells, block has 28/);
  });

  it('when an assertion goes missing', () => {
    const md = mutate(
      '  {"sheet": "BvA", "cell": "C6", "expected": 500000, "tolerance": 0.005},\n',
      '',
    );
    expect(() => extractContent(md)).toThrow(/Expected 277 assertions, found 276/);
  });

  it('when a json block appears twice in Appendix D (JSON scoping, OPEN_ISSUES #8)', () => {
    const md = mutate(
      'Each expected value was calculated independently',
      '```json\n[]\n```\n\nEach expected value was calculated independently',
    );
    expect(() => extractContent(md)).toThrow(/Appendix D: expected one json block, found 2/);
  });

  it('when a lesson formula disagrees with Appendix C', () => {
    const md = mutate('| `BvA!E6` | `=D6-C6` |', '| `BvA!E6` | `=C6-D6` |');
    expect(() => extractContent(md)).toThrow(/BvA!E6 is =C6-D6 in the lesson but differs/);
  });

  it('when an assertion targets a cell that is not in the workbook', () => {
    const md = mutate(
      '{"sheet": "BvA", "cell": "C6", "expected": 500000',
      '{"sheet": "BvA", "cell": "Z99", "expected": 500000',
    );
    expect(() => extractContent(md)).toThrow(/BvA!Z99 is not in Appendix C/);
  });
});

describe('committed src/data/', () => {
  it('matches what the Labs document generates (same check as `npm run extract -- --check`)', () => {
    for (const [path, content] of renderFiles(doc)) {
      const full = join(root, 'src', 'data', path);
      expect(existsSync(full), path).toBe(true);
      expect(readFileSync(full, 'utf8'), path).toBe(content);
    }
  });
});
