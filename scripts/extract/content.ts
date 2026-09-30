// Labs document → generated data (PRD §4.2). Pure: markdown in, file contents out.

import { createHash } from 'node:crypto';
import type {
  Assertion,
  ContentManifest,
  Exercise,
  Experiment,
  Lesson,
  WorkbookSpec,
} from '../../src/content/types.ts';
import { fences, requireSection, sections, toLines, type Section } from './markdown.ts';
import { checkAssertion, checkCell, checkExperiment, fail, parseJson } from './validate.ts';

/** Counts the PRD fixes (§4.2, §11 #1). A different count means the document or extractor changed. */
export const EXPECTED_COUNTS = { sheets: 15, assertions: 277, experiments: 12 } as const;

export interface ExtractedContent {
  workbook: WorkbookSpec;
  assertions: Assertion[];
  experiments: Experiment[];
  exercises: Exercise[];
  lessons: Lesson[];
  manifest: ContentManifest;
}

const SHEET_HEADING = /^Sheet `([^`]+)` \((\d+) cells\)$/;
const LAB_HEADING = /^Lab (\d+) — (.+)$/;
const PART_HEADING = /^Part (\d+) — (.+)$/;
const APPENDIX_HEADING = /^Appendix ([A-Z]) — (.+)$/;
/** Appendices that are data, not reading content (PRD §4.2 lessons row: Part 0, Labs, A–B). */
const DATA_APPENDICES = new Set(['C', 'D', 'E']);
const EXERCISE_ROW =
  /^\| `([A-Za-z][A-Za-z0-9_]*)!([A-Z]{1,3}[1-9]\d*)` \| `(=[^`]*)` \| ([^|]*)\|\s*$/;

export function slugify(title: string): string {
  return title
    .normalize('NFKD')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '');
}

function extractWorkbook(lines: string[], appendixC: Section): WorkbookSpec {
  const workbook: WorkbookSpec = { sheetOrder: [], sheets: {} };
  const sheetSections = sections(lines, 4, appendixC.start, appendixC.end);
  for (const s of sheetSections) {
    const m = SHEET_HEADING.exec(s.heading);
    if (!m) fail(`Appendix C: unexpected heading "#### ${s.heading}"`);
    const name = m[1]!;
    const declared = Number(m[2]);
    if (name in workbook.sheets) fail(`Appendix C: sheet ${name} appears twice`);
    const blocks = fences(lines, s.start + 1, s.end).filter((f) => f.lang === 'json');
    if (blocks.length !== 1)
      fail(`Appendix C: sheet ${name} needs one json block, found ${blocks.length}`);
    const raw = parseJson(`Sheet ${name}`, blocks[0]!.body);
    if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) {
      fail(`Sheet ${name}: must be an object of cells`);
    }
    const cells: WorkbookSpec['sheets'][string] = {};
    for (const [addr, cell] of Object.entries(raw)) cells[addr] = checkCell(name, addr, cell);
    const count = Object.keys(cells).length;
    if (count !== declared)
      fail(`Sheet ${name}: heading says ${declared} cells, block has ${count}`);
    workbook.sheetOrder.push(name);
    workbook.sheets[name] = cells;
  }
  return workbook;
}

/** The single json array in an appendix section (scoped, so Appendix B's examples never match). */
function singleJsonArray(lines: string[], section: Section, label: string): unknown[] {
  const blocks = fences(lines, section.start + 1, section.end).filter((f) => f.lang === 'json');
  if (blocks.length !== 1) fail(`${label}: expected one json block, found ${blocks.length}`);
  const raw = parseJson(label, blocks[0]!.body);
  if (!Array.isArray(raw)) fail(`${label}: json block must be an array`);
  return raw;
}

function checkRefs(workbook: WorkbookSpec, where: string, sheet: string, cell: string): void {
  if (!(sheet in workbook.sheets)) fail(`${where}: unknown sheet "${sheet}"`);
  if (!(cell in workbook.sheets[sheet]!)) fail(`${where}: ${sheet}!${cell} is not in Appendix C`);
}

function splitRef(ref: string): [string, string] {
  const i = ref.indexOf('!');
  return [ref.slice(0, i), ref.slice(i + 1)];
}

/** Section markdown, trimmed of trailing blank lines and the `---` separator between sections. */
function sectionMarkdown(lines: string[], s: Section): string {
  let end = s.end;
  while (end > s.start + 1 && /^\s*(---)?\s*$/.test(lines[end - 1]!)) end--;
  return lines.slice(s.start, end).join('\n') + '\n';
}

function extractLessons(lines: string[], top: Section[]): Lesson[] {
  const lessons: Lesson[] = [];
  for (const s of top) {
    let lab: RegExpExecArray | null;
    let part: RegExpExecArray | null;
    let appendix: RegExpExecArray | null;
    if ((lab = LAB_HEADING.exec(s.heading))) {
      const n = Number(lab[1]);
      lessons.push({
        id: `lab-${String(n).padStart(2, '0')}`,
        kind: 'lab',
        n,
        title: lab[2]!,
        slug: slugify(lab[2]!),
        markdown: sectionMarkdown(lines, s),
      });
    } else if ((part = PART_HEADING.exec(s.heading))) {
      lessons.push({
        id: `part-${part[1]}`,
        kind: 'part',
        n: null,
        title: part[2]!,
        slug: slugify(part[2]!),
        markdown: sectionMarkdown(lines, s),
      });
    } else if ((appendix = APPENDIX_HEADING.exec(s.heading))) {
      if (DATA_APPENDICES.has(appendix[1]!)) continue;
      lessons.push({
        id: `appendix-${appendix[1]!.toLowerCase()}`,
        kind: 'appendix',
        n: null,
        title: appendix[2]!,
        slug: slugify(appendix[2]!),
        markdown: sectionMarkdown(lines, s),
      });
    }
  }
  const labs = lessons.filter((l) => l.kind === 'lab').map((l) => l.n);
  const want = Array.from({ length: 12 }, (_, i) => i + 1);
  if (labs.join() !== want.join()) fail(`Expected Labs 1–12 in order, found ${labs.join(', ')}`);
  return lessons;
}

function extractExercises(lines: string[], top: Section[], workbook: WorkbookSpec): Exercise[] {
  const out: Exercise[] = [];
  const seen = new Set<string>();
  for (const s of top) {
    const lab = LAB_HEADING.exec(s.heading);
    if (!lab) continue;
    for (let i = s.start; i < s.end; i++) {
      const m = EXERCISE_ROW.exec(lines[i]!);
      if (!m) continue;
      const [, sheet, cell, formula, result] = m as unknown as [
        string,
        string,
        string,
        string,
        string,
      ];
      const where = `Lab ${lab[1]} formula table, line ${i + 1}`;
      checkRefs(workbook, where, sheet, cell);
      const spec = workbook.sheets[sheet]![cell]!;
      if (!('f' in spec) || spec.f !== formula) {
        fail(`${where}: ${sheet}!${cell} is ${formula} in the lesson but differs in Appendix C`);
      }
      const key = `${sheet}!${cell}`;
      if (seen.has(key)) fail(`${where}: ${key} is listed twice`);
      seen.add(key);
      out.push({ lab: Number(lab[1]), sheet, cell, formula, result: result.trim() });
    }
  }
  return out;
}

export function extractContent(markdown: string): ExtractedContent {
  const lines = toLines(markdown);
  const top = sections(lines, 2);

  const workbook = extractWorkbook(lines, requireSection(top, 'Appendix C —'));
  if (workbook.sheetOrder.length !== EXPECTED_COUNTS.sheets) {
    fail(`Expected ${EXPECTED_COUNTS.sheets} sheets, found ${workbook.sheetOrder.length}`);
  }

  const assertions = singleJsonArray(lines, requireSection(top, 'Appendix D —'), 'Appendix D').map(
    (raw, i) => checkAssertion(i, raw),
  );
  if (assertions.length !== EXPECTED_COUNTS.assertions) {
    fail(`Expected ${EXPECTED_COUNTS.assertions} assertions, found ${assertions.length}`);
  }
  assertions.forEach((a, i) => checkRefs(workbook, `Assertion #${i + 1}`, a.sheet, a.cell));

  const experiments = singleJsonArray(lines, requireSection(top, 'Appendix E —'), 'Appendix E').map(
    (raw, i) => checkExperiment(i, raw),
  );
  if (experiments.length !== EXPECTED_COUNTS.experiments) {
    fail(`Expected ${EXPECTED_COUNTS.experiments} experiments, found ${experiments.length}`);
  }
  for (const e of experiments) {
    for (const ref of [...Object.keys(e.changes), ...Object.keys(e.expect)]) {
      checkRefs(workbook, e.id, ...splitRef(ref));
    }
  }

  const lessons = extractLessons(lines, top);
  const exercises = extractExercises(lines, top, workbook);

  let cells = 0;
  let formulas = 0;
  for (const sheet of Object.values(workbook.sheets)) {
    for (const cell of Object.values(sheet)) {
      cells++;
      if ('f' in cell) formulas++;
    }
  }

  const manifest: ContentManifest = {
    sourceSha256: createHash('sha256').update(markdown).digest('hex'),
    counts: {
      sheets: workbook.sheetOrder.length,
      cells,
      formulas,
      assertions: assertions.length,
      experiments: experiments.length,
      exercises: exercises.length,
    },
    lessons: lessons.map(({ markdown: _markdown, ...summary }) => summary),
  };

  return { workbook, assertions, experiments, exercises, lessons, manifest };
}

/** The verification notice at the top of the document (PRD LS-4): the blockquote before the TOC. */
export function extractNotice(markdown: string): string {
  const lines = toLines(markdown);
  const toc = requireSection(sections(lines, 2), 'Table of Contents');
  const quote = lines.slice(0, toc.start).filter((l) => l.startsWith('>'));
  if (quote.length === 0) fail('No verification notice found before the Table of Contents');
  return quote.join('\n') + '\n';
}

const json = (x: unknown): string => JSON.stringify(x, null, 2) + '\n';

/** Every generated file, keyed by path relative to src/data/. */
export function renderFiles(markdown: string): Map<string, string> {
  const c = extractContent(markdown);
  const files = new Map<string, string>([
    ['workbook.json', json(c.workbook)],
    ['assertions.json', json(c.assertions)],
    ['experiments.json', json(c.experiments)],
    ['exercises.json', json(c.exercises)],
    ['manifest.json', json(c.manifest)],
    ['notice.json', json({ markdown: extractNotice(markdown) })],
  ]);
  for (const lesson of c.lessons) files.set(`lessons/${lesson.id}.json`, json(lesson));
  return files;
}
