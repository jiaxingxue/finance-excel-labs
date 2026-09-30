// Shapes of the generated files in src/data/ (PRD §4.2–4.4).
// Shared by the extractor (scripts/) and the app, so type-only exports here.

export type CellSpec =
  | { v: string | number | boolean; fmt?: string }
  | { v: { date: string }; fmt?: string }
  | { f: string; fmt?: string };

/** Cells keyed by A1 address. Blank cells are absent. */
export type SheetSpec = Record<string, CellSpec>;

export interface WorkbookSpec {
  sheetOrder: string[];
  sheets: Record<string, SheetSpec>;
}

export interface Assertion {
  sheet: string;
  cell: string;
  expected: number | string | boolean;
  tolerance: number;
}

export interface Experiment {
  id: string;
  lab: number;
  title: string;
  /** "Sheet!A1" → new value. A full "YYYY-MM-DD" string is a date; anything else is a literal. */
  changes: Record<string, number | string>;
  /** "Sheet!A1" → expected value. `null` means empty (blank or ""). */
  expect: Record<string, number | string | boolean | null>;
}

/** A row `| \`SHEET!ADDR\` | \`=FORMULA\` | RESULT |` in a lab's formula table (PRD §4.2). */
export interface Exercise {
  lab: number;
  sheet: string;
  cell: string;
  formula: string;
  /** The Result column as written in the lesson (display text, not a graded value). */
  result: string;
}

export type LessonKind = 'part' | 'lab' | 'appendix';

export interface Lesson {
  id: string;
  kind: LessonKind;
  /** Lab number for labs; null otherwise. */
  n: number | null;
  title: string;
  slug: string;
  /** Raw markdown of the section, heading included, verbatim from the Labs document. */
  markdown: string;
}

export type LessonSummary = Omit<Lesson, 'markdown'>;

export interface ContentManifest {
  /** SHA-256 of the Labs document the data was generated from. */
  sourceSha256: string;
  counts: {
    sheets: number;
    cells: number;
    formulas: number;
    assertions: number;
    experiments: number;
    exercises: number;
  };
  lessons: LessonSummary[];
}
