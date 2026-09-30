// Hand-authored, editorial lab configuration (PRD §4.5, §9.3).
// Lives outside src/data/ so that directory stays 100% generated (OPEN_ISSUES #3).
// Titles and slugs are not repeated here: they come from the Labs document via
// src/data/manifest.json. Control labels come from the workbook's own label cells.
// tests/labs-config.test.ts checks this file against the Labs document and the generated data.

export interface LabConfig {
  n: number;
  /** Sheet tabs shown for the lab; the first is the primary sheet shown on open. */
  sheets: string[];
  /** Sheets whose assertions this lab owns; Build mode clears their formulas (BD-1). */
  primarySheets: string[];
  /** What-if controls (WI-1), from Appendix B.4; ranges approved in OPEN_ISSUES #28. */
  controls: ControlGroup[];
  /** Display-only conditional formats (GR-12), from the lessons' "Excel UI features". */
  conditionalFormats: ConditionalFormat[];
}

// --- What-if controls ---------------------------------------------------------------------------

interface ControlBase {
  /** The input cell, e.g. "BvA!B2". */
  cell: string;
  /** Same-sheet cell whose text labels the control. Default: column A of the same row. */
  labelCell?: string;
}

export interface SliderControl extends ControlBase {
  type: 'slider';
  min: number;
  max: number;
  step: number;
  /** Excel format code for the value shown next to the slider. */
  fmt: string;
}

export interface NumberControl extends ControlBase {
  type: 'number';
  min?: number;
  max?: number;
  step: number;
  fmt: string;
}

export interface SelectControl extends ControlBase {
  type: 'select';
  /** Text values, e.g. periods "2026-08" (text, not dates: Part 0.2). */
  options: string[];
}

export interface SegmentedControl extends ControlBase {
  type: 'segmented';
  options: { value: number; label: string }[];
}

export interface DateControl extends ControlBase {
  type: 'date';
  /** YYYY-MM-DD */
  min: string;
  max: string;
}

/** A block of input cells edited as a small table: `rowLabels` down the side, `colLabels` on top. */
export interface TableControl {
  type: 'table';
  /** e.g. "PVM!B5:E7" */
  range: string;
  /** Same-sheet range holding one label per row, e.g. "A5:A7". */
  rowLabels: string;
  /** Same-sheet range holding one label per column. Omitted for a single-column series. */
  colLabels?: string;
  /** Per column, left to right. */
  columns: { min?: number; step: number; fmt: string }[];
  /** Long series start collapsed. */
  collapsed?: boolean;
}

/** A button that applies an Appendix E experiment's changes, and restores the cells on a second press. */
export interface ExperimentAction {
  type: 'experiment';
  /** Experiment id, e.g. "E9.1"; its `changes` come from src/data/experiments.json. */
  experiment: string;
  label: string;
  undoLabel: string;
}

export type Control =
  | SliderControl
  | NumberControl
  | SelectControl
  | SegmentedControl
  | DateControl
  | TableControl
  | ExperimentAction;

export interface ControlGroup {
  /** Same-sheet-qualified cell whose text titles the group, e.g. "Drivers!A4". */
  titleCell?: string;
  controls: Control[];
}

// --- Conditional formats ------------------------------------------------------------------------

/** Named styles; colors live in src/styles/tokens.css with dark-mode variants. */
export type FormatStyle = 'redFill' | 'orangeFill' | 'yellowFill' | 'greenFill' | 'bold';

/**
 * Excel's "Use a formula" rule: the formula is written for the range's top-left cell, and relative
 * references move with each cell, as in Excel. The engine evaluates it (PRD §9.3).
 */
export interface FormulaRule {
  type: 'formula';
  range: string;
  formula: string;
  style: FormatStyle;
}

/** Excel's two-color scale: lowest number → `low` color, highest → `high` color. */
export interface ColorScaleRule {
  type: 'colorScale';
  range: string;
  low: 'green';
  high: 'red';
}

export type ConditionalFormat = FormulaRule | ColorScaleRule;

// --- Labs ---------------------------------------------------------------------------------------

const PERIODS = ['2025-08', '2026-07', '2026-08'];

/**
 * Lab 10 driver sliders, one row per driver. Churn starts at 0.5%, not 0: the closed form in
 * Drivers!B21:D21 divides by churn.
 */
const DRIVER_SLIDERS: Record<
  '4' | '5' | '6',
  Pick<SliderControl, 'min' | 'max' | 'step' | 'fmt'>
> = {
  '4': { min: 0, max: 200, step: 5, fmt: '#,##0' },
  '5': { min: 0.005, max: 0.1, step: 0.005, fmt: '0.0%' },
  '6': { min: 50, max: 150, step: 1, fmt: '#,##0.00' },
};

export const labs: LabConfig[] = [
  { n: 1, sheets: ['Map', 'GL'], primarySheets: [], controls: [], conditionalFormats: [] },
  {
    n: 2,
    sheets: ['BvA', 'GL', 'Map'],
    primarySheets: ['BvA'],
    controls: [
      {
        controls: [
          { type: 'select', cell: 'BvA!B1', options: PERIODS },
          { type: 'slider', cell: 'BvA!B2', min: 0.01, max: 0.2, step: 0.01, fmt: '0.0%' },
          { type: 'slider', cell: 'BvA!B3', min: 0, max: 25000, step: 500, fmt: '#,##0' },
        ],
      },
    ],
    conditionalFormats: [
      { type: 'formula', range: 'BvA!A6:H12', formula: '=$G6="Unfavorable"', style: 'redFill' },
      { type: 'formula', range: 'BvA!A6:H12', formula: '=$H6=TRUE', style: 'bold' },
    ],
  },
  {
    n: 3,
    sheets: ['Flux', 'GL'],
    primarySheets: ['Flux'],
    controls: [
      {
        controls: [
          { type: 'slider', cell: 'Flux!B4', min: 0.01, max: 0.5, step: 0.01, fmt: '0.0%' },
          { type: 'slider', cell: 'Flux!B5', min: 0, max: 40000, step: 500, fmt: '#,##0' },
        ],
      },
    ],
    conditionalFormats: [],
  },
  {
    n: 4,
    sheets: ['PVM'],
    primarySheets: ['PVM'],
    controls: [
      {
        controls: [
          {
            type: 'table',
            range: 'PVM!B5:E7',
            rowLabels: 'A5:A7',
            colLabels: 'B4:E4',
            columns: [
              { min: 0, step: 1, fmt: '#,##0' },
              { min: 0, step: 0.01, fmt: '#,##0.00' },
              { min: 0, step: 1, fmt: '#,##0' },
              { min: 0, step: 0.01, fmt: '#,##0.00' },
            ],
          },
        ],
      },
    ],
    conditionalFormats: [],
  },
  {
    n: 5,
    sheets: ['CostVar'],
    primarySheets: ['CostVar'],
    controls: [
      {
        titleCell: 'CostVar!A1',
        controls: [
          { type: 'number', cell: 'CostVar!B3', min: 0, step: 100, fmt: '#,##0' },
          { type: 'number', cell: 'CostVar!B4', min: 0, step: 0.25, fmt: '#,##0.00' },
          { type: 'number', cell: 'CostVar!B5', min: 0, step: 100, fmt: '#,##0' },
          { type: 'number', cell: 'CostVar!B6', min: 0, step: 0.25, fmt: '#,##0.00' },
        ],
      },
      {
        titleCell: 'CostVar!A15',
        controls: [
          { type: 'number', cell: 'CostVar!B17', min: 0, step: 500, fmt: '#,##0' },
          { type: 'number', cell: 'CostVar!B18', min: 0, step: 0.5, fmt: '#,##0.00' },
          { type: 'number', cell: 'CostVar!B19', min: 0, step: 0.5, fmt: '#,##0.00' },
          { type: 'number', cell: 'CostVar!B20', min: 0, step: 1000, fmt: '#,##0' },
          { type: 'number', cell: 'CostVar!B21', min: 0, step: 500, fmt: '#,##0' },
          { type: 'number', cell: 'CostVar!B22', min: 0, step: 500, fmt: '#,##0' },
          { type: 'number', cell: 'CostVar!B23', min: 0, step: 500, fmt: '#,##0' },
          { type: 'number', cell: 'CostVar!B24', min: 0, step: 1000, fmt: '#,##0' },
        ],
      },
    ],
    conditionalFormats: [],
  },
  {
    n: 6,
    sheets: ['PivotLab', 'GL'],
    primarySheets: ['PivotLab'],
    controls: [
      {
        controls: [
          { type: 'select', cell: 'PivotLab!B2', options: PERIODS },
          { type: 'select', cell: 'PivotLab!B3', options: ['ACT', 'BUD'] },
        ],
      },
    ],
    conditionalFormats: [],
  },
  {
    n: 7,
    sheets: ['BankRec', 'Book', 'Bank'],
    primarySheets: ['Book', 'Bank', 'BankRec'],
    // Appendix B.4's "toggle to add a transaction row" is left out: the matching formulas use
    // fixed ranges (Bank!$D$2:$D$5), so a new row would be ignored (OPEN_ISSUES #29).
    controls: [
      {
        controls: [
          { type: 'number', cell: 'BankRec!B2', min: 0, max: 10, step: 0.5, fmt: '#,##0.00' },
          { type: 'number', cell: 'BankRec!B3', min: 0, max: 10, step: 1, fmt: '0' },
          { type: 'number', cell: 'BankRec!B5', step: 100, fmt: '#,##0.00' },
          { type: 'number', cell: 'BankRec!B6', step: 100, fmt: '#,##0.00' },
        ],
      },
    ],
    conditionalFormats: [
      { type: 'formula', range: 'Book!A2:G6', formula: '=$E2="Unmatched"', style: 'orangeFill' },
      { type: 'formula', range: 'Book!A2:G6', formula: '=$E2="Tolerance"', style: 'yellowFill' },
      // "Green fill when =$B$19=0 and red otherwise."
      { type: 'formula', range: 'BankRec!B20', formula: '=$B$19=0', style: 'greenFill' },
      { type: 'formula', range: 'BankRec!B20', formula: '=$B$19<>0', style: 'redFill' },
    ],
  },
  {
    n: 8,
    sheets: ['AR', 'ARSummary'],
    primarySheets: ['AR', 'ARSummary'],
    controls: [
      {
        controls: [
          { type: 'date', cell: 'AR!H2', labelCell: 'H1', min: '2026-05-01', max: '2026-12-31' },
        ],
      },
    ],
    conditionalFormats: [{ type: 'colorScale', range: 'AR!E2:E7', low: 'green', high: 'red' }],
  },
  {
    n: 9,
    sheets: ['Forecast'],
    primarySheets: ['Forecast'],
    controls: [
      {
        controls: [
          {
            type: 'experiment',
            experiment: 'E9.1',
            label: 'Inject outlier',
            undoLabel: 'Remove outlier',
          },
          {
            type: 'table',
            range: 'Forecast!C5:C34',
            rowLabels: 'A5:A34',
            colLabels: 'C4:C4',
            columns: [{ min: 0, step: 1000, fmt: '#,##0' }],
            collapsed: true,
          },
        ],
      },
    ],
    conditionalFormats: [],
  },
  {
    n: 10,
    sheets: ['Drivers'],
    primarySheets: ['Drivers'],
    controls: [
      {
        controls: [
          {
            type: 'segmented',
            cell: 'Drivers!C8',
            options: [
              { value: 1, label: 'Down' },
              { value: 2, label: 'Base' },
              { value: 3, label: 'Up' },
            ],
          },
        ],
      },
      ...(['4', '5', '6'] as const).map((row) => ({
        titleCell: `Drivers!A${row}`,
        controls: (['B', 'C', 'D'] as const).map((col): SliderControl => ({
          type: 'slider',
          cell: `Drivers!${col}${row}`,
          labelCell: `${col}3`,
          ...DRIVER_SLIDERS[row],
        })),
      })),
    ],
    conditionalFormats: [],
  },
  // Lab 11 is a sandbox for the simulated LAMBDA library (PRD §7.5); no workbook sheets.
  { n: 11, sheets: [], primarySheets: [], controls: [], conditionalFormats: [] },
  { n: 12, sheets: ['Checks'], primarySheets: ['Checks'], controls: [], conditionalFormats: [] },
];
