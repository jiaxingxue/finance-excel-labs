// Hand-authored, editorial lab configuration (PRD §4.5, §9.3).
// Lives outside src/data/ so that directory stays 100% generated (OPEN_ISSUES #3).
// Titles and slugs are not repeated here: they come from the Labs document via
// src/data/manifest.json. Later milestones add fills, controls, conditional formats, and charts.

export interface LabConfig {
  n: number;
  /** Sheet tabs shown for the lab; the first is the primary sheet shown on open. */
  sheets: string[];
  /** Sheets whose assertions this lab owns; Build mode clears their formulas (BD-1). */
  primarySheets: string[];
}

export const labs: LabConfig[] = [
  { n: 1, sheets: ['Map', 'GL'], primarySheets: [] },
  { n: 2, sheets: ['BvA', 'GL', 'Map'], primarySheets: ['BvA'] },
  { n: 3, sheets: ['Flux', 'GL'], primarySheets: ['Flux'] },
  { n: 4, sheets: ['PVM'], primarySheets: ['PVM'] },
  { n: 5, sheets: ['CostVar'], primarySheets: ['CostVar'] },
  { n: 6, sheets: ['PivotLab', 'GL'], primarySheets: ['PivotLab'] },
  { n: 7, sheets: ['BankRec', 'Book', 'Bank'], primarySheets: ['Book', 'Bank', 'BankRec'] },
  { n: 8, sheets: ['AR', 'ARSummary'], primarySheets: ['AR', 'ARSummary'] },
  { n: 9, sheets: ['Forecast'], primarySheets: ['Forecast'] },
  { n: 10, sheets: ['Drivers'], primarySheets: ['Drivers'] },
  // Lab 11 is a sandbox for the simulated LAMBDA library (PRD §7.5); no workbook sheets.
  { n: 11, sheets: [], primarySheets: [] },
  { n: 12, sheets: ['Checks'], primarySheets: ['Checks'] },
];
