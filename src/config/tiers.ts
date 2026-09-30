// Function tiers for the inspector's tier badge (EX-2). Hand-transcribed from the Labs document,
// Part 0.3 "Function tiers"; tests/tiers.test.ts checks TIER_TABLE against that table's text.

export type Tier = 'A' | 'B' | 'C' | 'D';

/** The "Examples used here" column of Part 0.3, row by row. */
export const TIER_TABLE: Record<Tier, readonly string[]> = {
  A: [
    'SUMIFS',
    'COUNTIFS',
    'AVERAGEIFS',
    'INDEX',
    'MATCH',
    'LOOKUP',
    'IF',
    'AND',
    'OR',
    'ABS',
    'ROUND',
    'SUMPRODUCT',
    'SLOPE',
    'INTERCEPT',
    'RSQ',
    'TREND',
    'EDATE',
    'MONTH',
    'TEXT',
    'N',
    'ISNUMBER',
    'IFERROR',
  ],
  B: ['FORECAST.LINEAR', 'FORECAST.ETS', 'IFS'],
  C: ['XLOOKUP', 'XMATCH', 'FILTER', 'UNIQUE', 'SORT', 'SEQUENCE', 'LET'],
  D: ['LAMBDA', 'SCAN', 'MAP', 'HSTACK', 'VSTACK', 'GROUPBY', 'PIVOTBY'],
};

/**
 * Functions the labs use that the Part 0.3 examples don't name. The classic ones have been in
 * Excel since long before 2010 (Tier A's definition); the ETS helpers are the "FORECAST.ETS
 * family" of Tier B.
 */
export const TIER_EXTRAS: Readonly<Record<string, Tier>> = {
  SUM: 'A',
  AVERAGE: 'A',
  COUNTIF: 'A',
  FORECAST: 'A',
  LINEST: 'A',
  'FORECAST.ETS.CONFINT': 'B',
  'FORECAST.ETS.SEASONALITY': 'B',
};

const TIERS = new Map<string, Tier>([
  ...Object.entries(TIER_TABLE).flatMap(([tier, names]) =>
    names.map((name): [string, Tier] => [name, tier as Tier]),
  ),
  ...Object.entries(TIER_EXTRAS),
]);

export function tierOf(fn: string): Tier | undefined {
  return TIERS.get(fn.toUpperCase());
}

const STRING_LITERAL = /"(?:[^"]|"")*"/g;
const FUNCTION_CALL = /\b([A-Za-z][A-Za-z0-9]*(?:\.[A-Za-z0-9]+)*)\s*\(/g;

/** Function names called by a formula, uppercased, in order of first use. */
export function formulaFunctions(formula: string): string[] {
  const code = formula.replace(STRING_LITERAL, '""');
  return [...new Set([...code.matchAll(FUNCTION_CALL)].map((m) => m[1]!.toUpperCase()))];
}

/**
 * The highest tier among a formula's functions: a formula is only as portable as its newest
 * function. Plain arithmetic (`=D6-C6`) runs everywhere, so it's Tier A. `null` when the formula
 * calls a function with no known tier.
 */
export function formulaTier(formula: string): Tier | null {
  let highest: Tier = 'A';
  for (const fn of formulaFunctions(formula)) {
    const tier = tierOf(fn);
    if (!tier) return null;
    if (tier > highest) highest = tier;
  }
  return highest;
}
