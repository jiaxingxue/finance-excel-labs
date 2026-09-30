import { describe, expect, it } from 'vitest';
import { isoDateToSerial } from '../src/engine/index.ts';
import { parseInput } from '../src/workspace/parseInput.ts';

describe('parseInput (GR-4)', () => {
  it.each([
    ['', null],
    ['   ', null],
    ['=SUM(A1:A3)', '=SUM(A1:A3)'],
    ['TRUE', true],
    ['false', false],
    ['1000', 1000],
    ['1,000', 1000],
    ['-2,900.5', -2900.5],
    ['5%', 0.05],
    ['.25', 0.25],
    ['1e3', 1000],
    ['2026-08', '2026-08'],
    ['Revenue', 'Revenue'],
    ['1,00', '1,00'],
    ['%', '%'],
  ])('%j → %j', (typed, content) => {
    expect(parseInput(typed)).toEqual(content);
  });

  it('turns a full YYYY-MM-DD date into a serial, like experiment changes', () => {
    expect(parseInput('2026-09-30')).toBe(isoDateToSerial('2026-09-30'));
    expect(parseInput('2026-02-30')).toBe('2026-02-30');
  });
});
