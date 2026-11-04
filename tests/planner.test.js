import { describe, expect, test } from 'vitest';
import { chooseMode, occurrenceIndex, planAutoConversion } from '../addin/src/planner.js';

describe('occurrenceIndex', () => {
  test('counts raw matches, including ones inside longer formulas', () => {
    const text = 'H2O a H2 a H2';
    expect(occurrenceIndex(text, 'H2', 6)).toBe(1);
    expect(occurrenceIndex(text, 'H2', 11)).toBe(2);
  });

  test('returns zero for the first match', () => {
    expect(occurrenceIndex('H2', 'H2', 0)).toBe(0);
  });
});

describe('chooseMode', () => {
  test('is lenient for a single selected token', () => {
    expect(chooseMode('B12')).toBe('lenient');
  });

  test('is strict when the selection has several words', () => {
    expect(chooseMode('vitamin B12 a H2O')).toBe('strict');
  });

  test('ignores surrounding whitespace', () => {
    expect(chooseMode(' B12 ')).toBe('lenient');
  });
});

describe('planAutoConversion', () => {
  test('converts formulas seen for the first time', () => {
    const { toConvert, seen } = planAutoConversion('Voda H2O a CO2', new Map());
    expect(toConvert.map((hit) => hit.source)).toEqual(['H2O', 'CO2']);
    expect(seen.get('H2O')).toBe(1);
  });

  test('does not mutate the previous state', () => {
    const previous = new Map([['H2O', 1]]);
    planAutoConversion('H2O H2O', previous);
    expect(previous.get('H2O')).toBe(1);
  });

  test('skips a formula the user reverted with undo', () => {
    const { toConvert } = planAutoConversion('Voda H2O', new Map([['H2O', 1]]));
    expect(toConvert).toEqual([]);
  });

  test('converts only the new copy, leaving one the user undid', () => {
    const { toConvert, seen } = planAutoConversion('H2O a H2O ', new Map([['H2O', 1]]));
    expect(toConvert.map((hit) => hit.start)).toEqual([6]);
    expect(seen.get('H2O')).toBe(2);
  });

  test('picks the copy next to the caret when it was typed before an older one', () => {
    const text = 'H2O a H2O';
    const { toConvert } = planAutoConversion(text, new Map([['H2O', 1]]), 4);
    expect(toConvert.map((hit) => hit.start)).toEqual([0]);
  });

  test('treats the hydrate dot and its typed form as one formula', () => {
    const { toConvert } = planAutoConversion('CuSO4.5H2O', new Map([['CuSO4·5H2O', 1]]));
    expect(toConvert).toEqual([]);
  });

  test('converts a formula retyped after being deleted', () => {
    const afterDelete = planAutoConversion('nic', new Map([['H2O', 1]]));
    const { toConvert } = planAutoConversion('H2O', afterDelete.seen);
    expect(toConvert).toHaveLength(1);
  });
});
