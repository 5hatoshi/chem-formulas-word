import { describe, expect, test } from 'vitest';
import { findFormulas, renderPlain } from '../addin/src/formula.js';

/** Shows a result compactly: _x_ marks subscript, ^x^ marks superscript. */
const show = (text, mode = 'strict') => {
  const hits = findFormulas(text, mode);
  return hits.map((hit) => renderPlain(hit.segments)).join(' | ');
};

describe('subscripts', () => {
  test.each([
    ['H2O', 'H_2_O'],
    ['H2SO4', 'H_2_SO_4_'],
    ['C6H12O6', 'C_6_H_12_O_6_'],
    ['CH3COOH', 'CH_3_COOH'],
    ['Ca(OH)2', 'Ca(OH)_2_'],
    ['Al2(SO4)3', 'Al_2_(SO_4_)_3_'],
    ['(NH4)2SO4', '(NH_4_)_2_SO_4_'],
    ['K3[Fe(CN)6]', 'K_3_[Fe(CN)_6_]'],
  ])('%s', (input, expected) => {
    expect(show(input)).toBe(expected);
  });

  test('keeps a leading coefficient at normal size', () => {
    expect(show('2H2O')).toBe('2H_2_O');
  });

  test('finds every formula in an equation', () => {
    expect(show('2H2 + O2 = 2H2O')).toBe('2H_2_ | O_2_ | 2H_2_O');
  });

  test('splits formulas joined by a plus without spaces', () => {
    expect(show('H2+O2')).toBe('H_2_ | O_2_');
  });
});

describe('charges', () => {
  test.each([
    ['Ca2+', 'Ca^2+^'],
    ['Fe3+', 'Fe^3+^'],
    ['S2-', 'S^2−^'],
    ['Na+', 'Na^+^'],
    ['Cl-', 'Cl^−^'],
    ['OH-', 'OH^−^'],
    ['H3O+', 'H_3_O^+^'],
    ['NH4+', 'NH_4_^+^'],
    ['NO3-', 'NO_3_^−^'],
    ['MnO4-', 'MnO_4_^−^'],
    ['SO42-', 'SO_4_^2−^'],
    ['PO43-', 'PO_4_^3−^'],
    ['Cr2O72-', 'Cr_2_O_7_^2−^'],
    ['Hg22+', 'Hg_2_^2+^'],
    ['Fe(CN)64-', 'Fe(CN)_6_^4−^'],
    ['[Cu(NH3)4]2+', '[Cu(NH_3_)_4_]^2+^'],
    ['[Ag(NH3)2]+', '[Ag(NH_3_)_2_]^+^'],
  ])('%s', (input, expected) => {
    expect(show(input)).toBe(expected);
  });

  test('accepts a charge before punctuation', () => {
    expect(show('ionty Na+, Cl- a Ca2+.')).toBe('Na^+^ | Cl^−^ | Ca^2+^');
  });

  test('does not read a hyphenated word as a charge', () => {
    expect(show('CO2-neutrální')).toBe('CO_2_');
  });

  test('does not read an arrow as a charge', () => {
    expect(show('H2O->')).toBe('H_2_O');
  });
});

describe('hydrate dot', () => {
  test.each([
    ['CuSO4.5H2O', 'CuSO_4_·5H_2_O'],
    ['CuSO4*5H2O', 'CuSO_4_·5H_2_O'],
    ['CuSO4·5H2O', 'CuSO_4_·5H_2_O'],
    ['CuSO4×5H2O', 'CuSO_4_·5H_2_O'],
    ['Na2CO3.10H2O', 'Na_2_CO_3_·10H_2_O'],
    ['CaSO4.1/2H2O', 'CaSO_4_·1/2H_2_O'],
    ['CuSO4 . 5H2O', 'CuSO_4_ · 5H_2_O'],
    ['CuSO4 * 5H2O', 'CuSO_4_ · 5H_2_O'],
    ['CuSO4 x 5H2O', 'CuSO_4_ · 5H_2_O'],
  ])('%s', (input, expected) => {
    expect(show(input)).toBe(expected);
  });

  test('leaves the full stop that ends a sentence', () => {
    expect(show('Vznikne H2O. 2H2 je plyn.')).toBe('H_2_O | 2H_2_');
  });

  test('leaves a full stop after a formula at the end of text', () => {
    const [hit] = findFormulas('Vznikne H2O.', 'strict');
    expect(hit.source).toBe('H2O');
  });

  test('requires water after a spaced x', () => {
    expect(show('CO2 x NO2')).toBe('CO_2_ | NO_2_');
  });
});

describe('text that must stay untouched in strict mode', () => {
  test.each([
    'papír A4',
    'soubor MP3',
    'model B2B',
    'Windows 10',
    'hra PS5',
    'vitamin B12',
    'norma ISO9001',
    'USB3 kabel',
    'CSS3 styly',
    'jazyk C++',
    'známka B+',
    'úroveň B2',
    'motor V8',
    'Formule F1',
    'Y2K',
    'W3C',
    '3D tisk',
    'Co je to',
    'NaCl',
    'www.H2O.cz',
    'H2O@firma.cz',
    'https://example.com/H2O',
    'Pro3',
    'kapitola 1.2',
  ])('%s', (input) => {
    expect(findFormulas(input, 'strict')).toEqual([]);
  });

  test('ignores letters and digits glued to the token', () => {
    expect(findFormulas('xH2O', 'strict')).toEqual([]);
    expect(findFormulas('H2Ox', 'strict')).toEqual([]);
    expect(findFormulas('čH2O', 'strict')).toEqual([]);
  });

  test.each(['O2', 'H2', 'N2', 'Cl2', 'Br2', 'I2', 'O3', 'S8', 'C60'])(
    'converts the common element molecule %s',
    (input) => {
      expect(findFormulas(input, 'strict')).toHaveLength(1);
    },
  );
});

describe('lenient mode for an explicitly selected token', () => {
  test.each([
    ['B12', 'B_12_'],
    ['B2B', 'B_2_B'],
    ['C4', 'C_4_'],
    ['K+', 'K^+^'],
  ])('%s', (input, expected) => {
    expect(show(input, 'lenient')).toBe(expected);
  });

  test('still rejects tokens that are not chemistry', () => {
    expect(findFormulas('A4', 'lenient')).toEqual([]);
    expect(findFormulas('MP3', 'lenient')).toEqual([]);
  });
});

describe('positions', () => {
  test('reports the source span in the original text', () => {
    const text = 'Voda je H2O a oxid je CO2.';
    const hits = findFormulas(text, 'strict');
    expect(hits.map((hit) => text.slice(hit.start, hit.end))).toEqual(['H2O', 'CO2']);
  });

  test('stops before an unmatched closing bracket', () => {
    const [hit] = findFormulas('(viz H2O)', 'strict');
    expect(hit.source).toBe('H2O');
  });

  test('handles a formula wrapped in brackets', () => {
    expect(show('(CO2)')).toBe('(CO_2_)');
  });
});

describe('ordinary text that looks like a formula', () => {
  const strictSources = (text) => findFormulas(text, 'strict').map((hit) => hit.source);

  test('leaves O2 alone in a name followed by a capitalised word', () => {
    expect(strictSources('Volejte O2 Arena.')).toEqual([]);
  });

  test.each(['O2 je plyn', 'H2 + O2 = 2H2O', 'Přidáme O2.'])('still converts oxygen in "%s"', (text) => {
    expect(strictSources(text)).toContain('O2');
  });

  test('leaves a Roman numeral followed by a dash alone', () => {
    expect(strictSources('Fáze I- a II- ')).toEqual([]);
  });

  test.each(['SN1', 'SN2'])('leaves the reaction name %s alone', (token) => {
    expect(strictSources(`reakce ${token} `)).toEqual([]);
  });

  test('still converts I- when it is selected on its own', () => {
    expect(findFormulas('I-', 'lenient')).toHaveLength(1);
  });
});
