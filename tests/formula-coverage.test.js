import { describe, expect, test } from 'vitest';
import { findFormulas } from '../addin/src/formula.js';

const MARKS = { normal: (text) => text, sub: (text) => `_${text}_`, sup: (text) => `^${text}^` };

/** Renders the one formula expected in the text, or the whole list when there are more. */
const render = (text) =>
  findFormulas(`${text} `, 'strict')
    .map((hit) => hit.segments.map(({ text: part, style }) => MARKS[style](part)).join(''))
    .join(' | ');

describe('complexes', () => {
  test.each([
    ['K4[Fe(CN)6]', 'K_4_[Fe(CN)_6_]'],
    ['[Cu(NH3)4]2+', '[Cu(NH_3_)_4_]^2+^'],
    ['[Co(NH3)5Cl]Cl2', '[Co(NH_3_)_5_Cl]Cl_2_'],
    ['K3[Fe(C2O4)3].3H2O', 'K_3_[Fe(C_2_O_4_)_3_]·3H_2_O'],
    ['[Pt(NH3)4][PtCl4]', '[Pt(NH_3_)_4_][PtCl_4_]'],
    ['[Ni(en)3]2+', '[Ni(en)_3_]^2+^'],
    ['[Co(en)3]Cl3', '[Co(en)_3_]Cl_3_'],
    ['[Fe(phen)3]2+', '[Fe(phen)_3_]^2+^'],
    ['[Cu(py)4]2+', '[Cu(py)_4_]^2+^'],
    ['[Fe(ox)3]3-', '[Fe(ox)_3_]^3−^'],
    ['[Ni(edta)]2-', '[Ni(edta)]^2−^'],
    ['[Co(acac)3]', '[Co(acac)_3_]'],
    ['[Ru(bipy)3]2+', '[Ru(bipy)_3_]^2+^'],
    ['[Co(EDTA)]-', '[Co(EDTA)]^−^'],
    ['trans-[Co(en)2Cl2]+', '[Co(en)_2_Cl_2_]^+^'],
    ['cis-[PtCl2(NH3)2]', '[PtCl_2_(NH_3_)_2_]'],
    ['[Fe(Cp)2]', '[Fe(Cp)_2_]'],
    ['[ML6]', '[ML_6_]'],
    ['[MX4]2-', '[MX_4_]^2−^'],
  ])('%s', (text, expected) => {
    expect(render(text)).toBe(expected);
  });
});

describe('general formulas and polymers', () => {
  test.each([
    ['CnH2n+2', 'C_n_H_2n+2_'],
    ['CnH2n', 'C_n_H_2n_'],
    ['CnH2n-2', 'C_n_H_2n-2_'],
    ['CnH2n+1OH', 'C_n_H_2n+1_OH'],
    ['CmH2m', 'C_m_H_2m_'],
    ['CxHy', 'C_x_H_y_'],
    ['Plyn CnH2n+ a', 'C_n_H_2n_'],
    ['CxHyOz', 'C_x_H_y_O_z_'],
    ['FexOy', 'Fe_x_O_y_'],
    ['NOx', 'NO_x_'],
    ['SOx', 'SO_x_'],
    ['(C6H10O5)n', '(C_6_H_10_O_5_)_n_'],
    ['(CH2-CH2)n', '(CH_2_-CH_2_)_n_'],
    ['CuSO4.nH2O', 'CuSO_4_·nH_2_O'],
    ['Na2SO4·xH2O', 'Na_2_SO_4_·xH_2_O'],
  ])('%s', (text, expected) => {
    expect(render(text)).toBe(expected);
  });
});

describe('organic compounds', () => {
  test.each([
    ['C6H12O6', 'C_6_H_12_O_6_'],
    ['CH3(CH2)4CH3', 'CH_3_(CH_2_)_4_CH_3_'],
    ['(C2H5)2O', '(C_2_H_5_)_2_O'],
    ['CH3CH(OH)COOH', 'CH_3_CH(OH)COOH'],
    ['CH3COO-', 'CH_3_COO^−^'],
    ['CH3-CH2-OH', 'CH_3_ | CH_2_'],
    ['CH2=CH2', 'CH_2_ | CH_2_'],
    ['Et2O', 'Et_2_O'],
    ['Et3N', 'Et_3_N'],
    ['Ph3P', 'Ph_3_P'],
    ['PPh3', 'PPh_3_'],
    ['Me2CO', 'Me_2_CO'],
    ['Bu4N+', 'Bu_4_N^+^'],
    ['Cp2Fe', 'Cp_2_Fe'],
  ])('%s', (text, expected) => {
    expect(render(text)).toBe(expected);
  });
});

describe('deuterated compounds', () => {
  test.each([
    ['D2O', 'D_2_O'],
    ['CDCl3', 'CDCl_3_'],
    ['CD3OD', 'CD_3_OD'],
    ['C6D6', 'C_6_D_6_'],
  ])('%s', (text, expected) => {
    expect(render(text)).toBe(expected);
  });
});

describe('ions and ways of writing a charge', () => {
  test.each([
    ['Ca++', 'Ca^++^'],
    ['SO4--', 'SO_4_^−−^'],
    ['Fe^3+', 'Fe^3+^'],
    ['SO4^2-', 'SO_4_^2−^'],
    ['NH4^+', 'NH_4_^+^'],
    ['Cu(2+)', 'Cu^2+^'],
    ['SO4(2-)', 'SO_4_^2−^'],
    ['Au3+', 'Au^3+^'],
    ['Bi3+', 'Bi^3+^'],
    ['Ce4+', 'Ce^4+^'],
    ['U4+', 'U^4+^'],
    ['Se2-', 'Se^2−^'],
    ['H-', 'H^−^'],
    ['2I-', '2I^−^'],
    ['Ca2+(aq)', 'Ca^2+^'],
    ['H2O(l)', 'H_2_O'],
  ])('%s', (text, expected) => {
    expect(render(text)).toBe(expected);
  });
});

describe('ordinary text stays as written', () => {
  test.each([
    'Audi A4 a RS6',
    'BMW X5 a M3',
    'buňky CD4 a CD8',
    'viz [C3] a [P1]',
    'body (S1) a (K2)',
    'rozsah (C2-C4)',
    'Galaxy S8+',
    'jazyk C++',
    'CS++ a BC--',
    'W10+',
    'K12 a Y3',
    'úroveň B2+ a C1+',
    'chřipka H1N1 a H5N1',
    'macOS OSx',
    'COVID19',
    'PhD studium',
    'Me2 a Et3',
    'Cox, Tin, Can, Sin, Pan a Con',
    'Sony, Canon, Fox a Box',
    'CoD4 a HDD2',
    'D2C a R2D2',
    '(py3) a (en2)',
    'Fáze I- a II-',
    'hodnota x a y',
    'iPhone 15',
    'B12, A4, MP3 a PS5',
  ])('%s', (text) => {
    expect(render(text)).toBe('');
  });
});

describe('reaction arrows', () => {
  test.each([
    ['2H2 + O2 -> 2H2O', '2H_2_ | O_2_ | → | 2H_2_O'],
    ['H2+O2->H2O', 'H_2_ | O_2_ | → | H_2_O'],
    ['N2 + 3H2 --> 2NH3', 'N_2_ | 3H_2_ | → | 2NH_3_'],
    ['N2 + 3H2 <=> 2NH3', 'N_2_ | 3H_2_ | ⇌ | 2NH_3_'],
    ['N2 + 3H2 <==> 2NH3', 'N_2_ | 3H_2_ | ⇌ | 2NH_3_'],
    ['A <-> B', '↔'],
    ['A <--> B', '↔'],
    ['A –> B', '→'],
    ['Ca2+->Ca', 'Ca^2+^ | →'],
  ])('%s', (text, expected) => {
    expect(render(text)).toBe(expected);
  });

  test.each(['p->x', 'https://a.cz/x->y', 'x <- 5', 'a < b', 'a > b', '- b'])('leaves "%s" alone', (text) => {
    expect(render(text)).toBe('');
  });
});

describe('gas and precipitate arrows', () => {
  test.each([
    ['CO2^', 'CO_2_↑'],
    ['NH3^', 'NH_3_↑'],
    ['H2^', 'H_2_↑'],
    ['HCl^', 'HCl↑'],
    ['CO^', 'CO↑'],
    ['BaSO4v', 'BaSO_4_↓'],
    ['AgClv', 'AgCl↓'],
    ['PbSv', 'PbS↓'],
    ['Fe(OH)3v', 'Fe(OH)_3_↓'],
    ['CaCO3 -> CaO + CO2^', 'CaCO_3_ | → | CO_2_↑'],
  ])('%s', (text, expected) => {
    expect(render(text)).toBe(expected);
  });

  test('keeps the source of a formula with an arrow for the Word search', () => {
    expect(findFormulas('CO2^ ', 'strict')[0].source).toBe('CO2^');
  });

  test('still reads a caret followed by a charge as a charge', () => {
    expect(render('SO4^2-')).toBe('SO_4_^2−^');
  });

  test.each(['Sv. Václav', 'Cv. 5', 'Cov', 'Kov', 'Nav', 'CO2v2', 'C^', 'O^', 'Fe^', 'BaSO4 v vodě'])(
    'leaves the arrow out of "%s"',
    (text) => {
      expect(render(text)).not.toMatch(/[↑↓]/);
    },
  );
});
