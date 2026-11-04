/**
 * Finds chemical formulas in plain text and splits them into styled segments.
 * Pure logic with no dependency on Word, so it can be tested directly.
 *
 * @typedef {'normal' | 'sub' | 'sup'} SegmentStyle
 * @typedef {{ text: string, style: SegmentStyle }} Segment
 * @typedef {{ start: number, end: number, source: string, segments: Segment[] }} FormulaHit
 * @typedef {'strict' | 'lenient'} Mode
 * @typedef {{ kind: 'element' | 'ligand' | 'group' | 'complex', symbol: string, digits: string, inner?: Unit[] }} Unit
 * @typedef {{ isBracket: boolean, isParenthesis: boolean, isComplex: boolean }} Scope
 * @typedef {{ end: number, segments: Segment[], units: Unit[], sign: string, coefficient: string, isExplicitCharge?: boolean, isDoubledCharge?: boolean }} Part
 */

const ELEMENTS = new Set(
  ('H He Li Be B C N O F Ne Na Mg Al Si P S Cl Ar K Ca Sc Ti V Cr Mn Fe Co Ni Cu Zn Ga Ge As Se Br Kr ' +
    'Rb Sr Y Zr Nb Mo Tc Ru Rh Pd Ag Cd In Sn Sb Te I Xe Cs Ba La Ce Pr Nd Pm Sm Eu Gd Tb Dy Ho Er Tm ' +
    'Yb Lu Hf Ta W Re Os Ir Pt Au Hg Tl Pb Bi Po At Rn Fr Ra Ac Th Pa U Np Pu Am Cm Bk Cf Es Fm Md No ' +
    'Lr Rf Db Sg Bh Hs Mt Ds Rg Cn Nh Fl Mc Lv Ts Og').split(' '),
);

/** Everyday tokens that happen to parse as formulas; never converted automatically. */
const NON_CHEMICAL_TOKENS = new Set(
  ('B2B B2C C2C P2P H2H C3PO W3C Y2K PS2 PS3 PS4 PS5 USB2 USB3 CSS2 CSS3 PHP5 PHP7 PHP8 SPF15 SPF30 SPF50 ' +
    'IOS14 IOS15 IOS16 IOS17 IOS18 OS2 PC2 PC3 PC4 PC5 SP1 SP2 HIV1 HIV2 SN1 SN2 H1N1 H3N2 H5N1 H5N8 H7N9 OSx H2Ox').split(' '),
);

const ELEMENT_MOLECULES = new Set(
  ['H2', 'D2', 'N2', 'O2', 'O3', 'F2', 'Cl2', 'Br2', 'I2', 'P2', 'P4', 'As4', 'S8', 'Se8', 'C60', 'C70'],
);
const METALS = new Set(
  ('Li Be Na Mg Al K Ca Sc Ti V Cr Mn Fe Co Ni Cu Zn Ga Ge Rb Sr Y Zr Nb Mo Tc Ru Rh Pd Ag Cd In Sn Sb ' +
    'Cs Ba La Ce Pr Nd Pm Sm Eu Gd Tb Dy Ho Er Tm Yb Lu Hf Ta W Re Os Ir Pt Au Hg Tl Pb Bi Po Fr Ra Ac ' +
    'Th Pa U Np Pu Am Cm').split(' '),
);
const COMMON_ANIONS = new Set(['O', 'S', 'N', 'P', 'Se', 'Te', 'As']);
// Everyday tokens such as CD4 (a cell marker) rule out treating D as hydrogen elsewhere.
const DEUTERATED = new Set(
  'D2 D2O HDO D2S D2SO4 ND3 CDCl3 CD2Cl2 CD3OD CD3CN C6D6 (CD3)2SO (CD3)2CO'.split(' '),
);
// Common organic groups, used in formulas just like element symbols.
const GROUP_ABBREVIATIONS = new Set('Me Et Ph Bu Bn Cp'.split(' '));
const LIGAND = /^(?:en|py|phen|bipy|bpy|ox|gly|acac|edta|EDTA|dien|tren|dmso|DMSO|tpy)(?![a-zA-Z])/;
// Placeholders for "any metal, ligand or halogen", only trusted inside a complex.
const COMPLEX_PLACEHOLDERS = new Set(['M', 'L', 'X']);
// Cn and Cm are also elements, but before hydrogen they mean a general formula.
const CARBON_WITH_VARIABLE = /^C(?=[nm]H)/;
const VARIABLE_COUNT = /^\d*[nmxyz](?:[+\-−]\d+)?(?![\p{Ll}\p{N}])/u;
const DIGITS_ONLY = /^\d*$/;
const BONDS = new Set(['-', '–', '=', '═', '≡']);
const EXPLICIT_CHARGE = /^(?:\^(\d*)([+\-−–])|\((\d*)([+\-−–])\))/;
const DOUBLE_CHARGE = /^(\+\+|--|−−)(?![\p{L}\p{N}_>=+\-−–])/u;
const LIGAND_NAMES_WITH_D = /EDTA|DMSO/g;
const DEUTERIUM = /D(?![a-z])/;
const COMMON_IONS_WITHOUT_DIGITS = new Set(
  'H+ H− Li+ Na+ K+ Rb+ Cs+ Ag+ Cu+ Au+ Tl+ F− Cl− Br− OH− CN− SCN− NO+'.split(' '),
);

// The Czech operator, as in "O2 Arena", rather than oxygen.
const BRAND_TOKEN = 'O2';
const CAPITALISED_WORD_NEXT = /^\s+\p{Lu}\p{Ll}/u;

const REACTION_ARROW = /^(?:<[-–]{1,2}>|<={1,2}>|[-–—]{1,2}>)/;
const LOWERCASE = /\p{Ll}/u;
const WORD_CHAR = /[\p{L}\p{N}_]/u;
const CHARGE_SIGNS = new Set(['+', '-', '−', '–']);
const NOT_AFTER_CHARGE = /[\p{L}\p{N}_>=+\-−–]/u;
const COEFFICIENT = /^\d+(?:[.,/]\d+)?(?=[A-Z([])/;
const HYDRATE_COEFFICIENT = /^(?:\d+(?:[.,/]\d+)?|[nxy])(?=[A-Z([])/;
const SPACED_SEPARATOR = /^\s+([.*·×x])\s+/;
const TIGHT_SEPARATOR = /^[.*·×]/;
const NON_PROSE_CHUNK = /:\/\/|^www\.|@/;
const MAX_STRICT_SUBSCRIPT_DIGITS = 2;
const HYDRATE_DOT = '·';
const MINUS = '−';
const CLOSING_BRACKET = { '(': ')', '[': ']' };
const BRACKET_KIND = { '(': 'group', '[': 'complex' };

const TOP_LEVEL = { isBracket: false, isParenthesis: false, isComplex: false };

/**
 * Reads how many of a unit follow: digits, or a general count such as n, 2n+2 or x.
 * @param {string} text @param {number} pos @returns {string}
 */
const readCount = (text, pos) => {
  const rest = text.slice(pos);
  return VARIABLE_COUNT.exec(rest)?.[0] ?? /^\d*/.exec(rest)[0];
};

/** @param {string} text @param {number} pos @param {Scope} scope @returns {string | null} */
const readSymbol = (text, pos, scope) => {
  if (CARBON_WITH_VARIABLE.test(text.slice(pos))) return 'C';
  const two = text.slice(pos, pos + 2);
  if (two.length === 2 && (ELEMENTS.has(two) || GROUP_ABBREVIATIONS.has(two))) return two;
  const one = text[pos];
  if (one === undefined) return null;
  return ELEMENTS.has(one) || one === 'D' || (scope.isComplex && COMPLEX_PLACEHOLDERS.has(one))
    ? one
    : null;
};

/** @param {string} count @returns {Segment[]} */
const subscript = (count) => (count ? [{ text: count, style: 'sub' }] : []);

/**
 * Parses a bracketed group such as (OH)2 or [Fe(CN)6]. A bond inside, as in
 * (CH2-CH2)n, needs a count after the bracket, because "(C2-C4)" is usually a range.
 * @param {string} text @param {number} pos @param {Scope} scope
 * @returns {{ end: number, segments: Segment[], unit: Unit } | null}
 */
const parseBracket = (text, pos, scope) => {
  const open = text[pos];
  const inner = parseUnits(text, pos + 1, {
    isBracket: true,
    isParenthesis: open === '(',
    isComplex: scope.isComplex || open === '[',
  });
  if (inner === null || text[inner.end] !== CLOSING_BRACKET[open]) return null;
  const close = text[inner.end];
  const digits = readCount(text, inner.end + 1);
  if (inner.hasBond && !digits) return null;
  return {
    end: inner.end + 1 + digits.length,
    segments: [
      { text: open, style: 'normal' },
      ...inner.segments,
      { text: close, style: 'normal' },
      ...subscript(digits),
    ],
    unit: { kind: BRACKET_KIND[open], symbol: '', digits, inner: inner.units },
  };
};

/**
 * Parses a ligand abbreviation, which only means something inside a bracket.
 * "(py3)" is left alone: a count belongs after the bracket, as in (py)3.
 * @param {string} text @param {number} pos @param {Scope} scope
 * @returns {{ end: number, segments: Segment[], unit: Unit } | null}
 */
const parseLigand = (text, pos, scope) => {
  if (!scope.isBracket) return null;
  const symbol = LIGAND.exec(text.slice(pos))?.[0];
  if (symbol === undefined) return null;
  const digits = readCount(text, pos + symbol.length);
  if (digits && scope.isParenthesis) return null;
  return {
    end: pos + symbol.length + digits.length,
    segments: [{ text: symbol, style: 'normal' }, ...subscript(digits)],
    unit: { kind: 'ligand', symbol, digits },
  };
};

/**
 * Parses one element, ligand or bracket with its count.
 * @param {string} text @param {number} pos @param {Scope} scope
 * @returns {{ end: number, segments: Segment[], unit: Unit } | null}
 */
const parseUnit = (text, pos, scope) => {
  if (text[pos] === '(' || text[pos] === '[') return parseBracket(text, pos, scope);
  const ligand = parseLigand(text, pos, scope);
  if (ligand !== null) return ligand;
  const symbol = readSymbol(text, pos, scope);
  if (symbol === null) return null;
  const digits = readCount(text, pos + symbol.length);
  return {
    end: pos + symbol.length + digits.length,
    segments: [{ text: symbol, style: 'normal' }, ...subscript(digits)],
    unit: { kind: 'element', symbol, digits },
  };
};

/**
 * Parses as many units as possible. Inside a bracket they may be joined by a bond.
 * @param {string} text @param {number} start @param {Scope} [scope]
 * @returns {{ end: number, segments: Segment[], units: Unit[], hasBond: boolean } | null}
 */
function parseUnits(text, start, scope = TOP_LEVEL) {
  let end = start;
  let segments = [];
  let units = [];
  let hasBond = false;
  for (let next = parseUnit(text, end, scope); next !== null; next = parseUnit(text, end, scope)) {
    end = next.end;
    segments = [...segments, ...next.segments];
    units = [...units, next.unit];
    const afterBond = scope.isBracket && BONDS.has(text[end]) ? parseUnit(text, end + 1, scope) : null;
    if (afterBond !== null) {
      segments = [...segments, { text: text[end], style: 'normal' }];
      end += 1;
      hasBond = true;
    }
  }
  return units.length > 0 ? { end, segments, units, hasBond } : null;
}

/** @param {string} text @param {number} pos @returns {boolean} */
const isChargeAt = (text, pos) => {
  if (!CHARGE_SIGNS.has(text[pos])) return false;
  const after = text[pos + 1];
  if (text[pos] === '+' && REACTION_ARROW.test(text.slice(pos + 1))) return true;
  return after === undefined || !NOT_AFTER_CHARGE.test(after);
};

/**
 * Decides whether trailing digits count atoms or belong to the charge:
 * Ca2+ is a charge, NH4+ is a count, and SO42- holds one of each.
 * @param {Unit[]} units @param {string} sign
 * @returns {{ countDigits: string, chargeText: string }}
 */
const splitChargeDigits = (units, sign) => {
  const last = units[units.length - 1];
  const { digits } = last;
  if (!DIGITS_ONLY.test(digits)) return { countDigits: digits, chargeText: sign };
  if (digits.length >= 2) {
    return { countDigits: digits.slice(0, -1), chargeText: digits.slice(-1) + sign };
  }
  const isSingleElement = units.length === 1 && last.kind === 'element';
  if (last.kind === 'complex' || isSingleElement) {
    return { countDigits: '', chargeText: digits + sign };
  }
  return { countDigits: digits, chargeText: sign };
};

/**
 * @param {Segment[]} segments @param {Unit[]} units @param {string} sign
 * @returns {Segment[]}
 */
const withCharge = (segments, units, sign) => {
  const { digits } = units[units.length - 1];
  const withoutDigits = digits ? segments.slice(0, -1) : segments;
  const { countDigits, chargeText } = splitChargeDigits(units, sign);
  return [...withoutDigits, ...subscript(countDigits), { text: chargeText, style: 'sup' }];
};

/**
 * Reads a charge written so that it can't be mistaken for a count: ^2-, (2-) or ++.
 * @param {string} text @param {number} pos
 * @returns {{ length: number, chargeText: string, isDoubled?: boolean } | null}
 */
const readExplicitCharge = (text, pos) => {
  const rest = text.slice(pos);
  const doubled = DOUBLE_CHARGE.exec(rest);
  if (doubled) {
    return { length: doubled[0].length, chargeText: doubled[0].replaceAll('-', MINUS), isDoubled: true };
  }
  const match = EXPLICIT_CHARGE.exec(rest);
  if (match === null) return null;
  const digits = match[1] ?? match[3];
  const sign = match[2] ?? match[4];
  return { length: match[0].length, chargeText: digits + (sign === '+' ? '+' : MINUS) };
};

/**
 * Parses an optional coefficient, the units and an optional charge.
 * @param {string} text @param {number} pos @param {RegExp} [coefficientPattern]
 * @returns {Part | null}
 */
const parsePart = (text, pos, coefficientPattern = COEFFICIENT) => {
  const coefficient = coefficientPattern.exec(text.slice(pos))?.[0] ?? '';
  const parsed = parseUnits(text, pos + coefficient.length);
  if (parsed === null) return null;
  const lead = coefficient ? [{ text: coefficient, style: 'normal' }] : [];
  const explicit = readExplicitCharge(text, parsed.end);
  if (explicit !== null) {
    return {
      end: parsed.end + explicit.length,
      segments: [...lead, ...parsed.segments, { text: explicit.chargeText, style: 'sup' }],
      units: parsed.units,
      sign: explicit.chargeText,
      coefficient,
      isExplicitCharge: !explicit.isDoubled,
      isDoubledCharge: Boolean(explicit.isDoubled),
    };
  }
  const hasVariableCountLast = !DIGITS_ONLY.test(parsed.units[parsed.units.length - 1].digits);
  if (hasVariableCountLast || !isChargeAt(text, parsed.end)) {
    return { ...parsed, segments: [...lead, ...parsed.segments], sign: '', coefficient };
  }
  const sign = text[parsed.end] === '+' ? '+' : MINUS;
  return {
    end: parsed.end + 1,
    segments: [...lead, ...withCharge(parsed.segments, parsed.units, sign)],
    units: parsed.units,
    sign,
    coefficient,
  };
};

/** @param {Part} part @returns {boolean} */
const hasDigit = (part) => part.coefficient !== '' || part.units.some((unit) => unit.digits !== '');

/** @param {Part} part @returns {boolean} */
const endsWithWater = (part) => {
  const [hydrogen, oxygen] = part.units.slice(-2);
  return hydrogen?.symbol === 'H' && hydrogen.digits === '2' && oxygen?.symbol === 'O' && !oxygen.digits;
};

/**
 * Reads the hydrate separator after a part, if one follows.
 * Requiring both spaces or neither keeps the full stop in "H2O. 2H2" untouched.
 * @param {string} text @param {number} pos
 * @returns {{ part: Part, separator: string } | null}
 */
const parseHydrate = (text, pos) => {
  const rest = text.slice(pos);
  const spaced = SPACED_SEPARATOR.exec(rest);
  const separatorLength = spaced ? spaced[0].length : TIGHT_SEPARATOR.test(rest) ? 1 : 0;
  if (separatorLength === 0) return null;

  const part = parsePart(text, pos + separatorLength, HYDRATE_COEFFICIENT);
  if (part === null || !hasDigit(part)) return null;
  if (spaced?.[1] === 'x' && !endsWithWater(part)) return null;

  const separator = spaced ? ` ${HYDRATE_DOT} ` : HYDRATE_DOT;
  return { part, separator };
};

/**
 * Parses a whole formula, including any hydrate parts.
 * @param {string} text @param {number} start
 * @returns {{ end: number, parts: Part[], segments: Segment[] } | null}
 */
const parseFormula = (text, start) => {
  const first = parsePart(text, start);
  if (first === null) return null;
  let parts = [first];
  let segments = first.segments;
  let end = first.end;
  for (let next = parseHydrate(text, end); next !== null; next = parseHydrate(text, end)) {
    parts = [...parts, next.part];
    segments = [...segments, { text: next.separator, style: 'normal' }, ...next.part.segments];
    end = next.part.end;
  }
  return { end, parts, segments };
};

/**
 * Looks through a bracket around one element, so "[C3]" counts as C3, which is
 * far more likely a citation or a label than a substance.
 * @param {Unit[]} units @returns {Unit | null} The lone element, if that is all there is.
 */
const findLoneElement = (units) => {
  if (units.length !== 1) return null;
  const [only] = units;
  if (only.kind === 'element') return only;
  return only.inner?.length === 1 && only.inner[0].kind === 'element' ? only.inner[0] : null;
};

/** @param {Part} part @returns {boolean} */
const isPlausibleSinglePart = (part) => {
  const { units, sign, coefficient, isExplicitCharge, isDoubledCharge } = part;
  if (isExplicitCharge) return true;
  const lone = findLoneElement(units);
  // "C++" is a programming language, while "Ca++" and "SO4--" are ions.
  if (isDoubledCharge) {
    if (lone !== null) return (sign[0] === '+' ? METALS : COMMON_ANIONS).has(lone.symbol);
    return units.some((unit) => unit.kind !== 'element' || unit.digits !== '');
  }

  if (lone !== null && !sign) return ELEMENT_MOLECULES.has(lone.symbol + lone.digits);
  if (!sign) return true;
  if (units.every((unit) => unit.kind === 'element' && !unit.digits)) {
    const name = units.map((unit) => unit.symbol).join('');
    // A lone "I-" is usually a Roman numeral, as in "Fáze I- a II-"; "2I-" is not.
    return COMMON_IONS_WITHOUT_DIGITS.has(name + sign) || (name === 'I' && coefficient !== '');
  }
  if (lone !== null) {
    // Only Hg22+ has both a count and a charge; "W10+" is more likely a name.
    if (lone.digits.length > 1 && lone.symbol !== 'Hg') return false;
    return (sign === '+' ? METALS : COMMON_ANIONS).has(lone.symbol);
  }
  return true;
};

/**
 * Strict mode guards automatic conversion, where a wrong guess rewrites prose.
 * @param {string} source @param {Part[]} parts @param {Segment[]} segments
 * @returns {boolean}
 */
const isPlausibleInProse = (source, parts, segments) => {
  if (NON_CHEMICAL_TOKENS.has(source)) return false;
  const hasDeuterium = DEUTERIUM.test(source.replace(LIGAND_NAMES_WITH_D, ''));
  if (hasDeuterium && !DEUTERATED.has(source)) return false;
  const hasLongCount = segments.some(
    ({ style, text }) =>
      style === 'sub' && DIGITS_ONLY.test(text) && text.length > MAX_STRICT_SUBSCRIPT_DIGITS,
  );
  if (hasLongCount) return false;
  return parts.length > 1 || isPlausibleSinglePart(parts[0]);
};

/** @param {string} text @param {string} source @param {number} end @returns {boolean} */
const isBrandName = (text, source, end) =>
  source === BRAND_TOKEN && CAPITALISED_WORD_NEXT.test(text.slice(end));

/** @param {Segment[]} segments @returns {Segment[]} */
const mergeSegments = (segments) =>
  segments.reduce((merged, segment) => {
    const previous = merged[merged.length - 1];
    if (previous?.style !== segment.style) return [...merged, segment];
    return [...merged.slice(0, -1), { ...previous, text: previous.text + segment.text }];
  }, /** @type {Segment[]} */ ([]));

/**
 * Reads a reaction arrow. "p->x" is left alone: letters on both sides mean code.
 * @param {string} text @param {number} pos @returns {{ end: number, arrow: string } | null}
 */
const readReactionArrow = (text, pos) => {
  const token = REACTION_ARROW.exec(text.slice(pos))?.[0];
  if (token === undefined) return null;
  const end = pos + token.length;
  if (LOWERCASE.test(text[pos - 1] ?? '') && LOWERCASE.test(text[end] ?? '')) return null;
  const arrow = token.startsWith('<=') ? '⇌' : token.startsWith('<') ? '↔' : '→';
  return { end, arrow };
};

/**
 * Reads the mark for a gas (^) or a precipitate (v) glued to a formula. Without a
 * count or two elements it is likely a word; "v" also needs a two-letter element,
 * because "Cv." and "Sv." are common abbreviations.
 * @param {string} text @param {number} pos @param {Part[]} parts @returns {string | null}
 */
const readPhaseArrow = (text, pos, parts) => {
  const mark = text[pos];
  if (mark !== '^' && mark !== 'v') return null;
  const after = text[pos + 1] ?? '';
  if (WORD_CHAR.test(after) || CHARGE_SIGNS.has(after)) return null;
  const units = parts.flatMap((part) => part.units);
  const hasCount = parts.some(hasDigit);
  if (mark === '^') return hasCount || units.length > 1 ? '↑' : null;
  const hasLongSymbol = units.some((unit) => unit.symbol.length === 2);
  return hasCount || (units.length > 1 && hasLongSymbol) ? '↓' : null;
};

/**
 * Lists the spans of URLs and e-mail addresses, which must stay as written.
 * @param {string} text @returns {Array<[number, number]>}
 */
const findProtectedSpans = (text) =>
  [...text.matchAll(/\S+/g)]
    .filter((match) => NON_PROSE_CHUNK.test(match[0]))
    .map((match) => [match.index, match.index + match[0].length]);

/**
 * @param {string} text @param {Mode} [mode]
 * @returns {FormulaHit[]} Formulas that would change, in text order.
 */
export function findFormulas(text, mode = 'strict') {
  const protectedSpans = findProtectedSpans(text);
  const isProtected = (pos) => protectedSpans.some(([from, to]) => pos >= from && pos < to);
  let hits = [];
  let pos = 0;

  while (pos < text.length) {
    const reactionArrow = isProtected(pos) ? null : readReactionArrow(text, pos);
    if (reactionArrow !== null) {
      const { end, arrow } = reactionArrow;
      hits = [...hits, { start: pos, end, source: text.slice(pos, end), segments: [{ text: arrow, style: 'normal' }] }];
      pos = end;
      continue;
    }

    const isWordStart = pos === 0 || !WORD_CHAR.test(text[pos - 1]);
    const parsed = isWordStart && !isProtected(pos) ? parseFormula(text, pos) : null;
    const phaseArrow = parsed === null ? null : readPhaseArrow(text, parsed.end, parsed.parts);
    const tokenEnd = parsed === null ? pos : parsed.end + (phaseArrow ? 1 : 0);
    const isWholeToken = parsed !== null && (tokenEnd === text.length || !WORD_CHAR.test(text[tokenEnd]));

    if (!isWholeToken) {
      pos += 1;
      continue;
    }

    const formula = text.slice(pos, parsed.end);
    const arrowSegments = phaseArrow ? [{ text: phaseArrow, style: 'normal' }] : [];
    const segments = mergeSegments([...parsed.segments, ...arrowSegments]);
    const isChanged =
      segments.some((segment) => segment.style !== 'normal') || parsed.parts.length > 1 || phaseArrow !== null;
    const isAccepted =
      mode === 'lenient' ||
      (isPlausibleInProse(formula, parsed.parts, segments) && !isBrandName(text, formula, tokenEnd));
    if (isChanged && isAccepted) {
      hits = [...hits, { start: pos, end: tokenEnd, source: text.slice(pos, tokenEnd), segments }];
    }
    pos = tokenEnd;
  }
  return hits;
}

/**
 * Renders segments as text, marking subscript as _x_ and superscript as ^x^.
 * @param {Segment[]} segments @returns {string}
 */
export function renderPlain(segments) {
  const marks = { normal: '', sub: '_', sup: '^' };
  return segments.map(({ text, style }) => marks[style] + text + marks[style]).join('');
}
