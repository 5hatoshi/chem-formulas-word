/**
 * Finds chemical formulas in plain text and splits them into styled segments.
 * Pure logic with no dependency on Word, so it can be tested directly.
 *
 * @typedef {'normal' | 'sub' | 'sup'} SegmentStyle
 * @typedef {{ text: string, style: SegmentStyle }} Segment
 * @typedef {{ start: number, end: number, source: string, segments: Segment[] }} FormulaHit
 * @typedef {'strict' | 'lenient'} Mode
 * @typedef {{ kind: 'element' | 'group' | 'complex', symbol: string, digits: string }} Unit
 * @typedef {{ end: number, segments: Segment[], units: Unit[], sign: string, coefficient: string }} Part
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
    'IOS14 IOS15 IOS16 IOS17 IOS18 OS2 PC2 SP1 SP2 HIV1 HIV2 SN1 SN2').split(' '),
);

const ELEMENT_MOLECULES = new Set(['H2', 'N2', 'O2', 'O3', 'F2', 'Cl2', 'Br2', 'I2', 'P4', 'S8', 'C60']);
const COMMON_CATIONS = new Set(
  'Mg Ca Sr Ba Fe Cu Zn Mn Co Ni Cr Al Pb Sn Hg Ag Ti Cd'.split(' '),
);
const COMMON_ANIONS = new Set(['O', 'S', 'N', 'P']);
const COMMON_IONS_WITHOUT_DIGITS = new Set(
  'H+ Li+ Na+ K+ Rb+ Cs+ Ag+ Cu+ F− Cl− Br− OH− CN− SCN− NO+'.split(' '),
);

// The Czech operator, as in "O2 Arena", rather than oxygen.
const BRAND_TOKEN = 'O2';
const CAPITALISED_WORD_NEXT = /^\s+\p{Lu}\p{Ll}/u;

const WORD_CHAR = /[\p{L}\p{N}_]/u;
const CHARGE_SIGNS = new Set(['+', '-', '−', '–']);
const NOT_AFTER_CHARGE = /[\p{L}\p{N}_>=+\-−–]/u;
const COEFFICIENT = /^\d+(?:[.,/]\d+)?(?=[A-Z([])/;
const SPACED_SEPARATOR = /^\s+([.*·×x])\s+/;
const TIGHT_SEPARATOR = /^[.*·×]/;
const NON_PROSE_CHUNK = /:\/\/|^www\.|@/;
const MAX_STRICT_SUBSCRIPT_DIGITS = 2;
const HYDRATE_DOT = '·';
const MINUS = '−';
const CLOSING_BRACKET = { '(': ')', '[': ']' };
const BRACKET_KIND = { '(': 'group', '[': 'complex' };

/** @param {string} text @param {number} pos @returns {string} */
const readDigits = (text, pos) => /^\d*/.exec(text.slice(pos))[0];

/** @param {string} text @param {number} pos @returns {string | null} */
const readElement = (text, pos) => {
  const two = text.slice(pos, pos + 2);
  if (two.length === 2 && ELEMENTS.has(two)) return two;
  const one = text[pos];
  return one !== undefined && ELEMENTS.has(one) ? one : null;
};

/** @param {string} digits @returns {Segment[]} */
const subscript = (digits) => (digits ? [{ text: digits, style: 'sub' }] : []);

/**
 * Parses a bracketed group such as (OH)2 or [Fe(CN)6].
 * @param {string} text @param {number} pos
 * @returns {{ end: number, segments: Segment[], unit: Unit } | null}
 */
const parseBracket = (text, pos) => {
  const open = text[pos];
  const inner = parseUnits(text, pos + 1);
  if (inner === null || text[inner.end] !== CLOSING_BRACKET[open]) return null;
  const close = text[inner.end];
  const digits = readDigits(text, inner.end + 1);
  return {
    end: inner.end + 1 + digits.length,
    segments: [
      { text: open, style: 'normal' },
      ...inner.segments,
      { text: close, style: 'normal' },
      ...subscript(digits),
    ],
    unit: { kind: BRACKET_KIND[open], symbol: '', digits },
  };
};

/**
 * Parses one element or bracket with its count.
 * @param {string} text @param {number} pos
 * @returns {{ end: number, segments: Segment[], unit: Unit } | null}
 */
const parseUnit = (text, pos) => {
  if (text[pos] === '(' || text[pos] === '[') return parseBracket(text, pos);
  const symbol = readElement(text, pos);
  if (symbol === null) return null;
  const digits = readDigits(text, pos + symbol.length);
  return {
    end: pos + symbol.length + digits.length,
    segments: [{ text: symbol, style: 'normal' }, ...subscript(digits)],
    unit: { kind: 'element', symbol, digits },
  };
};

/**
 * Parses as many units as possible.
 * @param {string} text @param {number} start
 * @returns {{ end: number, segments: Segment[], units: Unit[] } | null}
 */
function parseUnits(text, start) {
  let end = start;
  let segments = [];
  let units = [];
  for (let next = parseUnit(text, end); next !== null; next = parseUnit(text, end)) {
    end = next.end;
    segments = [...segments, ...next.segments];
    units = [...units, next.unit];
  }
  return units.length > 0 ? { end, segments, units } : null;
}

/** @param {string} text @param {number} pos @returns {boolean} */
const isChargeAt = (text, pos) => {
  if (!CHARGE_SIGNS.has(text[pos])) return false;
  const after = text[pos + 1];
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
 * Parses an optional coefficient, the units and an optional charge.
 * @param {string} text @param {number} pos @returns {Part | null}
 */
const parsePart = (text, pos) => {
  const coefficient = COEFFICIENT.exec(text.slice(pos))?.[0] ?? '';
  const parsed = parseUnits(text, pos + coefficient.length);
  if (parsed === null) return null;
  const lead = coefficient ? [{ text: coefficient, style: 'normal' }] : [];
  if (!isChargeAt(text, parsed.end)) {
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

  const part = parsePart(text, pos + separatorLength);
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

/** @param {Part} part @returns {boolean} */
const isPlausibleSinglePart = (part) => {
  const { units, sign } = part;
  const [only] = units;
  const isLoneElement = units.length === 1 && only.kind === 'element';

  if (isLoneElement && !sign) return ELEMENT_MOLECULES.has(only.symbol + only.digits);
  if (!sign) return true;
  if (units.every((unit) => unit.kind === 'element' && !unit.digits)) {
    const name = units.map((unit) => unit.symbol).join('');
    return COMMON_IONS_WITHOUT_DIGITS.has(name + sign);
  }
  if (isLoneElement) return (sign === '+' ? COMMON_CATIONS : COMMON_ANIONS).has(only.symbol);
  return true;
};

/**
 * Strict mode guards automatic conversion, where a wrong guess rewrites prose.
 * @param {string} source @param {Part[]} parts @param {Segment[]} segments
 * @returns {boolean}
 */
const isPlausibleInProse = (source, parts, segments) => {
  if (NON_CHEMICAL_TOKENS.has(source)) return false;
  const hasLongCount = segments.some(
    (segment) => segment.style === 'sub' && segment.text.length > MAX_STRICT_SUBSCRIPT_DIGITS,
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
    const isWordStart = pos === 0 || !WORD_CHAR.test(text[pos - 1]);
    const parsed = isWordStart && !isProtected(pos) ? parseFormula(text, pos) : null;
    const isWholeToken =
      parsed !== null && (parsed.end === text.length || !WORD_CHAR.test(text[parsed.end]));

    if (!isWholeToken) {
      pos += 1;
      continue;
    }

    const source = text.slice(pos, parsed.end);
    const segments = mergeSegments(parsed.segments);
    const isChanged = segments.some((segment) => segment.style !== 'normal') || parsed.parts.length > 1;
    const isAccepted =
      mode === 'lenient' ||
      (isPlausibleInProse(source, parsed.parts, segments) && !isBrandName(text, source, parsed.end));
    if (isChanged && isAccepted) {
      hits = [...hits, { start: pos, end: parsed.end, source, segments }];
    }
    pos = parsed.end;
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
