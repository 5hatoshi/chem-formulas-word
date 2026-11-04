/**
 * Decides what to convert, independently of the Word API.
 *
 * @typedef {import('./formula.js').FormulaHit} FormulaHit
 * @typedef {import('./formula.js').Mode} Mode
 */
import { findFormulas } from './formula.js';

/**
 * Word's search returns every raw match, so "H2" also matches inside "H2O".
 * This gives the position of the wanted match among those results.
 * @param {string} text @param {string} token @param {number} offset
 * @returns {number}
 */
export function occurrenceIndex(text, token, offset) {
  let count = 0;
  for (let at = text.indexOf(token); at !== -1 && at < offset; at = text.indexOf(token, at + token.length)) {
    count += 1;
  }
  return count;
}

/**
 * One selected token is an explicit request; a longer selection is mostly prose.
 * @param {string} selectedText @returns {Mode}
 */
export function chooseMode(selectedText) {
  return /\s/.test(selectedText.trim()) ? 'strict' : 'lenient';
}

/**
 * Identifies a formula however it was typed, so "." and "·" count as the same one.
 * @param {FormulaHit} hit @returns {string}
 */
export function formulaKey(hit) {
  return hit.segments.map((segment) => segment.text).join('');
}

/**
 * @param {FormulaHit[]} copies @param {number} caretOffset @param {number} amount
 * @returns {FormulaHit[]} The copies closest to the caret, which are the ones just typed.
 */
const closestToCaret = (copies, caretOffset, amount) =>
  [...copies]
    .sort((a, b) => Math.abs(caretOffset - a.end) - Math.abs(caretOffset - b.end))
    .slice(0, Math.max(amount, 0));

/**
 * Picks formulas to convert in a paragraph. Only as many copies convert as were
 * added since last time, so undoing a conversion is not reversed again.
 * @param {string} text
 * @param {ReadonlyMap<string, number>} previous Copies per formula at the last check.
 * @param {number} [caretOffset] Where the caret is in the text.
 * @returns {{ toConvert: FormulaHit[], seen: Map<string, number> }}
 */
export function planAutoConversion(text, previous, caretOffset = text.length) {
  const hits = findFormulas(text, 'strict');
  const seen = hits.reduce((counts, hit) => {
    const key = formulaKey(hit);
    return new Map(counts).set(key, (counts.get(key) ?? 0) + 1);
  }, new Map());
  const added = new Set(
    [...seen].flatMap(([key, count]) =>
      closestToCaret(
        hits.filter((hit) => formulaKey(hit) === key),
        caretOffset,
        count - (previous.get(key) ?? 0),
      ),
    ),
  );
  return { toConvert: hits.filter((hit) => added.has(hit)), seen };
}
