/**
 * Applies converted formulas to a Word document.
 *
 * @typedef {import('./formula.js').FormulaHit} FormulaHit
 * @typedef {import('./formula.js').Segment} Segment
 * @typedef {{ hit: FormulaHit, range: any }} LocatedHit
 */
import { findFormulas } from './formula.js';
import { chooseMode, formulaKey, occurrenceIndex, planAutoConversion } from './planner.js';

const CARET_OUTSIDE_PARAGRAPH = new Set(['Before', 'After', 'AdjacentBefore', 'AdjacentAfter']);
const DEFAULT_DELAY_MS = 0;
// Word reports changes only every second or two, so typing is also checked directly.
const DEFAULT_POLL_INTERVAL_MS = 0;
const POLL_QUIET_MS = 3000;
const ITEM_NOT_FOUND = 'ItemNotFound';
const CARET = '^';
const FONT_PROPERTIES = ['font/name', 'font/subscript', 'font/superscript'];

/**
 * Finds the Word range of each hit without a caret, with its font, in one round trip.
 * A token whose match count differs from the text (fields, hidden text) is skipped,
 * because its matches can't be told apart.
 * @param {any} context @param {any} scope @param {string} text @param {FormulaHit[]} hits
 * @returns {Promise<LocatedHit[]>}
 */
async function locatePlainHits(context, scope, text, hits) {
  const sources = [...new Set(hits.map((hit) => hit.source))];
  const searches = new Map(sources.map((source) => [source, scope.search(source, { matchCase: true })]));
  searches.forEach((results) => results.load(FONT_PROPERTIES.map((name) => `items/${name}`).join(',')));
  await context.sync();

  return hits.flatMap((hit) => {
    const { items } = searches.get(hit.source);
    if (items.length !== occurrenceIndex(text, hit.source, text.length)) return [];
    return [{ hit, range: items[occurrenceIndex(text, hit.source, hit.start)] }];
  });
}

/**
 * Prepares the searches around a hit's caret. A caret is special in a Word search and
 * was not converted in Word on the web, so the text is split at each caret instead,
 * where it is taken literally.
 * @param {any[]} pieces Parts of the text, each ending with a caret except the last.
 * @param {string} text @param {FormulaHit} hit
 * @returns {{ hit: FormulaHit, mark: any, heads: any, tails: any } | null} Null when
 *   Word split the text differently from what the text says.
 */
const searchAroundCaret = (pieces, text, hit) => {
  const caretAt = text.indexOf(CARET, hit.start);
  const pieceIndex = text.slice(0, caretAt).split(CARET).length - 1;
  const mark = pieces[pieceIndex];
  const tail = hit.source.slice(caretAt - hit.start + 1);
  const tailPiece = tail ? pieces[pieceIndex + 1] : mark;
  if (!mark || !tailPiece) return null;

  const heads = mark.search(hit.source.slice(0, caretAt - hit.start), { matchCase: true });
  const tails = tail ? tailPiece.search(tail, { matchCase: true }) : null;
  heads.load('items');
  tails?.load('items');
  return { hit, mark, heads, tails };
};

/**
 * Finds the Word range of each hit with a caret, with its font.
 * @param {any} context @param {any} scope @param {string} text @param {FormulaHit[]} hits
 * @returns {Promise<LocatedHit[]>}
 */
async function locateCaretHits(context, scope, text, hits) {
  const pieces = scope.getTextRanges([CARET], false);
  pieces.load('items');
  await context.sync();

  const found = hits.flatMap((hit) => searchAroundCaret(pieces.items, text, hit) ?? []);
  await context.sync();

  const located = found.flatMap(({ hit, mark, heads, tails }) => {
    const start = heads.items.at(-1);
    const end = tails ? tails.items[0] : mark;
    return start && end ? [{ hit, range: start.expandTo(end) }] : [];
  });
  located.forEach(({ range }) => range.load(FONT_PROPERTIES.join(',')));
  await context.sync();
  return located;
}

/**
 * @param {any} context @param {any} scope @param {string} text @param {FormulaHit[]} hits
 * @returns {Promise<LocatedHit[]>} The hits Word could find, in the order of the text.
 */
async function locateHits(context, scope, text, hits) {
  const hasCaret = (hit) => hit.source.includes(CARET);
  const plain = hits.filter((hit) => !hasCaret(hit));
  const withCaret = hits.filter(hasCaret);
  const located = [
    ...(plain.length > 0 ? await locatePlainHits(context, scope, text, plain) : []),
    ...(withCaret.length > 0 ? await locateCaretHits(context, scope, text, withCaret) : []),
  ];
  return located.sort((a, b) => a.hit.start - b.hit.start);
}

/**
 * Word reports null for mixed text. A new formula can be mixed when an earlier part
 * was converted before the rest was typed, as with CuSO4 and then CuSO4.5H2O.
 * @param {any} range @returns {boolean} False only when all of the text is raised.
 */
const isNotFullyRaised = (range) => range.font.subscript !== true && range.font.superscript !== true;

/**
 * Word links the two flags: clearing either one returns the text to the baseline,
 * so the flag being switched on has to be set last.
 * @param {any} range @param {Segment['style']} style
 */
const applyStyle = (range, style) => {
  if (style === 'sup') {
    range.font.subscript = false;
    range.font.superscript = true;
    return;
  }
  range.font.superscript = false;
  range.font.subscript = style === 'sub';
};

/**
 * Word can give inserted symbols such as arrows a fallback font, which later typing
 * then inherits, so the original font is set again. Mixed fonts have no single name.
 * @param {any} range @param {Segment[]} segments
 */
const writeSegments = (range, segments) => {
  const fontName = range.font.name;
  segments.reduce((previous, { text, style }, index) => {
    const written = index === 0 ? previous.insertText(text, 'Replace') : previous.insertText(text, 'After');
    if (fontName) written.font.name = fontName;
    applyStyle(written, style);
    return written;
  }, range);
};

/**
 * Writes from the end backwards so earlier ranges stay valid.
 * @param {any} context @param {LocatedHit[]} located @returns {Promise<number>}
 */
async function writeHits(context, located) {
  [...located].reverse().forEach(({ hit, range }) => writeSegments(range, hit.segments));
  await context.sync();
  return located.length;
}

/**
 * Converts the selection, or the current paragraph when nothing is selected.
 * @param {any} context @returns {Promise<number>} How many formulas changed.
 */
export async function convertSelection(context) {
  const selection = context.document.getSelection();
  selection.load('text');
  await context.sync();

  const isEmpty = selection.text.trim() === '';
  const scope = isEmpty ? selection.paragraphs.getFirst() : selection;
  if (isEmpty) {
    scope.load('text');
    await context.sync();
  }

  const mode = isEmpty ? 'strict' : chooseMode(selection.text);
  const hits = findFormulas(scope.text, mode);
  if (hits.length === 0) return 0;

  const located = await locateHits(context, scope, scope.text, hits);
  return writeHits(context, located.filter(({ range }) => isNotFullyRaised(range)));
}

/**
 * Converts formulas typed in a paragraph, leaving any the caret is still touching.
 * @param {any} context @param {string} paragraphId
 * @param {ReadonlyMap<string, number>} previousSeen
 * @returns {Promise<{ converted: number, seen: Map<string, number> }>}
 */
export async function convertParagraph(context, paragraphId, previousSeen) {
  const paragraph = context.document.getParagraphByUniqueLocalId(paragraphId);
  const selection = context.document.getSelection();
  const caretVsParagraph = selection.compareLocationWith(paragraph.getRange('Whole'));
  const textBeforeCaret = paragraph.getRange('Start').expandTo(selection.getRange('Start'));
  textBeforeCaret.load('text');
  paragraph.load('text');
  await context.sync();

  const { text } = paragraph;
  const isCaretOutside = CARET_OUTSIDE_PARAGRAPH.has(caretVsParagraph.value);
  const caretOffset = isCaretOutside ? text.length : textBeforeCaret.text.length;
  const { toConvert, seen } = planAutoConversion(text, previousSeen, caretOffset);
  if (toConvert.length === 0) return { converted: 0, seen };

  const isCaretTouching = (hit) =>
    !isCaretOutside && caretOffset >= hit.start && caretOffset <= hit.end;

  const located = await locateHits(
    context,
    paragraph,
    text,
    toConvert.filter((hit) => !isCaretTouching(hit)),
  );
  const ready = located.filter(({ range }) => isNotFullyRaised(range));
  const waitingKeys = new Set(toConvert.filter(isCaretTouching).map(formulaKey));

  const nextSeen = new Map(
    [...seen]
      .map(([key, count]) => [key, waitingKeys.has(key) ? previousSeen.get(key) : count])
      .filter(([, count]) => count !== undefined),
  );
  return { converted: await writeHits(context, ready), seen: nextSeen };
}

/**
 * Reads the paragraph holding the caret.
 * @param {any} context @returns {Promise<{ id: string, text: string }>}
 */
export async function readActiveParagraph(context) {
  const paragraph = context.document.getSelection().paragraphs.getFirst();
  paragraph.load('uniqueLocalId,text');
  await context.sync();
  return { id: paragraph.uniqueLocalId, text: paragraph.text };
}

/**
 * Debounces paragraph-change events and converts the affected paragraphs. An event
 * also starts direct checks of the caret's paragraph, which end after a quiet period.
 * The add-in's own edits raise events too; they find nothing left to do.
 * @param {{ runWord: (task: (context: any) => Promise<void>) => Promise<void>,
 *   readActive: () => Promise<{ id: string, text: string }>,
 *   onError: (error: unknown) => void, onConverted: (count: number) => void,
 *   delayMs?: number, pollIntervalMs?: number }} options
 */
export function createAutoConverter({
  runWord,
  readActive,
  onError,
  onConverted,
  delayMs = DEFAULT_DELAY_MS,
  pollIntervalMs = DEFAULT_POLL_INTERVAL_MS,
}) {
  let pendingIds = new Set();
  let seenByParagraph = new Map();
  let timer = null;
  let isRunning = false;
  let pollTimer = null;
  let lastActive = null;
  let lastChangeAt = 0;

  const convertOne = async (id) => {
    try {
      await runWord(async (context) => {
        const previous = seenByParagraph.get(id) ?? new Map();
        const { converted, seen } = await convertParagraph(context, id, previous);
        seenByParagraph = new Map(seenByParagraph).set(id, seen);
        if (converted > 0) onConverted(converted);
      });
    } catch (error) {
      if (error?.code !== ITEM_NOT_FOUND) onError(error);
    }
  };

  const schedule = () => {
    clearTimeout(timer);
    timer = setTimeout(flush, delayMs);
  };

  async function flush() {
    if (isRunning) return;
    isRunning = true;
    const ids = [...pendingIds];
    pendingIds = new Set();
    for (const id of ids) {
      await convertOne(id);
    }
    isRunning = false;
    if (pendingIds.size > 0) schedule();
  }

  const stopPolling = () => {
    clearTimeout(pollTimer);
    pollTimer = null;
    lastActive = null;
  };

  const noteActive = (active) => {
    const previous = lastActive;
    lastActive = active;
    if (!previous || (previous.id === active.id && previous.text === active.text)) return;
    lastChangeAt = Date.now();
    pendingIds = new Set([...pendingIds, previous.id, active.id]);
    clearTimeout(timer);
    flush();
  };

  async function poll() {
    try {
      const active = await readActive();
      if (pollTimer === null) return;
      noteActive(active);
    } catch (error) {
      stopPolling();
      if (error?.code !== ITEM_NOT_FOUND) onError(error);
      return;
    }
    if (Date.now() - lastChangeAt >= POLL_QUIET_MS) {
      stopPolling();
      return;
    }
    pollTimer = setTimeout(poll, pollIntervalMs);
  }

  const startPolling = () => {
    lastChangeAt = Date.now();
    if (pollTimer === null) pollTimer = setTimeout(poll, pollIntervalMs);
  };

  return {
    /** @param {{ source?: string, uniqueLocalIds?: string[] }} event */
    handleEvent(event) {
      if (event.source === 'Remote') return;
      pendingIds = new Set([...pendingIds, ...(event.uniqueLocalIds ?? [])]);
      schedule();
      startPolling();
    },
    stop() {
      clearTimeout(timer);
      stopPolling();
      pendingIds = new Set();
    },
  };
}
