import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { convertParagraph, convertSelection, createAutoConverter, onFormulaSkipped } from '../addin/src/word.js';

/** Models real Word: clearing either flag returns the text to the baseline. */
const createLinkedFont = () => {
  let position = 'baseline';
  const flag = (name) => ({
    get: () => position === name,
    set: (isOn) => {
      if (isOn) position = name;
      else position = 'baseline';
    },
  });
  return Object.defineProperties({}, { subscript: flag('sub'), superscript: flag('sup') });
};

/** Minimal stand-in for the parts of the Word API the add-in uses. */
const createFakeRange = (
  text,
  { plain = true, raised = false, log = [], fontName = 'Aptos' } = {},
) => {
  const plainFlag = plain ? false : null;
  const makeInserted = (inserted) => ({
    font: createLinkedFont(),
    insertText(next, location) {
      const range = makeInserted(next);
      log.push({ text: next, location, range });
      return range;
    },
    text: inserted,
  });
  return {
    text,
    load() {},
    font: {
      name: fontName,
      subscript: raised ? true : plainFlag,
      superscript: raised ? false : plainFlag,
      load() {},
    },
    insertText(next, location) {
      const range = makeInserted(next);
      log.push({ text: next, location, range });
      return range;
    },
  };
};

/**
 * A part of the text ending with "^", as Word's getTextRanges returns it. The words it
 * can find expand to the given range, and only to the expected end: real expandTo
 * covers both ranges, so expanding to the whole piece would take the text before too.
 */
const createCaretPiece = (words, expanded, expectedEnd) => {
  const pieceEnd = {};
  const piece = {
    getRange: (location) => {
      if (location !== 'End') throw new Error(`Unexpected location ${location}`);
      return pieceEnd;
    },
    search: (word) => ({
      load() {},
      items: word in words
        ? [
            {
              ...words[word],
              expandTo: (end) => {
                if (end !== (expectedEnd ?? pieceEnd)) throw new Error('Expanded to the wrong end');
                return expanded;
              },
            },
          ]
        : [],
    }),
  };
  return piece;
};

const createFakeContext = ({
  text,
  caretPieces = [],
  matches = {},
  selectionText = '',
  paragraphText = text,
  caretVsParagraph = 'Inside',
  caretOffset = paragraphText.length,
}) => {
  const scope = {
    text,
    load() {},
    getTextRanges: () => ({ items: caretPieces, load() {} }),
    search: (source) => ({
      items: matches[source] ?? [],
      // Real Word fails when a property is read without being loaded first.
      load(properties) {
        if (!properties.includes('items/font/name')) {
          throw Object.assign(new Error('PropertyNotLoaded'), { code: 'PropertyNotLoaded' });
        }
      },
    }),
  };
  const paragraph = {
    ...scope,
    text: paragraphText,
    getRange: (location) =>
      location === 'Start'
        ? { expandTo: () => ({ text: paragraphText.slice(0, caretOffset), load() {} }) }
        : { relation: caretVsParagraph },
  };
  const selection = {
    ...scope,
    text: selectionText,
    getRange: () => ({}),
    paragraphs: { getFirst: () => paragraph },
    // Real Word only accepts a range here and fails on a paragraph object.
    compareLocationWith: (range) => {
      if (range.relation === undefined) {
        throw Object.assign(new Error('GeneralException'), { code: 'GeneralException' });
      }
      return { value: range.relation };
    },
  };
  return {
    sync: vi.fn(async () => {}),
    document: {
      getSelection: () => selection,
      getParagraphByUniqueLocalId: () => paragraph,
    },
  };
};

const writtenStyles = (log) =>
  log.map(({ text, range }) => [text, range.font.subscript, range.font.superscript]);

describe('convertSelection', () => {
  test('writes each segment with its own style', async () => {
    const log = [];
    const context = createFakeContext({
      text: 'H2O',
      selectionText: 'H2O',
      matches: { H2O: [createFakeRange('H2O', { log })] },
    });

    const converted = await convertSelection(context);

    expect(converted).toBe(1);
    expect(log.map((entry) => entry.location)).toEqual(['Replace', 'After', 'After']);
    expect(writtenStyles(log)).toEqual([
      ['H', false, false],
      ['2', true, false],
      ['O', false, false],
    ]);
  });

  test('keeps a superscript charge after writing it', async () => {
    const log = [];
    const context = createFakeContext({
      text: 'Ca2+',
      selectionText: 'Ca2+',
      matches: { 'Ca2+': [createFakeRange('Ca2+', { log })] },
    });

    await convertSelection(context);

    expect(writtenStyles(log)).toEqual([
      ['Ca', false, false],
      ['2+', false, true],
    ]);
  });

  test('keeps the original font on every written segment, including an arrow', async () => {
    const log = [];
    const context = createFakeContext({
      text: 'CO2^',
      selectionText: 'CO2^',
      caretPieces: [createCaretPiece({ CO2: {} }, createFakeRange('CO2^', { log, fontName: 'Aptos' }))],
    });

    await convertSelection(context);

    expect(log.map(({ text, range }) => [text, range.font.name])).toEqual([
      ['CO', 'Aptos'],
      ['2', 'Aptos'],
      ['↑', 'Aptos'],
    ]);
  });

  test('leaves the font alone when the original text mixes fonts', async () => {
    const log = [];
    const context = createFakeContext({
      text: 'H2O',
      selectionText: 'H2O',
      matches: { H2O: [createFakeRange('H2O', { log, fontName: '' })] },
    });

    await convertSelection(context);

    expect(log.map(({ range }) => range.font.name)).toEqual([undefined, undefined, undefined]);
  });

  test('picks the right match when the token also appears inside a longer formula', async () => {
    const insideLonger = createFakeRange('H2');
    const standalone = createFakeRange('H2');
    const longer = createFakeRange('H2O');
    const context = createFakeContext({
      text: 'H2O a H2',
      selectionText: 'H2O a H2',
      matches: { H2: [insideLonger, standalone], H2O: [longer] },
    });
    insideLonger.insertText = vi.fn();
    standalone.insertText = vi.fn(() => ({ font: {}, insertText: () => ({ font: {} }) }));
    longer.insertText = vi.fn(() => ({
      font: {},
      insertText: () => ({ font: {}, insertText: () => ({ font: {} }) }),
    }));

    await convertSelection(context);

    expect(insideLonger.insertText).not.toHaveBeenCalled();
    expect(standalone.insertText).toHaveBeenCalledOnce();
  });

  test('skips text that is fully raised', async () => {
    const log = [];
    const context = createFakeContext({
      text: 'H2O',
      selectionText: 'H2O',
      matches: { H2O: [createFakeRange('H2O', { raised: true, log })] },
    });

    expect(await convertSelection(context)).toBe(0);
    expect(log).toEqual([]);
  });

  test('repairs a partly formatted formula', async () => {
    const log = [];
    const context = createFakeContext({
      text: 'CuSO4.5H2O',
      selectionText: 'CuSO4.5H2O',
      matches: { 'CuSO4.5H2O': [createFakeRange('CuSO4.5H2O', { plain: false, log })] },
    });

    expect(await convertSelection(context)).toBe(1);
    expect(log.map((entry) => entry.text).join('')).toBe('CuSO4·5H2O');
  });

  test('uses the current paragraph when nothing is selected', async () => {
    const log = [];
    const context = createFakeContext({
      text: 'Voda je H2O',
      selectionText: '',
      matches: { H2O: [createFakeRange('H2O', { log })] },
    });

    expect(await convertSelection(context)).toBe(1);
  });

  test('finds a caret in the middle of a formula without searching for it', async () => {
    const log = [];
    const tailStart = {};
    const tailPiece = { search: () => ({ items: [tailStart], load() {} }) };
    const context = createFakeContext({
      text: 'SO4^2-',
      selectionText: 'SO4^2-',
      caretPieces: [createCaretPiece({ SO4: {} }, createFakeRange('SO4^2-', { log }), tailStart), tailPiece],
    });

    expect(await convertSelection(context)).toBe(1);
    expect(log.map((entry) => entry.text).join('')).toBe('SO42−');
  });

  test('takes the part of the text that ends with the right caret', async () => {
    const firstLog = [];
    const secondLog = [];
    const context = createFakeContext({
      text: 'CO2^ H2^',
      selectionText: 'CO2^ H2^',
      caretPieces: [
        createCaretPiece({ CO2: {} }, createFakeRange('CO2^', { log: firstLog })),
        createCaretPiece({ H2: {} }, createFakeRange('H2^', { log: secondLog })),
      ],
    });

    expect(await convertSelection(context)).toBe(2);
    expect(firstLog.map((entry) => entry.text).join('')).toBe('CO2↑');
    expect(secondLog.map((entry) => entry.text).join('')).toBe('H2↑');
  });

  test('replaces only the formula when text comes before its caret', async () => {
    const log = [];
    const text = 'CaCO3 -> CaO + CO2^';
    const context = createFakeContext({
      text,
      selectionText: text,
      matches: {
        CaCO3: [createFakeRange('CaCO3')],
        '->': [createFakeRange('->')],
      },
      caretPieces: [createCaretPiece({ CO2: {} }, createFakeRange('CO2^', { log }))],
    });

    expect(await convertSelection(context)).toBe(3);
    expect(log.map((entry) => entry.text).join('')).toBe('CO2↑');
  });

  test('never overwrites a range holding more than the formula', async () => {
    const log = [];
    const text = 'CaCO3 -> CaO + CO2^';
    const context = createFakeContext({
      text,
      selectionText: text,
      caretPieces: [createCaretPiece({ CO2: {} }, createFakeRange(text, { log }))],
    });

    await convertSelection(context);

    expect(log).toEqual([]);
  });

  test('reports what Word returned instead of a skipped formula', async () => {
    const skipped = [];
    onFormulaSkipped((source, found) => skipped.push([source, found]));
    const context = createFakeContext({
      text: 'CO2^',
      selectionText: 'CO2^',
      caretPieces: [createCaretPiece({ CO2: {} }, createFakeRange('CO2^\r'))],
    });

    await convertSelection(context);
    onFormulaSkipped(() => {});

    expect(skipped).toEqual([['CO2^', 'CO2^\r']]);
  });

  test('skips a caret formula when Word splits the text differently', async () => {
    const context = createFakeContext({ text: 'CO2^', selectionText: 'CO2^', caretPieces: [] });

    expect(await convertSelection(context)).toBe(0);
  });

  test('skips a caret formula when the text before the caret is not found', async () => {
    const context = createFakeContext({
      text: 'CO2^',
      selectionText: 'CO2^',
      caretPieces: [createCaretPiece({}, createFakeRange('CO2^'))],
    });

    expect(await convertSelection(context)).toBe(0);
  });

  test('skips a token when the search count does not match the text', async () => {
    const context = createFakeContext({ text: 'H2O', selectionText: 'H2O', matches: { H2O: [] } });

    expect(await convertSelection(context)).toBe(0);
  });
});

describe('convertParagraph', () => {
  test('converts finished formulas and remembers them', async () => {
    const log = [];
    const context = createFakeContext({
      text: 'Voda H2O ',
      matches: { H2O: [createFakeRange('H2O', { log })] },
    });

    const { converted, seen } = await convertParagraph(context, 'p1', new Map());

    expect(converted).toBe(1);
    expect(seen.get('H2O')).toBe(1);
  });

  test.each([0, 1, 2])('waits while the caret is at offset %i of the formula', async (caretOffset) => {
    const log = [];
    const context = createFakeContext({
      text: 'H2 ',
      caretOffset,
      matches: { H2: [createFakeRange('H2', { log })] },
    });

    const { converted, seen } = await convertParagraph(context, 'p1', new Map());

    expect(converted).toBe(0);
    expect(log).toEqual([]);
    expect(seen.has('H2')).toBe(false);
  });

  test('converts a formula when the caret is elsewhere in the paragraph', async () => {
    const log = [];
    const context = createFakeContext({
      text: 'psal H2O dnes',
      caretOffset: 2,
      matches: { H2O: [createFakeRange('H2O', { log })] },
    });

    expect((await convertParagraph(context, 'p1', new Map())).converted).toBe(1);
  });

  test('waits on a formula ending the paragraph while the caret is in it', async () => {
    const log = [];
    const context = createFakeContext({
      text: 'H2',
      matches: { H2: [createFakeRange('H2', { log })] },
    });

    const { converted, seen } = await convertParagraph(context, 'p1', new Map());

    expect(converted).toBe(0);
    expect(log).toEqual([]);
    expect(seen.has('H2')).toBe(false);
  });

  test('converts a formula ending the paragraph once Enter moved the caret out', async () => {
    const log = [];
    const context = createFakeContext({
      text: 'H2O',
      caretVsParagraph: 'After',
      matches: { H2O: [createFakeRange('H2O', { log })] },
    });

    const { converted } = await convertParagraph(context, 'p1', new Map());

    expect(converted).toBe(1);
  });

  test('keeps the earlier count for a formula it had to wait on', async () => {
    const context = createFakeContext({
      text: 'H2 H2',
      matches: {
        H2: [createFakeRange('H2', { plain: false }), createFakeRange('H2')],
      },
    });

    const { seen } = await convertParagraph(context, 'p1', new Map([['H2', 1]]));

    expect(seen.get('H2')).toBe(1);
  });

  test('rewrites a formula that grew out of an already converted part', async () => {
    const log = [];
    const context = createFakeContext({
      text: 'CuSO4.5H2O ',
      matches: { 'CuSO4.5H2O': [createFakeRange('CuSO4.5H2O', { plain: false, log })] },
    });

    const { converted } = await convertParagraph(context, 'p1', new Map());

    expect(converted).toBe(1);
    expect(log.map((entry) => entry.text).join('')).toBe('CuSO4·5H2O');
  });

  test('leaves a new formula alone when all of it is already raised', async () => {
    const log = [];
    const context = createFakeContext({
      text: 'H2O ',
      matches: { H2O: [createFakeRange('H2O', { raised: true, log })] },
    });

    const { converted } = await convertParagraph(context, 'p1', new Map());

    expect(converted).toBe(0);
    expect(log).toEqual([]);
  });

  test('leaves a formula alone after the user undid its conversion', async () => {
    const log = [];
    const context = createFakeContext({
      text: 'H2O ',
      matches: { H2O: [createFakeRange('H2O', { log })] },
    });

    const { converted } = await convertParagraph(context, 'p1', new Map([['H2O', 1]]));

    expect(converted).toBe(0);
    expect(log).toEqual([]);
  });
});

describe('createAutoConverter', () => {
  const DELAY_MS = 100;

  beforeEach(() => vi.useFakeTimers());
  afterEach(() => vi.useRealTimers());

  const POLL_MS = 300;
  const QUIET_MS = 3000;
  const paragraphState = (id, text) => ({ id, text });

  const setup = (
    runWord = vi.fn(async () => {}),
    readActive = vi.fn(async () => paragraphState('p1', 'text')),
  ) => {
    const onError = vi.fn();
    const onConverted = vi.fn();
    const converter = createAutoConverter({
      runWord,
      readActive,
      onError,
      onConverted,
      delayMs: DELAY_MS,
      pollIntervalMs: POLL_MS,
    });
    return { converter, runWord, readActive, onError, onConverted };
  };

  test('converts a change it finds itself, before Word reports it', async () => {
    const readActive = vi
      .fn()
      .mockResolvedValueOnce(paragraphState('p1', 'H2O'))
      .mockResolvedValue(paragraphState('p1', 'H2O '));
    const { converter, runWord } = setup(undefined, readActive);

    converter.handleEvent({ source: 'Local', uniqueLocalIds: [] });
    await vi.advanceTimersByTimeAsync(POLL_MS);
    expect(runWord).not.toHaveBeenCalled();
    await vi.advanceTimersByTimeAsync(POLL_MS);

    expect(runWord).toHaveBeenCalledOnce();
  });

  test('also converts the paragraph the caret just left', async () => {
    const readActive = vi
      .fn()
      .mockResolvedValueOnce(paragraphState('p1', 'H2O'))
      .mockResolvedValue(paragraphState('p2', ''));
    const { converter, runWord } = setup(undefined, readActive);

    converter.handleEvent({ source: 'Local', uniqueLocalIds: [] });
    await vi.advanceTimersByTimeAsync(POLL_MS * 2);

    expect(runWord).toHaveBeenCalledTimes(2);
  });

  test('converts a paragraph the caret left after a quiet period', async () => {
    let active = paragraphState('p1', 'CO2^');
    const readActive = vi.fn(async () => active);
    const { converter, runWord } = setup(undefined, readActive);

    converter.handleEvent({ source: 'Local', uniqueLocalIds: [] });
    await vi.advanceTimersByTimeAsync(QUIET_MS + POLL_MS);
    expect(runWord).not.toHaveBeenCalled();

    active = paragraphState('p2', '');
    converter.handleEvent({ source: 'Local', uniqueLocalIds: ['p2'] });
    await vi.advanceTimersByTimeAsync(POLL_MS);

    expect(runWord.mock.calls.length).toBeGreaterThan(1);
  });

  test('does nothing while the text stays the same', async () => {
    const { converter, runWord } = setup();

    converter.handleEvent({ source: 'Local', uniqueLocalIds: [] });
    await vi.advanceTimersByTimeAsync(POLL_MS * 5);

    expect(runWord).not.toHaveBeenCalled();
  });

  test('stops asking Word after a quiet period', async () => {
    const { converter, readActive } = setup();

    converter.handleEvent({ source: 'Local', uniqueLocalIds: [] });
    await vi.advanceTimersByTimeAsync(QUIET_MS + POLL_MS);
    const callsWhenQuiet = readActive.mock.calls.length;
    await vi.advanceTimersByTimeAsync(POLL_MS * 5);

    expect(readActive.mock.calls.length).toBe(callsWhenQuiet);
  });

  test('stop ends the checking', async () => {
    const { converter, readActive } = setup();

    converter.handleEvent({ source: 'Local', uniqueLocalIds: [] });
    converter.stop();
    await vi.advanceTimersByTimeAsync(POLL_MS * 5);

    expect(readActive).not.toHaveBeenCalled();
  });

  test('reports a failed check once and stops checking', async () => {
    const failure = new Error('boom');
    const readActive = vi.fn(async () => { throw failure; });
    const { converter, onError } = setup(undefined, readActive);

    converter.handleEvent({ source: 'Local', uniqueLocalIds: [] });
    await vi.advanceTimersByTimeAsync(POLL_MS * 5);

    expect(onError).toHaveBeenCalledOnce();
    expect(onError).toHaveBeenCalledWith(failure);
    expect(readActive).toHaveBeenCalledOnce();
  });

  test('groups rapid changes into one run per paragraph', async () => {
    const { converter, runWord } = setup();

    converter.handleEvent({ source: 'Local', uniqueLocalIds: ['p1'] });
    converter.handleEvent({ source: 'Local', uniqueLocalIds: ['p1', 'p2'] });
    await vi.advanceTimersByTimeAsync(DELAY_MS);

    expect(runWord).toHaveBeenCalledTimes(2);
  });

  test('ignores changes made by a coauthor', async () => {
    const { converter, runWord } = setup();

    converter.handleEvent({ source: 'Remote', uniqueLocalIds: ['p1'] });
    await vi.advanceTimersByTimeAsync(DELAY_MS);

    expect(runWord).not.toHaveBeenCalled();
  });

  test('stays quiet when the paragraph no longer exists', async () => {
    const gone = Object.assign(new Error('gone'), { code: 'ItemNotFound' });
    const { converter, onError } = setup(vi.fn(async () => { throw gone; }));

    converter.handleEvent({ source: 'Local', uniqueLocalIds: ['p1'] });
    await vi.advanceTimersByTimeAsync(DELAY_MS);

    expect(onError).not.toHaveBeenCalled();
  });

  test('reports other failures', async () => {
    const failure = new Error('boom');
    const { converter, onError } = setup(vi.fn(async () => { throw failure; }));

    converter.handleEvent({ source: 'Local', uniqueLocalIds: ['p1'] });
    await vi.advanceTimersByTimeAsync(DELAY_MS);

    expect(onError).toHaveBeenCalledWith(failure);
  });

  test('cancels pending work when stopped', async () => {
    const { converter, runWord } = setup();

    converter.handleEvent({ source: 'Local', uniqueLocalIds: ['p1'] });
    converter.stop();
    await vi.advanceTimersByTimeAsync(DELAY_MS);

    expect(runWord).not.toHaveBeenCalled();
  });
});
