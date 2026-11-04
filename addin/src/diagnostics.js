/* global Office, Word */
/** Temporary measurements of how a Word platform handles a caret. Remove after use. */
import { BUILD, convertSelection } from './word.js';

const describeError = (error) =>
  `CHYBA ${error?.code ?? ''} ${error?.message ?? error} ${JSON.stringify(error?.debugInfo ?? {})}`;

/** @param {string} label @param {() => Promise<string>} measure @returns {Promise<string>} */
const probe = async (label, measure) => {
  try {
    return `${label}: ${await measure()}`;
  } catch (error) {
    return `${label}: ${describeError(error)}`;
  }
};

const readParagraph = () =>
  Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    paragraph.load('text');
    await context.sync();
    return paragraph.text;
  });

const countSearch = (query) =>
  Word.run(async (context) => {
    const results = context.document.getSelection().paragraphs.getFirst().search(query, { matchCase: true });
    results.load('items');
    await context.sync();
    return String(results.items.length);
  });

const readPieces = () =>
  Word.run(async (context) => {
    const pieces = context.document.getSelection().paragraphs.getFirst().getTextRanges(['^'], false);
    pieces.load('items/text');
    await context.sync();
    return JSON.stringify(pieces.items.map((piece) => piece.text));
  });

const describeCodes = (text) =>
  [...text]
    .filter((char) => !/[A-Za-z0-9 ]/.test(char))
    .map((char) => `${char}=U+${char.codePointAt(0).toString(16)}`)
    .join(' ');

/** @returns {Promise<string>} A report on the paragraph holding the caret. */
export async function measureCaret() {
  const text = await readParagraph();
  const token = text.split(/\s+/).find((word) => word.includes('^')) ?? 'CO2^';
  const bare = token.replaceAll('^', '');
  const measurements = [
    [`hledat "${bare}"`, () => countSearch(bare)],
    [`hledat "${token.replaceAll('^', '^^')}"`, () => countSearch(token.replaceAll('^', '^^'))],
    [`hledat "${token}"`, () => countSearch(token)],
    ['rozdělit u ^', readPieces],
    [
      'úprava',
      async () => {
        const converted = await Word.run(convertSelection);
        return `upraveno ${converted}, text ${JSON.stringify(await readParagraph())}`;
      },
    ],
  ];
  const lines = [];
  for (const [label, measure] of measurements) {
    lines.push(await probe(label, measure));
  }
  return [
    `verze ${BUILD}`,
    `${Office.context.platform} ${Office.context.diagnostics?.version ?? ''}`,
    `text ${JSON.stringify(text)}`,
    `znaky ${describeCodes(text)}`,
    ...lines,
  ].join('\n');
}
