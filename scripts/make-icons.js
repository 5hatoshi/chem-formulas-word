/** Generates the ribbon icons (a teal tile with "H₂") without any image dependency. */
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { crc32, deflateSync } from 'node:zlib';

const OUTPUT_DIR = join(dirname(fileURLToPath(import.meta.url)), '..', 'addin', 'assets');
const SIZES = [16, 32, 64, 80];
const GRID_SIZE = 16;
const TEAL = [15, 110, 110, 255];
const WHITE = [255, 255, 255, 255];
const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const COLOR_TYPE_RGBA = 6;
const BIT_DEPTH = 8;

const GLYPH = [
  '................',
  '................',
  '.##...##........',
  '.##...##........',
  '.##...##........',
  '.##...##........',
  '.#######........',
  '.#######........',
  '.##...##........',
  '.##...##..####..',
  '.##...##.....##.',
  '.##...##....##..',
  '.##...##...##...',
  '..........##....',
  '..........#####.',
  '................',
];

const chunk = (type, data) => {
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const checksum = Buffer.alloc(4);
  checksum.writeUInt32BE(crc32(body));
  return Buffer.concat([length, body, checksum]);
};

const pixelAt = (x, y, size) => {
  const row = Math.floor((y * GRID_SIZE) / size);
  const column = Math.floor((x * GRID_SIZE) / size);
  return GLYPH[row][column] === '#' ? WHITE : TEAL;
};

const makePng = (size) => {
  const rows = Array.from({ length: size }, (_, y) => {
    const pixels = Array.from({ length: size }, (__, x) => pixelAt(x, y, size)).flat();
    return Buffer.from([0, ...pixels]);
  });
  const header = Buffer.alloc(13);
  header.writeUInt32BE(size, 0);
  header.writeUInt32BE(size, 4);
  header.writeUInt8(BIT_DEPTH, 8);
  header.writeUInt8(COLOR_TYPE_RGBA, 9);
  return Buffer.concat([
    PNG_SIGNATURE,
    chunk('IHDR', header),
    chunk('IDAT', deflateSync(Buffer.concat(rows))),
    chunk('IEND', Buffer.alloc(0)),
  ]);
};

mkdirSync(OUTPUT_DIR, { recursive: true });
SIZES.forEach((size) => writeFileSync(join(OUTPUT_DIR, `icon-${size}.png`), makePng(size)));
