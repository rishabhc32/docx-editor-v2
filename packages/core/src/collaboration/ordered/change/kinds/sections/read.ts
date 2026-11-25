// What the section kinds share (task 9.8): a section's holder as a target, and the readers of a
// section's writes, which a checkpoint's manifest reads too.

import { NOTE_ID_MAX } from '../../../../../store/package/note-nodes.ts';
import type { OrderedNoteProperties, OrderedNotePropertiesOp } from '../../../model/types.ts';
import { at, bool, fieldPath, fieldsOf, int, Invalid, oneOf, type Fields } from '../../read.ts';
import type { OrderedChangeAddress } from '../../types.ts';
import type { OrderedTargetReader } from '../spec.ts';
import type { OrderedChangeFurnitureOptions, OrderedChangePageSetup } from './property.ts';

/**
 * Name a section's holder as a target: the body's own section item (`z:`) as a block item, else
 * a paragraph mark (`m:`, `k:`).
 */
export function touchSection(reader: OrderedTargetReader, holder: OrderedChangeAddress): void {
  reader.item(holder, holder.space.startsWith('z:') ? 'zero' : 'mark');
}

/** A page side from 1 twip up to 44 inches, a margin up to 22 inches (the engine's bounds). */
const PAGE = 63_360;
const MARGIN = 31_680;

const PAGE_FIELDS = ['pageWidthTwips', 'pageHeightTwips'] as const;
const MARGIN_FIELDS = [
  'marginTopTwips',
  'marginRightTwips',
  'marginBottomTwips',
  'marginLeftTwips',
] as const;

/** A page setup, each field bounded as the engine bounds it. */
export function readPage(value: unknown, path: string): OrderedChangePageSetup {
  const fields = fieldsOf(value, path, new Set([...PAGE_FIELDS, ...MARGIN_FIELDS, 'orientation']));
  const out: Record<string, number | string> = {};
  for (const name of PAGE_FIELDS) {
    if (fields.has(name)) {
      out[name] = int(fields.get(name), at(path, name), 1, PAGE);
    }
  }

  for (const name of MARGIN_FIELDS) {
    if (fields.has(name)) {
      out[name] = int(fields.get(name), at(path, name), 0, MARGIN);
    }
  }

  if (fields.has('orientation')) {
    out.orientation = oneOf(fields.get('orientation'), at(path, 'orientation'), [
      'portrait',
      'landscape',
    ] as const);
  }

  if (Object.keys(out).length === 0) {
    throw new Invalid('empty', path);
  }

  return out;
}

/** The furniture options, each distance bounded as a margin is. */
export function readFurniture(value: unknown, path: string): OrderedChangeFurnitureOptions {
  const fields = fieldsOf(
    value,
    path,
    new Set(['titlePage', 'headerDistanceTwips', 'footerDistanceTwips'])
  );
  const out: { titlePage?: boolean; headerDistanceTwips?: number; footerDistanceTwips?: number } =
    {};
  if (fields.has('titlePage')) {
    out.titlePage = bool(fields.get('titlePage'), at(path, 'titlePage'));
  }

  for (const name of ['headerDistanceTwips', 'footerDistanceTwips'] as const) {
    if (fields.has(name)) {
      out[name] = int(fields.get(name), at(path, name), 0, MARGIN);
    }
  }

  if (Object.keys(out).length === 0) {
    throw new Invalid('empty', path);
  }

  return out;
}

/** `ST_NumberFormat`: the number formats a note may take. */
const NUMBER_FORMATS = [
  'decimal',
  'upperRoman',
  'lowerRoman',
  'upperLetter',
  'lowerLetter',
  'ordinal',
  'cardinalText',
  'ordinalText',
  'hex',
  'chicago',
  'ideographDigital',
  'japaneseCounting',
  'aiueo',
  'iroha',
  'decimalFullWidth',
  'decimalHalfWidth',
  'japaneseLegal',
  'japaneseDigitalTenThousand',
  'decimalEnclosedCircle',
  'decimalFullWidth2',
  'aiueoFullWidth',
  'irohaFullWidth',
  'decimalZero',
  'bullet',
  'ganada',
  'chosung',
  'decimalEnclosedFullstop',
  'decimalEnclosedParen',
  'decimalEnclosedCircleChinese',
  'ideographEnclosedCircle',
  'ideographTraditional',
  'ideographZodiac',
  'ideographZodiacTraditional',
  'taiwaneseCounting',
  'ideographLegalTraditional',
  'taiwaneseCountingThousand',
  'taiwaneseDigital',
  'chineseCounting',
  'chineseLegalSimplified',
  'chineseCountingThousand',
  'koreanDigital',
  'koreanCounting',
  'koreanLegal',
  'koreanDigital2',
  'vietnameseCounting',
  'russianLower',
  'russianUpper',
  'none',
  'numberInDash',
  'hebrew1',
  'hebrew2',
  'arabicAlpha',
  'arabicAbjad',
  'hindiVowels',
  'hindiConsonants',
  'hindiNumbers',
  'hindiCounting',
  'thaiLetters',
  'thaiNumbers',
  'thaiCounting',
  'bahtText',
  'dollarText',
  'custom',
] as const;

/** Where footnotes and endnotes may go (`ST_FtnPos`, `ST_EdnPos`), and when they restart. */
const FOOTNOTE_POSITIONS = ['pageBottom', 'beneathText', 'sectEnd', 'docEnd'] as const;
const ENDNOTE_POSITIONS = ['sectEnd', 'docEnd'] as const;
const NUM_RESTARTS = ['continuous', 'eachSect', 'eachPage'] as const;

/** One kind's note properties: each field bounded, the start as the engine bounds it. */
function noteProperties(
  value: unknown,
  path: string,
  positions: readonly string[]
): OrderedNoteProperties {
  const f = fieldsOf(value, path, new Set(['position', 'numFmt', 'numStart', 'numRestart']));
  return {
    ...(f.has('position')
      ? { position: oneOf(f.get('position'), at(path, 'position'), positions) }
      : {}),
    ...(f.has('numFmt')
      ? { numFmt: oneOf(f.get('numFmt'), at(path, 'numFmt'), NUMBER_FORMATS) }
      : {}),
    ...(f.has('numStart')
      ? { numStart: int(f.get('numStart'), at(path, 'numStart'), 0, NOTE_ID_MAX) }
      : {}),
    ...(f.has('numRestart')
      ? { numRestart: oneOf(f.get('numRestart'), at(path, 'numRestart'), NUM_RESTARTS) }
      : {}),
  };
}

/**
 * The footnotes' and the endnotes' properties a write gives, at least one of them.
 *
 * @throws Invalid when they do not read.
 */
export function noteWrites(f: Fields): Omit<OrderedNotePropertiesOp, 'op'> {
  if (!f.has('footnote') && !f.has('endnote')) {
    throw new Invalid('missing-field', fieldPath(f, 'footnote'), 'footnote or endnote');
  }

  const read = (key: string, positions: readonly string[]) =>
    f.has(key) ? { [key]: noteProperties(f.get(key), fieldPath(f, key), positions) } : {};
  return { ...read('footnote', FOOTNOTE_POSITIONS), ...read('endnote', ENDNOTE_POSITIONS) };
}
