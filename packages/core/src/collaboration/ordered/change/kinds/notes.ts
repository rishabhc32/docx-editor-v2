// Change kinds of group 9.9 (notes and equations). A text box is a story the first kinds edit.

import { DEFAULT_OMML_LIMITS } from '../../../../store/package/omml-equation.ts';
import { replaceEquation } from '../../edits/equations.ts';
import { convertNotes } from '../../edits/note-convert.ts';
import { setNoteProperties } from '../../edits/note-properties.ts';
import { shownItemAt } from '../../edits/places.ts';
import { NOTE_ID_MAX } from '../../../../store/package/note-nodes.ts';
import type {
  OrderedDocument,
  OrderedItem,
  OrderedNoteProperties,
  OrderedStory,
} from '../../model/types.ts';
import { at, fieldPath, fieldsOf, int, Invalid, oneOf, plainText, required } from '../read.ts';
import type { OrderedChangeAddress, OrderedEditBase } from '../types.ts';
import { address, addresses } from '../values.ts';
import { runKindEdit } from './run.ts';
import { orderedKindSpec, type OrderedKindSpec } from './spec.ts';

/**
 * A new linear form for one equation of an equation object (`m:oMath`, or a display
 * `m:oMathPara`), or that equation removed from its display.
 *
 * @public
 */
export interface OrderedEquationChange extends OrderedEditBase {
  readonly kind: 'equation';
  /** The equation object's address. */
  readonly object: OrderedChangeAddress;
  /** Which equation of the object: 0 for an `m:oMath`, its place among a display's equations. */
  readonly equation: number;
  /** The new linear form; null removes the equation from a display that holds others. */
  readonly linear: string | null;
}

/**
 * Notes of one kind converted to the other kind (Word's Convert, for one note or for every note
 * of a kind). Word converts a note directly, so this is a direct edit only. `part` is the notes
 * story that holds the notes.
 *
 * @public
 */
export interface OrderedConvertNoteChange extends OrderedEditBase {
  readonly kind: 'convert-note';
  /** The notes' open items, in the order their copies take note ids. */
  readonly notes: readonly OrderedChangeAddress[];
}

/**
 * Where a section's properties are: the paragraph mark that ends the section (`mark`), or the
 * body's last section (`body`).
 *
 * @public
 */
export type OrderedNoteSection =
  | { readonly mark: OrderedChangeAddress }
  | { readonly body: OrderedChangeAddress };

/**
 * The notes' properties (number format, start, restart and place) of the document's settings
 * or of one section, as Word's Footnote and Endnote dialog writes them. A direct edit only.
 *
 * @public
 */
export interface OrderedNotePropertiesChange extends OrderedEditBase {
  readonly kind: 'note-properties';
  /** The section; absent for the document's settings. */
  readonly section?: OrderedNoteSection;
  readonly footnote?: OrderedNoteProperties;
  readonly endnote?: OrderedNoteProperties;
}

/**
 * The change kinds of group 9.9.
 *
 * @public
 */
export type OrderedNoteChange =
  | OrderedEquationChange
  | OrderedConvertNoteChange
  | OrderedNotePropertiesChange;

/** Most equations one display holds that a change names. */
const MAX_EQUATIONS = 255;

const EQUATION = orderedKindSpec<OrderedEquationChange>({
  kind: 'equation',
  fields: ['object', 'equation', 'linear'],
  read: (f) => {
    const linear = required(f, 'linear');
    return {
      object: address(required(f, 'object'), fieldPath(f, 'object')),
      equation: int(required(f, 'equation'), fieldPath(f, 'equation'), 0, MAX_EQUATIONS),
      linear:
        linear === null
          ? null
          : plainText(linear, fieldPath(f, 'linear'), 1, DEFAULT_OMML_LIMITS.maxTextLength),
    };
  },
  touches: (change, reader) => {
    reader.item(change.object, 'atom');
  },
  apply: (change, input) => {
    const object = shownItemAt(input.doc, input.story, change.object, 'atom');
    return runKindEdit(input, (ctx) =>
      replaceEquation(ctx, object, change.equation, change.linear)
    );
  },
  // It strikes the old equation, and its new one changes what a reader sees
  policy: { text: true, properties: false, strikes: true },
  publish: 'paragraphs',
});

const CONVERT_NOTE = orderedKindSpec<OrderedConvertNoteChange>({
  kind: 'convert-note',
  fields: ['notes'],
  read: (f) => ({ notes: addresses(required(f, 'notes'), fieldPath(f, 'notes')) }),
  touches: (change, reader) => {
    for (const note of change.notes) reader.item(note, 'open', ['note']);
  },
  apply: (change, input) => {
    const opens = change.notes.map((note) =>
      shownItemAt(input.doc, input.story, note, 'open', ['note'])
    );
    return runKindEdit(input, (ctx) => convertNotes(ctx, opens));
  },
  // The text a reader sees stays: a party's would change the package (the room's rule: `package`)
  policy: { text: false, properties: false },
  // It writes in both notes stories and where the references are
  publish: 'suggestions',
});

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

/** A section, by the item that holds its `w:sectPr`. */
function noteSection(value: unknown, path: string): OrderedNoteSection {
  const f = fieldsOf(value, path, new Set(['mark', 'body']));
  if (f.has('mark') === f.has('body'))
    throw new Invalid('invalid-value', path, 'one of mark and body');
  return f.has('mark')
    ? { mark: address(f.get('mark'), at(path, 'mark')) }
    : { body: address(f.get('body'), at(path, 'body')) };
}

/** The shown item that holds a section's `w:sectPr`; null for the document's settings. */
function sectionHolder(
  doc: OrderedDocument,
  story: OrderedStory,
  section: OrderedNoteSection | undefined
): OrderedItem | null {
  if (section === undefined) {
    return null;
  }

  if ('mark' in section) {
    return shownItemAt(doc, story, section.mark, 'mark');
  }

  return shownItemAt(doc, story, section.body, 'zero');
}

const NOTE_PROPERTIES = orderedKindSpec<OrderedNotePropertiesChange>({
  kind: 'note-properties',
  fields: ['section', 'footnote', 'endnote'],
  read: (f) => {
    if (!f.has('footnote') && !f.has('endnote'))
      throw new Invalid('missing-field', fieldPath(f, 'footnote'), 'footnote or endnote');
    return {
      ...(f.has('section')
        ? { section: noteSection(f.get('section'), fieldPath(f, 'section')) }
        : {}),
      ...(f.has('footnote')
        ? {
            footnote: noteProperties(
              f.get('footnote'),
              fieldPath(f, 'footnote'),
              FOOTNOTE_POSITIONS
            ),
          }
        : {}),
      ...(f.has('endnote')
        ? { endnote: noteProperties(f.get('endnote'), fieldPath(f, 'endnote'), ENDNOTE_POSITIONS) }
        : {}),
    };
  },
  touches: (change, reader) => {
    const section = change.section;
    if (section === undefined) return;
    if ('mark' in section) reader.item(section.mark, 'mark');
    else reader.item(section.body, 'zero');
  },
  apply: (change, input) => {
    const holder = sectionHolder(input.doc, input.story, change.section);
    const op = {
      op: 'setNoteProperties' as const,
      ...(change.footnote ? { footnote: change.footnote } : {}),
      ...(change.endnote ? { endnote: change.endnote } : {}),
    };
    return runKindEdit(input, (ctx) => setNoteProperties(ctx, holder, op));
  },
  // The text a reader sees stays: a party's would change the package (the room's rule: `package`)
  policy: { text: false, properties: true },
  // A section's mark, or the settings part, which the change's suggestion owns
  publish: 'suggestions',
});

/** The group's specs: one for each kind of {@link OrderedNoteChange}. */
export const NOTE_KINDS: { readonly [K in OrderedNoteChange['kind']]: OrderedKindSpec } = {
  equation: EQUATION,
  'convert-note': CONVERT_NOTE,
  'note-properties': NOTE_PROPERTIES,
};
