// Paragraph ops (task 6.1): Enter, joins, paragraph properties, lists and tab stops.

import { addressOf, paragraphAtIndex, placeAt } from '../paragraphs.ts';
import { currentListLevel, paragraphPropertyDelta } from '../properties.ts';
import { OrderedClientRefusal } from '../refusal.ts';
import { listLevelChanges } from './list-level.ts';
import { isNewList } from './paste/definitions.ts';
import { newListChanges } from './paste/new-list.ts';
import {
  PARTY_SUGGESTS,
  translator,
  type OrderedChangeBody,
  type OrderedOpOption,
  type OrderedOpTable,
  type OrderedTranslateBatch,
  type OrderedTranslateStep,
} from './types.ts';

/** Enter: a split, and for an Enter at a heading's end the tail's next style. */
export function splitStep(
  batch: OrderedTranslateBatch,
  paragraphId: string,
  offset: number,
  tail?: string | null
): OrderedTranslateStep {
  return () => {
    const p = batch.paragraph(paragraphId);
    const split: OrderedChangeBody = {
      kind: 'split',
      part: p.story.name,
      place: placeAt(p, offset),
    };
    if (tail === undefined) return [split];
    const style = tail ?? batch.defaultParagraphStyle;
    if (style === null)
      throw new OrderedClientRefusal('ordered-unsupported-op', 'splitParagraph.tailStyleId');

    // The list keeps the paragraph's own mark on the tail, so the tail is this node id still
    return [
      split,
      {
        kind: 'paragraph-property',
        part: p.story.name,
        marks: [addressOf(p.mark)],
        properties: [{ localName: 'pStyle', attributes: { val: style } }],
      },
    ];
  };
}

export const PARAGRAPH_OPS: OrderedOpTable<
  | 'splitParagraph'
  | 'splitParagraphMany'
  | 'proposeParagraphMerge'
  | 'joinParagraphs'
  | 'setParagraphMarkRevision'
  | 'setParagraphProperties'
  | 'setListLevel'
  | 'setListNumbering'
  | 'setParagraphTabStops'
> = {
  splitParagraph: {
    status: 'translated',
    group: '6.1',
    changes: ['split', 'paragraph-property'],
    party: PARTY_SUGGESTS,
    tracking: 'either',
    note: "Enter; with `tailStyleId`, the tail's next style. Tracked by its mark revision.",
    translate: translator('splitParagraph', (op, batch) =>
      splitStep(batch, op.paragraphId, op.offset, op.tailStyleId)
    ),
  },
  splitParagraphMany: {
    status: 'translated',
    group: '6.1',
    changes: ['split'],
    party: PARTY_SUGGESTS,
    note: 'A paste of several lines: every place read before the first split.',
    translate: translator('splitParagraphMany', (op, batch) => () => {
      const p = batch.paragraph(op.paragraphId);

      // Addresses do not move: every split place is read before the first split
      return op.offsets.map(
        (offset): OrderedChangeBody => ({
          kind: 'split',
          part: p.story.name,
          place: placeAt(p, offset),
        })
      );
    }),
  },
  proposeParagraphMerge: {
    status: 'translated',
    group: '6.1',
    changes: ['join'],
    party: PARTY_SUGGESTS,
    tracking: 'attribution',
    note: "Backspace at a paragraph's start in suggesting mode: the mark before it struck.",
    translate: translator('proposeParagraphMerge', (op, batch) => () => {
      const p = batch.paragraph(op.paragraphId);
      if (p.index === 0)
        throw new OrderedClientRefusal('ordered-refused', 'no-paragraph-before', op.paragraphId);
      const previous = paragraphAtIndex(batch.doc, p.story, p.index - 1);
      return [{ kind: 'join', part: p.story.name, mark: addressOf(previous.mark) }];
    }),
  },
  joinParagraphs: {
    status: 'translated',
    group: '6.1',
    changes: ['join'],
    party: {
      code: 'policy-direct',
      reason: 'text',
      why: 'a join with tracking off; suggesting mode sends `proposeParagraphMerge`',
    },
    tracking: 'direct',
    note: 'Backspace across paragraphs with tracking off.',
    translate: translator('joinParagraphs', (op, batch) => () => {
      const first = batch.paragraph(op.firstId);
      const second = batch.paragraph(op.secondId);
      if (second.story !== first.story || second.index !== first.index + 1)
        throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'not adjacent');
      return [{ kind: 'join', part: first.story.name, mark: addressOf(first.mark) }];
    }),
  },
  setParagraphMarkRevision: {
    status: 'translated',
    group: '6.1',
    changes: ['join'],
    party: PARTY_SUGGESTS,
    note: "A struck mark (`del`); an inserted one (`ins`) only beside its split, as Enter's record.",
    translate: translator('setParagraphMarkRevision', (op, batch) => {
      if (op.kind === 'ins')
        throw new OrderedClientRefusal('ordered-unsupported-op', 'setParagraphMarkRevision.ins');
      return () => {
        const p = batch.paragraph(op.paragraphId);
        return [{ kind: 'join', part: p.story.name, mark: addressOf(p.mark) }];
      };
    }),
  },
  setParagraphProperties: {
    status: 'translated',
    group: '6.1',
    changes: ['paragraph-property'],
    party: PARTY_SUGGESTS,
    note: 'Alignment, spacing, indents, a style: only the properties the press changes.',
    translate: translator('setParagraphProperties', (op, batch) => () => {
      const p = batch.paragraph(op.paragraphId);
      const properties = paragraphPropertyDelta(batch.doc, p.story, p.mark, op.properties);
      if (properties === null) return [];
      return [
        { kind: 'paragraph-property', part: p.story.name, marks: [addressOf(p.mark)], properties },
      ];
    }),
  },
  setListLevel: {
    status: 'translated',
    group: '6.1',
    changes: ['paragraph-property', 'list-level'],
    party: PARTY_SUGGESTS,
    note: "A list item's level (Tab and Shift+Tab in a list); a level the list does not declare travels as a `list-level` change.",
    translate: translator('setListLevel', (op, batch) => () => {
      const declared = listLevelChanges(op, batch);
      if (declared !== null) return declared;
      const p = batch.paragraph(op.paragraphId);
      return [
        {
          kind: 'paragraph-property',
          part: p.story.name,
          marks: [addressOf(p.mark)],
          properties: [
            {
              localName: 'numPr',
              children: [{ localName: 'ilvl', attributes: { val: String(op.level) } }],
            },
          ],
        },
      ];
    }),
  },
  setListNumbering: {
    status: 'translated',
    group: '6.1',
    changes: ['paragraph-property', 'new-list'],
    party: PARTY_SUGGESTS,
    note: 'A list on or off; a list the document does not define travels as a `new-list` change.',
    translate: translator('setListNumbering', (op, batch) => () => {
      // A list the document does not define yet: its definition travels with it (group 9.10)
      if (isNewList(batch.doc, op.numId)) return newListChanges(op, batch);
      const p = batch.paragraph(op.paragraphId);
      const level = op.level ?? currentListLevel(batch.doc, p.story, p.mark) ?? 0;
      return [
        {
          kind: 'paragraph-property',
          part: p.story.name,
          marks: [addressOf(p.mark)],
          properties: [
            {
              localName: 'numPr',
              children: [
                { localName: 'ilvl', attributes: { val: String(level) } },
                { localName: 'numId', attributes: { val: op.numId ?? '0' } },
              ],
            },
          ],
        },
      ];
    }),
  },
  setParagraphTabStops: {
    status: 'translated',
    group: '6.1',
    changes: ['paragraph-property'],
    party: PARTY_SUGGESTS,
    note: 'Tab stops; a stop in force that the press leaves out is cleared.',
    translate: translator('setParagraphTabStops', (op, batch) => () => {
      const p = batch.paragraph(op.paragraphId);
      const children = [
        ...op.stops.map((stop) => ({
          localName: 'tab',
          attributes: {
            val: stop.alignment,
            pos: String(stop.positionTwips),
            ...(stop.leader === undefined ? {} : { leader: stop.leader }),
          },
        })),
        // A stop in force that the press leaves out is cleared, as the engine's op does
        ...(op.inForcePositionsTwips ?? []).map((position) => ({
          localName: 'tab',
          attributes: { val: 'clear', pos: String(position) },
        })),
      ];
      return [
        {
          kind: 'paragraph-property',
          part: p.story.name,
          marks: [addressOf(p.mark)],
          properties: [{ localName: 'tabs', children }],
        },
      ];
    }),
  },
};

/** Options of the paragraph ops that change their meaning. */
export const PARAGRAPH_OPTIONS: readonly OrderedOpOption[] = [
  {
    op: 'setListNumbering',
    option: 'new-list-definition',
    status: 'translated',
    group: '9.10',
    note: 'A list number the document does not define: a `new-list` change with the definition the editor wrote (D14).',
  },
  {
    op: 'setListLevel',
    option: 'undeclared-level',
    status: 'translated',
    group: '9.10',
    note: "A level the list's definition does not declare: a `list-level` change that declares it in Word's default format (D14).",
  },
];
