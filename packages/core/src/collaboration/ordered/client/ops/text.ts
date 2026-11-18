// Text ops (task 6.1): typing, deletion, a tab, a break, and run formatting. Nothing in an op
// is dropped: a field the format has no place for is a refusal, except the placement hints
// `inside` and `bias`, whose meaning is the list's own placement rules (design D4).

import { ORDERED_CHANGE_LIMITS } from '../../change/types.ts';
import { addressOf, placeAt } from '../paragraphs.ts';
import { markFormatProperties, runFormatProperties } from '../properties.ts';
import { formFieldTyping } from './controls.ts';
import { translateResultRunFormat } from './fields.ts';
import {
  PARTY_SUGGESTS,
  translator,
  type OrderedChangeBody,
  type OrderedOpOption,
  type OrderedOpTable,
} from './types.ts';

/** UTF-16 chunks of at most the format's text cap, never between the halves of a pair. */
export function chunks(text: string): string[] {
  const cap = ORDERED_CHANGE_LIMITS.text;
  const out: string[] = [];
  for (let start = 0; start < text.length; ) {
    let end = Math.min(text.length, start + cap);
    const code = text.charCodeAt(end - 1);
    if (end < text.length && code >= 0xd800 && code <= 0xdbff) end -= 1;
    out.push(text.slice(start, end));
    start = end;
  }

  return out;
}

/** A tab or a break typed at a place. */
const breakAt = (name: 'insertTab' | 'insertHardBreak' | 'insertPageBreak') =>
  translator(name, (op, batch) => () => {
    const p = batch.paragraph(op.paragraphId);
    const place = placeAt(p, op.offset);
    if (op.op === 'insertTab') {
      return [{ kind: 'insert-tab', part: p.story.name, place }];
    }

    const kind = op.op === 'insertHardBreak' ? 'line' : 'page';
    return [{ kind: 'insert-break', part: p.story.name, place, break: kind }];
  });

export const TEXT_OPS: OrderedOpTable<
  | 'insertText'
  | 'deleteText'
  | 'insertTab'
  | 'insertHardBreak'
  | 'insertPageBreak'
  | 'setRunProperties'
  | 'setParagraphMarkProperties'
> = {
  insertText: {
    status: 'translated',
    group: '6.1',
    changes: ['insert'],
    party: PARTY_SUGGESTS,
    note: 'Typing; cut under the text cap, the pieces at one place.',
    translate: translator('insertText', (op, batch) => {
      if (op.textFormFieldId !== undefined) return formFieldTyping(op, batch, op.textFormFieldId);
      return () => {
        const p = batch.paragraph(op.paragraphId);
        const place = placeAt(p, op.offset);

        // Each chunk goes before the same place, so the chunks keep their order
        return chunks(op.text).map(
          (text): OrderedChangeBody => ({ kind: 'insert', part: p.story.name, place, text })
        );
      };
    }),
  },
  deleteText: {
    status: 'translated',
    group: '6.1',
    changes: ['delete'],
    party: PARTY_SUGGESTS,
    note: 'Backspace, Delete and a cut: one range.',
    translate: translator('deleteText', (op, batch) => {
      if (op.textFormFieldId !== undefined) return formFieldTyping(op, batch, op.textFormFieldId);
      return () => {
        const p = batch.paragraph(op.paragraphId);
        if (op.end <= op.start) return [];
        return [
          {
            kind: 'delete',
            part: p.story.name,
            from: placeAt(p, op.start),
            to: placeAt(p, op.end),
          },
        ];
      };
    }),
  },
  insertTab: {
    status: 'translated',
    group: '6.1',
    changes: ['insert-tab'],
    party: PARTY_SUGGESTS,
    note: 'The Tab key: a `w:tab` object.',
    translate: breakAt('insertTab'),
  },
  insertHardBreak: {
    status: 'translated',
    group: '6.1',
    changes: ['insert-break'],
    party: PARTY_SUGGESTS,
    note: 'Shift+Enter: a line break.',
    translate: breakAt('insertHardBreak'),
  },
  insertPageBreak: {
    status: 'translated',
    group: '6.1',
    changes: ['insert-break'],
    party: PARTY_SUGGESTS,
    note: 'Ctrl+Enter: a page break.',
    translate: breakAt('insertPageBreak'),
  },
  setRunProperties: {
    status: 'translated',
    group: '6.1',
    changes: ['format'],
    party: PARTY_SUGGESTS,
    note: 'Bold, italic, a font and the rest: the properties the op names (design D7).',
    translate: translator('setRunProperties', (op, batch) => {
      // A field result's run, named by its id (group 9.5); an empty list names no run, and the
      // engine formats the range as usual
      if (op.targetRunIds !== undefined && op.targetRunIds.length > 0)
        return () => translateResultRunFormat(op, batch);
      return () => {
        const p = batch.paragraph(op.paragraphId);
        if (op.end <= op.start) return [];
        const properties = runFormatProperties(
          batch.doc,
          p.story,
          p.view,
          op.start,
          op.end,
          op.properties
        );
        if (properties.length === 0) return [];
        const from = placeAt(p, op.start);
        const to = placeAt(p, op.end);
        return [{ kind: 'format', part: p.story.name, from, to, properties }];
      };
    }),
  },
  setParagraphMarkProperties: {
    status: 'translated',
    group: '6.1',
    changes: ['mark-format'],
    party: PARTY_SUGGESTS,
    note: "A format that covers a paragraph's end: the face of its mark.",
    translate: translator('setParagraphMarkProperties', (op, batch) => () => {
      const p = batch.paragraph(op.paragraphId);
      const properties = markFormatProperties(batch.doc, p.story, p.mark, op.properties);
      if (properties.length === 0) return [];
      return [{ kind: 'mark-format', part: p.story.name, marks: [addressOf(p.mark)], properties }];
    }),
  },
};

/** Options of the text ops that change their meaning. */
export const TEXT_OPTIONS: readonly OrderedOpOption[] = [
  {
    op: 'insertText',
    option: 'textFormFieldId',
    status: 'translated',
    group: '9.6',
    note: "Typing in a legacy text form field's result (forms protection): a `form-text` change.",
  },
  {
    op: 'deleteText',
    option: 'textFormFieldId',
    status: 'translated',
    group: '9.6',
    note: "A deletion in a legacy text form field's result: a `form-text` change.",
  },
  {
    op: 'setRunProperties',
    option: 'targetRunIds',
    status: 'translated',
    group: '9.5',
    note: "A format of a field result's run, named by its id: a format of the field's whole item with what the press changes there.",
  },
];
