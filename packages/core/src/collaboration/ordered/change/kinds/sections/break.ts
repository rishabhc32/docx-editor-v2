// The `section-break` change kind (task 9.8): a section ends at a paragraph mark, as the
// engine's `setSectionMark` writes it. The editor's break is Enter and this change in one
// batch: the mark is Enter's new mark. Suggested, the break rides on its author's new mark
// (Word marks the break's paragraph mark inserted, `w:pPr/w:rPr/w:ins`, ECMA-376 Part 1,
// §17.13.5.20, and a reject joins the paragraphs, which takes the break away). A type for the
// section that starts after it goes on that section untracked, as Word writes it (no
// `w:sectPrChange`; Word for Mac, 2026-10-29): a reject leaves it, and a take-back removes it.

import {
  insertSectionBreak,
  othersCopiedRecord,
  retypesNextSection,
} from '../../../edits/sections.ts';
import { address } from '../../values.ts';
import { at, oneOf, required, type Fields } from '../../read.ts';
import type { OrderedChangeAddress, OrderedEditBase } from '../../types.ts';
import type { OrderedDocument } from '../../../model/types.ts';
import { runKindEdit } from '../run.ts';
import { orderedKindSpec } from '../spec.ts';

/**
 * End a section at a paragraph mark. `type` is how the section that starts after it starts:
 * written on that section when it starts another way.
 *
 * @public
 */
export interface OrderedSectionBreakChange extends OrderedEditBase {
  readonly kind: 'section-break';
  readonly mark: OrderedChangeAddress;
  readonly type?: 'nextPage' | 'continuous';
}

function read(fields: Fields): Record<string, unknown> {
  const path = fields.path;
  const out: Record<string, unknown> = {
    mark: address(required(fields, 'mark'), at(path, 'mark')),
  };
  if (fields.has('type'))
    out.type = oneOf(fields.get('type'), at(path, 'type'), ['nextPage', 'continuous'] as const);
  return out;
}

/**
 * A suggesting person's break that gives the section after it another type, refused as the
 * room refuses it (`policy-direct`, `package`), or null. The room's party rule and the client's
 * gate (`suggests`) share it, so the editor refuses before it sends, with the room's name.
 */
export function retypingBreakRefusal(
  change: OrderedSectionBreakChange,
  doc: OrderedDocument
): { readonly code: 'policy-direct'; readonly reason: 'package'; readonly detail: string } | null {
  const story = doc.stories.get(change.part);
  if (!story || !retypesNextSection(doc, story, change.mark, change.type)) return null;
  return {
    code: 'policy-direct',
    reason: 'package',
    detail: 'the type of the section after the break is written untracked',
  };
}

export const SECTION_BREAK_SPEC = orderedKindSpec<OrderedSectionBreakChange>({
  kind: 'section-break',
  fields: ['mark', 'type'],
  read,
  touches: (change, reader) => void reader.item(change.mark, 'mark'),
  apply: (change, input) =>
    runKindEdit(input, (ctx) => insertSectionBreak(ctx, change.mark, change.type), {
      settleDirect: true,
    }),
  policy: {
    text: false,
    properties: false,
    // The host's check rejects a party's own suggestions and compares: a type for the section
    // after the break stays, so a break that retypes it changes the document as a direct edit
    // does (`package`). A break that keeps the type is a suggestion, unless the section it cuts
    // holds another person's pending record: the new section copies it inside the party's own
    // suggestion, and the check's take-back would take that record with it (`nested`, as Enter
    // in another person's suggestion, design D9).
    party(change, doc, maker) {
      if (change.kind !== 'section-break') {
        return null;
      }

      const retyped = retypingBreakRefusal(change, doc);
      if (retyped) {
        return { ok: false, ...retyped };
      }

      const story = doc.stories.get(change.part);
      if (!story) {
        return null;
      }

      const other = othersCopiedRecord(doc, story, change.mark, maker);
      if (other === null) {
        return null;
      }

      return {
        ok: false,
        code: 'policy-nested',
        reason: 'nested',
        detail: `the new section would copy ${other}'s pending section record`,
      };
    },
  },
  publish: 'story',
});
