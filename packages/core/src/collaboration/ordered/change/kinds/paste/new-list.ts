// The `new-list` change kind (group 9.10): a list the document does not define yet.

import { contextOf, suggestionFor } from '../../../edits/context.ts';
import { fragmentResourceNodes } from '../../../edits/fragment-nodes.ts';
import { addListDefinition } from '../../../edits/resources.ts';
import { OrderedRefusal } from '../../../model/refusal.ts';
import type { OrderedFragmentResolver } from '../../../edits/fragment-nodes.ts';
import {
  fragmentBudget,
  fragmentElement,
  fragmentElements,
  type OrderedFragmentList,
} from '../../fragment.ts';
import { at, fieldPath, fieldsOf, int, list, required } from '../../read.ts';
import {
  ORDERED_CHANGE_LIMITS,
  type OrderedChangeAddress,
  type OrderedEditBase,
} from '../../types.ts';
import { address } from '../../values.ts';
import { orderedKindSpec } from '../spec.ts';
import { pressParagraphs, type ParagraphPress } from './paragraph-presses.ts';

/** One paragraph a new list numbers: its mark and its level (0 to 8). @public */
export interface OrderedNewListParagraph {
  readonly mark: OrderedChangeAddress;
  readonly level: number;
}

/**
 * Set a new list on paragraphs: its definition (as a paste carries one) and the paragraphs it
 * numbers. The definition is a resource of the change's suggestion (design D14), numbered from
 * the change id; Word records no change in the numbering part, and a tracked paragraph keeps
 * its earlier properties in its `w:pPrChange` (ECMA-376 Part 1, §17.13.5.29).
 *
 * @public
 */
export interface OrderedNewListChange extends OrderedEditBase {
  readonly kind: 'new-list';
  readonly list: OrderedFragmentList;
  readonly paragraphs: readonly OrderedNewListParagraph[];
}

const NO_REFERENCES = { lists: 0, links: 0, notes: 0, media: 0 };

/** A definition names no list, link, picture or note of the change. */
const RESOLVE: OrderedFragmentResolver = {
  numId: (value) => {
    if (value.startsWith('#')) throw new OrderedRefusal('invalid-edit', 'a list names itself');
    return value;
  },
  relationship: () => {
    throw new OrderedRefusal('invalid-edit', 'a list definition holds no link');
  },
  note: () => {
    throw new OrderedRefusal('invalid-edit', 'a list definition holds no note');
  },
  media: () => {
    throw new OrderedRefusal('invalid-edit', 'a list definition holds no picture');
  },
  part: () => {
    throw new OrderedRefusal('invalid-edit', 'a list definition holds no chart');
  },
  object: () => {
    throw new OrderedRefusal('invalid-edit', 'a list definition holds no object');
  },
  drawingId: () => {
    throw new OrderedRefusal('invalid-edit', 'a list definition holds no drawing');
  },
};

export const NEW_LIST_SPEC = orderedKindSpec<OrderedNewListChange>({
  kind: 'new-list',
  fields: ['list', 'paragraphs'],
  read(fields) {
    const budget = fragmentBudget();

    // Every path starts at the change, so a refusal names its place in the message
    const own = fieldsOf(
      required(fields, 'list'),
      fieldPath(fields, 'list'),
      new Set(['abstract', 'overrides'])
    );
    const abstract = fragmentElement(
      required(own, 'abstract'),
      at(fieldPath(fields, 'list'), 'abstract'),
      'abstractNum',
      NO_REFERENCES,
      budget
    );
    const overrides = own.has('overrides')
      ? fragmentElements(
          own.get('overrides'),
          at(fieldPath(fields, 'list'), 'overrides'),
          'lvlOverride',
          NO_REFERENCES,
          budget
        )
      : [];
    const paragraphs = list(
      required(fields, 'paragraphs'),
      fieldPath(fields, 'paragraphs'),
      1,
      ORDERED_CHANGE_LIMITS.targets
    ).map((entry, k): OrderedNewListParagraph => {
      const path = at(fieldPath(fields, 'paragraphs'), k);
      const one = fieldsOf(entry, path, new Set(['mark', 'level']));
      return {
        mark: address(required(one, 'mark'), at(path, 'mark')),
        level: int(required(one, 'level'), at(path, 'level'), 0, 8),
      };
    });
    return { list: overrides.length > 0 ? { abstract, overrides } : { abstract }, paragraphs };
  },
  touches(change, reader) {
    for (const paragraph of change.paragraphs) reader.item(paragraph.mark, 'mark');
  },
  apply(change, { edit }) {
    const ctx = contextOf(edit, []);
    const sid = suggestionFor(ctx).id;
    const [abstract] = fragmentResourceNodes([change.list.abstract], 'abstractNum', RESOLVE);
    if (!abstract) throw new OrderedRefusal('engine-refused', 'a list without its definition');
    const overrides = fragmentResourceNodes(change.list.overrides ?? [], 'lvlOverride', RESOLVE);
    const numId = addListDefinition(ctx, sid, { abstract, overrides }).numId;
    const presses = change.paragraphs.map(
      (paragraph): ParagraphPress => ({
        mark: paragraph.mark,
        op: { op: 'setListNumbering', numId, level: paragraph.level },
      })
    );
    return pressParagraphs(edit, ctx, sid, presses);
  },
  policy: { text: false, properties: true },
  publish: 'paragraphs',
});
