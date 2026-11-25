// The `replace-story` change kind (group 9.10): automation's `replaceStoryBlocks`.

import type { OoxmlAttribute } from '../../../../../store/package/ooxml-tree.ts';
import { build } from '../../../../../store/store/tree-op-tracked.ts';
import { contextOf, suggestionFor } from '../../../edits/context.ts';
import { settle } from '../../../edits/direct.ts';
import { newParagraphIdentity } from '../../../edits/para-ids.ts';
import { createOrderedItem } from '../../../model/document.ts';
import { insertItems, relabel } from '../../../model/mutations.ts';
import { OrderedRefusal } from '../../../model/refusal.ts';
import type { OrderedItem } from '../../../model/types.ts';
import { refuseLockedWithin } from '../../../placement.ts';
import { at, fieldPath, Invalid, list, plainText, required } from '../../read.ts';
import { ORDERED_CHANGE_LIMITS, type OrderedEditBase } from '../../types.ts';
import { orderedKindSpec } from '../spec.ts';

/** Most paragraphs one replacement writes (the engine's own cap). */
const MAX_PARAGRAPHS = 10_000;

/**
 * Put plain paragraphs, one for each text, in place of a whole story. Every item is struck with
 * one settled deletion and the new paragraphs follow in the same suggestion, so the organiser's
 * undo (`revert`) shows the old story again. A direct edit only, as the engine refuses it in
 * suggesting mode.
 *
 * @public
 */
export interface OrderedReplaceStoryChange extends OrderedEditBase {
  readonly kind: 'replace-story';
  readonly paragraphs: readonly string[];
}

export const REPLACE_STORY_SPEC = orderedKindSpec<OrderedReplaceStoryChange>({
  kind: 'replace-story',
  fields: ['paragraphs'],
  read(fields) {
    let total = 0;
    const path = fieldPath(fields, 'paragraphs');
    const paragraphs = list(required(fields, 'paragraphs'), path, 1, MAX_PARAGRAPHS).map(
      (text, k) => {
        const checked = plainText(text, at(path, k), 0, ORDERED_CHANGE_LIMITS.text);
        total += checked.length;
        return checked;
      }
    );
    if (total > ORDERED_CHANGE_LIMITS.text * 4)
      throw new Invalid('too-long', path, `over ${ORDERED_CHANGE_LIMITS.text * 4}`);
    return { paragraphs };
  },
  // The whole story: a direct edit only, so no party rule reads its places
  touches: () => {},
  apply(change, { story, edit }) {
    if (edit.direct !== true)
      throw new OrderedRefusal('invalid-edit', 'a story replacement has no tracked form');
    const ctx = contextOf(edit, []);
    const sid = suggestionFor(ctx).id;
    const items = [...story.tree.items()];

    // Every control goes: a lock on any of them keeps its content or its place
    refuseLockedWithin(ctx, items, true);

    // The body's own section properties close the story and stay
    const tail = items.at(-1);
    const keep = tail && tail.kind === 'zero' && tail.block === true ? tail : null;
    for (const item of items)
      if (item !== keep && !item.dels.includes(sid))
        relabel(ctx.doc, story, item, { dels: [...item.dels, sid] });
    const fresh: OrderedItem[] = [];
    for (const text of change.paragraphs) {
      if (text.length > 0) {
        const space = ctx.ids.units();
        fresh.push(
          createOrderedItem({ space, kind: 'text', text, length: text.length, run: space, by: sid })
        );
      }

      const paragraph = newParagraphIdentity(ctx);
      const attributes: OoxmlAttribute[] = paragraph.attributes;
      fresh.push(
        createOrderedItem({
          space: paragraph.space,
          kind: 'mark',
          node: build(paragraph.space, 'paragraph', 'p', attributes, []),
          by: sid,
        })
      );
    }

    const after = keep ? story.tree.prev(keep) : story.tree.last();
    insertItems(ctx.doc, story, after, fresh);
    settle(ctx, sid);
    return { suggestions: [sid], replaced: [] };
  },
  policy: { text: true, properties: false, strikes: true },
  publish: 'story',
});
