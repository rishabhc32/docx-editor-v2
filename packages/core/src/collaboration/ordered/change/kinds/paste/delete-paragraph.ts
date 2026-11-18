// The `delete-paragraph` change kind (group 9.10): a whole paragraph removed with tracking off.

import { orderedViewParagraph } from '../../../compose/compose.ts';
import type { Ctx } from '../../../edits/context.ts';
import { directDelete, directJoin, settle } from '../../../edits/direct.ts';
import { strikeMark } from '../../../edits/marks.ts';
import { shownItemAt } from '../../../edits/places.ts';
import { OrderedRefusal } from '../../../model/refusal.ts';
import type { OrderedStory } from '../../../model/types.ts';
import { refuseLockedBlocks, refuseLockedWithin, siblings } from '../../../placement.ts';
import { fieldPath, required } from '../../read.ts';
import type { OrderedChangeAddress, OrderedEditBase } from '../../types.ts';
import { address } from '../../values.ts';
import { runKindEdit } from '../run.ts';
import { orderedKindSpec } from '../spec.ts';

/**
 * Remove a whole paragraph with tracking off: its text, objects and markers, and its mark, as
 * the engine's `deleteBlock` removes a `w:p`. The paragraph is named by its mark. A direct edit
 * only: Word tracks it as struck text and a struck mark, which `delete` says.
 *
 * @public
 */
export interface OrderedDeleteParagraphChange extends OrderedEditBase {
  readonly kind: 'delete-paragraph';
  readonly mark: OrderedChangeAddress;
}

/** Remove the shown paragraph a mark ends: its text, then its mark by a join or alone. */
function deleteParagraph(ctx: Ctx, story: OrderedStory, address: OrderedChangeAddress): void {
  const mark = shownItemAt(ctx.doc, story, address, 'mark');
  const index = story.tree.before(mark).marks;
  const p = orderedViewParagraph(ctx.doc, story.name, index);
  if (!p || p.mark !== mark) {
    throw new OrderedRefusal('hidden-target', mark.space.slice(0, 160));
  }

  if (story.tree.summary.marks < 2) {
    throw new OrderedRefusal('no-paragraph-after', 'the story keeps one paragraph');
  }

  // As the engine's `deleteBlock`: a lock around the paragraph, or any lock on a control in it
  refuseLockedBlocks(ctx, p.blocks);
  refuseLockedWithin(ctx, p.items, true);

  // The text first: every view offset of the paragraphs after it stays where it is
  if (p.chars.length > 0) {
    directDelete(ctx, p, 0, p.chars.length);
  }

  const after = orderedViewParagraph(ctx.doc, story.name, index);
  const next = orderedViewParagraph(ctx.doc, story.name, index + 1);
  if (!after) {
    throw new OrderedRefusal('malformed-story', 'the paragraph went');
  }

  // Beside a paragraph, a join; before a table or at a container's end, the mark alone goes
  if (next && siblings(ctx, after, next)) {
    directJoin(ctx, after, next);
  } else {
    settle(ctx.doc, strikeMark(ctx, after.mark, null, { evenStruck: true }));
  }
}

export const DELETE_PARAGRAPH_SPEC = orderedKindSpec<OrderedDeleteParagraphChange>({
  kind: 'delete-paragraph',
  fields: ['mark'],
  read: (fields) => ({ mark: address(required(fields, 'mark'), fieldPath(fields, 'mark')) }),
  touches(change, reader) {
    reader.item(change.mark, 'mark');
  },
  apply(change, input) {
    if (input.edit.direct !== true) {
      throw new OrderedRefusal('invalid-edit', 'a tracked paragraph deletion is struck text');
    }

    // The settled deletion and join make their own suggestions, which the run reports
    return runKindEdit(input, (ctx) => {
      deleteParagraph(ctx, input.story, change.mark);
      return null;
    });
  },
  policy: { text: true, properties: false, strikes: true },
  // A paragraph before a table or at a cell's end moves a block edge
  publish: 'story',
});
