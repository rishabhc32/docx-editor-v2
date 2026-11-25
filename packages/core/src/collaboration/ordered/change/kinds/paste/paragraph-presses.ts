// What the list kinds of a paste share: one paragraph op on each of several paragraphs.

import { paragraphOpKeeps } from '../../../compose/paragraph-ops.ts';
import { runOrderedEdit, type OrderedEditOutcome } from '../../../edits/apply.ts';
import {
  shownParagraphProperties,
  type Ctx,
  type OrderedEditContext,
  type OrderedReplacedFormat,
} from '../../../edits/context.ts';
import { settle } from '../../../edits/direct.ts';
import { resolveOrderedPlace } from '../../../edits/places.ts';
import type { OrderedParagraphOp } from '../../../model/types.ts';
import { refuseLockedBlocks } from '../../../placement.ts';
import type { OrderedChangeAddress } from '../../types.ts';

/** One paragraph's press: its mark, and the op it takes. */
export interface ParagraphPress {
  readonly mark: OrderedChangeAddress;
  readonly op: OrderedParagraphOp;
}

/**
 * Run each press as its own paragraph edit in the change's one suggestion `sid`, which a direct
 * edit settles at once. Returns every suggestion the presses made or joined.
 */
export function pressParagraphs(
  edit: OrderedEditContext,
  ctx: Ctx,
  sid: string,
  presses: readonly ParagraphPress[]
): OrderedEditOutcome {
  const suggestions = new Set([sid]);
  const replaced: OrderedReplacedFormat[] = [];
  for (const press of presses) {
    const outcome = runOrderedEdit(ctx.direct ? edit : { ...edit, join: sid }, {
      kind: 'paragraph',
      at: press.mark,
      ops: [press.op],
    });
    for (const id of outcome.suggestions) {
      suggestions.add(id);
    }

    replaced.push(...outcome.replaced);
  }

  if (ctx.direct) {
    settle(ctx, sid);
  }

  return { suggestions: [...suggestions], replaced };
}

/**
 * Whether a press leaves its paragraph as it shows (`paragraphOpKeeps`), so it writes no record.
 * A content lock around the paragraph refuses it first, as the press itself does.
 */
export function pressKeeps(ctx: Ctx, press: ParagraphPress): boolean {
  const { paragraph } = resolveOrderedPlace(ctx.doc, ctx.story.name, press.mark);
  refuseLockedBlocks(ctx, paragraph.blocks);
  let scratch = 0;
  const mint = () => `check#${(scratch += 1)}`;
  const shown = shownParagraphProperties(ctx, paragraph.mark);
  return paragraphOpKeeps(ctx.story.part, shown, press.op, mint, ctx.attribution);
}
