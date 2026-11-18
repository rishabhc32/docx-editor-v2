// What the list kinds of a paste share: one paragraph op on each of several paragraphs.

import { runOrderedEdit, type OrderedEditOutcome } from '../../../edits/apply.ts';
import type { Ctx, OrderedEditContext, OrderedReplacedFormat } from '../../../edits/context.ts';
import { settle } from '../../../edits/direct.ts';
import type { OrderedParagraphOp } from '../../../model/types.ts';
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
    settle(ctx.doc, sid);
  }

  return { suggestions: [...suggestions], replaced };
}
