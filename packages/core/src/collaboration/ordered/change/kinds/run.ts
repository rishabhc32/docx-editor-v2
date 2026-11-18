// What the change kinds' edits share: a run of one model edit, and a range of one paragraph.

import { editOutcome, type OrderedEditOutcome } from '../../edits/apply.ts';
import type { Ctx } from '../../edits/context.ts';
import { settle } from '../../edits/direct.ts';
import type { OrderedLinkRange } from '../../edits/links.ts';
import { OrderedRefusal } from '../../model/refusal.ts';
import { resolveChangePlace } from '../change-place.ts';
import type { OrderedChangePlace } from '../types.ts';
import type { OrderedKindApply } from './spec.ts';

/** How a kind's edit runs: `settleDirect` settles a direct edit's suggestion at once. */
interface KindEditOptions {
  readonly settleDirect?: boolean;
}

/**
 * Run a kind's edit as the model's edits run (`runOrderedEdit`): its context, then the
 * suggestion it made and every suggestion it added, and the records it took over.
 */
export function runKindEdit(
  input: OrderedKindApply,
  edit: (ctx: Ctx) => string | null,
  options: KindEditOptions = {}
): OrderedEditOutcome {
  return editOutcome(input.edit, (ctx) => {
    const made = edit(ctx);
    if (options.settleDirect === true && ctx.direct) {
      settle(ctx.doc, made);
    }

    return [made];
  });
}

/** A range of one paragraph, by view offsets. */
export function paragraphRange(
  ctx: Ctx,
  from: OrderedChangePlace,
  to: OrderedChangePlace
): OrderedLinkRange {
  const a = resolveChangePlace(ctx.doc, ctx.story.name, from);
  const b = resolveChangePlace(ctx.doc, ctx.story.name, to);
  if (a.index !== b.index) throw new OrderedRefusal('invalid-edit', 'a range of one paragraph');
  if (b.offset <= a.offset) throw new OrderedRefusal('empty-range');
  return { p: a.paragraph, start: a.offset, end: b.offset };
}
