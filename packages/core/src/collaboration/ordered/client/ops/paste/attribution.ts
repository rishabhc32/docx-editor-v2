import type { TreeDocOp } from '../../../../../store/store/tree-op-types.ts';
import { ORDERED_CHANGE_LIMITS } from '../../../change/types.ts';
import { statusOf } from '../../../model/document.ts';
import { triplesOfSites, TripleIndex } from '../review.ts';
import { mainStoryOf } from '../paste-blocks.ts';
import type { OrderedChangeBody, OrderedTranslateBatch } from '../types.ts';

type AttributionOp = Extract<TreeDocOp, { readonly op: 'setRevisionAttribution' }>;

/**
 * One `attribution` change (another author, and optionally another date) for each thousand
 * suggestions the op's revision marks name. The marks are read as revision triples, as review
 * decisions read them (`review.ts`). Compose writes one attribution on every mark of a suggestion.
 */
export function attributionChanges(
  op: AttributionOp,
  batch: OrderedTranslateBatch
): OrderedChangeBody[] {
  const index = new TripleIndex(batch.doc);
  const ids = new Set<string>();
  const triples = triplesOfSites(
    batch,
    op.sites.map((site) => site.nodeId)
  );
  for (const triple of triples)
    for (const id of index.suggestions(triple))
      if (statusOf(batch.doc, id) === 'pending') ids.add(id);
  const all = [...ids];
  const out: OrderedChangeBody[] = [];
  const part = mainStoryOf(batch.doc).name;
  for (let k = 0; k < all.length; k += ORDERED_CHANGE_LIMITS.targets)
    out.push({
      kind: 'attribution',
      part,
      suggestions: all.slice(k, k + ORDERED_CHANGE_LIMITS.targets),
      author: op.revision.author,
      ...(op.revision.date === undefined ? {} : { date: op.revision.date }),
    });
  return out;
}
