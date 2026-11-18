import type { TreeDocOp } from '../../../../../store/store/tree-op-types.ts';
import { insertedPending, isHidden } from '../../../model/document.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from '../../../model/types.ts';
import { addressOf, openOfNode } from '../../paragraphs.ts';
import type { OrderedChangeBody, OrderedTranslateBatch } from '../types.ts';

type DeleteBlockOp = Extract<TreeDocOp, { readonly op: 'deleteBlock' }>;

/** The shown table or row with this node id, in any story. */
function blockContainer(
  doc: OrderedDocument,
  nodeId: string
): { readonly story: OrderedStory; readonly open: OrderedItem } | null {
  for (const story of doc.stories.values()) {
    const open = openOfNode(doc, story, nodeId);
    if (!open || (open.token !== 'tbl' && open.token !== 'tr') || isHidden(doc, open)) continue;
    return { story, open };
  }

  return null;
}

/** The rows right inside a table that still show, base rows first and new rows last. */
function rowsOf(doc: OrderedDocument, story: OrderedStory, table: OrderedItem): OrderedItem[] {
  const rows: OrderedItem[] = [];
  let depth = 0;
  for (let item = story.tree.next(table); item; item = story.tree.next(item)) {
    if (item.kind === 'close' && item.pair === table.space) break;
    if (item.token === 'tbl' && item.kind === 'open') {
      depth += 1;
    }

    if (item.token === 'tbl' && item.kind === 'close') {
      depth -= 1;
    }

    if (depth === 0 && item.kind === 'open' && item.token === 'tr' && !isHidden(doc, item))
      rows.push(item);
  }

  // A new row taken back by its author never names the deletion's suggestion: it goes last, so
  // the first row's change makes the suggestion the others join
  const fresh = (row: OrderedItem) => (insertedPending(doc, row) ? 1 : 0);
  return [...rows].sort((a, b) => fresh(a) - fresh(b));
}

/**
 * The changes of one `deleteBlock`: rows of a table or one row, or one paragraph with its mark.
 * A table's deletion is every row's (the engine's tracked `deleteBlock` marks each row under one
 * id: one suggestion here), and compose drops a table whose rows are all gone, as the engine does.
 */
export function deleteBlockChanges(
  op: DeleteBlockOp,
  batch: OrderedTranslateBatch
): OrderedChangeBody[] {
  const block = blockContainer(batch.doc, op.blockId);
  if (block) {
    const part = block.story.name;
    const rows =
      block.open.token === 'tr' ? [block.open] : rowsOf(batch.doc, block.story, block.open);
    return rows.map((row) => ({ kind: 'delete-row', part, row: addressOf(row) }));
  }

  const p = batch.paragraph(op.blockId);
  return [{ kind: 'delete-paragraph', part: p.story.name, mark: addressOf(p.mark) }];
}
