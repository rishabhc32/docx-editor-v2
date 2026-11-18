// A picture's or a shape's element under its property records (group 9.4).

import { findNode, replaceNode } from '../../../store/package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { applyTreeOp } from '../../../store/store/tree-op-apply.ts';
import { setDocPrHyperlinkRelationship } from '../../../store/store/tree-op-drawings.ts';
import { cloneWithNewIds } from '../../../store/store/tree-op-nodes.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import type { DrawingTreeDocOp } from '../../../store/store/tree-op-types.ts';
import type { OrderedObjectOp, OrderedObjectRecord } from '../model/objects.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import type { ComposeContext } from './context.ts';
import { withStoryBlocks } from './paragraph-ops.ts';

/** The relationship id of a relationship resource a record names, or null when it has none. */
function relationshipIdOf(doc: OrderedDocument, key: string): string | null {
  const resource = doc.resources.added.get(key);
  return resource?.kind === 'relationship' ? resource.id : null;
}

/** The engine's op for one record, on the object with node id `id`. */
function drawingOpOf(
  doc: OrderedDocument,
  id: string,
  op: OrderedObjectOp
): DrawingTreeDocOp | null {
  switch (op.op) {
    case 'resize':
      return { op: 'resizeDrawing', drawingNodeId: id, extentEmu: { cx: op.cx, cy: op.cy } };
    case 'crop':
      return { op: 'cropDrawing', drawingNodeId: id, crop: op.crop };
    case 'position':
      return { op: 'positionDrawing', drawingNodeId: id, position: op.position };
    case 'wrap':
      return { op: 'setDrawingWrap', drawingNodeId: id, wrap: op.wrap };
    case 'metadata':
      return {
        op: 'setDrawingMetadata',
        drawingNodeId: id,
        title: op.title,
        description: op.description,
        // A new link is set on the result, as the editor's package write sets it
        ...(op.link === null ? { hyperlink: null } : {}),
      };
    case 'locks':
      return { op: 'setDrawingLocks', drawingNodeId: id, locks: op.locks };
    case 'transform':
      return { op: 'transformDrawing', drawingNodeId: id, action: op.action };
    case 'image': {
      const relationshipId = relationshipIdOf(doc, op.relationship);
      return relationshipId === null
        ? null
        : { op: 'replaceDrawingResource', drawingNodeId: id, relationshipId };
    }
  }
}

/** The first `w:drawing` at or under a node (an `mc:AlternateContent` wraps one). */
function drawingIn(node: OoxmlNode, depth = 0): OoxmlElement | null {
  if (node.kind === 'textValue' || depth > 8) return null;
  if (node.kind === 'drawing') return node;
  for (const child of node.children) {
    const found = drawingIn(child, depth + 1);
    if (found) return found;
  }

  return null;
}

/** The part with the object's `wp:docPr` link set to `relationshipId`. */
function withLink(part: OoxmlPart, id: string, relationshipId: string): OoxmlPart {
  const object = findNode(part, id);
  const drawing = object ? drawingIn(object) : null;
  const anchor = drawing?.children.find(
    (child) => child.kind === 'inlineDrawing' || child.kind === 'anchoredDrawing'
  );
  if (!drawing || !anchor || anchor.kind === 'textValue') return part;
  const updated = setDocPrHyperlinkRelationship(anchor, relationshipId);
  const next = {
    ...drawing,
    children: drawing.children.map((child) => (child.id === anchor.id ? updated : child)),
  } as OoxmlNode;
  const replaced = replaceNode(part, drawing.id, next, { deferValidation: true });
  return replaced.ok ? replaced.part : part;
}

/**
 * The object's element after running `ops` on it with the engine, on a scratch part of the
 * story that holds only the object, so the result is the element the editor writes for the same
 * ops. Each op the engine refuses is skipped (a lock another record set, a position on an object
 * made inline), and `refused` names why; the others still run.
 */
export function objectAfterOps(
  doc: OrderedDocument,
  part: OoxmlPart,
  node: OoxmlNode,
  ops: readonly OrderedObjectOp[],
  mint: () => string,
  refused?: (reason: string) => void
): OoxmlNode {
  const paragraph = build(mint(), 'paragraph', 'p', [], [build(mint(), 'run', 'r', [], [node])]);
  let scratch = withStoryBlocks(part, [paragraph]);
  for (const op of ops) {
    const drawingOp = drawingOpOf(doc, node.id, op);
    if (drawingOp === null) {
      refused?.('unknown-resource');
      continue;
    }

    const result = applyTreeOp(scratch, drawingOp, { deferValidation: true });
    if (!result.ok) {
      refused?.(result.reason);
      continue;
    }

    scratch = result.part;
    if (op.op === 'metadata' && typeof op.link === 'string') {
      const relationshipId = relationshipIdOf(doc, op.link);
      if (relationshipId !== null) scratch = withLink(scratch, node.id, relationshipId);
    }
  }

  return findNode(scratch, node.id) ?? node;
}

/** Composed objects by document and loaded element, by the accepted records' creation order. */
const COMPOSED = new WeakMap<OrderedDocument, WeakMap<OoxmlNode, Map<string, OoxmlNode>>>();

/**
 * The accepted object records of an item, in creation order. Records are only ever direct edits
 * (`model/objects.ts`), so none is pending.
 */
function acceptedRecords(doc: OrderedDocument, item: OrderedItem): OrderedObjectRecord[] {
  const out: OrderedObjectRecord[] = [];
  for (const record of item.props)
    if (record.kind === 'object' && doc.suggestions.get(record.suggestion)?.status === 'accepted')
      out.push(record);
  return out.sort((a, b) => a.seq - b.seq);
}

/**
 * An atom's element as compose writes it: its own, or, under accepted object records, the
 * engine's result of their ops. The object keeps its node id, so the editor's selection of it
 * holds, and the same records give the same element.
 */
export function objectNodeOf(ctx: ComposeContext, item: OrderedItem): OoxmlNode | null {
  const node = item.node;
  if (!node || item.props.length === 0) return node;
  const records = acceptedRecords(ctx.doc, item);
  if (records.length === 0) return node;
  let byNode = COMPOSED.get(ctx.doc);
  if (!byNode) {
    byNode = new WeakMap();
    COMPOSED.set(ctx.doc, byNode);
  }

  let bySeq = byNode.get(node);
  if (!bySeq) {
    bySeq = new Map();
    byNode.set(node, bySeq);
  }

  const key = records.map((record) => record.seq).join(',');
  const kept = bySeq.get(key);
  if (kept) return kept;
  const after = objectAfterOps(
    ctx.doc,
    ctx.story.part,
    node,
    records.map((record) => record.op),
    ctx.mint
  );
  const copy = after === node ? node : withObjectIds(after, node.id);
  bySeq.set(key, copy);
  return copy;
}

/**
 * The element under the object's own node id, its inner nodes under ids made from it: the
 * scratch part's minter does not know the story's ids, and the same element must take the same
 * ids in every compose (a scratch compose's ids included).
 */
function withObjectIds(node: OoxmlNode, id: string): OoxmlNode {
  let k = 0;
  const copy = cloneWithNewIds(node, () => `${id}#o${(k += 1)}`);
  return { ...copy, id } as OoxmlNode;
}
