// A custom node's payload as a resource (design D14), in its namespace's custom XML data store.

import { customNodeBindingTarget } from '../../../store/package/custom-node-payloads.ts';
import type { OrderedControlBinding, OrderedCustomPayload } from '../change/kinds/controls.ts';
import { isHidden } from '../model/document.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import {
  addOrderedResource,
  orderedResourceShows,
  type OrderedCustomXmlResource,
} from '../model/resources.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import type { Ctx } from './context.ts';

/** The store a namespace has: the file's, or the one its first custom node named. */
function storeOf(
  doc: OrderedDocument,
  namespaceUri: string
): { readonly storeItemId: string; readonly root: string } | null {
  const file = doc.resources.fileStores.get(namespaceUri);
  if (file) return file;
  for (const resource of doc.resources.added.values())
    if (resource.kind === 'custom-xml' && resource.namespaceUri === namespaceUri) return resource;
  return null;
}

/** Whether a store id names a store of another namespace. */
function storeOfOther(doc: OrderedDocument, storeItemId: string, namespaceUri: string): boolean {
  for (const [namespace, file] of doc.resources.fileStores)
    if (namespace !== namespaceUri && file.storeItemId === storeItemId) return true;
  for (const resource of doc.resources.added.values())
    if (
      resource.kind === 'custom-xml' &&
      resource.namespaceUri !== namespaceUri &&
      resource.storeItemId === storeItemId
    )
      return true;
  return false;
}

/** Whether another shown control binds the node already. */
function boundElsewhere(doc: OrderedDocument, namespaceUri: string, nodeId: string): boolean {
  for (const resource of doc.resources.added.values())
    if (
      resource.kind === 'custom-xml' &&
      resource.namespaceUri === namespaceUri &&
      resource.nodeId === nodeId &&
      orderedResourceShows(doc, resource) &&
      !isHidden(doc, resource.control)
    )
      return true;
  return false;
}

/**
 * The payload of a new custom node, for the suggestion that made its control (`open`).
 *
 * @throws OrderedRefusal `name-taken` when the binding's store is not its namespace's, or its
 *   node is bound by another shown control (two people who each minted `cx1`): every copy
 *   refuses alike what would make the store ambiguous.
 */
export function addCustomNodePayload(
  ctx: Ctx,
  made: { readonly suggestion: string; readonly open: OrderedItem },
  binding: OrderedControlBinding,
  payload: OrderedCustomPayload
): OrderedCustomXmlResource {
  const target = customNodeBindingTarget(binding);
  if (!target) throw new OrderedRefusal('invalid-edit', 'the binding of a custom node');
  const storeItemId = binding.storeItemId.toUpperCase();
  const store = storeOf(ctx.doc, target.namespaceUri);
  const ownStore = store
    ? store.storeItemId === storeItemId && store.root === target.rootLocalName
    : !storeOfOther(ctx.doc, storeItemId, target.namespaceUri);
  if (!ownStore || boundElsewhere(ctx.doc, target.namespaceUri, target.nodeId))
    throw new OrderedRefusal('name-taken', `${target.namespaceUri} ${target.nodeId}`);
  const resource: OrderedCustomXmlResource = {
    kind: 'custom-xml',
    key: ctx.ids.resource(),
    suggestion: made.suggestion,
    namespaceUri: target.namespaceUri,
    root: target.rootLocalName,
    storeItemId,
    nodeId: target.nodeId,
    label: payload.label,
    data: payload.data,
    control: made.open,
  };
  addOrderedResource(ctx.doc, resource);
  return resource;
}
