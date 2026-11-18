// Custom nodes' payloads in their custom XML data stores (design D14).

import {
  withoutCustomXmlNode,
  withCustomXmlNode,
} from '../../../store/package/custom-xml-nodes.ts';
import {
  findCustomXmlDataPart,
  withNamedCustomXmlDataPart,
} from '../../../store/package/custom-xml-part.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import {
  orderedResourceShows,
  orderedResourceToken,
  type OrderedCustomXmlResource,
} from '../model/resources.ts';
import type { OrderedDocument } from '../model/types.ts';

/** How a written control's binding names its node: the store's id, upper case, and the node. */
export const customXmlBindingKey = (storeItemId: string, nodeId: string): string =>
  `${storeItemId.toUpperCase()} ${nodeId}`;

/** Whether the document has custom nodes for compose to write or take out. */
export function hasCustomXml(doc: OrderedDocument): boolean {
  for (const store of doc.resources.fileStores.values()) if (store.boundNodes.size > 0) return true;
  for (const resource of doc.resources.added.values())
    if (resource.kind === 'custom-xml') return true;
  return false;
}

/**
 * The payload written for each bound node, by namespace: one whose suggestion shows and that a
 * written control binds. The last one a change added for a node wins, as the engine's update
 * replaces the node in place.
 */
function writtenPayloads(
  doc: OrderedDocument,
  bound: ReadonlySet<string>
): Map<string, Map<string, OrderedCustomXmlResource>> {
  const byNamespace = new Map<string, Map<string, OrderedCustomXmlResource>>();
  for (const resource of doc.resources.added.values()) {
    if (resource.kind !== 'custom-xml' || !orderedResourceShows(doc, resource)) continue;
    if (!bound.has(customXmlBindingKey(resource.storeItemId, resource.nodeId))) continue;
    let nodes = byNamespace.get(resource.namespaceUri);
    if (!nodes) byNamespace.set(resource.namespaceUri, (nodes = new Map()));
    nodes.set(resource.nodeId, resource);
  }

  return byNamespace;
}

/**
 * The namespace's store in the package, made from its first payload when the file has none: under
 * names from the payload's key and the `ds:itemID` its bindings quote.
 */
function storeFor(
  pkg: OoxmlPackage,
  first: OrderedCustomXmlResource
): { readonly pkg: OoxmlPackage; readonly partName: string } {
  const main = pkg.mainDocumentPart;
  const found = findCustomXmlDataPart(pkg, main, first.namespaceUri);
  if (found) return { pkg, partName: found.partName };
  const token = orderedResourceToken(first.key);
  const made = withNamedCustomXmlDataPart(
    pkg,
    main,
    { item: `/customXml/ordered_${token}.xml`, props: `/customXml/ordered_${token}_props.xml` },
    first.storeItemId,
    first.namespaceUri,
    first.root
  );
  if (!made.part) throw new OrderedRefusal('engine-refused', `a store for ${first.namespaceUri}`);
  return { pkg: made.pkg, partName: made.part.partName };
}

/**
 * The package with every bound payload in its store, and the file's nodes that no written
 * control binds any more taken out, as the engine's removal takes a node with its control.
 * `bound` holds the written stories' bindings (`customXmlBindingKey`).
 */
export function withCustomXml(
  pkg: OoxmlPackage,
  doc: OrderedDocument,
  bound: ReadonlySet<string>
): OoxmlPackage {
  let next = pkg;
  for (const [namespaceUri, file] of doc.resources.fileStores) {
    const store = findCustomXmlDataPart(next, next.mainDocumentPart, namespaceUri);
    if (!store) continue;
    for (const nodeId of file.boundNodes)
      if (!bound.has(customXmlBindingKey(file.storeItemId, nodeId)))
        next = withoutCustomXmlNode(next, store.partName, nodeId);
  }

  for (const nodes of writtenPayloads(doc, bound).values()) {
    const [first] = nodes.values();
    if (!first) continue;
    const store = storeFor(next, first);
    next = store.pkg;
    for (const node of nodes.values())
      next = withCustomXmlNode(next, store.partName, {
        id: node.nodeId,
        label: node.label,
        data: node.data,
      });
  }

  return next;
}
