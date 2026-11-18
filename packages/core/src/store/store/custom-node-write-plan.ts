// A custom-node write as ops and package edits, planned before anything commits.

import { boundCustomXmlNodeIdOf, customNodeBinding } from '../package/custom-node-payloads.ts';
import {
  readCustomXmlNode,
  withCustomXmlNode,
  withoutCustomXmlNode,
} from '../package/custom-xml-nodes.ts';
import { customXmlDataParts, withCustomXmlDataPart } from '../package/custom-xml-part.ts';
import { contentControlsIn } from '../package/content-control-nodes.ts';
import { runWithoutJournalCapture } from '../package/canonical-primitive-capture.ts';
import type { OoxmlPackage } from '../package/ooxml-package.ts';
import type { OoxmlPart } from '../package/ooxml-tree.ts';
import { PackageEditRefusal } from './package-edit-refusal.ts';
import type { TreeDocOp } from './tree-op-types.ts';
import type {
  CustomNodePayloadWrite,
  CustomNodeWriteRejection,
  InsertCustomNodeWrite,
} from './custom-node-writes.ts';

/**
 * The largest payload one node may carry, in UTF-16 code units.
 *
 * Same figure as the read side's (`parseCustomNodeData`) and for the same reason: far past any
 * legitimate chip, far short of anything that hurts. Checked on the WRITE too, so a document
 * cannot be authored here holding a payload the reader will later refuse to parse.
 */
export const MAX_CUSTOM_NODE_PAYLOAD_LENGTH = 256 * 1024;

/** The longest label a binding may paint. Word renders it as the control's whole content. */
export const MAX_CUSTOM_NODE_LABEL_LENGTH = 4_096;

/** A package edit of a custom-node write, applied before its ops in the same transaction. */
export type CustomNodePackageEdit = (pkg: OoxmlPackage) => OoxmlPackage;

/**
 * What a custom-node write commits: its package edits, then its ops, in one transaction. The
 * store write (`custom-node-writes.ts`) and the editor's gated write both commit it: the package
 * edits author the customXml store and its node, and the ops place the bound control. A
 * collaboration gate sees the ops with the package they apply to, so it can carry the node with
 * the control (`CollaborationGateContext.package`).
 */
export interface CustomNodeWritePlan {
  readonly ops: readonly TreeDocOp[];
  readonly packageEdits: readonly CustomNodePackageEdit[];
}

/** A plan, or why the write is refused before anything commits. */
export type CustomNodeWritePlanResult =
  | { readonly ok: true; readonly plan: CustomNodeWritePlan }
  | { readonly ok: false; readonly reason: CustomNodeWriteRejection; readonly detail?: string };

type Refused = Extract<CustomNodeWritePlanResult, { ok: false }>;

/** The data binding a control quotes for its payload's node. */
type PayloadBinding = NonNullable<ReturnType<typeof customNodeBinding>>;

/** The store authored with the payload's node in it, and the binding that names the node. */
type AuthoredPayload =
  | { readonly ok: true; readonly pkg: OoxmlPackage; readonly binding: PayloadBinding }
  | Refused;

/**
 * Author the store for `payload` (created if the document has none for the namespace) and its
 * node, read back. Every way this can fail is a refusal with its own reason.
 */
function authorPayload(
  pkg: OoxmlPackage,
  dataOwnerPartName: string,
  payload: CustomNodePayloadWrite
): AuthoredPayload {
  const authored = withCustomXmlDataPart(
    pkg,
    dataOwnerPartName,
    payload.namespaceUri,
    payload.rootLocalName
  );
  if (!authored.part) {
    return {
      ok: false,
      reason: 'store-not-authored',
      detail:
        `no customXml store could be created for ${payload.namespaceUri}. The namespace must ` +
        `be free of control characters and quotes, the root name must be an XML name, and a ` +
        `store this document already carries for that namespace must use the same root name ` +
        `(${payload.rootLocalName})`,
    };
  }

  // Built from the store's OWN root, never from what the caller asked for: the two agree by the
  // check above, and deriving from the part means an xpath can never name an element the file
  // does not have.
  const binding = customNodeBinding(authored.part, payload.rootLocalName, payload.nodeId);
  if (!binding) {
    return {
      ok: false,
      reason: 'unaddressable-payload',
      detail:
        `the node id ${JSON.stringify(payload.nodeId)} or the namespace ` +
        `${JSON.stringify(payload.namespaceUri)} cannot be written into an XPath. Ids and root ` +
        `names must match [A-Za-z_][\\w.-]{0,127}; a namespace may not contain a quote, an ` +
        `angle bracket or an ampersand`,
    };
  }

  const written = withCustomXmlNode(authored.pkg, authored.part.partName, {
    id: payload.nodeId,
    label: payload.label,
    data: payload.data,
  });

  // READ IT BACK. `withCustomXmlNode` answers the package unchanged on three paths, so without
  // this the body gets a control bound to a node that was never written, and the caller is
  // told `ok`.
  if (!readCustomXmlNode(written, authored.part.partName, payload.nodeId)) {
    return {
      ok: false,
      reason: 'store-not-authored',
      detail: 'the payload node could not be written into the store',
    };
  }

  return { ok: true, pkg: written, binding };
}

/** The payload's size refusal, or null. */
function payloadSizeRefusal(payload: CustomNodePayloadWrite | undefined): Refused | null {
  if (!payload) return null;
  if (payload.data.length > MAX_CUSTOM_NODE_PAYLOAD_LENGTH) {
    return { ok: false, reason: 'payload-too-large', detail: 'data' };
  }

  if (payload.label.length > MAX_CUSTOM_NODE_LABEL_LENGTH) {
    return { ok: false, reason: 'payload-too-large', detail: 'label' };
  }

  return null;
}

/**
 * Which store node, in which part, a control binds: across every store the story and the main
 * document part carry.
 *
 * Every store, not the first: a control's binding names one `w:storeItemID`, but a document may
 * hold several stores and only the one whose id matches answers. Resolving against all of them
 * means a removal cannot miss a payload because the store it lived in was not the store this
 * definition usually writes to. A header chip's store hangs off the main part (see
 * `planCustomNodeInsert`), and an older file may hang it off the header itself.
 */
function boundNodesOf(
  pkg: OoxmlPackage,
  story: OoxmlPart,
  controlNodeId: string
): readonly { readonly partName: string; readonly nodeId: string }[] {
  const control = contentControlsIn(story.root).find((entry) => entry.node.id === controlNodeId);
  if (!control) return [];
  const bound = new Map<string, string>();
  for (const owner of new Set([story.name, pkg.mainDocumentPart])) {
    for (const dataPart of customXmlDataParts(pkg, owner)) {
      const nodeId = boundCustomXmlNodeIdOf(control.node, dataPart.itemId);
      if (nodeId !== null) bound.set(dataPart.partName, nodeId);
    }
  }

  return [...bound].map(([partName, nodeId]) => ({ partName, nodeId }));
}

/**
 * The package edit that authors `payload`'s store and node, and the binding the control quotes,
 * planned on the package as `packageEdits` leave it. The edit authors the store again on the
 * package it meets and refuses the transaction when the store's `ds:itemID` is no longer the one
 * the binding quotes, so a control is never bound to a store that was not written.
 */
function payloadEdit(
  pkg: OoxmlPackage,
  packageEdits: readonly CustomNodePackageEdit[],
  dataOwnerPartName: string,
  payload: CustomNodePayloadWrite
):
  | { readonly ok: true; readonly edit: CustomNodePackageEdit; readonly binding: PayloadBinding }
  | Refused {
  // A trial on a copy, so it records nothing a journal would replay
  const planned = runWithoutJournalCapture(() =>
    authorPayload(
      packageEdits.reduce((edited, edit) => edit(edited), pkg),
      dataOwnerPartName,
      payload
    )
  );
  if (!planned.ok) {
    return planned;
  }

  const { binding } = planned;
  const edit = (current: OoxmlPackage): OoxmlPackage => {
    const authored = authorPayload(current, dataOwnerPartName, payload);
    if (!authored.ok) throw new PackageEditRefusal(authored.detail ?? authored.reason);
    if (authored.binding.storeItemId !== binding.storeItemId) {
      throw new PackageEditRefusal('the payload store changed');
    }

    return authored.pkg;
  };
  return { ok: true, edit, binding };
}

/** The edits that take the nodes a control binds out of their stores. */
function withoutBoundNodes(
  pkg: OoxmlPackage,
  story: OoxmlPart,
  controlNodeId: string
): CustomNodePackageEdit[] {
  // Resolved now, because once the control is gone nothing says what it pointed at.
  return boundNodesOf(pkg, story, controlNodeId).map(
    (entry) => (current: OoxmlPackage) =>
      withoutCustomXmlNode(current, entry.partName, entry.nodeId)
  );
}

/**
 * Plan one custom node: the store and its node, then the control that binds it.
 *
 * `dataOwnerPartName` is the part the customXml store hangs off, defaulting to the story being
 * written. A chip in a header is written against the HEADER story, but Word enumerates the data
 * store from the main document part: a store authored off a header is one Word never sees.
 */
export function planCustomNodeInsert(
  pkg: OoxmlPackage,
  story: OoxmlPart,
  write: InsertCustomNodeWrite,
  dataOwnerPartName: string = story.name
): CustomNodeWritePlanResult {
  const tooLarge = payloadSizeRefusal(write.payload);
  if (tooLarge) return tooLarge;

  // A reversed or empty span is not a wrap. Silently inserting instead would put the node beside
  // the text the caller believed it was replacing.
  const replaced = write.replaceUntil;
  if (replaced !== undefined && replaced <= write.offset) {
    return { ok: false, reason: 'invalid-range', detail: 'replaceUntil must be past offset' };
  }

  const ops: TreeDocOp[] = [];
  const packageEdits: CustomNodePackageEdit[] = [];

  // The control being rewritten goes first, at the offsets the caller measured against the tree
  // it could see. Its payload goes with it: a rewrite that kept the old node would leave the
  // store growing an entry per edit.
  if (write.replaceControlId !== undefined) {
    ops.push({ op: 'removeContentControl', controlId: write.replaceControlId, keepContent: false });
    packageEdits.push(...withoutBoundNodes(pkg, story, write.replaceControlId));
  }

  // The replaced text goes before the node, so the offsets the caller supplied still describe
  // the paragraph when the deletion is planned against it.
  if (replaced !== undefined) {
    ops.push({
      op: 'deleteText',
      paragraphId: write.paragraphId,
      start: write.offset,
      end: replaced,
    });
  }

  const control: Extract<TreeDocOp, { op: 'insertInlineContentControl' }> = {
    op: 'insertInlineContentControl',
    paragraphId: write.paragraphId,
    offset: write.offset,
    tag: write.tag,
    text: write.text,
    ...(write.alias === undefined ? {} : { alias: write.alias }),
    ...(write.lock === undefined ? {} : { lock: write.lock }),
  };
  const payload = write.payload;
  if (!payload) return { ok: true, plan: { ops: [...ops, control], packageEdits } };

  // The store first: a transaction that dies at the body edit takes the unbound store with it
  const planned = payloadEdit(pkg, packageEdits, dataOwnerPartName, payload);
  if (!planned.ok) return planned;
  packageEdits.push(planned.edit);
  return {
    ok: true,
    plan: { ops: [...ops, { ...control, dataBinding: planned.binding }], packageEdits },
  };
}

/**
 * Plan the removal of a custom node and, in the same transaction, the payload it bound.
 *
 * The sweep would collect the node eventually (that is what makes deletion in Word survivable),
 * but "eventually" is the next open, and a document saved in between carries a payload for a
 * chip that is gone. Doing it here means the ordinary case is exact and the sweep is a backstop.
 */
export function planCustomNodeRemoval(
  pkg: OoxmlPackage,
  story: OoxmlPart,
  controlNodeId: string
): CustomNodeWritePlan {
  return {
    ops: [{ op: 'removeContentControl', controlId: controlNodeId, keepContent: false }],
    packageEdits: withoutBoundNodes(pkg, story, controlNodeId),
  };
}
