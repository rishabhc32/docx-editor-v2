// Authoring a custom node that carries a payload, and taking one back.
//
// ONE TRANSACTION, three writes: the customXml data part (created if the document has none for
// the namespace), the node inside it, and the `w:sdt` in the body whose `w:dataBinding` names
// that node. Partial application is the failure this must not have — a control bound to a store
// that was never written is a document Word offers to repair, and repairing it means throwing
// the control away.
//
// The order is deliberate. The store goes in first, so the id the binding quotes is one the
// package already holds by the time the body references it; the body edit is last, so a refusal
// anywhere (a locked paragraph, an offset out of range, a protected document) abandons the store
// write with it rather than leaving a payload nothing points at.
//
// WHY THE STORE IS THE SOURCE OF TRUTH. A bound control with no type child is read-only in Word:
// the text is painted from the xpath and a user cannot type into it (verified,
// `sdt-custom-node-databinding-word-roundtrip.docx`). This engine refuses content edits inside a
// bound control for its own reasons — see `contentControlBindingRefusal` — so the two agree. The
// page and the payload cannot drift, and nothing here has to reconcile them.

import {
  boundCustomXmlNodeIdOf,
  boundCustomXmlNodeIdsInPackage,
} from '../package/custom-node-payloads.ts';
import {
  customXmlNodes,
  type CustomXmlNode,
  withoutOrphanCustomXmlNodes,
} from '../package/custom-xml-nodes.ts';
import { customXmlDataParts, findCustomXmlDataPart } from '../package/custom-xml-part.ts';
import { contentControlsIn } from '../package/content-control-nodes.ts';
import type { OoxmlPackage } from '../package/ooxml-package.ts';
import type { OoxmlPart } from '../package/ooxml-tree.ts';
import {
  MAX_CUSTOM_NODE_LABEL_LENGTH,
  MAX_CUSTOM_NODE_PAYLOAD_LENGTH,
  planCustomNodeInsert,
  planCustomNodeRemoval,
  type CustomNodeWritePlan,
} from './custom-node-write-plan.ts';
import type { TransactResult, TreeDocumentStore, TreeModelChange } from './tree-store.ts';
import type { TreeOpRejection } from './tree-op-types.ts';

/**
 * The payload half of an insert: which store, which node, and what it holds.
 *
 * `data` is opaque here. The lane that owns a schema is the one that declared it, and a store
 * that parsed payloads would be a second opinion about what a host's node means.
 */
export interface CustomNodePayloadWrite {
  /** Namespace of the store's root element — what identifies one store among several. */
  readonly namespaceUri: string;
  /** Local name of that root. An NCName; anything else refuses. */
  readonly rootLocalName: string;
  /** The node's own id, which the binding's xpath quotes. */
  readonly nodeId: string;
  /** The text the control shows. Word paints this from the store, so an empty one is an empty chip. */
  readonly label: string;
  /** The payload, serialized. JSON by convention; never parsed here. */
  readonly data: string;
}

/** Where the control goes, what it says, and the payload it carries. */
export interface InsertCustomNodeWrite {
  readonly paragraphId: string;
  readonly offset: number;
  /**
   * Wrap rather than insert: the text from `offset` to here is removed first.
   *
   * The node's label REPLACES the words it covered, because that is what turning a stretch of a
   * sentence into a citation means. Removed inside the same transaction, so a refused insert
   * leaves the text where it was.
   */
  readonly replaceUntil?: number;
  /**
   * Rewrite an existing control: it and the payload it bound go first, in this transaction.
   *
   * An UPDATE is remove-and-reinsert, because the tag codec has no in-place rewrite and
   * pretending it did would put a second write path beside this one. Doing both halves here is
   * what keeps it one undo step, and what stops an update from leaving a label and a payload
   * that disagree.
   */
  readonly replaceControlId?: string;
  readonly tag: string;
  readonly text: string;
  readonly alias?: string;
  /** Defaults to none. Callers that want Word's own "cannot type into it" pass `contentLocked`. */
  readonly lock?: 'sdtLocked' | 'sdtContentLocked' | 'contentLocked';
  /** Omitted authors an ordinary tagged control with no store — the pre-payload behaviour. */
  readonly payload?: CustomNodePayloadWrite;
}

/**
 * Why a payload write was refused.
 *
 * The tree rejections pass through unchanged, so a locked paragraph refuses a bound insert for
 * the same named reason it refuses a plain one. The three added here are the payload's own.
 */
export type CustomNodeWriteRejection =
  | TreeOpRejection
  /** The id, root name or namespace cannot be spelled in an XPath, so no binding could name it. */
  | 'unaddressable-payload'
  /** The store could not be authored — see `withCustomXmlDataPart` for every way that happens. */
  | 'store-not-authored'
  /** The payload or the label is past the cap. */
  | 'payload-too-large';

export type CustomNodeWriteResult =
  | {
      readonly ok: true;
      readonly change: TreeModelChange | null;
      /**
       * The control this write authored, when it authored one. A rewrite replaces the control
       * rather than editing it, so the id the caller passed in names nothing afterwards.
       */
      readonly nodeId?: string;
    }
  | { readonly ok: false; readonly reason: CustomNodeWriteRejection; readonly detail?: string };

export { MAX_CUSTOM_NODE_LABEL_LENGTH, MAX_CUSTOM_NODE_PAYLOAD_LENGTH };

/** Commit a plan on one story store: its package edits, then its ops, in one transaction. */
function transactPlan(store: TreeDocumentStore, plan: CustomNodeWritePlan): TransactResult {
  return store.transact((ctx) => {
    for (const edit of plan.packageEdits) {
      if (!ctx.applyPackage(edit)) return;
    }
    for (const op of plan.ops) {
      if (!ctx.apply(op)) return;
    }
  });
}

/** A transaction's refusal as a write's. */
function refusedWrite(
  result: Extract<TransactResult, { ok: false }>
): Extract<CustomNodeWriteResult, { ok: false }> {
  return {
    ok: false,
    reason: result.reason,
    ...(result.detail === undefined ? {} : { detail: result.detail }),
  };
}

/**
 * Insert one custom node, with its payload, as a single transaction.
 *
 * Answers the store transaction's own change so the caller can publish it — a payload write is a
 * package write reaching through a story store, exactly as a comment is, and the coordinator
 * needs the change to know which paragraphs went dirty.
 */
export function insertCustomNodeWrite(
  store: TreeDocumentStore,
  write: InsertCustomNodeWrite,
  /**
   * The part the customXml store hangs off, defaulting to the story being written.
   *
   * A chip in a header is written against the HEADER store, but Word enumerates the data
   * store from the main document part — a store authored off a header is one Word never
   * sees. So the caller passes the main part while the control itself lands in the story.
   */
  dataOwnerPartName?: string
): CustomNodeWriteResult {
  const planned = planCustomNodeInsert(store.package, store.part, write, dataOwnerPartName);
  if (!planned.ok) return planned;
  // Ids, not nodes: the transaction rebuilds the tree, so nothing else carries across it.
  const controlsBefore = contentControlIdsIn(store.part);
  const result = transactPlan(store, planned.plan);
  if (!result.ok) return refusedWrite(result);
  const nodeId = authoredControlId(store.part, controlsBefore);
  return { ok: true, change: result.change, ...(nodeId === undefined ? {} : { nodeId }) };
}

/** The ids of every content control a story holds. */
export function contentControlIdsIn(part: OoxmlPart): ReadonlySet<string> {
  return new Set(contentControlsIn(part.root).map((entry) => entry.node.id));
}

/**
 * The content control a write added to the story. Not from `change.created`, which reports
 * touched paragraphs rather than minted nodes and is empty for an inline insert.
 */
export function authoredControlId(
  part: OoxmlPart,
  before: ReadonlySet<string>
): string | undefined {
  for (const entry of contentControlsIn(part.root)) {
    if (!before.has(entry.node.id)) return entry.node.id;
  }
  return undefined;
}

/** Remove a control and, in the same transaction, the payload it bound. */
export function removeCustomNodeWrite(
  store: TreeDocumentStore,
  controlNodeId: string
): CustomNodeWriteResult {
  const result = transactPlan(
    store,
    planCustomNodeRemoval(store.package, store.part, controlNodeId)
  );
  if (!result.ok) return refusedWrite(result);
  return { ok: true, change: result.change };
}

/**
 * What one sweep collected, or why it collected nothing.
 *
 * `ok: false` is NOT "the document was already tidy" — that is `ok: true` with an empty
 * `removed`. It means a store this sweep was asked to tidy refused the rewrite, which a caller
 * that keeps saving into the same document should know about rather than silently retry forever.
 */
export type CustomNodeSweepResult =
  | {
      readonly ok: true;
      readonly pkg: OoxmlPackage;
      /** Node ids removed, across every store swept. Empty means there were no orphans. */
      readonly removed: readonly string[];
    }
  | { readonly ok: false; readonly reason: string };

/**
 * Drop every payload no control binds, in the stores whose namespaces a host claims.
 *
 * ON OPEN, NOT ON SAVE. A chip cut to the clipboard is unbound for as long as it sits there, so a
 * save mid-cut would destroy the payload the user is about to paste. On open the only unbound
 * nodes are ones a control genuinely lost — deleted here, or deleted in Word, which is the case
 * nothing else can collect.
 *
 * `namespaces` is the claim, and it is what keeps this off other people's stores: Word's own Cover
 * Page Properties store rides in most templates, and a sweep that walked every customXml part
 * would be deleting from it on the strength of a name collision.
 */
export function sweepCustomNodePayloads(
  pkg: OoxmlPackage,
  storyPartName: string,
  namespaces: readonly string[]
): CustomNodeSweepResult {
  if (!pkg.parts.has(storyPartName) || namespaces.length === 0) {
    return { ok: true, pkg, removed: [] };
  }
  let next = pkg;
  const removed: string[] = [];
  for (const namespaceUri of namespaces) {
    const dataPart = findCustomXmlDataPart(next, storyPartName, namespaceUri);
    if (!dataPart) continue;
    // EVERY story in the package. A control in a header binds a payload as readily as one in
    // the body, and reading the body alone made this delete a payload a header was still
    // painting — on open, with no undo entry to bring it back.
    const referenced = boundCustomXmlNodeIdsInPackage(next, dataPart.itemId);
    const swept = withoutOrphanCustomXmlNodes(next, dataPart.partName, referenced);
    // An unchanged package means "nothing to collect" OR "the rewrite was refused", and only
    // the first is a success. Told apart by asking whether anything unbound is still in there.
    if (
      swept.pkg === next &&
      customXmlNodes(next, dataPart.partName).some((n) => !referenced.has(n.id))
    ) {
      return { ok: false, reason: `the payload store for ${namespaceUri} refused the sweep` };
    }
    next = swept.pkg;
    removed.push(...swept.removed);
  }
  return { ok: true, pkg: next, removed };
}

/** One control's payload, as the store holds it. Both strings are untrusted file input. */
export interface CustomNodePayloadRead {
  /** The store node's own id. */
  readonly nodeId: string;
  /** The text Word paints the control from. */
  readonly label: string;
  /** The payload as authored. JSON by convention, unparsed here. */
  readonly data: string;
}

/**
 * The customXml stores one part relates, with each store's nodes indexed by id.
 *
 * Memoized per package, because the owner is a constant and the readers under it are not
 * memoized themselves. {@link customNodePayloadsByControl} is called once per STORY when the
 * review queue rebuilds, and the queue rebuilds per package revision — so on a document with a
 * dozen furniture parts this parsed the same store a dozen times per keystroke.
 */
const customXmlStoreNodesByPackage = new WeakMap<
  OoxmlPackage,
  Map<string, readonly { readonly itemId: string; readonly nodes: Map<string, CustomXmlNode> }[]>
>();

function customXmlStoreNodes(
  pkg: OoxmlPackage,
  dataOwnerPartName: string
): readonly { readonly itemId: string; readonly nodes: Map<string, CustomXmlNode> }[] {
  let byOwner = customXmlStoreNodesByPackage.get(pkg);
  if (!byOwner) {
    byOwner = new Map();
    customXmlStoreNodesByPackage.set(pkg, byOwner);
  }
  const cached = byOwner.get(dataOwnerPartName);
  if (cached) return cached;
  const built = customXmlDataParts(pkg, dataOwnerPartName).map((store) => ({
    itemId: store.itemId,
    nodes: new Map(customXmlNodes(pkg, store.partName).map((node) => [node.id, node])),
  }));
  byOwner.set(dataOwnerPartName, built);
  return built;
}

/**
 * The payload every control in a story binds to, keyed by the CONTROL's canonical node id.
 *
 * Keyed by the control rather than by the store node because that is the question a reader
 * actually has — "what does this chip carry" — and because two stores may each hold a `cx1`.
 * Resolved here rather than by a capability package: the stores are package parts, and a
 * derivation that only gets story parts has no way to reach them.
 *
 * Every store the story relates to, so a document carrying two definitions' payloads answers
 * for both without anyone naming a namespace.
 */
export function customNodePayloadsByControl(
  pkg: OoxmlPackage,
  storyPartName: string,
  /**
   * The part that RELATES the customXml store, when it is not the story's own.
   *
   * Two different questions wear one name here: which part holds the CONTROLS, and which part
   * relates the STORE they bind into. Word only reads a store hung off the main document part,
   * so {@link writeCustomNode} already takes them separately — this read did not, and answered
   * with the body's controls whatever story it was asked about. A chip in a header therefore
   * read as having no payload at all, and an update that meant to keep the existing data
   * dropped it.
   */
  dataOwnerPartName: string = storyPartName
): ReadonlyMap<string, CustomNodePayloadRead> {
  const found = new Map<string, CustomNodePayloadRead>();
  const story = pkg.parts.get(storyPartName);
  if (!story) return found;
  const nodesByStore = customXmlStoreNodes(pkg, dataOwnerPartName);
  if (nodesByStore.length === 0) return found;
  for (const entry of contentControlsIn(story.root)) {
    for (const store of nodesByStore) {
      const nodeId = boundCustomXmlNodeIdOf(entry.node, store.itemId);
      if (nodeId === null) continue;
      const node = store.nodes.get(nodeId);
      // A binding naming a node the store does not hold is a control Word paints from nothing.
      // Left out rather than reported as an empty payload, so a reader can tell "no payload"
      // from "a payload that says nothing".
      if (!node) continue;
      found.set(entry.node.id, { nodeId, label: node.label, data: node.data });
      break;
    }
  }
  return found;
}

/** Every payload one store holds, for a caller resolving a control's `data`. */
export function customNodePayloadsOf(
  pkg: OoxmlPackage,
  storyPartName: string,
  namespaceUri: string
): ReadonlyMap<string, { readonly label: string; readonly data: string }> {
  const found = new Map<string, { readonly label: string; readonly data: string }>();
  const dataPart = findCustomXmlDataPart(pkg, storyPartName, namespaceUri);
  if (!dataPart) return found;
  for (const node of customXmlNodes(pkg, dataPart.partName)) {
    found.set(node.id, { label: node.label, data: node.data });
  }
  return found;
}

/**
 * What the session answers for a sweep: the ids collected, or the reason none were.
 *
 * Narrower than {@link CustomNodeSweepResult}, which also carries the rewritten package — the
 * session has already installed that, and handing it back would invite a caller to install it
 * twice.
 */
export type CustomNodeSweepOutcome =
  | { readonly ok: true; readonly removed: readonly string[] }
  | { readonly ok: false; readonly reason: string };
