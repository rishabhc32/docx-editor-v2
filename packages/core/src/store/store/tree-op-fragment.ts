// `insertFragment` — the clipboard fragment's atomic landing (rich-clipboard-fidelity 2.1).
//
// The op carries BLOCKS ONLY, staying JSON-safe like every TreeDocOp; the resource merge
// travels through the package-edit path in the same transaction (clipboard-fragment-merge).
// Word's paragraph-mark rule at the seam: a merged edge paragraph takes the properties of
// the mark that ENDS it — the fragment's first-paragraph mark on the leading merge, the
// host's original mark on the trailing merge. A single-paragraph fragment whose source
// mark was not covered splices inline and leaves host properties untouched.

import type {
  OoxmlElement,
  OoxmlNode,
  OoxmlParagraphNode,
  OoxmlPart,
} from '../package/ooxml-tree.ts';
import {
  carryIndexToRebuiltRoot,
  createNodeIdAllocator,
  findNode,
  insertChildren,
  replaceChildren,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import { usedParaIds, w14PrefixInScopeAt, withFreshParaIds } from '../package/para-id.ts';
import { actorScopedSeed } from '../package/actor-scoped-ids.ts';
import {
  cloneWithNewIds,
  paragraphPropertiesNodeOf,
  parentOf,
  TEXT_DEPS,
} from './tree-op-nodes.ts';
import {
  indivisibleAt,
  isParagraph,
  paragraphLength,
  paragraphOffsetIndex,
  splitsSurrogate,
} from './tree-op-segments.ts';
import { rejectContentEdit } from './tree-op-validate-controls.ts';
import { applyTreeOp } from './tree-op-apply.ts';
import { clearPlaceholder, placeholderControlForInsertion } from './tree-op-content-controls.ts';
import { promptInsertionOffset } from './content-control-prompt-restore.ts';
import { withParagraphDiff } from './content-control-value-content.ts';
import { keepingPrompt } from './tracked-content-control-value.ts';
import { insideOwnInsertion } from './tree-op-tracked-properties.ts';
import { equationsOfAtom, isOmmlDisplay } from '../package/omml-display.ts';
import {
  acrossMarks,
  insideContainersAt,
  insideLandingRefusal,
  inlinePromptBlocks,
  landInside,
  landIntoControl,
} from './fragment-landing-inside.ts';
import type { TreeOpEffect, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';
import type { InsertFragmentOp } from './fragment-op-types.ts';
import { recordSetNamespaceBinding } from '../package/canonical-primitive-capture.ts';
import { invalidRevisionAttribution } from './tree-op-revision-attribution.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import {
  trackedFragmentBlocks,
  trackedFragmentRefusal,
  withInsertedMark,
  type FragmentWriter,
} from './tracked-fragment.ts';

export type { InsertFragmentOp } from './fragment-op-types.ts';

/** Fragment budgets: a hostile payload cannot force unbounded tree work. */
export const MAX_FRAGMENT_INSERT_BLOCKS = 50_000;
export const MAX_FRAGMENT_DEPTH = 64;
export const MAX_FRAGMENT_NODES = 500_000;

const INSERTABLE_BLOCK_KINDS = new Set(['paragraph', 'table', 'contentControl']);

function fragmentShape(
  node: OoxmlNode,
  depth: number,
  budget: { nodes: number }
): TreeOpRejection | null {
  if (depth > MAX_FRAGMENT_DEPTH) return 'fragment-too-deep';
  budget.nodes += 1;
  if (budget.nodes > MAX_FRAGMENT_NODES) return 'fragment-resource-budget';
  if (node.kind === 'textValue') {
    return typeof node.value === 'string' ? null : 'fragment-invalid-block';
  }
  if (typeof node.localName !== 'string' || typeof node.namespaceUri !== 'string') {
    return 'fragment-invalid-block';
  }
  if (!Array.isArray(node.children) || !Array.isArray(node.attributes)) {
    return 'fragment-invalid-block';
  }
  for (const child of node.children) {
    const refused = fragmentShape(child, depth + 1, budget);
    if (refused) return refused;
  }
  return null;
}

export function validateInsertFragment(
  part: OoxmlPart,
  op: InsertFragmentOp
): TreeOpRejection | null {
  const paragraph = findNode(part, op.paragraphId);
  if (!paragraph) return 'unknown-paragraph';
  if (!isParagraph(paragraph)) return 'not-a-paragraph';
  if (!Number.isInteger(op.offset) || op.offset < 0 || op.offset > paragraphLength(paragraph)) {
    return 'offset-out-of-range';
  }
  if (splitsSurrogate(paragraph, op.offset)) return 'splits-surrogate-pair';
  // Inside a link, a control or an insertion, the paste lands as Word lands it there
  if (indivisibleAt(paragraph, op.offset)) {
    const inside = insideContainersAt(paragraph, op.offset);
    const refused = insideLandingRefusal(inside, op.blocks, op.lastMarkCovered === true);
    if (refused) return refused;
  }
  if (!Array.isArray(op.blocks) || op.blocks.length === 0) return 'fragment-invalid-block';
  if (op.blocks.length > MAX_FRAGMENT_INSERT_BLOCKS) return 'fragment-resource-budget';
  const budget = { nodes: 0 };
  for (const block of op.blocks) {
    if (block.kind === 'textValue' || !INSERTABLE_BLOCK_KINDS.has(block.kind)) {
      return 'fragment-invalid-block';
    }
    const refused = fragmentShape(block, 1, budget);
    if (refused) return refused;
  }
  if (op.revision !== undefined) {
    if (invalidRevisionAttribution(op.revision)) return 'invalidArgs';
    const untrackable = trackedFragmentRefusal(op.blocks);
    if (untrackable) return untrackable;
  }
  return rejectContentEdit(part, paragraph as OoxmlParagraphNode, op.offset, op.offset);
}

/**
 * Prefix bindings the incoming nodes need that the part root does not declare.
 *
 * Fragment content legitimately uses prefixes the TARGET root never bound (`wp:`, `a:`,
 * `pic:`, `w14:` — the source bound them on ITS root, which does not travel). The commit
 * validator requires every prefix bound in scope, so the missing ones are added to the
 * root; a prefix already bound to a DIFFERENT uri stays untouched and the incoming nodes
 * keep working through their own local bindings when they carry any.
 */
export function withRequiredNamespaceBindings(
  part: OoxmlPart,
  nodes: readonly OoxmlNode[]
): OoxmlPart {
  const rootBindings = new Map(
    part.root.namespaceBindings.map((binding) => [binding.prefix, binding.namespaceUri])
  );
  const additions = new Map<string, string>();
  const visit = (node: OoxmlNode, scope: ReadonlyMap<string, string>): void => {
    if (node.kind === 'textValue') return;
    let local = scope;
    if (node.namespaceBindings.length > 0) {
      const merged = new Map(scope);
      for (const binding of node.namespaceBindings) {
        merged.set(binding.prefix, binding.namespaceUri);
      }
      local = merged;
    }
    const need = (prefix: string | undefined, namespaceUri: string): void => {
      if (prefix === undefined || prefix.length === 0 || namespaceUri.length === 0) return;
      // `xml` and `xmlns` are bound by the XML spec itself, never by a declaration.
      // Treating `xml` as unbound declared it redundantly and rebuilt the root for
      // virtually every paste — `xml:space="preserve"` is in every real document copy.
      if (prefix === 'xml' || prefix === 'xmlns') return;
      if (local.get(prefix) === namespaceUri) return;
      if (rootBindings.get(prefix) === namespaceUri) return;
      if (rootBindings.has(prefix) || additions.has(prefix)) {
        if (additions.get(prefix) === namespaceUri) return;
        // Bound to something else: leave it; the serializer resolves through local bindings.
        if (rootBindings.get(prefix) !== undefined && rootBindings.get(prefix) !== namespaceUri) {
          return;
        }
        return;
      }
      additions.set(prefix, namespaceUri);
    };
    need(node.prefix, node.namespaceUri);
    for (const attribute of node.attributes) {
      need(attribute.prefix, attribute.namespaceUri);
    }
    for (const child of node.children) visit(child, local);
  };
  for (const node of nodes) visit(node, rootBindings);
  if (additions.size === 0) return part;
  const root = {
    ...part.root,
    namespaceBindings: [
      ...part.root.namespaceBindings,
      ...[...additions].map(([prefix, namespaceUri]) => ({ prefix, namespaceUri })),
    ],
  } as typeof part.root;
  // Same id, same children — only the bindings changed. Carry the index like every op
  // executor does, or the rebuilt root re-walks the whole part on its next lookup and
  // restarts the mint frontier underneath ids already handed out.
  carryIndexToRebuiltRoot(part.root, root);
  for (const [prefix, namespaceUri] of additions) {
    recordSetNamespaceBinding(part.root.id, prefix, namespaceUri);
  }
  return Object.freeze({ ...part, root }) as OoxmlPart;
}

/** Keep local bindings when a paste removes the node's original parent. */
function withParentBindings(node: OoxmlNode, parent: OoxmlElement): OoxmlNode {
  if (node.kind === 'textValue' || parent.namespaceBindings.length === 0) return node;
  const ownPrefixes = new Set(node.namespaceBindings.map((binding) => binding.prefix));
  const inherited = parent.namespaceBindings.filter((binding) => !ownPrefixes.has(binding.prefix));
  if (inherited.length === 0) return node;
  return { ...node, namespaceBindings: [...inherited, ...node.namespaceBindings] };
}

function inlineChildrenOf(paragraph: OoxmlElement): readonly OoxmlNode[] {
  const pPr = paragraphPropertiesNodeOf(paragraph);
  return paragraph.children
    .filter((child) => child !== pPr)
    .map((child) => withParentBindings(child, paragraph));
}

/**
 * A display equation (`m:oMathPara`) is a paragraph of its own. Merged beside existing
 * text, it becomes its inline equations, the same as Word pastes a display into a line.
 */
function besideText(
  pasted: readonly OoxmlNode[],
  neighbour: readonly OoxmlNode[]
): readonly OoxmlNode[] {
  const hasContent = (nodes: readonly OoxmlNode[]): boolean =>
    paragraphLength({ children: nodes } as unknown as OoxmlParagraphNode) > 0;
  if (!pasted.some(isOmmlDisplay) || !hasContent(neighbour)) return pasted;
  return pasted.flatMap((node): OoxmlNode[] =>
    isOmmlDisplay(node)
      ? equationsOfAtom(node).map((equation) => withParentBindings(equation, node))
      : [node]
  );
}

/** The head paragraph with its mark stamped as the paste's insertion. */
function withHeadMarkInserted(
  part: OoxmlPart,
  headId: string,
  writer: FragmentWriter,
  options?: EditOptions
): OoxmlPart | null {
  const head = findNode(part, headId);
  if (!head || head.kind !== 'paragraph') return null;
  const pPr = paragraphPropertiesNodeOf(head);
  const replaced = replaceChildren(
    part,
    head.id,
    [withInsertedMark(pPr, writer), ...head.children.filter((child) => child !== pPr)],
    { ...options, deferValidation: true }
  );
  return replaced.ok ? replaced.part : null;
}

function rebuiltParagraph(
  pPr: OoxmlElement | undefined,
  inline: readonly OoxmlNode[]
): readonly OoxmlNode[] {
  return pPr ? [pPr, ...inline] : inline;
}

export function applyInsertFragment(
  hostPart: OoxmlPart,
  op: InsertFragmentOp,
  options?: EditOptions
): TreeOpResult {
  const prompt = placeholderControlForInsertion(hostPart, op.paragraphId, op.offset);
  return prompt
    ? landOverPrompt(hostPart, op, prompt, options)
    : landFragment(hostPart, op, options);
}

/**
 * A paste at a control's prompt REPLACES the prompt, as typed text does (`tree-op-apply.ts`)
 * and as Word pastes: landed inside it, the pasted blocks became prompt text, still marked
 * `w:showingPlcHdr`. The prompt goes without a revision and stays in the glossary, so a reject
 * or a deletion that empties the control shows it again.
 */
function landOverPrompt(
  part: OoxmlPart,
  op: InsertFragmentOp,
  prompt: { readonly control: OoxmlNode; readonly offset: number },
  options?: EditOptions
): TreeOpResult {
  const cleared = clearPlaceholder(part, prompt.control.id, options, op.paragraphId);
  if (!cleared) return { ok: false, reason: 'unsupported' };
  const emptied = withoutEmptyRuns(part, cleared, op.paragraphId, options);
  const offset = promptInsertionOffset(emptied, op.paragraphId, prompt.offset);
  // At an inline control's prompt the paste lands in the control, as one line: Word turns
  // each pasted mark into a space there
  const line = inlinePromptBlocks(
    findNode(part, op.paragraphId),
    prompt.control,
    op.blocks,
    op.lastMarkCovered === true,
    createNodeIdAllocator(part, 'paste')
  );
  const result = line
    ? landFragment(
        emptied,
        { ...op, blocks: line, lastMarkCovered: false, offset },
        options,
        prompt.control.id
      )
    : landFragment(emptied, { ...op, offset }, options);
  const control = findNode(emptied, prompt.control.id);
  if (!result.ok || !control) return result;
  const replaced = { ...result, effect: withParagraphDiff(result.effect, prompt.control, control) };
  // The prompt stays the control's, in the glossary, as typing over it keeps it; not under a
  // tracked paste inside the author's own new control, whose reject takes the control too
  return op.revision && insideOwnInsertion(part, prompt.control.id, op.revision.author)
    ? replaced
    : keepingPrompt(part, prompt.control, replaced, options);
}

/**
 * The emptied prompt's paragraph without the empty run `clearPlaceholder` leaves for typing.
 * Only a run that `before` did not hold goes: a field's runs after its begin, or a comment
 * reference, have no offset length either and stay.
 */
export function withoutEmptyRuns(
  before: OoxmlPart,
  cleared: OoxmlPart,
  paragraphId: string,
  options?: EditOptions
): OoxmlPart {
  const paragraph = findNode(cleared, paragraphId);
  if (!paragraph || paragraph.kind !== 'paragraph') return cleared;
  const index = paragraphOffsetIndex(paragraph);
  const kept = paragraph.children.filter(
    (child) => child.kind !== 'run' || index.lengthOf(child) > 0 || findNode(before, child.id)
  );
  if (kept.length === paragraph.children.length) return cleared;
  const replaced = replaceChildren(cleared, paragraph.id, kept, options);
  return replaced.ok ? replaced.part : cleared;
}

/** `into`: the control whose prompt the paste cleared, which takes the paste as its content. */
function landFragment(
  hostPart: OoxmlPart,
  op: InsertFragmentOp,
  options?: EditOptions,
  into?: string
): TreeOpResult {
  const host = findNode(hostPart, op.paragraphId);
  if (!host || host.kind !== 'paragraph') return { ok: false, reason: 'unknown-paragraph' };

  // Clone every block with fresh node ids and fresh paragraph identities: the fragment's
  // ids belong to another document, and a second paste of the same payload must not
  // collide with the first.
  //
  // The clones stay DETACHED until the split/merge sequence below lands them, so their ids
  // are invisible to every later allocator's in-tree check — the frontier alone kept them
  // apart, and losing it (an orphaned index after a root rebuild) re-issued a clone's id to
  // the split's tail and refused the whole paste as duplicate ids. A dedicated family makes
  // the disjointness structural instead of frontier-dependent; in-tree occupancy (a prior
  // paste) is still skipped by the allocator's index check.
  const nextId = createNodeIdAllocator(hostPart, 'paste');
  const paraIds = new Set(usedParaIds(hostPart.root as OoxmlElement));
  const counter = { value: 0 };
  // The clones land beside the host paragraph (or inside it), so its scope decides the prefix.
  const hostPrefix = w14PrefixInScopeAt(hostPart, host);
  const cloned = op.blocks.map((block, index) =>
    withFreshParaIds(
      cloneWithNewIds(block, nextId),
      paraIds,
      actorScopedSeed(`${op.paragraphId}:${op.offset}:${index}`),
      counter,
      hostPrefix
    )
  );
  // Tracked: the whole paste is one insertion, so it takes ONE id from the part as it stands
  // after whatever this transaction already wrote (a replaced selection's struck words).
  const writer = op.revision
    ? { revisionId: nextRevisionId(hostPart)(), revision: op.revision, nextId }
    : null;
  // Inside a link or a control the paste goes into it, its paragraph marks spaces in a control
  const inside = indivisibleAt(host, op.offset) ? insideContainersAt(host, op.offset) : null;
  const intoContainer = inside !== null && (inside.link || inside.control);
  const shaped = intoContainer ? acrossMarks(cloned, op.lastMarkCovered === true, nextId) : cloned;
  const blocks = writer ? trackedFragmentBlocks(shaped, writer) : shaped;
  if (!blocks) return { ok: false, reason: 'fragment-untrackable-content' };

  // Prefixes the fragment uses must be bound at the target root before commit validation.
  const part = withRequiredNamespaceBindings(hostPart, blocks);

  const first = blocks[0]!;
  const last = blocks[blocks.length - 1]!;
  if (intoContainer || into !== undefined) {
    const inline = besideText(inlineChildrenOf(first as OoxmlElement), inlineChildrenOf(host));
    return into === undefined
      ? landInside(part, host, op.offset, inline, options)
      : landIntoControl(part, host.id, into, inline, options);
  }
  const inlineOnly =
    blocks.length === 1 && first.kind === 'paragraph' && op.lastMarkCovered !== true;

  if (inlineOnly) {
    // Pure inline splice: split, append the fragment's inline content to the head, join.
    // Reusing split+join keeps every offset/atom/wrapper rule identical to typing lanes.
    const split = applyTreeOp(
      part,
      { op: 'splitParagraph', paragraphId: host.id, offset: op.offset },
      options
    );
    if (!split.ok) return split;
    const tailId = split.effect.split?.tail;
    if (!tailId) return { ok: false, reason: 'tree-invariant' };
    const head = findNode(split.part, host.id);
    if (!head || head.kind !== 'paragraph') return { ok: false, reason: 'tree-invariant' };
    const appended = replaceChildren(
      split.part,
      head.id,
      [
        ...head.children,
        ...besideText(inlineChildrenOf(first as OoxmlElement), inlineChildrenOf(host)),
      ],
      { ...options, deferValidation: true }
    );
    if (!appended.ok) return { ok: false, reason: 'tree-invariant' };
    const joined = applyTreeOp(
      appended.part,
      { op: 'joinParagraphs', firstId: head.id, secondId: tailId },
      options
    );
    if (!joined.ok) return joined;
    const effect: TreeOpEffect = {
      dirty: [host.id],
      created: [],
      deleted: [],
      dependencyKeys: TEXT_DEPS,
      impact: 'flow-structural',
    };
    return { ok: true, part: joined.part, effect };
  }

  // Structural landing: split the host, merge paragraph edges, insert the rest between.
  const split = applyTreeOp(
    part,
    { op: 'splitParagraph', paragraphId: host.id, offset: op.offset },
    options
  );
  if (!split.ok) return split;
  const tailId = split.effect.split?.tail;
  if (!tailId) return { ok: false, reason: 'tree-invariant' };
  let current = split.part;
  // A fragment that opens with a table keeps the host's properties on the head, but the head
  // now ends on the mark the split wrote, which the paste inserts.
  if (writer && first.kind !== 'paragraph') {
    const stamped = withHeadMarkInserted(current, host.id, writer, options);
    if (!stamped) return { ok: false, reason: 'tree-invariant' };
    current = stamped;
  }

  const middle: OoxmlNode[] = [...blocks];
  const created: string[] = [tailId];

  // Leading merge: the merged paragraph ends at the fragment's first mark, so it takes the
  // fragment's paragraph properties and keeps the host's leading text.
  if (first.kind === 'paragraph') {
    middle.shift();
    const head = findNode(current, host.id);
    if (!head || head.kind !== 'paragraph') return { ok: false, reason: 'tree-invariant' };
    const sourcePPr = paragraphPropertiesNodeOf(first as OoxmlElement);
    const fragmentPPr = sourcePPr
      ? (withParentBindings(sourcePPr, first as OoxmlElement) as OoxmlElement)
      : undefined;
    const headInline = inlineChildrenOf(head);
    const replaced = replaceChildren(
      current,
      head.id,
      rebuiltParagraph(fragmentPPr, [
        ...headInline,
        ...besideText(inlineChildrenOf(first as OoxmlElement), headInline),
      ]),
      { ...options, deferValidation: true }
    );
    if (!replaced.ok) return { ok: false, reason: 'tree-invariant' };
    current = replaced.part;
  }

  // Trailing merge: the merged paragraph ends at the HOST's original mark (now on the
  // tail), so the host properties stay and the fragment's last inline content leads it.
  if (
    middle.length > 0 &&
    last.kind === 'paragraph' &&
    middle[middle.length - 1] === last &&
    op.lastMarkCovered !== true
  ) {
    middle.pop();
    const tail = findNode(current, tailId);
    if (!tail || tail.kind !== 'paragraph') return { ok: false, reason: 'tree-invariant' };
    const tailPPr = paragraphPropertiesNodeOf(tail);
    const tailInline = inlineChildrenOf(tail);
    const replaced = replaceChildren(
      current,
      tail.id,
      rebuiltParagraph(tailPPr, [
        ...besideText(inlineChildrenOf(last as OoxmlElement), tailInline),
        ...tailInline,
      ]),
      { ...options, deferValidation: true }
    );
    if (!replaced.ok) return { ok: false, reason: 'tree-invariant' };
    current = replaced.part;
  }

  if (middle.length > 0) {
    const parent = parentOf(current, tailId);
    if (!parent) return { ok: false, reason: 'tree-invariant' };
    const tailIndex = parent.children.findIndex((child) => child.id === tailId);
    if (tailIndex === -1) return { ok: false, reason: 'tree-invariant' };
    const inserted = insertChildren(current, parent.id, tailIndex, middle, {
      ...options,
      deferValidation: true,
    });
    if (!inserted.ok) return { ok: false, reason: 'tree-invariant' };
    current = inserted.part;
    for (const block of middle) created.push(block.id);
  }

  const effect: TreeOpEffect = {
    dirty: [host.id, tailId],
    created,
    deleted: [],
    ...(op.lastMarkCovered === true ? { caret: { paragraphId: tailId } } : {}),
    dependencyKeys: TEXT_DEPS,
    impact: 'flow-structural',
  };
  return { ok: true, part: current, effect };
}
