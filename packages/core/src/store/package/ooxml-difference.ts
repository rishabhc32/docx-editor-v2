// Where two canonical trees differ, answered in the size of the difference.

import { WML_NAMESPACE_URI, canonicalQNameAttributeValue } from './ooxml-shared.ts';
import {
  DEFAULT_FINGERPRINT_BINDINGS,
  extendedBindings,
  inheritedPreserveOf,
  normalizedFingerprintAttributes,
  significantChildren,
  sortedAttributes,
} from './ooxml-serialize.ts';
import type { OoxmlAttribute, OoxmlElement, OoxmlNode, OoxmlPart } from './ooxml-tree.ts';

/** How far the difference between two canonical trees reaches. */
export type CanonicalDifferenceReach =
  | 'equal'
  /** Only `w:t` character data differs, and only inside paragraphs. */
  | 'paragraph-text'
  /** Element structure or attributes differ, but only inside paragraphs. */
  | 'paragraph-content'
  /**
   * Paragraphs were added to or dropped from a block container whose other blocks kept their
   * ids; every other difference sits inside paragraphs. See `createdParagraphIds`.
   */
  | 'paragraph-sequence'
  /** At least one difference sits above every paragraph. */
  | 'outside-paragraph';

/** How far the renamed nodes inside kept paragraphs reach. */
export type CanonicalRenameReach = 'equal' | 'paragraph-text' | 'paragraph-content';

/** What an incremental comparison of two canonical trees found. */
export interface CanonicalTreeDifference {
  /** True when the two trees say the same thing, by the same rules as the fingerprint. */
  readonly equal: boolean;
  /**
   * True when the walk stopped without an answer. `equal` and `reach` mean nothing then,
   * and the caller must ask the full oracle instead.
   */
  readonly undecided: boolean;
  readonly reach: CanonicalDifferenceReach;
  /** Paragraph ids from `next` that own a difference. Complete only when decided. */
  readonly paragraphIds: readonly string[];
  /** True when every node pair the walk compared carried the same canonical id. */
  readonly idsPreserved: boolean;
  /**
   * Paragraph ids from `next` whose content kept its shape but changed ids, and whether an
   * id changed anywhere else. A collaboration view renames inline ids it shares with another
   * paragraph, and the editor has to take those ids even when the content is equal.
   */
  readonly renamedParagraphIds: readonly string[];
  readonly renamedOutsideParagraphs: boolean;
  /**
   * `paragraph-text` when every node renamed inside a kept paragraph is text flow (a run,
   * its text, a revision wrapper, run properties), `paragraph-content` when one is anything
   * else (a control, a link, a field), `equal` when none was renamed.
   */
  readonly renamedReach: CanonicalRenameReach;
  /**
   * Paragraphs in `next` that `previous` did not hold, and the reverse, where a block
   * container's paragraph sequence changed while every other block kept its id (a split, a
   * join, a paragraph replaced under another id). Empty otherwise.
   */
  readonly createdParagraphIds: readonly string[];
  readonly deletedParagraphIds: readonly string[];
  /** Node pairs the walk compared. An identity-equal subtree counts as one. */
  readonly visited: number;
}

const REACH_RANK: Record<CanonicalDifferenceReach, number> = {
  equal: 0,
  'paragraph-text': 1,
  'paragraph-content': 2,
  'paragraph-sequence': 3,
  'outside-paragraph': 4,
};

/**
 * Node pairs one comparison may look at before it gives up.
 *
 * The ceiling is high because giving up is the expensive outcome, not the cheap one: the
 * caller's fallback is the fingerprint oracle, which allocates a nested array for every node
 * of both trees and then stringifies it. Finishing the walk is cheaper than that even when
 * the two trees share nothing. The budget is here so recursion over a hostile tree cannot
 * run unbounded, not to steer ordinary work toward the oracle.
 */
const DEFAULT_DIFFERENCE_BUDGET = 250_000;

interface DifferenceWalk {
  readonly budget: number;
  visited: number;
  undecided: boolean;
  reach: CanonicalDifferenceReach;
  idsPreserved: boolean;
  readonly paragraphIds: Set<string>;
  readonly renamedParagraphIds: Set<string>;
  renamedOutsideParagraphs: boolean;
  renamedReach: CanonicalRenameReach;
  readonly created: Set<string>;
  readonly deleted: Set<string>;
}

function newWalk(budget: number): DifferenceWalk {
  return {
    budget,
    visited: 0,
    undecided: false,
    reach: 'equal',
    idsPreserved: true,
    paragraphIds: new Set<string>(),
    renamedParagraphIds: new Set<string>(),
    renamedOutsideParagraphs: false,
    renamedReach: 'equal',
    created: new Set<string>(),
    deleted: new Set<string>(),
  };
}

/** True when the two elements declare the same prefixes, in the same order. */
export function namespaceDeclarationsMatch(left: OoxmlElement, right: OoxmlElement): boolean {
  if (left.namespaceBindings === right.namespaceBindings) return true;
  if (left.namespaceBindings.length !== right.namespaceBindings.length) return false;
  return left.namespaceBindings.every((binding, index) => {
    const other = right.namespaceBindings[index]!;
    return binding.prefix === other.prefix && binding.namespaceUri === other.namespaceUri;
  });
}

function fingerprintAttributesMatch(
  left: readonly OoxmlAttribute[],
  right: readonly OoxmlAttribute[],
  bindings: ReadonlyMap<string, string>,
  owner: OoxmlElement
): boolean {
  if (left.length !== right.length) return false;
  const leftSorted = sortedAttributes(left);
  const rightSorted = sortedAttributes(right);
  return leftSorted.every((attribute, index) => {
    const other = rightSorted[index]!;
    return (
      attribute.namespaceUri === other.namespaceUri &&
      attribute.localName === other.localName &&
      canonicalQNameAttributeValue(attribute, bindings, owner.namespaceUri, owner.localName) ===
        canonicalQNameAttributeValue(other, bindings, owner.namespaceUri, owner.localName)
    );
  });
}

/**
 * Elements whose content only ever re-breaks the paragraph that holds it.
 *
 * A local `insertText` publishes `text-local` however it lands (merged into a run, split
 * across two, or wrapped in a tracked insertion), so a received one has to be allowed to say
 * the same thing. Everything absent from this list re-flows more than one paragraph, or
 * might: a drawing changes height, a `w:br` moves a page boundary, a note reference reaches
 * another story. The list is an allowlist on purpose. An element nobody has classified yet
 * reads as flow-affecting and costs a wider invalidation, which is the safe direction.
 */
const TEXT_FLOW_ELEMENTS = new Set(['r', 't', 'delText', 'ins', 'del', 'rPr']);

function isTextFlowElement(node: OoxmlNode): boolean {
  return (
    node.kind !== 'textValue' &&
    node.namespaceUri === WML_NAMESPACE_URI &&
    TEXT_FLOW_ELEMENTS.has(node.localName)
  );
}

/** Count one node pair against the budget; false, with the walk undecided, once it is spent. */
function spend(walk: DifferenceWalk): boolean {
  walk.visited += 1;
  if (walk.visited > walk.budget) {
    walk.undecided = true;
    return false;
  }

  return true;
}

/** The reach of a difference inside a paragraph: text flow only, or other content. */
function textOrContent(textOnly: boolean): 'paragraph-text' | 'paragraph-content' {
  return textOnly ? 'paragraph-text' : 'paragraph-content';
}

function isTextFlowSubtree(node: OoxmlNode, walk: DifferenceWalk): boolean {
  if (!spend(walk)) return false;
  if (node.kind === 'textValue') return true;
  if (!isTextFlowElement(node)) return false;

  // Run properties are formatting only: no flow content can hang under them.
  if (node.localName === 'rPr') return true;
  return node.children.every((child) => isTextFlowSubtree(child, walk));
}

/** True when every child either side does not share by identity only re-breaks its paragraph. */
function replacedChildrenAreTextFlow(
  left: readonly OoxmlNode[],
  right: readonly OoxmlNode[],
  walk: DifferenceWalk
): boolean {
  const inLeft = new Set(left);
  const inRight = new Set(right);
  return (
    left.every((child) => inRight.has(child) || isTextFlowSubtree(child, walk)) &&
    right.every((child) => inLeft.has(child) || isTextFlowSubtree(child, walk))
  );
}

function widened<T extends CanonicalDifferenceReach>(current: T, next: T): T {
  return REACH_RANK[next] > REACH_RANK[current] ? next : current;
}

function recordDifference(
  walk: DifferenceWalk,
  paragraphId: string | null,
  reach: 'paragraph-text' | 'paragraph-content'
): void {
  if (paragraphId === null) {
    walk.reach = 'outside-paragraph';
    return;
  }

  walk.paragraphIds.add(paragraphId);
  walk.reach = widened(walk.reach, reach);
}

/**
 * What the nodes under comparison are: a text element's characters, a run's formatting (its
 * `w:rPr` and everything in it, which only re-breaks the paragraph), or anything else.
 */
type Within = 'content' | 'characters' | 'formatting';

/** Note a node pair under different ids: inside a kept paragraph, or anywhere else. */
function recordRename(
  walk: DifferenceWalk,
  left: OoxmlNode,
  right: OoxmlNode,
  paragraphId: string | null,
  within: Within
): void {
  walk.idsPreserved = false;
  if (paragraphId === null || left.kind === 'paragraph') {
    walk.renamedOutsideParagraphs = true;
    return;
  }

  walk.renamedParagraphIds.add(paragraphId);
  const textFlow = (node: OoxmlNode): boolean =>
    node.kind === 'textValue' || isTextFlowElement(node);
  walk.renamedReach = widened(
    walk.renamedReach,
    textOrContent(within === 'formatting' || (textFlow(left) && textFlow(right)))
  );
}

/** What a child of `parent` is, for the walk below it. */
function withinChild(parent: OoxmlElement, child: OoxmlNode, within: Within): Within {
  if (within === 'formatting') return within;
  if (parent.kind === 'run' && child.kind === 'runProperties') return 'formatting';
  return parent.kind === 'text' || parent.kind === 'deletedText' ? 'characters' : 'content';
}

/** True when the two block lists hold the same ids in the same places. */
function sameChildIds(left: readonly OoxmlNode[], right: readonly OoxmlNode[]): boolean {
  if (left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) {
    const leftNode = left[index]!;
    const rightNode = right[index]!;
    if (leftNode !== rightNode && leftNode.id !== rightNode.id) return false;
  }

  return true;
}

/** How many blocks at the start and at the end of two block lists keep their ids. */
function sequenceEdges(
  left: readonly OoxmlNode[],
  right: readonly OoxmlNode[]
): { readonly prefix: number; readonly suffix: number } {
  const limit = Math.min(left.length, right.length);
  let prefix = 0;
  while (prefix < limit && left[prefix]!.id === right[prefix]!.id) {
    prefix += 1;
  }

  let suffix = 0;
  while (
    suffix < limit - prefix &&
    left[left.length - 1 - suffix]!.id === right[right.length - 1 - suffix]!.id
  ) {
    suffix += 1;
  }

  return { prefix, suffix };
}

/**
 * Where each left paragraph sits by id, and the right ids; null when an id repeats on one side,
 * or when a paragraph that stays passes another one that stays (a move).
 */
function stayingIndex(
  leftMiddle: readonly OoxmlNode[],
  rightMiddle: readonly OoxmlNode[]
): { readonly leftAt: ReadonlyMap<string, number>; readonly rightIds: ReadonlySet<string> } | null {
  const leftAt = new Map<string, number>();
  for (const [index, node] of leftMiddle.entries()) {
    if (leftAt.has(node.id)) return null;
    leftAt.set(node.id, index);
  }

  const rightIds = new Set<string>();
  let last = -1;
  for (const node of rightMiddle) {
    if (rightIds.has(node.id)) return null;
    rightIds.add(node.id);
    const at = leftAt.get(node.id);
    if (at === undefined) continue;
    if (at <= last) return null;
    last = at;
  }

  return { leftAt, rightIds };
}

/**
 * A block container whose children moved: the blocks before and after the change keep their
 * ids (compared in place), and the blocks between are all paragraphs on both sides, the ones
 * that stay keeping their order. The paragraphs that stay are compared; the others are named
 * created or deleted. Returns false, having recorded nothing, for any other shape.
 */
function diffParagraphSequence(
  left: readonly OoxmlNode[],
  right: readonly OoxmlNode[],
  preserve: boolean,
  bindings: ReadonlyMap<string, string>,
  walk: DifferenceWalk
): boolean {
  const { prefix, suffix } = sequenceEdges(left, right);
  const leftMiddle = left.slice(prefix, left.length - suffix);
  const rightMiddle = right.slice(prefix, right.length - suffix);
  const allParagraphs = (nodes: readonly OoxmlNode[]): boolean =>
    nodes.every((node) => node.kind === 'paragraph');
  if (!allParagraphs(leftMiddle) || !allParagraphs(rightMiddle)) return false;
  const staying = stayingIndex(leftMiddle, rightMiddle);
  if (staying === null) return false;
  const { leftAt, rightIds } = staying;
  walk.idsPreserved = false;
  const compareInPlace = (leftNode: OoxmlNode, rightNode: OoxmlNode): void =>
    diffNodes(leftNode, rightNode, preserve, bindings, null, 'content', walk);
  for (let index = 0; index < prefix && !walk.undecided; index += 1) {
    compareInPlace(left[index]!, right[index]!);
  }

  for (let index = 1; index <= suffix && !walk.undecided; index += 1) {
    compareInPlace(left[left.length - index]!, right[right.length - index]!);
  }

  // The paragraphs that stay are named by id; the verdict compares the paragraphs in place,
  // because the fingerprint reads them in place
  const naming = newWalk(walk.budget - walk.visited);
  for (const node of rightMiddle) {
    if (naming.undecided) break;
    const at = leftAt.get(node.id);
    if (at === undefined) {
      walk.created.add(node.id);
    } else {
      diffNodes(leftMiddle[at]!, node, preserve, bindings, null, 'content', naming);
    }
  }

  absorbNames(walk, naming);
  for (const node of leftMiddle) {
    if (!rightIds.has(node.id)) walk.deleted.add(node.id);
  }

  if (!walk.undecided && !middlesSayTheSame(leftMiddle, rightMiddle, preserve, bindings, walk)) {
    walk.reach = widened(walk.reach, 'paragraph-sequence');
  }

  return true;
}

/** Take the paragraphs a side walk named, and its cost, but not its verdict. */
function absorbNames(walk: DifferenceWalk, side: DifferenceWalk): void {
  walk.visited += side.visited;
  if (side.undecided) walk.undecided = true;
  for (const id of side.paragraphIds) {
    walk.paragraphIds.add(id);
  }

  for (const id of side.renamedParagraphIds) {
    walk.renamedParagraphIds.add(id);
  }

  if (side.renamedOutsideParagraphs) walk.renamedOutsideParagraphs = true;
  walk.renamedReach = widened(walk.renamedReach, side.renamedReach);
}

/**
 * Whether the paragraphs between say the same thing in place, ids aside, so `equal` keeps
 * agreeing with the fingerprint oracle. Counted against the same budget.
 */
function middlesSayTheSame(
  left: readonly OoxmlNode[],
  right: readonly OoxmlNode[],
  preserve: boolean,
  bindings: ReadonlyMap<string, string>,
  walk: DifferenceWalk
): boolean {
  if (left.length !== right.length) return false;
  const scratch = newWalk(walk.budget - walk.visited);
  for (let index = 0; index < left.length; index += 1) {
    diffNodes(left[index]!, right[index]!, preserve, bindings, null, 'content', scratch);
    if (scratch.undecided || scratch.reach !== 'equal') break;
  }

  walk.visited += scratch.visited;
  if (scratch.undecided) walk.undecided = true;
  return !scratch.undecided && scratch.reach === 'equal';
}

/** Compare a node pair where one side, or both, is character data. */
function diffTextValues(
  left: OoxmlNode,
  right: OoxmlNode,
  paragraphId: string | null,
  within: Within,
  walk: DifferenceWalk
): void {
  if (left.kind !== 'textValue' || right.kind !== 'textValue') {
    recordDifference(walk, paragraphId, textOrContent(within === 'formatting'));
    return;
  }

  if (left.value !== right.value) {
    recordDifference(walk, paragraphId, textOrContent(within !== 'content'));
  }
}

/** Compare the children of two elements whose names and attributes match. */
function diffChildren(
  left: OoxmlElement,
  right: OoxmlElement,
  preserve: boolean,
  bindings: ReadonlyMap<string, string>,
  scope: string | null,
  within: Within,
  walk: DifferenceWalk
): void {
  const leftChildren = significantChildren(left, preserve);
  const rightChildren = significantChildren(right, preserve);

  // Above the paragraphs, a split or a join changes a container's block list: the blocks
  // around it keep their ids, so the difference can still be named paragraph by paragraph.
  if (
    scope === null &&
    !sameChildIds(leftChildren, rightChildren) &&
    diffParagraphSequence(leftChildren, rightChildren, preserve, bindings, walk)
  ) {
    return;
  }

  if (leftChildren.length !== rightChildren.length) {
    const textFlowOnly =
      within === 'formatting' ||
      ((left.kind === 'paragraph' || isTextFlowElement(left)) &&
        replacedChildrenAreTextFlow(leftChildren, rightChildren, walk) &&
        !walk.undecided);
    recordDifference(walk, scope, textOrContent(textFlowOnly));
    return;
  }

  for (let index = 0; index < leftChildren.length; index += 1) {
    const child = leftChildren[index]!;
    diffNodes(
      child,
      rightChildren[index]!,
      preserve,
      bindings,
      scope,
      withinChild(left, child, within),
      walk
    );
    if (walk.undecided) return;
  }
}

/**
 * Compare two canonical subtrees in lockstep, skipping the parts that never moved.
 *
 * The engine hands identical node objects back for content it did not rebuild, so an
 * identity match proves a whole subtree is equal for free. Everything else mirrors
 * `fingerprintNode` exactly (same attribute normalization, same `xml:space` inheritance,
 * same QName canonicalization, same whitespace-significance rule), so the two agree on
 * every pair of trees.
 */
function diffNodes(
  left: OoxmlNode,
  right: OoxmlNode,
  inheritedPreserve: boolean,
  inheritedBindings: ReadonlyMap<string, string>,
  paragraphId: string | null,
  within: Within,
  walk: DifferenceWalk
): void {
  if (walk.undecided || !spend(walk)) return;

  // An identity match is only sound because every ancestor pair matched: the inherited
  // bindings and `xml:space` state both sides carry into this subtree are the same.
  if (left === right) return;
  if (left.id !== right.id) recordRename(walk, left, right, paragraphId, within);
  if (left.kind === 'textValue' || right.kind === 'textValue') {
    diffTextValues(left, right, paragraphId, within, walk);
    return;
  }

  const formatting = within === 'formatting';
  if (
    left.kind !== right.kind ||
    left.namespaceUri !== right.namespaceUri ||
    left.localName !== right.localName
  ) {
    const textFlowOnly =
      formatting ||
      (isTextFlowSubtree(left, walk) && isTextFlowSubtree(right, walk) && !walk.undecided);
    recordDifference(walk, paragraphId, textOrContent(textFlowOnly));
    return;
  }

  // Two prefix sets can resolve every QName the same way and still differ as declarations.
  // Proving that costs more than the walk saves, so an unmatched pair goes to the oracle.
  if (!namespaceDeclarationsMatch(left, right)) {
    walk.undecided = true;
    return;
  }

  const scope = left.kind === 'paragraph' ? right.id : paragraphId;
  const bindings = extendedBindings(inheritedBindings, left);
  const leftAttributes = normalizedFingerprintAttributes(left);
  const rightAttributes = normalizedFingerprintAttributes(right);
  if (!fingerprintAttributesMatch(leftAttributes, rightAttributes, bindings, left)) {
    // Only this element's own attributes moved, so its subtree does not need classifying.
    recordDifference(walk, scope, textOrContent(formatting || isTextFlowElement(left)));
    return;
  }

  const preserve = inheritedPreserveOf(leftAttributes, inheritedPreserve);
  diffChildren(left, right, preserve, bindings, scope, within, walk);
}

/** How `canonicalTreeDifference` walks. */
export interface CanonicalTreeDifferenceOptions {
  /** Node pairs the walk may compare before it gives up (`undecided`). */
  readonly budget?: number;
  /**
   * The namespace bindings in scope above both roots, for two subtrees of parts
   * (`inheritedNamespaceBindings`). Word declares `wps` on `w:document` and names it in an
   * `mc:Choice`'s `Requires` deep inside a paragraph: only these resolve such a prefix. Both
   * roots must sit under this one scope. Two parts need none.
   */
  readonly bindings?: ReadonlyMap<string, string> | undefined;
}

/**
 * Where two canonical trees differ, answered in the size of the difference.
 *
 * This agrees with `ooxmlTreesEqual` on the verdict whenever it returns a decided result, and
 * it reports which paragraphs own the difference so a caller does not have to declare that
 * everything changed. The editor installs a received package only when it says something new,
 * and tells layout which paragraphs changed, so a received keystroke costs one paragraph, not
 * the document. Two subtrees compare under the scope `options.bindings` gives them.
 */
export function canonicalTreeDifference(
  previous: OoxmlPart | OoxmlNode,
  next: OoxmlPart | OoxmlNode,
  options: CanonicalTreeDifferenceOptions = {}
): CanonicalTreeDifference {
  const walk = newWalk(options.budget ?? DEFAULT_DIFFERENCE_BUDGET);
  diffNodes(
    'root' in previous ? previous.root : previous,
    'root' in next ? next.root : next,
    false,
    options.bindings ?? DEFAULT_FINGERPRINT_BINDINGS,
    null,
    'content',
    walk
  );

  // One id on both lists moved between containers: no paragraph-level answer describes that
  for (const id of walk.created) {
    if (walk.deleted.has(id)) walk.reach = 'outside-paragraph';
  }

  return {
    equal: !walk.undecided && REACH_RANK[walk.reach] === 0,
    undecided: walk.undecided,
    reach: walk.reach,
    paragraphIds: [...walk.paragraphIds],
    idsPreserved: walk.idsPreserved,
    renamedParagraphIds: [...walk.renamedParagraphIds],
    renamedOutsideParagraphs: walk.renamedOutsideParagraphs,
    renamedReach: walk.renamedReach,
    createdParagraphIds: [...walk.created],
    deletedParagraphIds: [...walk.deleted],
    visited: walk.visited,
  };
}
