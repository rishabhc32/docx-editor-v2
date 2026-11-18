// The paragraphs a write reached, found in the list and in the part the editor was last given,
// and replaced there (design D7, task 6.8).

import { canonicalTreeDifference } from '../../../store/package/ooxml-difference.ts';
import {
  findNode,
  parentNodeOf,
  replaceChildren,
  replaceNode,
} from '../../../store/package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { inheritedNamespaceBindings } from '../../../store/package/wml-namespace.ts';
import { composeOrderedParagraph } from '../compose/compose.ts';
import { ComposeContext, type OrderedRetainedPrompt } from '../compose/context.ts';
import { hideBlockPrompt } from '../compose/paragraph.ts';
import { showsPlaceholder, withChildren } from '../ooxml-nodes.ts';
import { isHidden, itemAt } from '../model/document.ts';
import {
  BLOCK_TOKENS,
  type OrderedDocument,
  type OrderedItem,
  type OrderedItemAddress,
  type OrderedStory,
} from '../model/types.ts';

/** The part with the written paragraphs composed again, and the prompts they hide. */
export interface RegionPublish {
  readonly part: OoxmlPart;
  readonly prompts: readonly OrderedRetainedPrompt[];
}

/** One region: shown marks `from` to `to` (story indexes), and the hidden marks inside it. */
interface Region {
  readonly from: number;
  readonly to: number;
  readonly hidden: readonly OrderedItem[];
}

/** The node id of a mark's paragraph. */
function paragraphIdOf(mark: OrderedItem): string | null {
  const shell = mark.node;
  return shell && shell.kind !== 'textValue' ? shell.id : null;
}

/** The `index`-th shown mark of the story, or null past its end. */
function shownMark(story: OrderedStory, index: number): OrderedItem | null {
  if (index < 0) return null;
  const hit = story.tree.seek((s) => s.marks, index);
  return hit && hit.item.kind === 'mark' ? hit.item : null;
}

/** An item that starts or ends a block container (a table, row, cell or block control). */
const blockEdge = (item: OrderedItem): boolean =>
  ((item.kind === 'open' || item.kind === 'close') &&
    item.token !== undefined &&
    BLOCK_TOKENS.has(item.token)) ||
  (item.kind === 'zero' && item.block === true);

/** No block edge between two marks of the story: their paragraphs share one container. */
function sameContainer(story: OrderedStory, from: OrderedItem, to: OrderedItem): boolean {
  if (from === to) return true;
  for (let item = story.tree.next(from); item && item !== to; item = story.tree.next(item))
    if (blockEdge(item)) return false;
  return true;
}

/**
 * The part with the paragraphs the `seeds` (paragraph marks, shown or hidden, old or new)
 * reach composed again from the list. Null when a region cannot be found in both (a region that
 * crosses a table edge or a control, a paragraph the part does not hold where the list says, a
 * control a decision emptied): the caller then composes the whole story.
 */
export function withParagraphsComposed(
  doc: OrderedDocument,
  story: OrderedStory,
  part: OoxmlPart,
  seeds: readonly OrderedItemAddress[]
): RegionPublish | null {
  const marks = new Set<OrderedItem>();
  for (const address of seeds) {
    const mark = itemAt(story, address);
    if (!mark || mark.kind !== 'mark' || mark.offset !== address.offset) return null;
    marks.add(mark);
  }

  const prompts: OrderedRetainedPrompt[] = [];
  const covered = new Set<OrderedItem>();
  let current = part;
  for (const seed of marks) {
    if (covered.has(seed)) continue;

    // A paragraph the part shows and no seed names is unchanged: the region stops there
    const clean = (mark: OrderedItem): boolean => {
      const id = paragraphIdOf(mark);
      return !marks.has(mark) && id !== null && findNode(current, id)?.kind === 'paragraph';
    };

    // A paragraph the part does not show yet (Enter, or a join taken back) took its words from
    // the paragraph after it, which changed though no seed names it
    const shows = (mark: OrderedItem): boolean => {
      const id = paragraphIdOf(mark);
      return id !== null && findNode(current, id)?.kind === 'paragraph';
    };
    const region = regionOf(doc, story, seed, marks, clean, shows);
    if (!region) return null;
    for (let k = region.from; k <= region.to; k += 1) {
      const mark = shownMark(story, k);
      if (mark) covered.add(mark);
    }

    for (const mark of region.hidden) covered.add(mark);
    const next = replaced(doc, story, current, region, prompts);
    if (!next) return null;
    current = next;
  }

  return { part: current, prompts };
}

/**
 * The shown marks around a seed that are not clean, and the hidden seeds among them. A split or
 * a join reaches the paragraphs around it: the region runs to the nearest paragraphs on both
 * sides that the part shows unchanged.
 */
function regionOf(
  doc: OrderedDocument,
  story: OrderedStory,
  seed: OrderedItem,
  seeds: ReadonlySet<OrderedItem>,
  clean: (mark: OrderedItem) => boolean,
  shows: (mark: OrderedItem) => boolean
): Region | null {
  // A hidden mark's paragraph now ends at the next shown mark
  const at = story.tree.before(seed).marks;
  if (!shownMark(story, at)) return null;
  let from = at;
  let to = at;
  for (
    let prev = shownMark(story, from - 1);
    prev && !clean(prev);
    prev = shownMark(story, from - 1)
  )
    from -= 1;
  for (
    let next = shownMark(story, to + 1);
    next && (!clean(next) || !shows(shownMark(story, to) ?? next));
    next = shownMark(story, to + 1)
  )
    to += 1;
  const first = shownMark(story, from);
  const last = shownMark(story, to);
  if (!first || !last || !sameContainer(story, first, last)) return null;
  const hidden: OrderedItem[] = [];
  for (const mark of seeds) {
    if (!isHidden(doc, mark)) continue;
    const index = story.tree.before(mark).marks;
    if (index >= from && index <= to) hidden.push(mark);
  }

  return { from, to, hidden };
}

/**
 * The part with a region's paragraphs replaced by the list's, or null when they do not line
 * up.
 */
function replaced(
  doc: OrderedDocument,
  story: OrderedStory,
  part: OoxmlPart,
  region: Region,
  prompts: OrderedRetainedPrompt[]
): OoxmlPart | null {
  const old: OoxmlNode[] = [];
  const shown: OrderedItem[] = [];
  for (let k = region.from; k <= region.to; k += 1) {
    const mark = shownMark(story, k);
    if (!mark) {
      return null;
    }

    shown.push(mark);
  }

  for (const mark of [...shown, ...region.hidden]) {
    const id = paragraphIdOf(mark);
    const node = id === null ? null : findNode(part, id);
    if (node) {
      old.push(node);
    }
  }

  if (old.length === 0) {
    return null;
  }

  const fresh = freshNodes(doc, story, part, region.from, shown, prompts);
  return fresh === null ? null : spliced(part, old, fresh);
}

/**
 * The list's paragraphs for the region's shown marks, from story index `from`, and the prompts
 * they hide. Null when one does not compose alone.
 */
function freshNodes(
  doc: OrderedDocument,
  story: OrderedStory,
  part: OoxmlPart,
  from: number,
  shown: readonly OrderedItem[],
  prompts: OrderedRetainedPrompt[]
): OoxmlNode[] | null {
  const fresh: OoxmlNode[] = [];
  for (const [k, mark] of shown.entries()) {
    const composed = composeOrderedParagraph(doc, story.name, from + k);
    if (!composed || composed.view.mark !== mark) {
      return null;
    }

    if (composed.emptied.length > 0 || composed.reached.length > 0) {
      return null;
    }

    // A paragraph that shows nothing in a block control: a decision may have emptied the
    // control, whose prompt only the whole story's drawing shows again (`blocks.ts`)
    if (composed.view.chars.length === 0 && composed.view.blocks.some(isBlockControl)) {
      return null;
    }

    if (composed.view.blocks.some((block) => promptFlips(doc, story, part, block))) {
      return null;
    }

    prompts.push(...composed.prompts);
    fresh.push(keptNode(part, composed.view.id, composed.node));
  }

  return fresh;
}

/**
 * The node a composed paragraph shows as: what the part shows already keeps its node, so the
 * editor sees nothing new there. The two compare under the part's scope: a shape's
 * `mc:Choice` names prefixes of the root.
 */
function keptNode(part: OoxmlPart, id: string, composed: OoxmlNode): OoxmlNode {
  const before = findNode(part, id);
  const bindings = before ? inheritedNamespaceBindings(part, before) : undefined;
  if (!before || before === composed || !bindings) {
    return composed;
  }

  if (canonicalTreeDifference(before, composed, { bindings }).equal) {
    return before;
  }

  return withEqualChildrenKept(before, composed, bindings);
}

/** The part with `old`, children side by side of one parent, replaced by `fresh`. */
function spliced(
  part: OoxmlPart,
  old: readonly OoxmlNode[],
  fresh: readonly OoxmlNode[]
): OoxmlPart | null {
  const parent = parentNodeOf(part, old[0]!.id);
  if (!parent) {
    return null;
  }

  const range = contiguous(parent, old);
  if (!range) {
    return null;
  }

  const [only] = old;
  const [single] = fresh;
  if (old.length === 1 && fresh.length === 1 && only && single) {
    if (only === single) {
      return part;
    }

    const edited = replaceNode(part, only.id, single, { deferValidation: true });
    return edited.ok ? edited.part : null;
  }

  const children = [
    ...parent.children.slice(0, range.start),
    ...fresh,
    ...parent.children.slice(range.end + 1),
  ];
  const edited = replaceChildren(part, parent.id, children, { deferValidation: true });
  return edited.ok ? edited.part : null;
}

/**
 * `fresh` with each child that says what the child of `before` under the same id says kept
 * as that child: the paragraph's properties and the runs a key did not reach. The editor's
 * store sees by identity what stayed, so a key that merges two insertions into one is still
 * text in one paragraph, not a change to the paragraph's structure.
 */
function withEqualChildrenKept(
  before: OoxmlNode,
  fresh: OoxmlNode,
  bindings: ReadonlyMap<string, string>
): OoxmlNode {
  if (before.kind === 'textValue' || fresh.kind === 'textValue') return fresh;
  const old = new Map<string, OoxmlNode>();
  for (const child of before.children) old.set(child.id, child);
  let kept = false;
  const children = fresh.children.map((child) => {
    const earlier = old.get(child.id);
    if (!earlier || earlier === child) return child;
    if (!canonicalTreeDifference(earlier, child, { bindings }).equal) return child;
    kept = true;
    return earlier;
  });
  return kept ? withChildren(fresh, children) : fresh;
}

const isBlockControl = (block: OrderedItem): boolean => block.token === 'sdtBlock';

/**
 * Whether a block control around the region shows its prompt in the list and not in the part,
 * or the other way round: its `w:sdtPr` changes (the prompt's flag, the glossary block that
 * keeps the prompt), and the control sits outside the region, so its story is drawn whole.
 */
function promptFlips(
  doc: OrderedDocument,
  story: OrderedStory,
  part: OoxmlPart,
  block: OrderedItem
): boolean {
  if (!isBlockControl(block)) {
    return false;
  }

  const shell = block.node;
  const shown = shell && shell.kind !== 'textValue' ? findNode(part, shell.id) : null;
  if (!shown || shown.kind === 'textValue') return block.prompt === true;
  const flagged = showsPlaceholder(shown);

  // A control whose prompt went for good (a direct emptying taken back) shows none any more
  if (block.prompt !== true) return flagged;
  const ctx = new ComposeContext(doc, story, {});
  hideBlockPrompt(ctx, block);
  return flagged === ctx.unprompted.has(block.space);
}

/** Where `nodes` sit in `parent`, when they are its children side by side and nothing else. */
function contiguous(
  parent: OoxmlElement,
  nodes: readonly OoxmlNode[]
): { readonly start: number; readonly end: number } | null {
  const wanted = new Set(nodes);
  let start = -1;
  let end = -1;
  let found = 0;
  for (const [index, child] of parent.children.entries()) {
    if (!wanted.has(child)) continue;
    if (start < 0) start = index;
    end = index;
    found += 1;
  }

  if (found !== wanted.size || end - start + 1 !== found) return null;
  return { start, end };
}
