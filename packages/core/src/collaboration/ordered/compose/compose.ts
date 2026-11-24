// Items back to Word: a whole story, its part, or one paragraph.

import { parentNodeOf, findNode } from '../../../store/package/ooxml-edit.ts';
import type { OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { isHidden, itemAt, statusOf, storyNamed } from '../model/document.ts';
import {
  tokenSpace,
  type OrderedDocument,
  type OrderedItem,
  type OrderedStory,
} from '../model/types.ts';
import { composeBlocks, isBlockToken, isLeftOut, joinsNothing, showsPastJoin } from './blocks.ts';
import {
  ComposeContext,
  type ComposeOptions,
  type HeldTakeovers,
  type OrderedComposeOptions,
  type OrderedRetainedPrompt,
  type OrderedViewParagraph,
} from './context.ts';
import { isNotesStory, notesWithHiddenReferences } from './note-references.ts';
import { withStoryBlocks } from './paragraph-ops.ts';
import { composeParagraph, hideBlockPrompt, viewOfParagraph } from './paragraph.ts';
import { withControlRecords, withRestoredPrompts } from './prompts.ts';
import { withCheckpointNamespace } from '../model/checkpoint-labels.ts';

/** A composed story. @public */
export interface OrderedComposedStory {
  /** The story's blocks: a body's children, or a header's, footer's or notes part's root's. */
  readonly children: readonly OoxmlNode[];
  /** The view: each shown paragraph and the units it shows. */
  readonly paragraphs: readonly OrderedViewParagraph[];
  /** Controls a decision emptied: their prompt shows again in the part. */
  readonly emptied: readonly string[];
  /** Controls a decision reached: their value record follows their text in the part. */
  readonly reached: readonly string[];
  /** Placeholder blocks the caller writes to the glossary: prompts the part no longer shows. */
  readonly prompts: readonly OrderedRetainedPrompt[];
}

/** One composed paragraph. @public */
export interface OrderedComposedParagraph {
  /** The paragraph node: the loaded node itself when untouched, with its id and `w14:paraId`. */
  readonly node: OoxmlNode;
  readonly view: OrderedViewParagraph;
  readonly emptied: readonly string[];
  readonly reached: readonly string[];
  readonly prompts: readonly OrderedRetainedPrompt[];
}

/**
 * Compose a story's blocks and view.
 *
 * @public
 */
export function composeOrderedStory(
  doc: OrderedDocument,
  name: string,
  options: OrderedComposeOptions = {}
): OrderedComposedStory {
  return composeStory(doc, name, options).composed;
}

/** A story's drawing: the composed story, and what the drawing leaves past it. */
interface StoryDrawing {
  readonly composed: OrderedComposedStory;
  /** The emptied controls a decision on a suggestion emptied. */
  readonly decisionEmptied: ReadonlySet<string>;
  /** Controls whose value record the words a reject took away stated. */
  readonly stale: ReadonlySet<string>;
  /** The takeovers each change record holds, by the record's node id. */
  readonly takeovers: ReadonlyMap<string, HeldTakeovers>;
  /** Whether the drawing put a checkpoint's attribute on an element (a checkpoint's only). */
  readonly marked: boolean;
}

/** A story's drawing. */
function composeStory(doc: OrderedDocument, name: string, options: ComposeOptions): StoryDrawing {
  const ctx = new ComposeContext(doc, storyNamed(doc, name), options);
  const { children, paragraphs } = composeBlocks(ctx);
  const composed = {
    children,
    paragraphs,
    emptied: ctx.emptied,
    reached: ctx.reached,
    prompts: ctx.prompts,
  };
  return {
    composed,
    decisionEmptied: ctx.decisionEmptied,
    stale: ctx.stale,
    takeovers: ctx.takeovers,
    marked: ctx.markedAny,
  };
}

/**
 * Compose a story's part: its blocks in the loaded part, emptied controls showing their
 * prompt again, and value records following the text. The glossary blocks in
 * `composed.prompts` are for the caller to write.
 *
 * @public
 */
export function composeOrderedPart(
  doc: OrderedDocument,
  name: string,
  options: OrderedComposeOptions = {}
): { readonly part: OoxmlPart; readonly composed: OrderedComposedStory } {
  const { part, composed } = composeStoryPart(doc, name, options);
  return { part, composed };
}

/** A story's part as {@link composeOrderedPart} composes it, and the takeovers its records hold. */
export function composeStoryPart(
  doc: OrderedDocument,
  name: string,
  options: ComposeOptions = {}
): {
  readonly part: OoxmlPart;
  readonly composed: OrderedComposedStory;
  readonly takeovers: ReadonlyMap<string, HeldTakeovers>;
} {
  const { composed, decisionEmptied, stale, takeovers, marked } = composeStory(doc, name, options);
  let part = withStoryBlocks(storyNamed(doc, name).part, composed.children);

  // A checkpoint's prefix bound before the edits below, which check the part
  if (marked) {
    part = { ...part, root: withCheckpointNamespace(part.root) };
  }

  // A prompt's records show their decisions as the rest of the drawing does
  const status = (suggestion: string) => {
    const now = statusOf(doc, suggestion);
    return now === 'pending' && options.pending === 'reject' ? 'rejected' : now;
  };
  part = withRestoredPrompts(doc, part, composed.emptied, decisionEmptied, status);
  part = withControlRecords(part, composed.reached, stale);
  return { part, composed, takeovers };
}

/** Shown paragraphs in a story: the tree's mark count. O(1). @public */
export function orderedParagraphCount(doc: OrderedDocument, name: string): number {
  return storyNamed(doc, name).tree.summary.marks;
}

/**
 * Compose one shown paragraph, the `index`-th of the story, as the story's compose draws it:
 * the mark is a seek in the tree (O(log n)), and its items are the ones back to the shown mark
 * before it. Pending suggestions show (the index counts the marks they show). Null past the end.
 *
 * @public
 */
export function composeOrderedParagraph(
  doc: OrderedDocument,
  name: string,
  index: number
): OrderedComposedParagraph | null {
  const found = paragraphAt(doc, storyNamed(doc, name), index);
  if (!found) return null;
  const { ctx, buffer, mark, blocks } = found;
  const drawn = composeParagraph(ctx, buffer, mark, blocks);
  return {
    node: drawn.node,
    view: drawn.view,
    emptied: ctx.emptied,
    reached: ctx.reached,
    prompts: ctx.prompts,
  };
}

/**
 * The view of one shown paragraph, the `index`-th of the story, without drawing its node: what
 * an edit reads to place itself. Null past the end.
 *
 * @public
 */
export function orderedViewParagraph(
  doc: OrderedDocument,
  name: string,
  index: number
): OrderedViewParagraph | null {
  const found = paragraphAt(doc, storyNamed(doc, name), index);
  if (!found) return null;
  return viewOfParagraph(found.ctx, found.buffer, found.mark, found.blocks).view;
}

/**
 * The `index`-th shown paragraph's mark, its items back to the shown mark before it, and its
 * blocks.
 */
function paragraphAt(
  doc: OrderedDocument,
  story: OrderedStory,
  index: number
): {
  ctx: ComposeContext;
  buffer: OrderedItem[];
  mark: OrderedItem;
  blocks: OrderedItem[];
} | null {
  const hit = story.tree.seek((s) => s.marks, index);
  if (!hit) return null;
  const mark = hit.item;
  const buffer: OrderedItem[] = [];
  const ctx = new ComposeContext(doc, story, {});
  let orphans: ReadonlySet<string> | undefined;
  const orphansOf = () => (orphans ??= notesWithHiddenReferences(doc));
  const notes = isNotesStory(story);

  // Past a mark that joins nothing, only what shows goes on (`composeBlocks`)
  let carried = false;
  for (let item = story.tree.prev(mark); item; item = story.tree.prev(item)) {
    if (item.kind === 'mark' && !isHidden(doc, item)) break;

    // A block the story's drawing leaves out takes its items with it (`composeBlocks`): a cell
    // taken back, a note whose references show nowhere
    if (item.kind === 'close' && isBlockToken(item) && item.pair !== undefined) {
      const open = itemAt(story, { space: item.pair, offset: 0 });
      if (open?.kind === 'open' && isLeftOut(ctx, open, () => (notes ? orphansOf() : null))) {
        item = open;
        continue;
      }
    }

    if (item.kind === 'mark' && joinsNothing(doc, item)) carried = true;
    if (isBlockToken(item) || (item.kind === 'zero' && item.block)) continue;
    if (carried && !showsPastJoin(ctx, item)) continue;
    buffer.push(item);
  }

  buffer.reverse();
  const blocks = enclosingBlocks(doc, story, mark);
  for (const block of blocks) if (block.prompt) hideBlockPrompt(ctx, block);
  return { ctx, buffer, mark, blocks };
}

/** The block containers (open items) around a paragraph's mark, outermost first. */
function enclosingBlocks(
  doc: OrderedDocument,
  story: OrderedStory,
  mark: OrderedItem
): OrderedItem[] {
  const out: OrderedItem[] = [];
  const home = mark.home;
  if (home !== undefined && findNode(story.part, home)) {
    // A loaded paragraph: its ancestors in the loaded part
    for (
      let node = parentNodeOf(story.part, home);
      node;
      node = parentNodeOf(story.part, node.id)
    ) {
      const open = story.spaces.get(tokenSpace(doc.file, node.id))?.[0];

      // A block control removed with its content kept no longer holds the paragraph
      if (open && open.kind === 'open' && isBlockToken(open) && !isHidden(doc, open))
        out.push(open);
    }

    return out.reverse();
  }

  // Otherwise walk back over the items, past closed containers
  let depth = 0;
  for (let item = story.tree.prev(mark); item; item = story.tree.prev(item)) {
    if (!isBlockToken(item) || isHidden(doc, item)) continue;
    if (item.kind === 'close') depth += 1;
    else if (depth > 0) depth -= 1;
    else out.push(item);
  }

  return out.reverse();
}
