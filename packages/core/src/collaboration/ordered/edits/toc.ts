// A table of contents (group 9.5): a new one before a body paragraph, and its result made again.

import type { OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import {
  buildTocContentControl,
  buildTocEntryParagraph,
  type TocEntryPlan,
} from '../../../store/package/toc-build.ts';
import {
  instrTextValue,
  isDeletedInstrText,
  isInstrText,
} from '../../../store/package/field-nodes.ts';
import {
  parseTocInstruction,
  TOC_MAX_INSTRUCTION_CHARS,
  type TocInstruction,
} from '../../../store/package/toc-instruction.ts';
import { withHeadingSymbols } from '../../../store/store/toc-heading-symbols.ts';
import { composeOrderedParagraph } from '../compose/compose.ts';
import { withStoryBlocks } from '../compose/paragraph-ops.ts';
import { createOrderedItem, isHidden } from '../model/document.ts';
import { insertItems, relabel } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { BLOCK_TOKENS, type OrderedItem, type OrderedItemAddress } from '../model/types.ts';
import { attributeOf, isWml, withChildren } from '../ooxml-nodes.ts';
import { firstItemOf } from '../placement.ts';
import { suggestionFor, type Ctx } from './context.ts';
import { settle } from './direct.ts';
import { bookmarkParagraph } from './fields.ts';
import { resolveOrderedPlace, type OrderedResolvedPlace } from './places.ts';
import { templateItems } from './templates.ts';

/**
 * One row of a table of contents: its outline level (0 to 8), its text, the heading it points
 * to (the heading paragraph's mark), the heading's bookmark and the page number it shows. The
 * change carries them, computed once by the editor that made it from its own layout, so every
 * copy writes the same rows and never lays the document out to compute them.
 *
 * @public
 */
export interface OrderedTocRow {
  readonly level: number;
  readonly text: string;
  readonly heading: OrderedItemAddress;
  readonly bookmark: string;
  readonly page: string;
}

/**
 * A bookmark the edit adds around a whole heading paragraph (named by its mark), for the rows
 * that point to it.
 *
 * @public
 */
export interface OrderedTocBookmark {
  readonly heading: OrderedItemAddress;
  readonly name: string;
}

/** What both edits share: the field's instruction, its rows and the heading bookmarks. */
export interface OrderedTocContent {
  readonly instruction: string;
  readonly rows: readonly OrderedTocRow[];
  readonly bookmarks: readonly OrderedTocBookmark[];
}

/** A change's TOC field instruction, parsed; refused when it is not one. */
function tocInstruction(instruction: string): TocInstruction {
  const toc = parseTocInstruction(instruction);
  if (!toc) throw new OrderedRefusal('invalid-edit', 'a TOC field instruction');
  return toc;
}

/**
 * Refuses a suggested edit: the engine has no tracked form for a table of contents, and the
 * host's rule refuses a party's untracked change.
 */
function refuseTracked(ctx: Ctx): void {
  if (!ctx.direct) {
    throw new OrderedRefusal('invalid-edit', 'a table of contents has no tracked form');
  }
}

const heading = (ctx: Ctx, mark: OrderedItemAddress): OrderedResolvedPlace =>
  resolveOrderedPlace(ctx.doc, ctx.story.name, mark);

/**
 * A row as the engine's builder reads it, with its heading's symbols, read from the heading
 * paragraph as the list shows it, as the engine reads them from its tree.
 */
function entryPlan(ctx: Ctx, row: OrderedTocRow): TocEntryPlan {
  const drawn = composeOrderedParagraph(ctx.doc, ctx.story.name, heading(ctx, row.heading).index);
  if (!drawn) throw new OrderedRefusal('hidden-target', row.heading.space.slice(0, 160));
  const plan: TocEntryPlan = {
    level: row.level,
    text: row.text,
    headingParagraphId: drawn.node.id,
    bookmarkName: row.bookmark,
    pageNumberText: row.page,
  };
  return withHeadingSymbols(withStoryBlocks(ctx.story.part, [drawn.node]), plan);
}

/** Add the heading bookmarks, as the engine does before it builds the rows. */
function addBookmarks(ctx: Ctx, content: OrderedTocContent, suggestion: string): void {
  for (const bookmark of content.bookmarks)
    bookmarkParagraph(ctx, heading(ctx, bookmark.heading), bookmark.name, suggestion);
}

/** Items for blocks the engine built: every one is the suggestion's insertion. */
function insertedItems(ctx: Ctx, nodes: readonly OoxmlNode[], suggestion: string): OrderedItem[] {
  return nodes.flatMap((node) => {
    if (node.kind === 'textValue') throw new OrderedRefusal('engine-refused', 'no block');
    return templateItems(ctx, node, suggestion).map((item) =>
      createOrderedItem({ ...item, by: suggestion })
    );
  });
}

/**
 * A table of contents in a block control before a body paragraph, with its first result and
 * its heading bookmarks, as the engine's `insertToc` builds it.
 */
export function insertToc(
  ctx: Ctx,
  before: OrderedResolvedPlace,
  alias: string,
  content: OrderedTocContent
): string {
  refuseTracked(ctx);
  const toc = tocInstruction(content.instruction);
  const p = before.paragraph;
  if (ctx.story.part.root.localName !== 'document' || p.blocks.length > 0)
    throw new OrderedRefusal('invalid-edit', "a table of contents goes between the body's blocks");
  const suggestion = suggestionFor(ctx);
  const plans = content.rows.map((row) => entryPlan(ctx, row));
  addBookmarks(ctx, content, suggestion.id);
  const control = buildTocContentControl(() => ctx.ids.node(), plans, toc, alias);
  const items = insertedItems(ctx, [control], suggestion.id);
  insertItems(ctx.doc, ctx.story, ctx.story.tree.prev(firstItemOf(p)), items);
  settle(ctx, suggestion.id);
  return suggestion.id;
}

const fieldCharType = (item: OrderedItem): string | undefined => {
  const node = item.node;
  return item.kind === 'zero' && node && node.kind !== 'textValue' && isWml(node, 'fldChar')
    ? attributeOf(node, 'fldCharType')
    : undefined;
};

/** Whether a container's edge counts: one that Enter made counts only while it shows. */
const edgeCounts = (ctx: Ctx, item: OrderedItem): boolean =>
  (item.kind === 'open' || item.kind === 'close') &&
  (item.space === item.pair || !isHidden(ctx.doc, item));

/**
 * The instruction of the field whose separator this is, read back to its begin: the live text,
 * or the deleted text when no live text is left. Null when no begin is before it.
 */
function instructionBefore(ctx: Ctx, separate: OrderedItem): string | null {
  const tree = ctx.story.tree;
  let depth = 0;
  let live = '';
  let deleted = '';
  for (let at = tree.prev(separate); at; at = tree.prev(at)) {
    if (isHidden(ctx.doc, at)) continue;
    const type = fieldCharType(at);
    if (type === 'end') depth += 1;
    else if (type === 'begin' && depth === 0) return live || deleted;
    else if (type === 'begin') depth -= 1;
    // A separator of its own field: this separator is no field's own
    else if (type === 'separate' && depth === 0) return null;
    const node = at.node;
    if (depth > 0 || at.kind !== 'zero' || !node || !isInstrText(node)) continue;
    const text = instrTextValue(node);
    if (isDeletedInstrText(node)) deleted = text + deleted;
    else live = text + live;
    if (live.length + deleted.length > TOC_MAX_INSTRUCTION_CHARS) return null;
  }

  return null;
}

/**
 * Refuses a separator and an end that are not one TOC field's own pair, or whose result
 * crosses the edge of an inline container. The rows are paragraphs that go right after the
 * separator, so a separator inside a link, or an end inside another field or a link, leaves a
 * container's open item in one paragraph and its close item in another: a part no copy
 * composes. A copy refuses that change, so the room never orders it.
 */
function refuseOtherField(ctx: Ctx, separate: OrderedItem, end: OrderedItem): void {
  const instruction = instructionBefore(ctx, separate);
  if (instruction === null || !parseTocInstruction(instruction))
    throw new OrderedRefusal('not-one-field', 'the separator of no TOC field');
  refuseSeparatorInContainer(ctx, separate);
  refuseEndOutsideResult(ctx, separate, end);
}

/** Refuses a separator inside a container: back to the paragraph's start, each one closes. */
function refuseSeparatorInContainer(ctx: Ctx, separate: OrderedItem): void {
  const tree = ctx.story.tree;
  const closed: string[] = [];
  for (let at = tree.prev(separate); at; at = tree.prev(at)) {
    if (at.kind === 'mark' && !isHidden(ctx.doc, at)) break;
    if (at.token !== undefined && BLOCK_TOKENS.has(at.token)) break;
    if (!edgeCounts(ctx, at)) continue;
    if (at.kind === 'close') closed.push(at.pair ?? '');
    else if (closed.pop() !== at.pair)
      throw new OrderedRefusal('not-one-field', 'the separator inside a container');
  }
}

/**
 * Refuses an end that is not the first end of no field after the separator, or a result that
 * does not hold every container in it whole.
 */
function refuseEndOutsideResult(ctx: Ctx, separate: OrderedItem, end: OrderedItem): void {
  const tree = ctx.story.tree;
  const open: string[] = [];
  let depth = 0;
  for (let at = tree.next(separate); at; at = tree.next(at)) {
    if (edgeCounts(ctx, at)) {
      if (at.kind === 'open') open.push(at.pair ?? '');
      else if (open.pop() !== at.pair)
        throw new OrderedRefusal('not-one-field', 'the result crosses the edge of a container');
      continue;
    }

    if (isHidden(ctx.doc, at)) continue;
    const type = fieldCharType(at);
    if (type === 'begin') depth += 1;
    else if (type === 'end' && depth > 0) depth -= 1;
    else if (type === 'end') {
      if (at !== end) throw new OrderedRefusal('not-one-field', 'the end of another field');
      if (open.length > 0) throw new OrderedRefusal('not-one-field', 'the end inside a container');
      return;
    }

    if (at === end) throw new OrderedRefusal('not-one-field', 'the end of another field');
  }

  throw new OrderedRefusal('not-one-field', 'no end after the separator');
}

/** A result paragraph's properties under a contents style (`TOC1` to `TOC9`), by style. */
function rowTemplates(marks: readonly OrderedItem[]): Map<string, OoxmlNode> {
  const out = new Map<string, OoxmlNode>();
  for (const mark of marks) {
    const shell = mark.node;
    const pPr =
      shell && shell.kind !== 'textValue' ? shell.children.find((c) => isWml(c, 'pPr')) : undefined;
    if (!pPr || pPr.kind === 'textValue') continue;
    const pStyle = pPr.children.find((c) => isWml(c, 'pStyle'));
    const style = pStyle && pStyle.kind !== 'textValue' ? attributeOf(pStyle, 'val') : undefined;
    if (!style?.startsWith('TOC') || out.has(style)) continue;

    // A section's properties stay with the paragraph that ends the section
    out.set(
      style,
      withChildren(
        pPr,
        pPr.children.filter((c) => !isWml(c, 'sectPr'))
      )
    );
  }

  return out;
}

/**
 * The old result's shown items between the separator and the end, and the shown marks among
 * them with the first one after the end.
 */
function oldResult(
  ctx: Ctx,
  separate: OrderedItem,
  end: OrderedItem
): { readonly old: OrderedItem[]; readonly marks: OrderedItem[] } {
  const tree = ctx.story.tree;
  const old: OrderedItem[] = [];
  const marks: OrderedItem[] = [];
  let at = tree.next(separate);
  for (; at && at !== end; at = tree.next(at)) {
    if (at.token !== undefined && BLOCK_TOKENS.has(at.token))
      throw new OrderedRefusal('invalid-edit', 'a table of contents result of paragraphs');
    if (isHidden(ctx.doc, at)) continue;
    if (at.kind === 'mark') marks.push(at);
    old.push(at);
  }

  if (at !== end) throw new OrderedRefusal('invalid-edit', 'the separator comes before the end');
  for (let after = tree.next(end); after; after = tree.next(after)) {
    if (after.kind === 'mark' && !isHidden(ctx.doc, after)) {
      marks.push(after);
      break;
    }
  }

  return { old, marks };
}

/**
 * A table of contents' result made again, as the engine's `replaceTocResult` makes it: every
 * unit between the field's separator and its end is struck, and the new rows go right after the
 * separator, the first one in the separator's paragraph. The engine treats a result as
 * generated and read-only, so no one's suggestion is inside it. The organiser's undo (`revert`)
 * shows the old result again exactly.
 */
export function replaceTocResult(
  ctx: Ctx,
  separate: OrderedItem,
  end: OrderedItem,
  content: OrderedTocContent
): string {
  refuseTracked(ctx);
  const toc = tocInstruction(content.instruction);
  if (fieldCharType(separate) !== 'separate' || fieldCharType(end) !== 'end')
    throw new OrderedRefusal('invalid-edit', "a field's separator and its end");
  for (const item of [separate, end])
    if (isHidden(ctx.doc, item))
      throw new OrderedRefusal('hidden-target', item.space.slice(0, 160));
  refuseOtherField(ctx, separate, end);
  const { old, marks } = oldResult(ctx, separate, end);
  const suggestion = suggestionFor(ctx);
  const plans = content.rows.map((row) => entryPlan(ctx, row));
  addBookmarks(ctx, content, suggestion.id);

  // With no rows the separator's paragraph stays, as the engine keeps it before the end's
  const kept = plans.length === 0 && marks[0] !== marks[marks.length - 1] ? marks[0] : undefined;
  for (const item of old)
    if (item !== kept) relabel(ctx.doc, ctx.story, item, { dels: [...item.dels, suggestion.id] });
  const templates = rowTemplates(marks);
  const rows = plans.map((plan) =>
    buildTocEntryParagraph(
      () => ctx.ids.node(),
      plan,
      toc,
      templates.get(`TOC${Math.min(plan.level + 1, 9)}`)
    )
  );
  insertItems(ctx.doc, ctx.story, separate, insertedItems(ctx, rows, suggestion.id));
  settle(ctx, suggestion.id);
  return suggestion.id;
}
