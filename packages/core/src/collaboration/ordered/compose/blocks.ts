// A story's items back to its blocks: paragraphs, tables, rows, cells, block controls, notes.
// An untouched block is the loaded node itself; a touched one is drawn again from its open
// item's shell and its children.

import type { OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import { insertedPending, isLabelled } from '../model/document.ts';
import { isDroppedMark, isFilledPrompt } from '../model/prompt-marks.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { REVISION_ROLE } from '../model/revision-ids.ts';
import { BLOCK_TOKENS, type OrderedDocument, type OrderedItem } from '../model/types.ts';
import { isWml, withChildren } from '../ooxml-nodes.ts';
import type { ComposeContext, OrderedViewParagraph } from './context.ts';
import { controlContent, isValueRecord } from './inline.ts';
import { isNotesStory, noteKeyOf, notesWithHiddenReferences } from './note-references.ts';
import {
  composeParagraph,
  hideBlockPrompt,
  promptControlOf,
  type DrawnParagraph,
} from './paragraph.ts';
import { containerPropertiesAfter } from './properties.ts';
import { pendingHeightOf, withRowHeightAnchors } from './row-heights.ts';
import { resetUnrecordedCells } from './row-cells.ts';
import { bodySectionNode, isSectionItem } from './sections.ts';
import { hidePrompt } from './prompts.ts';
import {
  isRow,
  removedCellStub,
  removedColumnCell,
  withoutRemovedCells,
  type RemovedCell,
} from './table-columns.ts';

interface Builder {
  readonly open: OrderedItem | null;
  readonly children: OoxmlNode[];
  touched: boolean;
  /** A unit sits in it, shown or not. */
  hadUnits: boolean;
  /** A unit shows in it. */
  drawn: boolean;
  /** A decision on a suggestion took a unit in it away (`ComposeContext.decidedAway`). */
  decided: boolean;
  /** A decision changed whether a unit in it reads as text (`ComposeContext.decided`). */
  reached: boolean;
  /** A unit in it shows struck by a pending suggestion. */
  struck: boolean;
  /** A table's cells a column decision removed (`table-columns.ts`). */
  removed?: RemovedCell[];
  /** A row's drawn cells and their open items (`row-cells.ts`). */
  cells?: Map<OoxmlNode, OrderedItem>;
  /** A table's rows that show a pending height with no row record, to its suggestion. */
  heights?: Map<string, string>;
}

export const isBlockToken = (item: OrderedItem): boolean =>
  (item.kind === 'open' || item.kind === 'close') &&
  item.token !== undefined &&
  BLOCK_TOKENS.has(item.token);

/**
 * Whether the story's drawing leaves a block out with every item in it: it does not show, or it
 * is a note whose references show nowhere (`orphans`, read only for a note). A block control that
 * went with its content kept is not left out: its content draws in its place.
 */
export function isLeftOut(
  ctx: ComposeContext,
  open: OrderedItem,
  orphans: () => ReadonlySet<string> | null
): boolean {
  if (!ctx.gone(open)) {
    const note = noteKeyOf(open);
    return note !== undefined && orphans()?.has(note) === true;
  }

  return !(open.token === 'sdtBlock' && open.pair !== undefined);
}

/**
 * Whether a paragraph mark that does not show joins nothing: a prompt's paragraph that content
 * replaced (its control's last mark would join past the control's end, `change/settle-prompts.ts`),
 * or a later paragraph of a block control its emptying dropped. Only what shows, which a join
 * brought from before it, goes on into the next paragraph.
 */
export const joinsNothing = (doc: OrderedDocument, mark: OrderedItem): boolean =>
  mark.prompt === true || isDroppedMark(doc, mark);

/**
 * Whether an item before a mark that joins nothing goes on into the next paragraph: only what
 * shows does. A block prompt's unit shows by its control's rule, which the story's drawing
 * settles as it opens the control (`hideBlockPrompt`); one paragraph's walk back settles it here.
 */
export function showsPastJoin(ctx: ComposeContext, item: OrderedItem): boolean {
  const control = item.prompt === true ? promptControlOf(ctx.story, item) : null;
  if (control?.token === 'sdtBlock' && !ctx.gone(control)) hideBlockPrompt(ctx, control);
  return !ctx.gone(item);
}

/** Draw the whole story: its blocks and its view paragraphs. */
export function composeBlocks(ctx: ComposeContext): {
  children: OoxmlNode[];
  paragraphs: OrderedViewParagraph[];
} {
  const paragraphs: OrderedViewParagraph[] = [];
  const root = newBuilder(null);
  const stack: Builder[] = [root];
  const top = (): Builder => stack[stack.length - 1] ?? root;

  // Notes whose references show nowhere go with them
  const orphans = isNotesStory(ctx.story) ? notesWithHiddenReferences(ctx.doc) : null;
  let buffer: OrderedItem[] = [];
  let skip: string | null = null;

  // Block controls that went with their content kept (Word's "Remove content control"): their
  // content draws in their place, each item showing as its own labels say
  const unwrapped = new Set<string>();
  let deferred: OoxmlNode[] = [];
  for (const item of ctx.story.tree.items()) {
    if (skip !== null) {
      if (item.kind === 'close' && item.pair === skip) skip = null;
      continue;
    }

    if (isBlockToken(item) && item.kind === 'open') {
      if (item.token === 'sdtBlock' && item.pair !== undefined && ctx.gone(item)) {
        unwrapped.add(item.pair);
        for (const b of stack) b.touched = true;
        continue;
      }

      if (isLeftOut(ctx, item, () => orphans)) {
        skip = item.pair ?? null;
        for (const b of stack) b.touched = true;
        removedCell(ctx, item, stack);
        continue;
      }

      if (item.prompt) hideBlockPrompt(ctx, item);
      stack.push(newBuilder(item));
      continue;
    }

    if (isBlockToken(item)) {
      if (item.pair !== undefined && unwrapped.delete(item.pair)) continue;
      const builder = stack.pop();
      if (!builder?.open)
        throw new OrderedRefusal('malformed-story', 'a close item without its open item');
      closeBlock(ctx, builder, builder.open, top());
      continue;
    }

    if (item.kind === 'zero' && item.block) {
      if (ctx.gone(item)) continue;

      // A body section under its records
      const node = isSectionItem(item) ? bodySectionNode(ctx, item) : item.node;
      if (!node) continue;
      if (buffer.length > 0) deferred.push(node);
      else top().children.push(node);
      continue;
    }

    if (item.kind === 'mark' && ctx.gone(item)) {
      if (ctx.decidedAway(item)) for (const b of stack) b.decided = true;
      top().touched = true;

      // A paragraph that joins nothing draws nothing: only what shows goes on. A join: the
      // paragraph goes on into the next one
      if (joinsNothing(ctx.doc, item)) buffer = buffer.filter((x) => showsPastJoin(ctx, x));
      else buffer.push(item);
      continue;
    }

    if (item.kind === 'mark') {
      const blocks = stack.flatMap((b) => (b.open ? [b.open] : []));
      const drawn = composeParagraph(ctx, buffer, item, blocks);
      top().children.push(drawn.node, ...deferred);
      for (const b of stack) noteParagraph(b, drawn);
      paragraphs.push(drawn.view);
      buffer = [];
      deferred = [];
      continue;
    }

    buffer.push(item);
  }

  if (buffer.some((item) => !ctx.gone(item) && item.kind !== 'mark'))
    throw new OrderedRefusal(
      'malformed-story',
      `${ctx.story.name}: items after the last paragraph mark`
    );
  return { children: root.children, paragraphs };
}

/** A builder for the story's root (`open` null) or for a block. */
function newBuilder(open: OrderedItem | null): Builder {
  return {
    open,
    children: [],
    touched: open !== null && isLabelled(open),
    hadUnits: false,
    drawn: false,
    decided: false,
    reached: false,
    struck: false,
  };
}

/** A block's close: its element drawn into the builder around it, `into`. */
function closeBlock(ctx: ComposeContext, builder: Builder, open: OrderedItem, into: Builder): void {
  // A table whose every row is gone goes too (the engine removes it)
  if (open.token === 'tbl' && !builder.children.some(isRow)) {
    into.touched = true;
    return;
  }

  if (open.token === 'tr' && resetUnrecordedCells(ctx, open, builder.children, builder.cells))
    builder.touched = true;
  let element = blockElement(ctx, builder);
  if (open.token === 'tc') (into.cells ??= new Map()).set(element, open);
  const height = open.token === 'tr' ? pendingHeightOf(ctx, open) : undefined;
  if (height !== undefined) (into.heights ??= new Map()).set(element.id, height);
  if (open.token === 'tbl') element = withRowHeightAnchors(ctx, element, builder.heights);
  into.children.push(
    builder.removed && element.kind !== 'textValue'
      ? withoutRemovedCells(ctx, element, builder.removed)
      : element
  );
  if (builder.touched) into.touched = true;
}

/**
 * What a paragraph drawn in a block tells it: touched, held units, drew units, decided, reached,
 * struck.
 */
function noteParagraph(builder: Builder, drawn: DrawnParagraph): void {
  if (drawn.touched) builder.touched = true;
  if (drawn.hadUnits) builder.hadUnits = true;
  if (drawn.drawn) builder.drawn = true;
  if (drawn.decided) builder.decided = true;
  if (drawn.reached) builder.reached = true;
  if (drawn.struck) builder.struck = true;
}

/** A cell a column decision removed: drawn for its table to remove by the engine's rule. */
function removedCell(ctx: ComposeContext, item: OrderedItem, stack: readonly Builder[]): void {
  const row = stack[stack.length - 1];
  const table = stack[stack.length - 2];
  if (item.token !== 'tc' || row?.open?.token !== 'tr' || !table?.open) return;
  if (table.open.token !== 'tbl' || !ctx.gone(item)) return;
  const removed = removedColumnCell(ctx, item, table.open);
  const stub = removed ? removedCellStub(ctx, item) : null;
  if (!removed || !stub) return;
  row.children.push(stub);
  (table.removed ??= []).push({ ...removed, id: stub.id });
}

function blockElement(ctx: ComposeContext, builder: Builder): OoxmlNode {
  const open = builder.open;
  const shell = open?.node;
  if (!open || !shell || shell.kind === 'textValue')
    throw new OrderedRefusal('malformed-story', 'a block without its shell');
  const original = ctx.doc.originals.get(shell.id);
  if (!builder.touched && !ctx.options.rebuild && original) return original;
  let props = containerPropertiesAfter(ctx, open, shell.children);
  if (open.token === 'tr') props = withMarker(ctx, open, props, 'trPr', ['ins', 'del']);
  if (open.token === 'tc') props = withMarker(ctx, open, props, 'tcPr', ['cellIns', 'cellDel']);
  if (open.token === 'sdtBlock') {
    // A block control a decision emptied shows its prompt again
    // One whose prompt Enter removed for good shows a prompt again only after a decision, as
    // the engine's: Enter itself leaves it empty, with no prompt (policy 4)
    // A decision that takes a paragraph of an empty control away empties it too, as the engine's
    const removed = ctx.promptRemoved.has(open.space);
    const prompts = removed ? builder.decided : !ctx.unprompted.has(open.space);
    const emptied = (builder.hadUnits || builder.decided) && !builder.drawn && prompts;
    if (emptied) {
      ctx.emptied.push(shell.id);
      if (builder.decided) ctx.decisionEmptied.add(shell.id);
    }

    // Its value record follows what it shows, as an inline control's (`closeBox`)
    const valued = !builder.struck && open.props.some(isValueRecord);
    if (builder.reached || emptied || valued) ctx.reached.push(shell.id);

    // Enter removes a prompt for good, with no glossary block, unless a suggestion had filled it
    // before and its words went for good: the engine kept it in the glossary then
    const retain = !removed || isFilledPrompt(ctx.doc, open.space);
    if (ctx.unprompted.has(open.space)) props = hidePrompt(ctx, props, original, { retain });
    else if (ctx.filledPrompts.has(open.space))
      props = hidePrompt(ctx, props, original, { hide: false });
    return withChildren(shell, [...props, controlContent(ctx, original, builder.children)]);
  }

  return withChildren(shell, [...props, ...builder.children]);
}

/**
 * The suggestion a row's or a cell's insertion marker names. A cell of a new row shows the row's
 * marker, though the row's marker covers its content.
 */
function insertedMarkerBy(
  ctx: ComposeContext,
  open: OrderedItem,
  container: 'trPr' | 'tcPr'
): string | null {
  if (container !== 'tcPr' || ctx.options.pending === 'reject') return ctx.insertedBy(open);
  if (insertedPending(ctx.doc, open)) return open.by;
  return null;
}

/** A row's or a cell's own revision markers in its properties. */
function withMarker(
  ctx: ComposeContext,
  open: OrderedItem,
  props: readonly OoxmlNode[],
  container: 'trPr' | 'tcPr',
  [insertName, deleteName]: readonly [string, string]
): readonly OoxmlNode[] {
  const by = insertedMarkerBy(ctx, open, container);
  const del = ctx.struckBy(open);
  if (!by && !del) return props;
  const markers: OoxmlNode[] = [];
  if (by)
    markers.push(
      build(ctx.mint(), 'generic', insertName, ctx.revisionMark(by, REVISION_ROLE.rowInsert), [])
    );
  if (del)
    markers.push(
      build(ctx.mint(), 'generic', deleteName, ctx.revisionMark(del, REVISION_ROLE.rowDelete), [])
    );
  const properties = props.find((c) => isWml(c, container));
  if (!properties || properties.kind === 'textValue') {
    // A row's `w:trPr` follows its `w:tblPrEx`; a cell's `w:tcPr` comes first
    const made = build(ctx.mint(), 'generic', container, [], markers);
    const after = props.findIndex((c) => isWml(c, 'tblPrEx')) + 1;
    return [...props.slice(0, after), made, ...props.slice(after)];
  }

  // The markers go before the properties' change record, the last child in the schema
  const children = [...properties.children];
  const at = children.findIndex((c) => c.kind !== 'textValue' && c.localName.endsWith('Change'));
  children.splice(at < 0 ? children.length : at, 0, ...markers);
  return props.map((c) => (c === properties ? withChildren(properties, children) : c));
}
