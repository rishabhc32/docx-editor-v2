// What every edit runs with: the document and its story, who makes the edit, and its id source.

import type { ContentLockPolicy } from '../../../store/package/ooxml-edit.ts';
import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { cloneWithNewIds } from '../../../store/store/tree-op-nodes.ts';
import type { OrderedIdMinter } from '../change/ids.ts';
import { ComposeContext } from '../compose/context.ts';
import { faceAfter, paragraphPropertiesAfter } from '../compose/properties.ts';
import { pendingBy, storyNamed } from '../model/document.ts';
import { newSuggestion, relabel } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type {
  OrderedAttribution,
  OrderedDocument,
  OrderedFormatRecord,
  OrderedItem,
  OrderedStory,
  OrderedSuggestion,
} from '../model/types.ts';
import { childWhere, isParagraphProperties, isWml, withChildren } from '../ooxml-nodes.ts';

/** A format record that a later format took over (policy 2), reported to the host. @public */
export interface OrderedReplacedFormat {
  /** The suggestion whose record was taken over. */
  readonly suggestion: string;
  /** The suggestion that took it. */
  readonly by: string;
}

/**
 * Who makes an edit, in which story, and where its new ids come from.
 *
 * @public
 */
export interface OrderedEditContext {
  readonly doc: OrderedDocument;
  /** The story part the edit's places are in, such as `/word/document.xml`. */
  readonly story: string;
  readonly attribution: OrderedAttribution;
  /**
   * Where new ids come from: the minter of the change the edit belongs to, never a counter of
   * the document. Every copy that applies the change passes a new minter for the same change
   * id, so every copy makes the same ids.
   */
  readonly ids: OrderedIdMinter;
  /**
   * A pending suggestion of the same author that the edit joins instead of making its own: a
   * change made of several edits (a replacement as a deletion and an insertion) is one card.
   */
  readonly join?: string;
  /**
   * The organiser edits with tracking off: the edit is a suggestion accepted at once, except
   * inside another author's pending suggestion, where it changes that suggestion.
   */
  readonly direct?: boolean;
  /** Policy 2: told of each format record the edit took over, once the edit holds. */
  readonly onReplaced?: (replaced: OrderedReplacedFormat) => void;
  /** The room's relaxation of content-control locks, as the editor's option gives one. */
  readonly contentLockPolicy?: ContentLockPolicy;
}

/** An edit's working state: the context with its story resolved, and what the edit reports. */
export interface Ctx {
  readonly doc: OrderedDocument;
  readonly story: OrderedStory;
  readonly attribution: OrderedAttribution;
  readonly ids: OrderedIdMinter;
  readonly join: string | undefined;
  /** A direct edit: never extends a pending suggestion of its author's; it joins only `join`. */
  readonly direct: boolean;
  /**
   * The joined suggestion struck its words in an earlier transaction (a selection deleted, then
   * typed over): typed text goes in front of them, not after them (`siteAt`, rule 13).
   */
  readonly earlierStrike?: boolean;
  /** Policy 2's takeovers so far, reported once the edit holds. */
  readonly replaced: OrderedReplacedFormat[];
  readonly contentLockPolicy: ContentLockPolicy | undefined;
}

/** The working state of an edit. */
export function contextOf(context: OrderedEditContext, replaced: OrderedReplacedFormat[]): Ctx {
  if (typeof context.attribution.author !== 'string')
    throw new OrderedRefusal('invalid-edit', 'an author is required');
  return {
    doc: context.doc,
    story: storyNamed(context.doc, context.story),
    attribution: context.attribution,
    ids: context.ids,
    join: context.join,
    direct: context.direct === true,
    replaced,
    contentLockPolicy: context.contentLockPolicy,
  };
}

/** The same edit made as another author's own change (a direct edit inside their suggestion). */
export function asAuthor(ctx: Ctx, suggestion: OrderedSuggestion): Ctx {
  return { ...ctx, attribution: suggestion.attribution, join: suggestion.id, direct: false };
}

/**
 * The suggestion, when it is pending and the edit's author made it. A direct edit has none: it
 * never changes a pending suggestion as its author's own edit (takes text back, changes its
 * face or its paragraph's properties); it lays its own settled record on those units, which its
 * undo (`revert`) turns off exactly (task 6.7, 9.12).
 */
export function pendingOwn(ctx: Ctx, suggestion: string | null): OrderedSuggestion | null {
  if (suggestion === null || ctx.direct) {
    return null;
  }

  return pendingBy(ctx.doc, suggestion, ctx.attribution);
}

/**
 * Take back an item for its own author, naming the edit's change, so that the undo of the change
 * shows it again (`restore`). An item taken back already keeps the change that took it.
 */
export function takeBack(ctx: Ctx, story: OrderedStory, item: OrderedItem): void {
  if (!item.retracted) relabel(ctx.doc, story, item, { retracted: true, taken: ctx.ids.changeId });
}

/** Strike items with the edit's suggestion; an item it strikes already stays as it is. */
export function strikeItems(
  ctx: Ctx,
  story: OrderedStory,
  items: readonly OrderedItem[],
  suggestion: string
): void {
  for (const item of items) {
    if (!item.dels.includes(suggestion)) {
      relabel(ctx.doc, story, item, { dels: [...item.dels, suggestion] });
    }
  }
}

/** Hide an item in a container the suggestion removes (`within.dels`), as a struck row's. */
export function holdWithin(
  ctx: Ctx,
  story: OrderedStory,
  item: OrderedItem,
  suggestion: string
): void {
  relabel(ctx.doc, story, item, {
    within: { by: item.within.by, dels: [...item.within.dels, suggestion] },
  });
}

/** The suggestion an edit's labels go to: the joined one, the one it extends, or a new one. */
export function suggestionFor(
  ctx: Ctx,
  extend: OrderedSuggestion | null = null
): OrderedSuggestion {
  // The edit's own suggestion (a direct one too: a move's two halves are one suggestion)
  const joined = ctx.join === undefined ? null : pendingBy(ctx.doc, ctx.join, ctx.attribution);
  if (joined) {
    return joined;
  }

  // A direct edit is settled at once: it never grows a pending suggestion of its author's
  if (!ctx.direct && extend) return extend;
  return newSuggestion(ctx.doc, ctx.attribution, ctx.ids.suggestion());
}

/** A copy of a node with ids from the edit's id source. */
export function withNewIds<T extends OoxmlElement>(ctx: Ctx, node: T): T {
  const copy = cloneWithNewIds(node, () => ctx.ids.node());

  // cloneWithNewIds keeps the node's kind, so the copy is a T as the node is
  return copy as T;
}

/**
 * A compose context over a story, to read what shows now. Its ids are scratch ones
 * (`<prefix>#1`, ...), so `doc.composeIds` stays as it is and every copy reads the same.
 */
export function scratchCompose(
  doc: OrderedDocument,
  story: OrderedStory,
  prefix: string
): ComposeContext {
  let scratch = 0;
  return new ComposeContext(doc, story, {}, () => `${prefix}#${(scratch += 1)}`);
}

/**
 * Whether a unit's format records draw a change record (`w:rPrChange`): their pending ones leave
 * its face other than it was before them.
 */
export function recordsFormat(
  ctx: Ctx,
  fmts: readonly OrderedFormatRecord[],
  face: OoxmlElement | null
): boolean {
  const drawn = faceAfter(scratchCompose(ctx.doc, ctx.story, 'read'), fmts, face);
  return drawn !== null && drawn.children.some((c) => isWml(c, 'rPrChange'));
}

/** An item's face as it shows now, its change record (`w:rPrChange`) included. */
export function shownFace(ctx: Ctx, item: OrderedItem): OoxmlElement | null {
  return faceAfter(scratchCompose(ctx.doc, ctx.story, 'read'), item.fmts, item.face);
}

/** A paragraph mark's properties as they show now, the change record included. */
export function shownParagraphProperties(ctx: Ctx, mark: OrderedItem): OoxmlElement | null {
  const shell = mark.node;
  const base =
    shell && shell.kind !== 'textValue' ? childWhere(shell, isParagraphProperties) : undefined;
  return paragraphPropertiesAfter(scratchCompose(ctx.doc, ctx.story, 'read'), mark, base ?? null);
}

/**
 * The face a unit shows now: its run properties under its accepted and pending format records,
 * without the change record, as a copy with the edit's own ids (a new run must not share nodes
 * with its neighbour's), unless `copy` is false (a read only). Null when it has no properties.
 */
export function currentFace(ctx: Ctx, item: OrderedItem, copy = true): OoxmlElement | null {
  const face = shownFace(ctx, item);
  if (!face) return null;
  const children = face.children.filter((c) => !isWml(c, 'rPrChange'));
  if (children.length === 0) return null;
  const shown = withChildren(face, children);
  return copy ? withNewIds(ctx, shown) : shown;
}

/** A face without the placeholder style a prompt shows in (typed text takes the rest). */
export function withoutPlaceholderStyle(face: OoxmlElement | null): OoxmlElement | null {
  if (!face) return null;
  const children = face.children.filter(
    (c) =>
      c.kind === 'textValue' ||
      !(
        isWml(c, 'rStyle') &&
        c.attributes.some((a) => a.localName === 'val' && a.value === 'PlaceholderText')
      )
  );
  return children.length > 0 ? withChildren(face, children) : null;
}
