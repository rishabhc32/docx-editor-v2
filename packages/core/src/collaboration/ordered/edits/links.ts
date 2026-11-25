// Links (group 9.5): a new link around words, a link taken off its words, and a link's new
// target, each as the engine's own op writes it.

import { RELATIONSHIPS_NAMESPACE_URI } from '../../../store/package/ooxml-shared.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlAttribute,
  type OoxmlElement,
} from '../../../store/package/ooxml-tree.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import type { OrderedChangeStyle } from '../change/resources.ts';
import { BOOKMARK_NAME } from '../change/values.ts';
import { withStyleWrite } from '../compose/properties.ts';
import type { OrderedViewParagraph } from '../compose/context.ts';
import { isLinkSegment, segmentsOf } from '../containers.ts';
import { createOrderedItem, insertedPending, isHidden } from '../model/document.ts';
import { insertItems, isolateUnits, nextSeq, relabel, replaceItems } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import {
  holdsRunFormat,
  isUnit,
  type OrderedFormatRecord,
  type OrderedItem,
  type OrderedStyleWrite,
  type OrderedSuggestion,
} from '../model/types.ts';
import { relationshipIdAttribute, withChildren, withId, wmlAttribute } from '../ooxml-nodes.ts';
import {
  closeOf,
  itemsInside,
  refuseLocked,
  refuseLockedAt,
  refuseLockedWithin,
} from '../placement.ts';
import {
  currentFace,
  pendingOwn,
  strikeItems,
  suggestionFor,
  takeBack,
  type Ctx,
} from './context.ts';
import { settle } from './direct.ts';
import { addRelationship, addStyle } from './resources.ts';
import { groupBySpace, sameFace, struckAlready, withFormat } from './text.ts';

/** Where a link points: an address, or a bookmark in the document. @public */
export type OrderedLinkTarget = { readonly url: string } | { readonly anchor: string };

/** What a new link or a new target gives besides where it points. */
export interface OrderedLinkLook {
  readonly tooltip?: string;
  /** The character style the linked words take (`Hyperlink`). */
  readonly style?: string;
  /** The style's definition, added when the document has none with its id. */
  readonly newStyle?: OrderedChangeStyle;
}

/** Whether an attribute names where a link points (and so goes with a new target). */
const isTarget = (a: OoxmlAttribute) =>
  (a.namespaceUri === RELATIONSHIPS_NAMESPACE_URI && a.localName === 'id') ||
  (a.namespaceUri === WML_NAMESPACE_URI && a.localName === 'anchor');

/**
 * The attributes a target gives, in the engine's order (`r:id` or `w:anchor`, then
 * `w:tooltip`). An address is the suggestion's own relationship resource (design D14), so a
 * rejected or taken-back link leaves no relationship in the file.
 */
function targetAttributes(
  ctx: Ctx,
  suggestion: string,
  target: OrderedLinkTarget,
  tooltip: string | undefined
): OoxmlAttribute[] {
  const out: OoxmlAttribute[] = [];
  if ('anchor' in target) {
    if (!BOOKMARK_NAME.test(target.anchor))
      throw new OrderedRefusal('invalid-edit', 'a bookmark name of letters, digits and _');
    out.push(wmlAttribute('anchor', target.anchor));
  } else {
    // `addRelationship` keeps only an address `sanitizeHref` keeps as it is
    const relationship = addRelationship(ctx, suggestion, 'hyperlink', { external: target.url });
    out.push(relationshipIdAttribute(relationship.id));
  }

  if (tooltip !== undefined) out.push(wmlAttribute('tooltip', tooltip));
  return out;
}

/**
 * A link's container element, its node id the open item's space, so the editor names the link
 * by an id the list finds (`openOfNode`). `r:` is declared on it when a relationship names it.
 */
function linkShell(space: string, attributes: readonly OoxmlAttribute[]): OoxmlElement {
  const shell = build(space, 'hyperlink', 'hyperlink', [...attributes], []);
  return attributes.some((a) => a.namespaceUri === RELATIONSHIPS_NAMESPACE_URI)
    ? { ...shell, namespaceBindings: [{ prefix: 'r', namespaceUri: RELATIONSHIPS_NAMESPACE_URI }] }
    : shell;
}

/** A link's open and close items under a new space, with the labels given. */
function linkItems(
  ctx: Ctx,
  attributes: readonly OoxmlAttribute[],
  labels: Pick<OrderedItem, 'by' | 'within'>,
  bindings: OoxmlElement['namespaceBindings'] = []
): { readonly open: OrderedItem; readonly close: OrderedItem } {
  const space = ctx.ids.container();
  const shell = linkShell(space, attributes);
  const extra = bindings.filter(
    (b) => !shell.namespaceBindings.some((own) => own.prefix === b.prefix)
  );
  const node =
    extra.length > 0
      ? { ...shell, namespaceBindings: [...shell.namespaceBindings, ...extra] }
      : shell;
  const edge = { token: 'hyperlink' as const, pair: space, ...labels };
  return {
    open: createOrderedItem({ space, kind: 'open', node, ...edge }),
    close: createOrderedItem({ space, offset: 1, kind: 'close', ...edge }),
  };
}

/** A link's own element. */
function shellOf(open: OrderedItem): OoxmlElement {
  const shell = open.node;
  if (!shell || shell.kind === 'textValue')
    throw new OrderedRefusal('malformed-story', 'a link with no node');
  return shell;
}

/**
 * The item a container's open item goes after to come right before `first`: past the markers
 * and hidden items before it, so a marker at the range's start is inside the container (the
 * engine's rule for a new link), but never past a unit, a mark or a container's edge that shows.
 */
function openAfter(ctx: Ctx, first: OrderedItem): OrderedItem | null {
  let at = ctx.story.tree.prev(first);
  while (
    at &&
    ((at.kind === 'zero' && at.inRun !== true && at.block !== true) ||
      (at.kind !== 'mark' && isHidden(ctx.doc, at)))
  )
    at = ctx.story.tree.prev(at);
  return at;
}

/** The units of the view characters `[start, end)`, cut at the range's edges, in order. */
function cutUnits(ctx: Ctx, p: OrderedViewParagraph, start: number, end: number): OrderedItem[] {
  const out: OrderedItem[] = [];
  for (const group of groupBySpace(p.chars.slice(start, end)))
    out.push(...isolateUnits(ctx.doc, ctx.story, group.space, group.start, group.end));
  return out;
}

/** The shown units of a container, in order. */
function shownUnits(ctx: Ctx, open: OrderedItem): OrderedItem[] {
  return itemsInside(ctx.story, open, ctx.doc).filter(
    (item) => isUnit(item) && !isHidden(ctx.doc, item)
  );
}

/** Whether a unit is its author's own pending insertion, which they change in place. */
const ownUnit = (ctx: Ctx, item: OrderedItem): OrderedSuggestion | null =>
  insertedPending(ctx.doc, item) ? pendingOwn(ctx, item.by) : null;

/** A face with a character style written, under the edit's ids; null when nothing is left. */
function styledFace(
  ctx: Ctx,
  face: OoxmlElement | null,
  write: OrderedStyleWrite
): OoxmlElement | null {
  const children = withStyleWrite(face?.children ?? [], write, () => ctx.ids.node());
  if (children.length === 0) return null;
  const base = face ?? build(ctx.ids.node(), 'runProperties', 'rPr', [], []);
  return withId(withChildren(base, children), ctx.ids.node());
}

/** Whether a unit shows the character style now. */
function shows(ctx: Ctx, item: OrderedItem, style: string): boolean {
  const face = currentFace(ctx, item, false);
  return (face?.children ?? []).some(
    (c) =>
      c.kind !== 'textValue' &&
      c.localName === 'rStyle' &&
      c.attributes.some((a) => a.localName === 'val' && a.value === style)
  );
}

/** The characters `[start, end)` share their containers and are in no link. */
function checkLinkable(p: OrderedViewParagraph, start: number, end: number): void {
  const chars = p.chars.slice(start, end);
  const first = chars[0];
  if (!first || end <= start) throw new OrderedRefusal('empty-range');
  if (chars.some((c) => c.containers !== first.containers))
    throw new OrderedRefusal('invalid-edit', 'a link crosses the edge of a container');

  // Word writes no link inside a link, and the engine refuses one
  if (segmentsOf(first.containers).some(isLinkSegment))
    throw new OrderedRefusal('invalid-edit', 'a link inside a link');
}

/**
 * A new link around the view characters `[start, end)` of one paragraph, as the engine's
 * `insertHyperlink` makes it, with the link's character style. Suggesting, the characters are
 * the suggestion's own pending copy of the words it struck (the editor's proposal: Word has no
 * revision for a link, ECMA-376 `CT_Hyperlink`); the link joins that suggestion.
 */
export function insertLink(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number,
  target: OrderedLinkTarget,
  look: OrderedLinkLook
): string {
  refuseLocked(ctx, p, start, end);
  checkLinkable(p, start, end);
  const units = cutUnits(ctx, p, start, end);
  const first = units[0];
  const last = units[units.length - 1];
  if (!first || !last) throw new OrderedRefusal('empty-range');
  let own: OrderedSuggestion | null = null;
  if (!ctx.direct) {
    // Word has no tracked form for a link around words: a suggested link wraps its own copy
    for (const unit of units) {
      const owner = ownUnit(ctx, unit);
      if (!owner)
        throw new OrderedRefusal('invalid-edit', "a suggested link wraps its suggestion's words");
      own ??= owner;
    }
  }

  const suggestion = suggestionFor(ctx, own);
  const attributes = [
    wmlAttribute('history', '1'),
    ...targetAttributes(ctx, suggestion.id, target, look.tooltip),
  ];
  const { open, close } = linkItems(ctx, attributes, { by: suggestion.id, within: first.within });
  insertItems(ctx.doc, ctx.story, last, [close]);
  insertItems(ctx.doc, ctx.story, openAfter(ctx, first), [open]);
  if (look.style !== undefined) styleUnits(ctx, units, { set: look.style }, suggestion);
  if (look.newStyle !== undefined) addStyle(ctx, suggestion.id, look.newStyle);
  if (ctx.direct) {
    settle(ctx, suggestion.id);
  }

  return suggestion.id;
}

/**
 * Put a link's words in new containers: the old open and close items are struck by the
 * suggestion, and each part (`from` to `to` of the shown units) goes in a new link with its own
 * attributes, or in none (`null`). Parts are in order and cover every shown unit. Nothing is
 * removed, so the organiser's undo (`revert`) shows the old link again exactly.
 */
function rewrap(
  ctx: Ctx,
  open: OrderedItem,
  close: OrderedItem,
  parts: readonly {
    readonly first: OrderedItem;
    readonly last: OrderedItem;
    readonly attributes: readonly OoxmlAttribute[] | null;
  }[],
  suggestion: OrderedSuggestion
): void {
  strikeItems(ctx, ctx.story, [open, close], suggestion.id);
  const bindings = shellOf(open).namespaceBindings;
  for (const [k, part] of parts.entries()) {
    if (part.attributes === null) continue;
    const labels = { by: suggestion.id, within: open.within };
    const made = linkItems(ctx, part.attributes, labels, bindings);
    insertItems(ctx.doc, ctx.story, part.last, [made.close]);
    insertItems(ctx.doc, ctx.story, k === 0 ? open : openAfter(ctx, part.first), [made.open]);
  }
}

/** The shown units of a link before, in and after a range of them. */
function partsOf(
  units: readonly OrderedItem[],
  range: readonly OrderedItem[] | null
): { before: OrderedItem[]; inside: OrderedItem[]; after: OrderedItem[] } {
  if (range === null) return { before: [], inside: [...units], after: [] };
  const set = new Set(range);
  const at = units.findIndex((u) => set.has(u));
  const end = at + range.length;
  if (at < 0 || units.slice(at, end).some((u) => !set.has(u)))
    throw new OrderedRefusal('not-in-link', 'the range is not in the link');
  return { before: units.slice(0, at), inside: units.slice(at, end), after: units.slice(end) };
}

/** The parts of a link's words, each with the attributes of its new link (null: none). */
function partList(
  before: readonly OrderedItem[],
  inside: readonly OrderedItem[],
  after: readonly OrderedItem[],
  first: readonly OoxmlAttribute[],
  middle: readonly OoxmlAttribute[] | null,
  last: readonly OoxmlAttribute[]
) {
  const out: {
    first: OrderedItem;
    last: OrderedItem;
    attributes: readonly OoxmlAttribute[] | null;
  }[] = [];
  const add = (units: readonly OrderedItem[], attributes: readonly OoxmlAttribute[] | null) => {
    const a = units[0];
    const b = units[units.length - 1];
    if (a && b) out.push({ first: a, last: b, attributes });
  };
  add(before, first);
  add(inside, middle);
  add(after, last);
  return out;
}

/** The range's units in the link: null for the whole link. */
function rangeIn(
  ctx: Ctx,
  open: OrderedItem,
  range: { readonly p: OrderedViewParagraph; readonly start: number; readonly end: number } | null
): OrderedItem[] | null {
  if (range === null) return null;
  if (range.end <= range.start) throw new OrderedRefusal('empty-range');
  const units = cutUnits(ctx, range.p, range.start, range.end);

  // Every unit in the link (the cut above may split items the link holds)
  const inLink = new Set(shownUnits(ctx, open));
  if (units.some((u) => !inLink.has(u)))
    throw new OrderedRefusal('not-in-link', 'the range is not in the link');
  return units.length === inLink.size ? null : units;
}

/** A range of a link's words, by view offsets of the paragraph that holds it. */
export interface OrderedLinkRange {
  readonly p: OrderedViewParagraph;
  readonly start: number;
  readonly end: number;
}

/**
 * Take a link off its words (the engine's `removeHyperlink`), or off part of them, with the
 * link's character style (`style`). Suggesting: a link the suggestion inserted goes in place;
 * the suggestion's own pending copy at the end of a link leaves it (the editor's proposal of an
 * unlink or a new target: the copy goes right after the link, where the engine's split leaves
 * it).
 */
export function removeLink(
  ctx: Ctx,
  open: OrderedItem,
  range: OrderedLinkRange | null,
  style: string | undefined
): string {
  // The engine's removal rearranges the link: the controls around it and in it are asked
  refuseLockedAt(ctx, open);
  refuseLockedWithin(ctx, itemsInside(ctx.story, open, ctx.doc), false);
  const close = closeOf(ctx.story, open, ctx.doc);
  const cut = rangeIn(ctx, open, range);
  const { before, inside, after } = partsOf(shownUnits(ctx, open), cut);
  const unstyle: OrderedStyleWrite | null = style === undefined ? null : { remove: style };
  if (ctx.direct) {
    const suggestion = suggestionFor(ctx);
    const kept = shellOf(open).attributes;
    rewrap(ctx, open, close, partList(before, inside, after, kept, null, kept), suggestion);
    if (unstyle) styleUnits(ctx, inside, unstyle, suggestion);
    settle(ctx, suggestion.id);
    return suggestion.id;
  }

  const ownLink = pendingOwn(ctx, open.by);
  if (cut === null && ownLink && insertedPending(ctx.doc, open)) {
    if (inside.some((u) => !ownUnit(ctx, u)))
      throw new OrderedRefusal('invalid-edit', "a suggested link holds its suggestion's words");
    for (const edge of [open, close]) takeBack(ctx, ctx.story, edge);
    if (unstyle) styleUnits(ctx, inside, unstyle, ownLink);
    return ownLink.id;
  }

  return unlinkOwnCopy(ctx, close, inside, after, unstyle);
}

/**
 * Write a character style on shown units (each still in the tree): an own pending unit changes
 * its face; any other unit takes a record of the suggestion, which a direct edit settles (the
 * engine writes the style straight into the run).
 */
function styleUnits(
  ctx: Ctx,
  items: readonly OrderedItem[],
  write: OrderedStyleWrite,
  suggestion: OrderedSuggestion
): void {
  const seq = nextSeq(ctx.doc);
  for (const item of items) {
    // A unit with no run (an equation, an element kept whole) takes no character style
    if (!holdsRunFormat(item)) {
      continue;
    }

    if ('remove' in write && !shows(ctx, item, write.remove)) continue;
    if (ownUnit(ctx, item)) {
      replaceItems(
        ctx.doc,
        ctx.story,
        [item],
        [{ ...item, face: styledFace(ctx, item.face, write), leaf: undefined }]
      );
      continue;
    }

    const record: OrderedFormatRecord = {
      kind: 'properties',
      id: `f:${suggestion.id}#${seq}`,
      suggestion: suggestion.id,
      seq,
      properties: [],
      style: write,
    };
    relabel(ctx.doc, ctx.story, item, { fmts: withFormat(ctx, item, record) });
  }
}

/**
 * The suggestion's own pending copy at the end of a link leaves it: the copy is taken back
 * where it was typed and typed again right after the link, without the link's style.
 */
function unlinkOwnCopy(
  ctx: Ctx,
  close: OrderedItem,
  inside: readonly OrderedItem[],
  after: readonly OrderedItem[],
  unstyle: OrderedStyleWrite | null
): string {
  const owners = new Set(inside.map((u) => ownUnit(ctx, u)));
  const [owner] = owners;
  if (after.length > 0 || owners.size !== 1 || !owner)
    throw new OrderedRefusal(
      'invalid-edit',
      "a suggested unlink takes its suggestion's own words from the link's end"
    );
  const space = ctx.ids.units();
  let offset = 0;
  const copies = inside.map((unit) => {
    const copy = createOrderedItem({
      ...unit,
      space,
      offset,
      run: space,
      face: unstyle ? styledFace(ctx, unit.face, unstyle) : unit.face,
      retracted: false,
    });
    offset += unit.length;
    return copy;
  });
  for (const unit of inside) takeBack(ctx, ctx.story, unit);
  insertItems(ctx.doc, ctx.story, close, copies);
  return owner.id;
}

/**
 * A link's new target, tooltip, or both, as the engine's `setHyperlinkTarget` writes it: every
 * other attribute stays, and a part of the link (`range`) becomes a link of its own. Suggesting,
 * a link the suggestion inserted changes in place; another link's words are struck and a copy
 * goes in a new link right after it (`proposeLinkTarget`), as the engine proposes.
 */
export function retargetLink(
  ctx: Ctx,
  open: OrderedItem,
  target: OrderedLinkTarget,
  tooltip: string | undefined,
  range: OrderedLinkRange | null
): string {
  refuseLockedAt(ctx, open);
  const close = closeOf(ctx.story, open, ctx.doc);
  const shell = shellOf(open);
  const ownLink = pendingOwn(ctx, open.by);
  if (!ctx.direct && !(ownLink && insertedPending(ctx.doc, open))) {
    if (range !== null)
      throw new OrderedRefusal('invalid-edit', 'a suggested target for part of a link');
    return proposeLinkTarget(ctx, open, target, tooltip);
  }

  const suggestion = suggestionFor(ctx, ownLink);
  const retained = shell.attributes.filter(
    (a) =>
      !isTarget(a) &&
      !(tooltip !== undefined && a.namespaceUri === WML_NAMESPACE_URI && a.localName === 'tooltip')
  );
  const aimed = [...retained, ...targetAttributes(ctx, suggestion.id, target, tooltip)];
  const cut = rangeIn(ctx, open, range);
  const { before, inside, after } = partsOf(shownUnits(ctx, open), cut);
  if (inside.length === 0) throw new OrderedRefusal('hidden-target', `${open.space} shows no text`);
  if (ownLink) {
    // The suggestion's own link: it was never anyone else's, so it changes in place
    if (before.length > 0 || after.length > 0)
      throw new OrderedRefusal('invalid-edit', 'a suggested target for part of a link');
    for (const edge of [open, close]) takeBack(ctx, ctx.story, edge);
    const made = linkItems(ctx, aimed, { by: suggestion.id, within: open.within });
    insertItems(ctx.doc, ctx.story, ctx.story.tree.prev(close), [made.close]);
    insertItems(ctx.doc, ctx.story, open, [made.open]);
    return suggestion.id;
  }

  const parts = partList(before, inside, after, shell.attributes, aimed, shell.attributes);
  rewrap(ctx, open, close, parts, suggestion);
  settle(ctx, suggestion.id);
  return suggestion.id;
}

/**
 * A new target proposed for another person's or the file's link, as the engine proposes it: the
 * old link's words struck, and a copy of them in a new link to the target, inserted right after
 * it. Only a link of plain text in one format: the copy writes its words again.
 */
function proposeLinkTarget(
  ctx: Ctx,
  open: OrderedItem,
  target: OrderedLinkTarget,
  tooltip: string | undefined
): string {
  // The words that show (struck ones too): never text a decision took out
  const shown = shownUnits(ctx, open);
  const first = shown[0];
  if (
    shown.some((item) => item.kind !== 'text') ||
    shown.some((item) => !sameFace(item.face, first?.face ?? null))
  )
    throw new OrderedRefusal('link-not-plain', 'only a link of plain text in one format');
  const suggestion = suggestionFor(ctx);
  strikeItems(
    ctx,
    ctx.story,
    shown.filter((item) => !struckAlready(ctx, item.dels)),
    suggestion.id
  );
  const attributes = [
    wmlAttribute('history', '1'),
    ...targetAttributes(ctx, suggestion.id, target, tooltip),
  ];
  const labels = { by: suggestion.id, within: open.within };
  const made = linkItems(ctx, attributes, labels);
  const text = shown.map((item) => item.text).join('');
  const textSpace = ctx.ids.units();
  const words =
    text.length > 0
      ? [
          createOrderedItem({
            space: textSpace,
            kind: 'text',
            length: text.length,
            text,
            // The Hyperlink character style, as the engine's link command writes link text
            face: styledFace(ctx, first?.face ?? null, { set: 'Hyperlink' }),
            run: textSpace,
            ...labels,
          }),
        ]
      : [];
  insertItems(ctx.doc, ctx.story, closeOf(ctx.story, open, ctx.doc), [
    made.open,
    ...words,
    made.close,
  ]);
  return suggestion.id;
}
