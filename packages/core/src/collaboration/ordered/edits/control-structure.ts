// New, removed and changed content controls (task 9.6), each checked by the engine's own op on
// the view first. Nothing moves (design D1): new edges go in, removed ones are struck.

import type { ContentControlLock } from '../../../store/package/content-control-nodes.ts';
import { findNode } from '../../../store/package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { canTrackContentControl } from '../../../store/store/tracked-content-control-insert.ts';
import type { InsertableContentControlKind } from '../../../store/store/tree-op-content-controls.ts';
import { parentOf } from '../../../store/store/tree-op-nodes.ts';
import type { TreeDocOp } from '../../../store/store/tree-op-types.ts';
import type { OrderedViewChar, OrderedViewParagraph } from '../compose/context.ts';
import { ownItems } from '../compose/paragraph.ts';
import { segmentsOf } from '../containers.ts';
import { createOrderedItem, isHidden, itemAt, suggestionOrNull } from '../model/document.ts';
import { insertItems, relabel, splitItem } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { isUnit, type OrderedItem } from '../model/types.ts';
import { childNamed, isControlContent, withAttributes, withChildren } from '../ooxml-nodes.ts';
import { closeOf, itemsInside, staysInside } from '../placement.ts';
import {
  addedControl,
  changedProperties,
  contentUnits,
  controlEdges,
  engineOnView,
  sdtPrRecord,
  strikeBlock,
  type ControlLabels,
} from './control-shell.ts';
import { holdWithin, strikeItems, suggestionFor, withNewIds, type Ctx } from './context.ts';
import { isParagraphIdentity } from './para-ids.ts';
import { shellId, templateItems, viewNode } from './templates.ts';
import { deleteText } from './text.ts';

/** A control's metadata as the engine's ops take it. */
export interface OrderedControlMetadata {
  readonly tag?: string;
  readonly alias?: string;
  readonly lock?: ContentControlLock;
}

/**
 * A new control at a caret: an empty one of a type, showing its type's prompt (the engine's
 * `insertContentControl` at a caret), or a custom node, a control holding `text` (the engine's
 * `insertInlineContentControl`).
 */
export type OrderedNewControl =
  | (OrderedControlMetadata & { readonly type: InsertableContentControlKind })
  | {
      readonly type: 'custom';
      readonly tag: string;
      readonly text: string;
      readonly alias?: string;
      readonly lock?: 'sdtLocked' | 'sdtContentLocked' | 'contentLocked';
      readonly dataBinding?: {
        readonly prefixMappings: string;
        readonly xpath: string;
        readonly storeItemId: string;
      };
    };

const metadataOf = (control: OrderedControlMetadata) => ({
  ...(control.tag === undefined ? {} : { tag: control.tag }),
  ...(control.alias === undefined ? {} : { alias: control.alias }),
  ...(control.lock === undefined ? {} : { lock: control.lock }),
});

/**
 * A direct-only edit sent as a suggestion: Word records no change of a control's properties,
 * and the engine has a tracked form only for a new control around text (`wrapControl`).
 */
function directOnly(ctx: Ctx, what: string): void {
  if (!ctx.direct)
    throw new OrderedRefusal('invalid-edit', `${what} has no tracked form: a direct edit only`);
}

/** The item a text unit's address names, cut so it starts there. */
function unitFrom(ctx: Ctx, char: OrderedViewChar): OrderedItem {
  const item = itemAt(ctx.story, char.address);
  if (!item) throw new OrderedRefusal('unknown-address', char.address.space);
  if (item.kind !== 'text' || char.address.offset === item.offset) return item;
  return splitItem(ctx.doc, ctx.story, item, char.address.offset)?.[1] ?? item;
}

/** The item a text unit's address names, cut so it ends there. */
function unitTo(ctx: Ctx, char: OrderedViewChar): OrderedItem {
  const item = itemAt(ctx.story, char.address);
  if (!item) throw new OrderedRefusal('unknown-address', char.address.space);
  const end = char.address.offset + 1;
  if (item.kind !== 'text' || end === item.offset + item.length) return item;
  return splitItem(ctx.doc, ctx.story, item, end)?.[0] ?? item;
}

const depthOf = (char: OrderedViewChar): number => segmentsOf(char.containers).length;

/**
 * Where the engine puts a new control at a caret: after the unit before it and the ends of the
 * containers that end there, before any marker at the caret; at a paragraph's start, before
 * everything. Inside a link or a control (both sides in one) the engine refuses, or splits the
 * container for a custom node: refused by name here.
 */
function caretAnchor(ctx: Ctx, p: OrderedViewParagraph, offset: number): OrderedItem | null {
  const previous = offset > 0 ? p.chars[offset - 1] : undefined;
  const next = p.chars[offset];
  if (previous && next && previous.containers !== '' && staysInside(previous, next))
    throw new OrderedRefusal('engine-refused', 'a new control inside a link or a control');
  if (!previous) {
    const first = p.items[0] ?? p.mark;
    return ctx.story.tree.prev(first);
  }

  let at = unitTo(ctx, previous);
  for (let item = ctx.story.tree.next(at); item; item = ctx.story.tree.next(item)) {
    const hiddenUnit = isUnit(item) && isHidden(ctx.doc, item);
    if (item.kind !== 'close' && !hiddenUnit) break;
    at = item;
  }

  return at;
}

/** A new control at a caret, made by the engine on the view: its suggestion and its open item. */
export function insertControl(
  ctx: Ctx,
  p: OrderedViewParagraph,
  offset: number,
  control: OrderedNewControl
): { readonly suggestion: string; readonly open: OrderedItem } {
  const op: TreeDocOp =
    control.type === 'custom'
      ? {
          op: 'insertInlineContentControl',
          paragraphId: p.id,
          offset,
          tag: control.tag,
          text: control.text,
          ...(control.alias === undefined ? {} : { alias: control.alias }),
          ...(control.lock === undefined ? {} : { lock: control.lock }),
          ...(control.dataBinding === undefined ? {} : { dataBinding: control.dataBinding }),
        }
      : {
          op: 'insertContentControl',
          paragraphId: p.id,
          start: offset,
          end: offset,
          type: control.type,
          ...metadataOf(control),
        };
  directOnly(ctx, 'a new control at a caret');
  const { view, result } = engineOnView(ctx, op);
  const element = addedControl(view, result, p.id);
  const anchor = caretAnchor(ctx, p, offset);
  const suggestion = suggestionFor(ctx);
  const labels: ControlLabels = { suggestion: suggestion.id, within: p.mark.within };
  const prompt = control.type !== 'custom';
  const { open, close } = controlEdges(ctx, element, labels, prompt);
  const units = contentUnits(ctx, element, labels, prompt);
  insertItems(ctx.doc, ctx.story, anchor, [open, ...units, close]);
  return { suggestion: suggestion.id, open };
}

/**
 * The first (`prev`) or last (`next`) item of the top-level element the character is in: the
 * outermost container's open or close item, walking from `from`.
 */
function topEdge(
  ctx: Ctx,
  char: OrderedViewChar,
  from: OrderedItem,
  direction: 'prev' | 'next'
): OrderedItem {
  const [leaves, enters] = direction === 'prev' ? ['open', 'close'] : ['close', 'open'];
  let at = from;
  for (let depth = depthOf(char); depth > 0; ) {
    const step = ctx.story.tree[direction](at);
    if (!step) {
      break;
    }

    at = step;
    if (step.kind === leaves) {
      depth -= 1;
    } else if (step.kind === enters) {
      depth += 1;
    }
  }

  return at;
}

/**
 * A new control around the view characters `[start, end)` of one paragraph. Directly: an open
 * item before them and a close item after them, at the paragraph's own level, as the engine
 * wraps whole runs and containers. As a suggestion, the engine's tracked form: the words struck,
 * and a new control holding a copy of them inserted right after (both in one suggestion), so a
 * reject leaves the words as they were and an accept leaves the control.
 */
export function wrapControl(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number,
  control: OrderedControlMetadata & { readonly type: InsertableContentControlKind }
): string {
  const first = p.chars[start];
  const last = p.chars[end - 1];
  if (!first || !last || end <= start) throw new OrderedRefusal('empty-range');
  const revision = ctx.direct ? undefined : ctx.attribution;
  const op: TreeDocOp = {
    op: 'insertContentControl',
    paragraphId: p.id,
    start,
    end,
    type: control.type,
    ...metadataOf(control),
    ...(revision === undefined
      ? {}
      : {
          revision: {
            author: revision.author,
            ...(revision.date === undefined ? {} : { date: revision.date }),
          },
        }),
  };
  const { view, result } = engineOnView(ctx, op);
  if (revision !== undefined && !canTrackContentControl(view, p.id, start, end))
    throw new OrderedRefusal('engine-refused', 'insertContentControl: indivisible-content');
  const element = addedControl(view, result, p.id);
  if (revision !== undefined) {
    return wrapAsSuggestion(ctx, p, start, end, last, element);
  }

  return wrapDirect(ctx, p, first, last, element);
}

/** The tracked form: the words struck, and a new control holding a copy right after them. */
function wrapAsSuggestion(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number,
  last: OrderedViewChar,
  element: OoxmlElement
): string {
  const struck = deleteText(ctx, p, start, end, suggestionOrNull(ctx.doc, ctx.join));
  const suggestion = suggestionOrNull(ctx.doc, struck) ?? suggestionFor(ctx);
  const labels: ControlLabels = { suggestion: suggestion.id, within: p.mark.within };
  const { open, close } = controlEdges(ctx, element, labels, false);
  const copies = contentUnits(ctx, element, labels, false);
  insertItems(ctx.doc, ctx.story, unitTo(ctx, last), [open, ...copies, close]);
  return suggestion.id;
}

/** Directly: an open item before the characters and a close item after them, at top level. */
function wrapDirect(
  ctx: Ctx,
  p: OrderedViewParagraph,
  first: OrderedViewChar,
  last: OrderedViewChar,
  element: OoxmlElement
): string {
  const suggestion = suggestionFor(ctx);
  const labels: ControlLabels = { suggestion: suggestion.id, within: p.mark.within };
  const { open, close } = controlEdges(ctx, element, labels, false);

  // Cut both ends first: a cut replaces the item it cuts, so each is read again by its address
  unitFrom(ctx, first);
  unitTo(ctx, last);
  const before = topEdge(ctx, first, unitFrom(ctx, first), 'prev');
  const afterLast = topEdge(ctx, last, unitTo(ctx, last), 'next');
  insertItems(ctx.doc, ctx.story, ctx.story.tree.prev(before), [open]);
  insertItems(ctx.doc, ctx.story, afterLast, [close]);
  return suggestion.id;
}

/** A shown control's open item, by its address. */
export function shownControl(ctx: Ctx, open: OrderedItem): OrderedItem {
  if (open.kind !== 'open' || (open.token !== 'sdt' && open.token !== 'sdtBlock'))
    throw new OrderedRefusal('unknown-control', open.space);
  if (isHidden(ctx.doc, open)) throw new OrderedRefusal('hidden-target', open.space);
  return open;
}

/** A plain copy of a prompt unit, as words the removal of its control leaves: no prompt. */
function promptWords(ctx: Ctx, unit: OrderedItem, suggestion: string): OrderedItem {
  const space = ctx.ids.units();
  const node = unit.node && unit.node.kind !== 'textValue' ? withNewIds(ctx, unit.node) : null;
  return createOrderedItem({
    space,
    kind: unit.kind,
    length: unit.length,
    text: unit.text,
    node,
    face: unit.face ? withNewIds(ctx, unit.face) : null,
    fmts: unit.fmts,
    run: space,
    ...(unit.textShell ? { textShell: withNewIds(ctx, unit.textShell) } : {}),
    ...(unit.inRun ? { inRun: true } : {}),
    ...(unit.bare ? { bare: true } : {}),
    by: suggestion,
    within: unit.within,
  });
}

/**
 * Remove a control, keeping its content (Word's "Remove content control"): its open and close
 * items are struck, so its content draws in the control's place.
 */
export function removeControl(ctx: Ctx, open: OrderedItem): string {
  directOnly(ctx, 'removing a control');
  shownControl(ctx, open);
  engineOnView(ctx, { op: 'removeContentControl', controlId: shellId(open), keepContent: true });
  const suggestion = suggestionFor(ctx);

  // Its own items: a control inside it keeps its own prompt, and goes on as it is
  const inside = [...ownItems(itemsInside(ctx.story, open, ctx.doc), open)];

  // A prompt's marks stay: the paragraphs content took over hold it (`change/settle-prompts.ts`)
  const prompt = inside.filter((item) => item.prompt === true && item.kind !== 'mark');
  const units = prompt.filter((item) => isUnit(item) && !isHidden(ctx.doc, item));
  const filled = inside.some((item) => isUnit(item) && !item.prompt && !isHidden(ctx.doc, item));

  // A prompt that content replaced goes for good with the control (the engine's unwrap keeps
  // only what the control shows). One the control still shows becomes its words: each unit
  // goes, and a plain copy takes its place, so later edits treat it as text, as the engine
  // does, and taking the removal back shows the prompt again.
  strikeItems(
    ctx,
    ctx.story,
    [open, ...(filled ? prompt : units), closeOf(ctx.story, open, ctx.doc)],
    suggestion.id
  );
  if (!filled)
    for (const unit of units)
      insertItems(ctx.doc, ctx.story, unit, [promptWords(ctx, unit, suggestion.id)]);
  return suggestion.id;
}

/**
 * Remove a control with its content: every item from its open item to its close item is struck
 * (another person's pending insertion in it too, as a direct deletion strikes it, so taking the
 * edit back shows it again as it was).
 */
export function deleteControl(ctx: Ctx, open: OrderedItem): string {
  directOnly(ctx, 'removing a control with its content');
  const suggestion = suggestionFor(ctx);
  strikeControl(ctx, open, suggestion.id);
  return suggestion.id;
}

/**
 * Strike a control with its content by `suggestion`, as the engine removes it: its own op on the
 * view refuses a lock that keeps the control, or a control in it.
 */
export function strikeControl(ctx: Ctx, open: OrderedItem, suggestion: string): void {
  shownControl(ctx, open);
  engineOnView(ctx, { op: 'removeContentControl', controlId: shellId(open), keepContent: false });

  // A block control goes as a struck row does: its content through `within`
  if (open.token === 'sdtBlock') {
    strikeBlock(ctx, open, suggestion);
    return;
  }

  const close = closeOf(ctx.story, open, ctx.doc);
  strikeItems(ctx, ctx.story, [open, ...itemsInside(ctx.story, open, ctx.doc), close], suggestion);
}

/**
 * A control's tag, title (`w:alias`) and lock, as the engine writes them: a record of the
 * `w:sdtPr` children the op changed, and only those. Null when it changed none.
 */
export function setControlProperties(
  ctx: Ctx,
  open: OrderedItem,
  edits: {
    readonly tag?: string | null;
    readonly alias?: string | null;
    readonly lock?: ContentControlLock;
  }
): string | null {
  directOnly(ctx, "a control's properties");
  shownControl(ctx, open);
  const id = shellId(open);
  const { view, result } = engineOnView(ctx, {
    op: 'setContentControlProperties',
    controlId: id,
    ...edits,
  });
  const before = childNamed(viewNode(view, id), 'sdtPr');
  const after = childNamed(viewNode(result, id), 'sdtPr');
  if (!after) throw new OrderedRefusal('engine-refused', 'the control lost its properties');
  const merge = changedProperties(before, after, result);
  if (merge.length === 0) {
    return null;
  }

  const suggestion = suggestionFor(ctx);
  const record = sdtPrRecord(ctx, suggestion.id, after, merge);
  relabel(ctx.doc, ctx.story, open, { props: [...open.props, record] });
  return suggestion.id;
}

/** A copy of a node without `w14:paraId` and `w14:textId`: the change's minter gives new ones. */
function withoutParagraphIds(node: OoxmlElement, depth = 0): OoxmlElement {
  if (depth > 64) return node;
  const children = node.children.map((child) =>
    child.kind === 'textValue' ? child : withoutParagraphIds(child, depth + 1)
  );
  const attributes = node.attributes.filter((a) => !isParagraphIdentity(a));
  return withAttributes(withChildren(node, children), attributes);
}

/**
 * A gallery control's content replaced by a glossary block's body, as the engine inserts it:
 * what the control showed is struck with the edit's suggestion (a direct edit settles it), and
 * the body's runs (an inline control) or blocks (a block control) are inserted before its end.
 * A temporary control goes too, leaving the body. The block is the glossary's, named: every
 * copy reads it from the same Word file.
 */
export function insertBuildingBlock(ctx: Ctx, open: OrderedItem, name: string): string {
  directOnly(ctx, 'a building block');
  shownControl(ctx, open);
  const blocks = ctx.doc.placeholders.get(name);
  if (!blocks)
    throw new OrderedRefusal(
      'engine-refused',
      `insertBuildingBlock: no block ${name.slice(0, 64)}`
    );
  const id = shellId(open);
  const op: TreeDocOp = { op: 'insertBuildingBlock', controlId: id, name, blocks };
  const { view, result } = engineOnView(ctx, op);
  const after = findNode(result, id);
  const filled = after ? ensureElement(after) : unwrappedControl(view, result, id);
  const suggestion = suggestionFor(ctx);
  const close = closeOf(ctx.story, open, ctx.doc);
  const inside = itemsInside(ctx.story, open, ctx.doc).filter((item) => !isHidden(ctx.doc, item));

  // A block control's old blocks go through `within`, as a struck row's content does
  if (open.token === 'sdtBlock') {
    for (const item of inside) {
      holdWithin(ctx, ctx.story, item, suggestion.id);
    }
  } else {
    strikeItems(ctx, ctx.story, inside, suggestion.id);
  }

  const labels: ControlLabels = { suggestion: suggestion.id, within: open.within };
  const items =
    open.token === 'sdt'
      ? contentUnits(ctx, filled, labels, false)
      : bodyItems(ctx, filled, labels);
  insertItems(ctx.doc, ctx.story, ctx.story.tree.prev(close), items);

  // A temporary control unwraps once filled, its content kept
  if (!after) {
    strikeItems(ctx, ctx.story, [open, close], suggestion.id);
  }

  return suggestion.id;
}

function ensureElement(node: OoxmlNode | null): OoxmlElement {
  if (!node || node.kind === 'textValue')
    throw new OrderedRefusal('engine-refused', 'insertBuildingBlock: no control');
  return node;
}

/**
 * A temporary control the engine filled and then removed, keeping its content in its place: the
 * view's control holding the nodes the engine put in its parent.
 */
function unwrappedControl(view: OoxmlPart, result: OoxmlPart, id: string): OoxmlElement {
  const control = viewNode(view, id);
  const parent = ensureElement(parentOf(view, id));
  const filledParent = ensureElement(findNode(result, parent.id));
  const old = new Set(parent.children.map((child) => child.id));
  const kept = filledParent.children.filter((child) => !old.has(child.id));
  return withChildren(
    control,
    control.children.map((child) =>
      child.kind !== 'textValue' && isControlContent(child) ? withChildren(child, kept) : child
    )
  );
}

/** A block control's new blocks as items, every one inserted by the edit's suggestion. */
function bodyItems(ctx: Ctx, control: OoxmlElement, labels: ControlLabels): OrderedItem[] {
  const content = control.children.find(isControlContent);
  if (!content || content.kind === 'textValue') return [];

  // The new blocks go with the edit through `within`: words typed in them later too
  const within = { by: [...labels.within.by, labels.suggestion], dels: labels.within.dels };
  return content.children.flatMap((block) =>
    block.kind === 'textValue'
      ? []
      : templateItems(ctx, withoutParagraphIds(block), labels.suggestion).map(
          (item): OrderedItem => ({ ...item, by: item.by ?? labels.suggestion, within })
        )
  );
}
