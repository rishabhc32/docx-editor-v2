// Content the engine keeps from edits: a table of contents, a control bound to custom XML, and a
// content-locked control. An edit that reaches it is refused by name, as the engine refuses it.

import type { OoxmlElement } from '../../store/package/ooxml-tree.ts';
import {
  lockForbidsEdit,
  lockForbidsRemoval,
  lockInForce,
  lockPropertiesIn,
  ownContentLockOf,
  type ControlLockProperties,
} from '../../store/package/content-control-nodes.ts';
import { formattingPassesContentLock } from '../../store/store/content-lock-policy.ts';
import { refusesTypedContent } from '../../store/store/tree-op-nodes.ts';
import type { OoxmlProperty } from '../../store/store/tree-op-types.ts';
import { orderedViewParagraph } from './compose/compose.ts';
import {
  ComposeContext,
  type OrderedViewChar,
  type OrderedViewParagraph,
} from './compose/context.ts';
import { containerPropertiesAfter } from './compose/properties.ts';
import { isControlSegment, openOfSegment, segmentsOf, sharedDepth } from './containers.ts';
import type { Ctx } from './edits/context.ts';
import { isHidden } from './model/document.ts';
import { isControl } from './model/prompt-marks.ts';
import { OrderedRefusal } from './model/refusal.ts';
import type { OrderedItem } from './model/types.ts';
import { attributeOf, childNamed, isWml, withChildren } from './ooxml-nodes.ts';
import { reachesTocField } from './toc-field-range.ts';

/** A control's properties (`w:sdtPr`), or undefined. */
function propertiesOf(open: OrderedItem): OoxmlElement | undefined {
  const node = open.node;
  return node && node.kind !== 'textValue' ? childNamed(node, 'sdtPr') : undefined;
}

/** A control whose content comes from custom XML (`w:dataBinding`). */
function isBound(open: OrderedItem): boolean {
  return propertiesOf(open)?.children.some((c) => isWml(c, 'dataBinding')) === true;
}

/** A block control that holds a table of contents (Word's gallery of that name). */
export function isTableOfContents(open: OrderedItem): boolean {
  const properties = propertiesOf(open);
  const docPartObj = properties === undefined ? undefined : childNamed(properties, 'docPartObj');
  const gallery = docPartObj === undefined ? undefined : childNamed(docPartObj, 'docPartGallery');
  return gallery !== undefined && attributeOf(gallery, 'val') === 'Table of Contents';
}

/**
 * A control's element with its properties as they show now: the file's, or the organiser's later
 * change of them (`control-property`), a record on the open item that compose writes. Undefined
 * for an item that is no control.
 */
function shownControl(ctx: Ctx, open: OrderedItem): OoxmlElement | undefined {
  const node = open.node;
  if ((open.token !== 'sdt' && open.token !== 'sdtBlock') || !node || node.kind === 'textValue') {
    return undefined;
  }

  if (open.props.length === 0) {
    return node;
  }

  const compose = new ComposeContext(ctx.doc, ctx.story, {});
  return withChildren(node, containerPropertiesAfter(compose, open, node.children));
}

/** What the lock rule reads of a control (its `w:lock` and its type) as it shows now. */
function lockPropertiesOf(ctx: Ctx, open: OrderedItem): ControlLockProperties {
  const shown = shownControl(ctx, open);
  return lockPropertiesIn(shown === undefined ? undefined : childNamed(shown, 'sdtPr'));
}

/** The refusal of an edit in a control's own content that its lock keeps. */
function lockedRefusal(control: ControlLockProperties | undefined): OrderedRefusal {
  const what =
    control?.type === 'group'
      ? "a group control's own content"
      : 'the content of a content-locked control';
  return new OrderedRefusal('locked', what);
}

/**
 * Refuse a change at a place whose lock in force keeps it (the engine's `lockInForce`, ECMA-376
 * Part 1 §17.5.2.23): `containers` are the containers around the place, outermost first (a
 * paragraph's blocks, then the inline ones). A content lock keeps the control's own content, not
 * a control nested in it.
 */
export function refuseLockedPlace(ctx: Ctx, containers: readonly OrderedItem[]): void {
  const chain = lockingChain(ctx, containers);
  if (chain !== null) {
    throw lockedRefusal(chain.at(-1));
  }
}

/** The controls around a place (`containers`, outermost first) when their lock keeps it, or null. */
function lockingChain(
  ctx: Ctx,
  containers: readonly OrderedItem[]
): ControlLockProperties[] | null {
  const chain = containers.filter(isControl).map((open) => lockPropertiesOf(ctx, open));
  return lockForbidsEdit(lockInForce(chain)) ? chain : null;
}

/**
 * The inline control that typed text enters at a caret on its leading edge when no control is
 * named, as the engine enters it: only where the place beside it is locked (a group's own content,
 * say), so the control is the one place there that takes the text, and only when typing may enter
 * it (`refusesTyping`). Elsewhere typed text goes beside a control (placement rule 1).
 */
export function controlEnteredAt(
  ctx: Ctx,
  p: OrderedViewParagraph,
  offset: number
): OrderedItem | undefined {
  const before = segmentsOf(p.chars[offset - 1]?.containers ?? '');
  const after = segmentsOf(p.chars[offset]?.containers ?? '');
  const shared = before.slice(0, sharedDepth(before, after));
  const starting = after[shared.length];
  const entered =
    starting !== undefined && isControlSegment(starting)
      ? openOfSegment(ctx.story, starting)
      : null;
  if (entered === null || refusesTyping(ctx, entered)) {
    return undefined;
  }

  return lockingChain(ctx, [...p.blocks, ...controlsOf(ctx, shared)]) === null
    ? undefined
    : entered;
}

/** The open items of the inline controls that container segments name, outermost first. */
function controlsOf(ctx: Ctx, segments: readonly string[]): OrderedItem[] {
  return segments
    .filter(isControlSegment)
    .map((segment) => openOfSegment(ctx.story, segment))
    .filter((open): open is OrderedItem => open !== null);
}

/**
 * The inline controls a deletion of view characters `[start, end)` in one paragraph takes whole,
 * by their segments, as the editor removes them (`coveredLockedControlOps`): each outermost
 * control it covers whole whose content the lock in force keeps (a content lock, a group). The
 * engine's removal of the control then answers for its characters (`strikeControl`).
 */
export function controlsTakenWhole(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number
): Map<string, OrderedItem> {
  const placesOf = (chars: readonly OrderedViewChar[]) =>
    new Set(chars.map((char) => char.containers));
  const outside = new Set(
    [...placesOf([...p.chars.slice(0, start), ...p.chars.slice(end)])].flatMap((containers) =>
      segmentsOf(containers)
    )
  );
  const taken = new Map<string, OrderedItem>();
  for (const containers of placesOf(p.chars.slice(start, end))) {
    const segments = segmentsOf(containers);
    for (const [index, segment] of segments.entries()) {
      if (taken.has(segment)) {
        break;
      }

      if (!isControlSegment(segment) || outside.has(segment)) {
        continue;
      }

      const line = controlsOf(ctx, segments.slice(0, index + 1));
      const open = line.at(-1);
      if (open !== undefined && lockingChain(ctx, [...p.blocks, ...line]) !== null) {
        taken.set(segment, open);
        break;
      }
    }
  }

  return taken;
}

/** The inline controls a view character sits in that `kept` names, by their segments. */
function keptControls(char: OrderedViewChar | undefined, kept: (segment: string) => boolean) {
  if (!char) {
    return [];
  }

  return segmentsOf(char.containers).filter(
    (segment) => isControlSegment(segment) && kept(segment)
  );
}

/**
 * Whether an edit of view characters `[start, end)` (a caret when they are equal) reaches inside
 * an inline control `kept` names: a character in it, or a caret with it on both sides. At its
 * edges a caret is outside it (placement rules 1 and 5).
 */
function reachesInside(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number,
  kept: (open: OrderedItem) => boolean
): boolean {
  // Each control once, however many characters of a long range sit in it
  const known = new Map<string, boolean>();
  const isKept = (segment: string): boolean => {
    let answer = known.get(segment);
    if (answer === undefined) {
      const open = openOfSegment(ctx.story, segment);
      answer = open !== null && kept(open);
      known.set(segment, answer);
    }

    return answer;
  };
  if (end > start) {
    return p.chars.slice(start, end).some((char) => keptControls(char, isKept).length > 0);
  }

  const before = keptControls(p.chars[start - 1], isKept);
  return before.length > 0 && keptControls(p.chars[start], isKept).some((c) => before.includes(c));
}

/**
 * Refuse a change of `items` (a row's, a table's, a paragraph's) when a control among them that
 * shows keeps its own content, or, when the change removes them (`removes`), when a control among
 * them may not be removed (`sdtLocked`, `sdtContentLocked`).
 */
export function refuseLockedWithin(
  ctx: Ctx,
  items: readonly OrderedItem[],
  removes: boolean
): void {
  for (const item of items) {
    if (!isControl(item) || isHidden(ctx.doc, item)) {
      continue;
    }

    const control = lockPropertiesOf(ctx, item);
    const lock = ownContentLockOf(control);
    if (lockForbidsEdit(lock)) {
      throw lockedRefusal(control);
    }

    if (removes && lockForbidsRemoval(lock)) {
      throw new OrderedRefusal('locked', 'a control that may not be removed');
    }
  }
}

/** The shown paragraph that holds an item (a mark, a unit, a container's edge), as drawn now. */
export function paragraphOfItem(ctx: Ctx, item: OrderedItem): OrderedViewParagraph | null {
  const p = orderedViewParagraph(ctx.doc, ctx.story.name, ctx.story.tree.before(item).marks);
  return p && (p.mark === item || p.items.includes(item)) ? p : null;
}

/**
 * Refuse a change of one item at a place whose lock in force keeps it: a paragraph's mark (a
 * section break), an object (a picture's size, an equation), a note reference (a conversion), a
 * link's open item (its target). For a container's open item only the containers around it
 * count, as the engine asks a node's enclosing controls.
 */
export function refuseLockedAt(ctx: Ctx, item: OrderedItem): void {
  const p = paragraphOfItem(ctx, item);
  if (!p) {
    return;
  }

  const own = `${item.token === 'hyperlink' ? 'h' : 'c'}:${item.space}`;
  const char =
    item.kind === 'mark'
      ? undefined
      : p.chars.find((c) =>
          item.kind === 'open' ? segmentsOf(c.containers).includes(own) : c.item === item
        );
  const segments = char ? segmentsOf(char.containers) : [];
  const around = item.kind === 'open' ? segments.slice(0, segments.indexOf(own)) : segments;
  refuseLockedPlace(ctx, [...p.blocks, ...controlsOf(ctx, around)]);
}

/**
 * Refuse a change of the table a paragraph is in (a row, a column, a cell's properties) when the
 * lock in force around the table keeps it. The engine's own op on the table meets the controls
 * in it.
 */
export function refuseLockedTable(ctx: Ctx, p: OrderedViewParagraph): void {
  const table = p.blocks.map((block) => block.token).lastIndexOf('tbl');
  refuseLockedPlace(ctx, table < 0 ? p.blocks : p.blocks.slice(0, table));
}

/**
 * Refuse an edit of view characters `[start, end)` (a caret when they are equal) whose content
 * the lock in force keeps (`refuseLockedPlace`): each character in the controls around it. A caret
 * is in the controls on both its sides, or in `inside`, the control the keyboard types in. At an
 * inline control's edge, typed text goes beside it (`promptSite` never enters it).
 */
export function refuseContentLock(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number,
  inside?: OrderedItem
): void {
  const refuseIn = (segments: readonly string[]) =>
    refuseLockedPlace(ctx, [...p.blocks, ...controlsOf(ctx, segments)]);
  if (end > start) {
    // Each place once, however many characters of a long range sit in it
    const places = new Set(p.chars.slice(start, end).map((char) => char.containers));
    for (const containers of places) {
      refuseIn(segmentsOf(containers));
    }

    return;
  }

  const before = segmentsOf(p.chars[start - 1]?.containers ?? '');
  const after = segmentsOf(p.chars[start]?.containers ?? '');
  const shared = before.slice(0, sharedDepth(before, after));
  if (inside === undefined) {
    refuseIn(shared);
    return;
  }

  const own = `c:${inside.space}`;
  const side = [after, before].find((segments) => segments.includes(own));
  refuseIn(side ? side.slice(0, side.indexOf(own) + 1) : [...shared, own]);
}

/**
 * Whether a format of `properties` passes a content lock: a direct format that the room's policy
 * lets past it, as the engine lets an untracked restyle (`restylePassesContentLock`). A suggested
 * format stays refused, as the engine refuses a tracked restyle there.
 */
export function formatPassesContentLock(ctx: Ctx, properties: readonly OoxmlProperty[]): boolean {
  return ctx.direct && formattingPassesContentLock(properties, ctx.contentLockPolicy);
}

/**
 * Whether typing never enters a control, by the engine's own rule (`refusesTypedContent`): bound,
 * its own content locked, or a checkbox or picture control.
 */
export function refusesTyping(ctx: Ctx, open: OrderedItem): boolean {
  const shown = shownControl(ctx, open);
  return shown !== undefined && refusesTypedContent(shown);
}

/**
 * Refuse an edit of view characters `[start, end)` (a caret when they are equal) that reaches
 * content the engine keeps from edits: a table of contents, a bound control's content, a locked
 * control's content, a group control's own content. A format names its `properties`: a locked
 * control's content may then take it (`formatPassesContentLock`). Typing names the control the
 * keyboard types in (`inside`), where its text goes.
 */
export function refuseLocked(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number,
  options: { readonly properties?: readonly OoxmlProperty[]; readonly inside?: OrderedItem } = {}
): void {
  const { properties, inside } = options;
  const toc = p.blocks.find(isTableOfContents);
  if (toc && reachesTocField(ctx, toc, p, start, end)) {
    throw new OrderedRefusal('not-covered', 'a table of contents result is generated');
  }

  if (reachesInside(ctx, p, start, end, isBound)) {
    throw new OrderedRefusal('not-covered', 'a control bound to custom XML');
  }

  if (properties === undefined || !formatPassesContentLock(ctx, properties)) {
    refuseContentLock(ctx, p, start, end, inside);
  }
}
