// Where an edit puts its items: the engine's placement rules, in one place, so a view and a
// saved file read as the engine's suggesting mode writes them. The matrix suite in
// `__tests__/matrix/` checks the numbered rules against the engine. Rules 8 and 9 are view rules
// that compose draws (`compose/inline.ts`, `compose/prompts.ts`).

import type { OoxmlElement, OoxmlNode } from '../../store/package/ooxml-tree.ts';
import {
  containerIsComplexScript,
  withComplexScriptCompanions,
} from '../../store/store/complex-script-companions.ts';
import {
  declaredLockOf,
  lockFlagsFromValue,
  type ContentControlLockFlags,
} from '../../store/store/tree-op-nodes.ts';
import type { OoxmlProperty } from '../../store/store/tree-op-types.ts';
import { formattingPassesContentLock } from '../../store/store/content-lock-policy.ts';
import { orderedViewParagraph } from './compose/compose.ts';
import {
  ComposeContext,
  type OrderedViewChar,
  type OrderedViewParagraph,
} from './compose/context.ts';
import { keptLinkEdges } from './compose/kept-links.ts';
import { showsPrompt } from './compose/paragraph.ts';
import { containerPropertiesAfter } from './compose/properties.ts';
import {
  closesContainer,
  innermostSegment,
  isControlSegment,
  openOfSegment,
  segmentsOf,
  sharedDepth,
} from './containers.ts';
import { currentFace, withoutPlaceholderStyle, type Ctx } from './edits/context.ts';
import { isHidden, itemAt, pendingByOther } from './model/document.ts';
import { insertItems, splitItem } from './model/mutations.ts';
import { OrderedRefusal } from './model/refusal.ts';
import {
  isUnit,
  type OrderedDocument,
  type OrderedItem,
  type OrderedStory,
} from './model/types.ts';
import {
  attributeOf,
  childNamed,
  isParagraphProperties,
  isWml,
  withChildren,
} from './ooxml-nodes.ts';
import { reachesTocField } from './toc-field-range.ts';

/** Where new items go. */
export interface Site {
  /** Right after this unit (split there), then past markers as the flags say. */
  readonly after: { readonly item: OrderedItem; readonly offset: number } | null;
  /** Or right before this item (split at the neighbour's unit when it is text). */
  readonly before: OrderedItem | null;
  /** The character whose face new text takes, and whose suggestion it may extend. */
  readonly neighbour: OrderedViewChar | undefined;
  /** The face new text takes when it is not the neighbour's. */
  readonly face?: OoxmlElement | null;
  /** Past the markers with no width right after `after` (rule 1). */
  readonly pastMarkers?: boolean;
  /** Past range ends too (rule 2). */
  readonly pastRangeEnds?: boolean;
}

// ---- markers ----

function zeroNamed(item: OrderedItem, test: (name: string) => boolean): boolean {
  const node = item.node;
  return (
    item.kind === 'zero' &&
    !item.block &&
    !item.inRun &&
    node !== null &&
    node.kind !== 'textValue' &&
    test(node.localName)
  );
}

/** A bookmark's start or end between runs: new text goes past it (rule 1). */
export function isBookmarkMarker(item: OrderedItem): boolean {
  return zeroNamed(item, (name) => name === 'bookmarkStart' || name === 'bookmarkEnd');
}

const RANGE_ENDS = new Set(['commentRangeEnd', 'permEnd', 'moveFromRangeEnd', 'moveToRangeEnd']);

/** A range's end between runs: a paragraph break leaves it in the first paragraph (rule 2). */
export function isRangeEnd(item: OrderedItem): boolean {
  return zeroNamed(item, (name) => RANGE_ENDS.has(name));
}

/** A field's end in a run of its own, after a result the caret can reach (rule 1). */
export function isFieldEnd(item: OrderedItem): boolean {
  const node = item.node;
  return (
    item.kind === 'zero' &&
    item.inRun === true &&
    node !== null &&
    node.kind !== 'textValue' &&
    node.localName === 'fldChar' &&
    attributeOf(node, 'fldCharType') === 'end'
  );
}

/**
 * The paragraph's first item, past the bookmark markers that open it when asked (rule 1). At a
 * paragraph start typed text goes past bookmark starts only.
 */
export function firstItemOf(p: OrderedViewParagraph, pastBookmarks = false): OrderedItem {
  const first = pastBookmarks ? p.items.find((item) => !isBookmarkMarker(item)) : p.items[0];
  return first ?? p.mark;
}

// ---- containers ----

/** The containers two characters share, as the view names them. */
export function commonContainers(a: OrderedViewChar, b: OrderedViewChar): string {
  const outer = segmentsOf(a.containers);
  return outer.slice(0, sharedDepth(outer, segmentsOf(b.containers))).join('>');
}

/**
 * Whether the next character is in every link or control the previous one is in. Containers
 * compare by whole segments: a control id can be a prefix of another's (`#0.1`, `#0.10`).
 */
export function staysInside(previous: OrderedViewChar, next: OrderedViewChar): boolean {
  const outer = previous.containers;
  return outer === '' || next.containers === outer || next.containers.startsWith(`${outer}>`);
}

/** The items between an open item and its close item, in order. */
export function itemsInside(
  story: OrderedStory,
  open: OrderedItem,
  doc?: OrderedDocument
): OrderedItem[] {
  const out: OrderedItem[] = [];
  for (
    let item = story.tree.next(open);
    item && !closesContainer(item, open, doc);
    item = story.tree.next(item)
  )
    out.push(item);
  return out;
}

/** The close item of an open item. */
export function closeOf(
  story: OrderedStory,
  open: OrderedItem,
  doc?: OrderedDocument
): OrderedItem {
  for (let item = story.tree.next(open); item; item = story.tree.next(item))
    if (closesContainer(item, open, doc)) return item;
  throw new OrderedRefusal('malformed-story', `no close item for ${open.space}`);
}

/**
 * The close item of the last container `prev` sits in that the next character is not in, when
 * only markers, close items and hidden units lie between them: the caret is at their end. It
 * reads the items of `p` as the view does: an item that does not show ends and starts nothing.
 * An Enter taken back leaves its split's close, mark and open there, hidden; counting that
 * close would put a break inside the container, before its own close, with no split of it.
 */
function closeAfter(
  ctx: Ctx,
  p: OrderedViewParagraph,
  prev: OrderedViewChar,
  nextContainers: string
): OrderedItem | null {
  const inside = segmentsOf(prev.containers);
  let toClose = inside.length - sharedDepth(inside, segmentsOf(nextContainers));
  if (toClose === 0) {
    return null;
  }

  let kept: ReadonlySet<OrderedItem> | null = null;

  // As `viewOfParagraph` reads it: a hidden edge counts only as a rejected new link kept whole
  const shows = (item: OrderedItem): boolean => {
    if (!isHidden(ctx.doc, item)) return true;
    if (item.kind !== 'open' && item.kind !== 'close') return false;
    kept ??= keptLinkEdges(new ComposeContext(ctx.doc, ctx.story, {}), p.items);
    return kept.has(item);
  };
  for (let item = ctx.story.tree.next(prev.item); item; item = ctx.story.tree.next(item)) {
    if (item.kind === 'zero' || !shows(item)) {
      continue;
    }

    if (item.kind !== 'close') {
      return null;
    }

    toClose -= 1;
    if (toClose === 0) {
      return item;
    }
  }

  return null;
}

/** Before the open items of the containers `char` sits in that start right before it. */
function outsideStart(p: OrderedViewParagraph, char: OrderedViewChar): OrderedItem {
  const at = p.items.indexOf(char.item);
  let k = at;
  while (k > 0 && p.items[k - 1]?.kind === 'open') k -= 1;
  return p.items[k] ?? char.item;
}

// ---- faces ----

/**
 * The face of the text right after an object: the run the engine types into (rule 4). Typed text
 * takes the next run's face, unless a link or control starts there.
 */
function followingFace(
  ctx: Ctx,
  p: OrderedViewParagraph,
  previous: OrderedViewChar,
  offset: number
): OoxmlElement | null {
  const next = p.chars[offset];

  // The next run's face, text or object, unless it starts a link or control (no face then)
  if (next && next.containers !== previous.containers) return null;
  if (next && isUnit(next.item)) return currentFace(ctx, next.item);

  // An empty run after the object: its face
  const at = p.items.indexOf(previous.item);
  const empty =
    at >= 0
      ? p.items.slice(at + 1).find((item) => item.kind === 'zero' && item.run !== previous.item.run)
      : undefined;
  return empty ? currentFace(ctx, empty) : null;
}

/** Whether a pending insertion by another author holds the character. */
function foreignInsertion(ctx: Ctx, char: OrderedViewChar): boolean {
  return pendingByOther(ctx.doc, char.item.by, ctx.attribution);
}

// ---- sites ----

/**
 * Text typed at the front of words the same edit struck replaces them (rule 13, the engine's
 * `replacesThisEdit`; a link proposal aims there): it goes right after them, inside the link or
 * control that holds them. A strike of an earlier edit keeps the caret before it. Null when the
 * caret is not at such words.
 */
function replacingSite(ctx: Ctx, chars: readonly OrderedViewChar[], offset: number): Site | null {
  const join = ctx.direct || ctx.earlierStrike === true ? undefined : ctx.join;
  if (join === undefined || !chars[offset]?.deleted || !chars[offset]?.item.dels.includes(join)) {
    return null;
  }

  let end = offset;
  while (end < chars.length && chars[end]?.item.dels.includes(join)) {
    end += 1;
  }

  const last = chars[end - 1];
  if (!last) {
    return null;
  }

  return {
    after: { item: last.item, offset: last.address.offset },
    before: null,
    neighbour: last,
    pastMarkers: true,
  };
}

/** The end of the struck text a caret is in: the engine moves the caret there first (rule 5). */
function pastStruck(chars: readonly OrderedViewChar[], offset: number): number {
  let end = offset;
  while (end < chars.length && chars[end]?.deleted) {
    end += 1;
  }

  return end;
}

/**
 * Where typed text goes at a view offset (rules 1, 4, 5 and 13).
 *
 * Rule 1: typed text goes right after the run before the caret, past bookmark starts and ends
 * and a field's end (a form field's result is text the caret can reach), and before every other
 * marker (comment and permission ranges, proofing marks, empty runs, a positional tab in its own
 * run).
 *
 * Rule 5: leaving a link or control, typed text goes right after its end. After another
 * author's insertion, it takes the face of the next text that shows. Beside struck text (inside
 * it, right after it, or before it at a paragraph's start) it takes the struck run's face, the
 * one before the caret first, without its change record.
 */
export function siteAt(ctx: Ctx, p: OrderedViewParagraph, at: number, skipStruck = true): Site {
  const chars = p.chars;
  if (skipStruck) {
    const replacing = replacingSite(ctx, chars, at);
    if (replacing) {
      return replacing;
    }
  }

  // In struck text the new text takes the face of the struck character before the caret
  const inStruck = skipStruck && chars[at - 1]?.deleted === true && chars[at]?.deleted === true;
  const offset = inStruck ? pastStruck(chars, at) : at;
  const previous = offset > 0 ? chars[offset - 1] : undefined;
  const next = chars[offset];

  // At the start: past the bookmark markers that open the paragraph, before anything else
  const startOfParagraph: Site = {
    after: null,
    before: firstItemOf(p, true),
    neighbour: undefined,
  };
  if (!previous) {
    if (!next) {
      return startOfParagraph;
    }

    // A paragraph that starts in a link or control: outside it, with no face
    if (next.containers !== '') {
      return { ...startOfParagraph, neighbour: next, face: null };
    }

    return { after: null, before: next.item, neighbour: next };
  }

  const afterObject = previous.ch === '￼' ? followingFace(ctx, p, previous, offset) : undefined;
  const afterPrevious = { item: previous.item, offset: previous.address.offset };

  // After an object, at the paragraph's end or in the same containers: the face rule 4 gives
  if (afterObject !== undefined && (!next || staysInside(previous, next))) {
    return {
      after: afterPrevious,
      before: null,
      neighbour: previous,
      face: afterObject,
      pastMarkers: true,
    };
  }

  if (!next) {
    // At the end: outside a link or control the paragraph ends in; else right after the last
    // character, past bookmark ends, before other markers
    if (previous.containers !== '') {
      return { after: null, before: p.mark, neighbour: previous };
    }

    return { after: afterPrevious, before: null, neighbour: previous, pastMarkers: true };
  }

  if (!staysInside(previous, next)) {
    // Leaving a link or control: right after its end, past bookmarks, before other markers;
    // the next run's face unless it opens a container here
    const close = closeAfter(ctx, p, previous, next.containers);
    if (close) {
      const opensHere = next.containers.length > commonContainers(previous, next).length;
      return {
        after: { item: close, offset: close.offset },
        before: null,
        neighbour: next,
        face: opensHere ? null : currentFace(ctx, next.item),
        pastMarkers: true,
      };
    }

    // Between two of them if the next one opens here
    const outside = next.containers !== '';
    return { after: null, before: outside ? outsideStart(p, next) : next.item, neighbour: next };
  }

  // Between two characters: right after the previous character's run (with the markers in it)
  // and past bookmark markers, before any other marker. After struck text: the struck run's
  // face (the neighbour's)
  const between: Site = {
    after: afterPrevious,
    before: null,
    neighbour: previous,
    pastMarkers: true,
  };
  if (previous.deleted || !foreignInsertion(ctx, previous)) {
    return between;
  }

  // After another author's pending insertion: the face of the next text that shows
  const sameContainers = next.containers === previous.containers;
  return { ...between, face: sameContainers ? currentFace(ctx, next.item) : null };
}

/**
 * Typing at a control's shown prompt, at its edge or in it (rule 7): into the control, after
 * the prompt, in the face of the prompt's first words without the placeholder style, wherever
 * the caret is in it (the engine empties the prompt into its first run). The typed text replaces
 * the prompt; deleting the prompt does nothing (that skip is in `edits/text.ts`). Null when no
 * prompt is there. In a control's prompt that holds another control showing its own, the outer
 * prompt is replaced, as the engine replaces the outermost (`placeholderControlForInsertion`).
 */
export function promptSite(ctx: Ctx, p: OrderedViewParagraph, offset: number): Site | null {
  const at = [p.chars[offset], offset > 0 ? p.chars[offset - 1] : undefined].find(
    (char) => char?.item.prompt === true
  );
  if (!at) {
    return null;
  }

  const shown = new ComposeContext(ctx.doc, ctx.story, {});
  const showsOwn = (segment: string) => {
    const open = isControlSegment(segment) ? openOfSegment(ctx.story, segment) : null;
    return open !== null && showsPrompt(shown, open);
  };
  const segment = segmentsOf(at.containers).find(showsOwn) ?? innermostSegment(at.containers);
  const first = p.chars.find(
    (char) =>
      char.item.prompt === true &&
      char.item.kind === 'text' &&
      innermostSegment(char.containers) === segment
  );
  const face = withoutPlaceholderStyle(currentFace(ctx, (first ?? at).item));
  if (!isControlSegment(segment)) {
    // A block control's prompt: after the paragraph's last prompt unit
    const last = [...p.items]
      .reverse()
      .find((item) => item.prompt === true && item.kind !== 'mark');
    if (!last) {
      return null;
    }

    return {
      after: { item: last, offset: last.offset + Math.max(0, last.length - 1) },
      before: null,
      neighbour: undefined,
      face,
    };
  }

  const open = openOfSegment(ctx.story, segment);

  // A control typing never enters keeps its prompt: text at its edge goes beside it
  if (!open || open.kind !== 'open' || refusesTyping(ctx, open)) {
    return null;
  }

  return { after: null, before: closeOf(ctx.story, open, ctx.doc), neighbour: undefined, face };
}

/**
 * Typing at an edge of the inline control the keyboard types in (the engine's `inside`): into
 * it, after its last character or before its first, where rule 5 puts other text beside it.
 */
export function insideSite(
  ctx: Ctx,
  p: OrderedViewParagraph,
  offset: number,
  open: OrderedItem
): Site | null {
  const segment = `c:${open.space}`;
  const holds = (char: OrderedViewChar | undefined): char is OrderedViewChar =>
    char !== undefined && segmentsOf(char.containers).includes(segment);
  const previous = offset > 0 ? p.chars[offset - 1] : undefined;
  const next = p.chars[offset];

  // Strictly inside it `siteAt` places the text in it; away from it (another person's edit took
  // the caret out before this one arrived) beside it
  if (holds(previous) === holds(next)) {
    return null;
  }

  // At its end: right after its last character, or after a control inside it that ends there
  if (holds(previous)) {
    if (innermostSegment(previous.containers) !== segment) {
      return { after: null, before: closeOf(ctx.story, open, ctx.doc), neighbour: previous };
    }

    const after = { item: previous.item, offset: previous.address.offset };
    return { after, before: null, neighbour: previous, pastMarkers: true };
  }

  // At its start: before its first character, or before a control inside it that starts there
  const first = next && innermostSegment(next.containers) === segment ? next.item : null;
  return { after: null, before: first ?? ctx.story.tree.next(open), neighbour: next };
}

/** Where a paragraph break goes, and the inline containers it splits (outermost first). */
export interface BreakSite {
  readonly site: Site;
  readonly split: readonly OrderedItem[];
}

/**
 * Where a paragraph break goes (rules 2 and 6). Enter at the end of a link or control puts the
 * break after it; else the break goes where typed text goes, and also past range ends (comment
 * and permission), which stay in the first paragraph. Inside links and controls it splits them,
 * as the engine does: the break goes between the two characters and `split` names the containers
 * both sit in. Refused in a place where the characters on its two sides sit in no container they
 * share.
 */
export function breakSite(ctx: Ctx, p: OrderedViewParagraph, offset: number): BreakSite {
  const prev = offset > 0 ? p.chars[offset - 1] : undefined;
  const next = p.chars[offset];
  const close = prev?.containers ? closeAfter(ctx, p, prev, next?.containers ?? '') : null;
  if (close)
    return {
      site: {
        after: { item: close, offset: close.offset },
        before: null,
        neighbour: prev,
        pastMarkers: true,
        pastRangeEnds: true,
      },
      split: [],
    };
  if (!prev?.containers) {
    const found = siteAt(ctx, p, offset, true);
    if (found.neighbour?.containers && found.after)
      throw new OrderedRefusal('break-in-container', 'Enter beside struck text in a container');
    return { site: found.pastMarkers ? { ...found, pastRangeEnds: true } : found, split: [] };
  }

  const shared = next ? commonContainers(prev, next) : '';
  if (!shared) throw new OrderedRefusal('break-in-container', 'Enter at a container edge');
  const split = segmentsOf(shared).map((segment) => {
    const open = openOfSegment(ctx.story, segment);
    if (!open || open.kind !== 'open') {
      throw new OrderedRefusal('malformed-story', `no open item for ${segment}`);
    }

    return open;
  });

  // Right after the character before it: inside a container, the markers at the caret go down
  // with it, as the engine's split sends them to the second part
  return {
    site: {
      after: { item: prev.item, offset: prev.address.offset },
      before: null,
      neighbour: prev,
    },
    split,
  };
}

/** Put new items at a site. */
export function insertAtSite(ctx: Ctx, site: Site, items: readonly OrderedItem[]): void {
  const { doc, story } = ctx;
  if (site.after) {
    const address = { space: site.after.item.space, offset: site.after.offset };
    let target = itemAt(story, address);
    if (!target) throw new OrderedRefusal('unknown-address', `${address.space}|${address.offset}`);
    if (target.kind === 'text' && address.offset + 1 < target.offset + target.length)
      target = splitItem(doc, story, target, address.offset + 1)?.[0] ?? target;
    if (site.pastMarkers) {
      // Past the rest of the character's run and past bookmark markers; nothing else
      const run = target.run;
      for (
        let next = story.tree.next(target);
        next &&
        next.kind === 'zero' &&
        !next.block &&
        ((next.inRun === true && next.run === run) ||
          isBookmarkMarker(next) ||
          isFieldEnd(next) ||
          (site.pastRangeEnds === true && isRangeEnd(next)));
        next = story.tree.next(next)
      )
        target = next;
    }

    insertItems(doc, story, target, items);
    return;
  }

  let before = site.before;
  if (!before) throw new RangeError('A site with no place.');
  if (before.kind === 'text') {
    // A unit inside an item: split there
    const char = site.neighbour && site.neighbour.item === before ? site.neighbour : null;
    const at = char ? char.address.offset : before.offset;
    const current = itemAt(story, { space: before.space, offset: at });
    if (!current) throw new OrderedRefusal('unknown-address', `${before.space}|${at}`);
    before = at > current.offset ? (splitItem(doc, story, current, at)?.[1] ?? current) : current;
  }

  insertItems(doc, story, story.tree.prev(before), items);
}

// ---- markers that go with their run (rule 3) ----

/**
 * The markers inside runs (a rendered page break, a positional tab) that go with the units an
 * edit reached, as the engine splits runs: a marker between struck units goes with the unit
 * before it in its run, and one at the range's end stays with the text after it. A marker that
 * starts its run stays out of a deletion; under a format it goes with the unit after it
 * (`either`), as the run carries the properties.
 */
export function inRunMarkers(
  ctx: Ctx,
  chars: readonly OrderedViewChar[],
  reached: (item: OrderedItem) => boolean,
  either: boolean
): OrderedItem[] {
  const tree = ctx.story.tree;
  const firstChar = chars[0];
  const lastChar = chars[chars.length - 1];
  if (!firstChar || !lastChar) return [];
  const isMarker = (item: OrderedItem | null, run: string): boolean =>
    item !== null && item.kind === 'zero' && item.inRun === true && item.run === run;
  let start = itemAt(ctx.story, firstChar.address);
  const end = itemAt(ctx.story, lastChar.address);
  if (!start || !end) return [];
  if (either)
    for (let prev = tree.prev(start); prev && isMarker(prev, start.run); prev = tree.prev(prev))
      start = prev;
  const unitBeside = (
    item: OrderedItem,
    step: (x: OrderedItem) => OrderedItem | null
  ): OrderedItem | null => {
    let n = step(item);
    while (n && isMarker(n, item.run)) n = step(n);
    return n && n.run === item.run && isUnit(n) ? n : null;
  };
  const out: OrderedItem[] = [];
  for (const item of itemsFromTo(ctx.story, start, end)) {
    if (item.kind !== 'zero' || item.inRun !== true || reached(item)) {
      continue;
    }

    const before = unitBeside(item, (x) => tree.prev(x));
    const after = before ? null : unitBeside(item, (x) => tree.next(x));
    if ((before && reached(before)) || (either && after && reached(after))) out.push(item);
  }

  return out;
}

// ---- formats (rule 10) ----

/**
 * The properties a format writes on a run (rule 10): a right-to-left run takes the complex-script
 * partner (`w:bCs` with `w:b`).
 */
export function withComplexScript(
  face: OoxmlElement | null,
  properties: readonly OoxmlProperty[]
): readonly OoxmlProperty[] {
  return withComplexScriptCompanions(containerIsComplexScript(face ?? undefined), properties);
}

// ---- new marks (rule 11) ----

/** A mark's paragraph shell without its section break, which stays on the section's last mark. */
export function withoutSectionBreak(node: OoxmlNode): OoxmlNode {
  if (node.kind === 'textValue' || !isParagraphProperties(node)) return node;
  const children = node.children.filter((c) => !isWml(c, 'sectPr'));
  return children.length === node.children.length ? node : withChildren(node, children);
}

// ---- joins (rule 6) ----

/**
 * Two paragraphs that follow each other in one container, with nothing between them. Backspace
 * at the edge of a cell or a block control does nothing.
 */
export function siblings(ctx: Ctx, a: OrderedViewParagraph, b: OrderedViewParagraph): boolean {
  if (a.blocks.at(-1) !== b.blocks.at(-1)) return false;
  const end = b.items[0] ?? b.mark;
  for (
    let item = ctx.story.tree.next(a.mark);
    item && item !== end;
    item = ctx.story.tree.next(item)
  ) {
    if (item.kind === 'zero' && item.block) return false;
    if (item.kind === 'open' || item.kind === 'close') return false;
  }

  return true;
}

// ---- direct deletions (rule 12) ----

// A comment's markers are not among them: a comment's range stays at the deletion's place when
// its words are deleted directly (the engine removes the comment with them)
const BOOKMARK_MARKERS: ReadonlySet<string> = new Set(['bookmarkStart', 'bookmarkEnd']);

/** The items from `from` to `to`, both included, in order. */
function* itemsFromTo(story: OrderedStory, from: OrderedItem, to: OrderedItem) {
  for (
    let item: OrderedItem | null = from;
    item;
    item = item === to ? null : story.tree.next(item)
  ) {
    yield item;
  }
}

/** A unit or a paragraph mark: a deletion's covered markers lie between them. */
const isUnitOrMark = (item: OrderedItem): boolean => isUnit(item) || item.kind === 'mark';

/**
 * The markers a direct deletion of `[first, last]` covers whole (a bookmark), from the unit
 * before it to the unit after it, as the engine deletes them. A comment's range and a permission
 * range stay; a tracked deletion leaves them all.
 */
export function coveredMarkers(ctx: Ctx, first: OrderedItem, last: OrderedItem): OrderedItem[] {
  const tree = ctx.story.tree;
  let from: OrderedItem = first;
  for (let prev = tree.prev(from); prev && !isUnitOrMark(prev); prev = tree.prev(prev)) {
    from = prev;
  }

  let to: OrderedItem = last;
  for (let next = tree.next(to); next && !isUnitOrMark(next); next = tree.next(next)) {
    to = next;
  }

  // Each bookmark by its id: its markers here, and their names
  const found = new Map<string, { readonly items: OrderedItem[]; readonly names: Set<string> }>();
  for (const item of itemsFromTo(ctx.story, from, to)) {
    const node = item.node;
    if (item.kind !== 'zero' || item.block || !node || node.kind === 'textValue') {
      continue;
    }

    const id = attributeOf(node, 'id');
    if (!BOOKMARK_MARKERS.has(node.localName) || id === undefined) {
      continue;
    }

    const entry = found.get(id) ?? { items: [], names: new Set<string>() };
    entry.items.push(item);
    entry.names.add(node.localName);
    found.set(id, entry);
  }

  const out: OrderedItem[] = [];
  for (const { items, names } of found.values()) {
    if (names.has('bookmarkStart') && names.has('bookmarkEnd')) {
      for (const item of items) {
        out.push(item);
      }
    }
  }

  for (const item of coveredFields(ctx, from, to)) {
    out.push(item);
  }

  return out;
}

const FIELD_CHROME = new Set(['fldChar', 'instrText']);

/** The begin, instruction, separator and end markers of each field `[from, to]` holds whole. */
function coveredFields(ctx: Ctx, from: OrderedItem, to: OrderedItem): OrderedItem[] {
  const out: OrderedItem[] = [];
  const open: OrderedItem[][] = [];
  for (const item of itemsFromTo(ctx.story, from, to)) {
    const node = item.node;
    if (item.kind !== 'zero' || !node || node.kind === 'textValue') {
      continue;
    }

    if (!FIELD_CHROME.has(node.localName)) {
      continue;
    }

    const type = attributeOf(node, 'fldCharType');
    if (type === 'begin') {
      open.push([item]);
    } else if (type === 'end') {
      const field = open.pop();
      if (field) {
        for (const part of [...field, item]) {
          out.push(part);
        }
      }
    } else {
      open.at(-1)?.push(item);
    }
  }

  return out;
}

// ---- content the engine keeps from edits ----

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
 * A control's own lock (`w:lock`) as it shows now: the file's, or the organiser's later change
 * of it (`control-property`), a record on the open item that compose writes.
 */
function lockOf(ctx: Ctx, open: OrderedItem): ContentControlLockFlags {
  const node = open.node;
  if ((open.token !== 'sdt' && open.token !== 'sdtBlock') || !node || node.kind === 'textValue')
    return lockFlagsFromValue(undefined);
  if (open.props.length === 0) return declaredLockOf(node);
  const compose = new ComposeContext(ctx.doc, ctx.story, {});
  return declaredLockOf(withChildren(node, containerPropertiesAfter(compose, open, node.children)));
}

/** A control whose lock keeps its content from edits (`contentLocked`, `sdtContentLocked`). */
const contentLocked = (ctx: Ctx) => (open: OrderedItem) => lockOf(ctx, open).content;

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
  if (end > start)
    return p.chars.slice(start, end).some((char) => keptControls(char, isKept).length > 0);
  const before = keptControls(p.chars[start - 1], isKept);
  return before.length > 0 && keptControls(p.chars[start], isKept).some((c) => before.includes(c));
}

const LOCKED = 'the content of a content-locked control';

/**
 * Refuse a change inside block containers (a paragraph's, or the ones around a table) when one
 * of them is a control whose lock keeps its content (`locked`).
 */
export function refuseLockedBlocks(ctx: Ctx, blocks: readonly OrderedItem[]): void {
  if (blocks.some(contentLocked(ctx))) throw new OrderedRefusal('locked', LOCKED);
}

/**
 * Refuse a change of `items` (a row's, a table's, a paragraph's) when a control among them that
 * shows keeps its content, or, when the change removes them (`removes`), when a control among
 * them may not be removed (`sdtLocked`, `sdtContentLocked`).
 */
export function refuseLockedWithin(
  ctx: Ctx,
  items: readonly OrderedItem[],
  removes: boolean
): void {
  for (const item of items) {
    if (item.kind !== 'open' || isHidden(ctx.doc, item)) continue;
    const lock = lockOf(ctx, item);
    if (lock.content) throw new OrderedRefusal('locked', LOCKED);
    if (removes && lock.wrapper)
      throw new OrderedRefusal('locked', 'a control that may not be removed');
  }
}

/** The shown paragraph that holds an item (a mark, a unit, a container's edge), as drawn now. */
function paragraphOfItem(ctx: Ctx, item: OrderedItem): OrderedViewParagraph | null {
  const p = orderedViewParagraph(ctx.doc, ctx.story.name, ctx.story.tree.before(item).marks);
  return p && (p.mark === item || p.items.includes(item)) ? p : null;
}

/**
 * Refuse a change of one item when a control around it keeps its content: a paragraph's mark (a
 * section break), an object (a picture's size, an equation), a note reference (a conversion), a
 * link's open item (its target). For a container's open item only the controls around it count,
 * as the engine asks a node's enclosing controls.
 */
export function refuseLockedAt(ctx: Ctx, item: OrderedItem): void {
  const p = paragraphOfItem(ctx, item);
  if (!p) return;
  refuseLockedBlocks(ctx, p.blocks);
  if (item.kind === 'mark') return;
  const own = `${item.token === 'hyperlink' ? 'h' : 'c'}:${item.space}`;
  const char = p.chars.find((c) =>
    item.kind === 'open' ? segmentsOf(c.containers).includes(own) : c.item === item
  );
  if (!char) {
    return;
  }

  const segments = segmentsOf(char.containers);
  const around = item.kind === 'open' ? segments.slice(0, segments.indexOf(own)) : segments;
  const locked = contentLocked(ctx);
  for (const segment of around) {
    const open = isControlSegment(segment) ? openOfSegment(ctx.story, segment) : null;
    if (open && locked(open)) {
      throw new OrderedRefusal('locked', LOCKED);
    }
  }
}

/**
 * Refuse a change of the table a paragraph is in (a row, a column, a cell's properties) when a
 * control around the table keeps its content. The engine's own op on the table meets the
 * controls in it.
 */
export function refuseLockedTable(ctx: Ctx, p: OrderedViewParagraph): void {
  const table = p.blocks.map((block) => block.token).lastIndexOf('tbl');
  refuseLockedBlocks(ctx, table < 0 ? p.blocks : p.blocks.slice(0, table));
}

/**
 * Refuse an edit of view characters `[start, end)` (a caret when they are equal) whose content a
 * control's lock keeps (ECMA-376 Part 1 §17.5.2.23, the engine's `locked`): the paragraph is in a
 * `contentLocked` or `sdtContentLocked` block control, or the edit reaches inside such an inline
 * control. A lock holds for the controls inside the control too (`resolveContentControlLock`).
 * At an inline control's edge, typed text goes beside it (`promptSite` never enters it).
 */
export function refuseContentLock(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number
): void {
  refuseLockedBlocks(ctx, p.blocks);
  if (reachesInside(ctx, p, start, end, contentLocked(ctx)))
    throw new OrderedRefusal('locked', LOCKED);
}

/**
 * Whether a format of `properties` passes a content lock: a direct format that the room's policy
 * lets past it, as the engine lets an untracked restyle (`restylePassesContentLock`). A suggested
 * format stays refused, as the engine refuses a tracked restyle there.
 */
export function formatPassesContentLock(ctx: Ctx, properties: readonly OoxmlProperty[]): boolean {
  return ctx.direct && formattingPassesContentLock(properties, ctx.contentLockPolicy);
}

/** Whether typing never enters a control (the engine's `refusesTypedContent`): bound or locked. */
function refusesTyping(ctx: Ctx, open: OrderedItem): boolean {
  return isBound(open) || lockOf(ctx, open).content;
}

/**
 * Refuse an edit of view characters `[start, end)` (a caret when they are equal) that reaches
 * content the engine keeps from edits: a table of contents, a bound control's content, a locked
 * control's content. A format names its `properties`: a locked control's content may then take
 * it (`formatPassesContentLock`).
 */
export function refuseLocked(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number,
  properties?: readonly OoxmlProperty[]
): void {
  const toc = p.blocks.find(isTableOfContents);
  if (toc && reachesTocField(ctx, toc, p, start, end))
    throw new OrderedRefusal('not-covered', 'a table of contents result is generated');
  if (reachesInside(ctx, p, start, end, isBound))
    throw new OrderedRefusal('not-covered', 'a control bound to custom XML');
  if (properties === undefined || !formatPassesContentLock(ctx, properties))
    refuseContentLock(ctx, p, start, end);
}
