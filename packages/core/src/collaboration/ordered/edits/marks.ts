// Paragraph marks: Enter adds one, Backspace and Delete across paragraphs strike one, and a
// paragraph's own presses (alignment, indents, lists, tab stops) are records on its mark.

import type { OoxmlAttribute, OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { contentControlPropertiesOf } from '../../../store/package/content-control-nodes.ts';
import { propertiesHoldLineBreaks } from '../../../store/store/tree-op-inline-elements.ts';
import { enterDividesControl } from '../../../store/store/tree-op-split-controls.ts';
import {
  ComposeContext,
  type OrderedViewChar,
  type OrderedViewParagraph,
} from '../compose/context.ts';
import { paragraphOpKeeps, paragraphOpRefusal } from '../compose/paragraph-ops.ts';
import { ownItems, showsBlockPrompt } from '../compose/paragraph.ts';
import { innermostSegment, isControlSegment, openOfSegment } from '../containers.ts';
import { createOrderedItem, insertedPending, isSplitEdge } from '../model/document.ts';
import { nextSeq, relabel } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { orderedListOf } from '../model/resources.ts';
import { recordUndo } from '../model/undo-record.ts';
import {
  isUnit,
  type OrderedFormatRecord,
  type OrderedFormatWrite,
  type OrderedItem,
  type OrderedParagraphOp,
  type OrderedPropertyRecord,
  type OrderedSuggestion,
} from '../model/types.ts';
import { attributeOf, childNamed, isWml, withChildren, withId } from '../ooxml-nodes.ts';
import {
  breakSite,
  commonContainers,
  formatPassesContentLock,
  insertAtSite,
  itemsInside,
  refuseContentLock,
  refuseLocked,
  refuseLockedBlocks,
  siblings,
  withoutSectionBreak,
  type Site,
} from '../placement.ts';
import {
  currentFace,
  pendingOwn,
  shownParagraphProperties,
  strikeItems,
  suggestionFor,
  takeBack,
  withNewIds,
  type Ctx,
} from './context.ts';
import { newParagraphIdentity, withParagraphIdentity } from './para-ids.ts';
import {
  formatKeeps,
  noteOwnAddition,
  ownFormat,
  struckAlready,
  unitsOf,
  withPressFormat,
  withPropertyRecord,
} from './text.ts';

const MULTI_LINE = new Set(['1', 'true', 'on']);

/** Whether a control's open item allows one line only: a plain text control, not multi-line. */
function singleLine(open: OrderedItem): boolean {
  const shell = open.node;
  if (!shell || shell.kind === 'textValue') return false;
  const sdtPr = childNamed(shell, 'sdtPr');
  const text = sdtPr ? childNamed(sdtPr, 'text') : undefined;
  if (!text) return false;
  const multi = attributeOf(text, 'multiLine');
  return multi === undefined || !MULTI_LINE.has(multi);
}

/**
 * Policy 4: Enter inside a control's prompt removes the prompt, then breaks. A block control's
 * prompt goes for good, its other paragraphs too, with no revision and no glossary block, as
 * Word's Return replaces the selected prompt (Word 16.114, `~$p4-enter-tracked.docx`); once the
 * control is empty it shows a prompt as any control does. An inline control cannot hold a
 * break, and a control that allows one line refuses it by name. `containers` are those Enter
 * falls in (`promptAround`).
 */
export function promptBreakSite(ctx: Ctx, p: OrderedViewParagraph, containers: string): Site {
  const segment = innermostSegment(containers);
  if (isControlSegment(segment)) {
    const open = openOfSegment(ctx.story, segment);
    if (open && singleLine(open))
      throw new OrderedRefusal('single-line-control', 'Enter in a single-line control');
    throw new OrderedRefusal('break-in-container', "Enter in an inline control's prompt");
  }

  const control = [...p.blocks]
    .reverse()
    .find((block) => block.token === 'sdtBlock' && block.prompt);

  // The block control's own prompt: a control inside it keeps its own (`ownItems`)
  const own = control
    ? new Set([...ownItems(itemsInside(ctx.story, control), control)].filter((item) => item.prompt))
    : null;
  const last = [...p.items]
    .reverse()
    .find((item) => item.kind !== 'mark' && (own ? own.has(item) : item.prompt === true));
  if (!last) throw new OrderedRefusal('break-in-container', "Enter in a control's prompt");

  // The whole prompt goes, as Word's Return over the selected prompt: its units, and the marks
  // of its other paragraphs (the caret's paragraph keeps its mark, as the engine keeps it)
  for (const item of own ?? [])
    if (isUnit(item) || (item.kind === 'mark' && item !== p.mark))
      relabel(ctx.doc, ctx.story, item, { gone: true });
  return {
    after: { item: last, offset: last.offset + Math.max(0, last.length - 1) },
    before: null,
    neighbour: undefined,
  };
}

/** A mark's paragraph shell for a new mark: new ids, its own paragraph ids, no section break. */
export function newMarkShell(
  ctx: Ctx,
  shell: OoxmlElement,
  paragraph: { readonly space: string; readonly attributes: readonly OoxmlAttribute[] }
): OoxmlElement {
  const children = shell.children.map((child): OoxmlNode => {
    const kept = withoutSectionBreak(child);
    return kept.kind === 'textValue' ? kept : withNewIds(ctx, kept);
  });
  return withId(
    withParagraphIdentity(withChildren(shell, children), paragraph.attributes),
    paragraph.space
  );
}

/**
 * The mark's records as the new mark takes them: element records lose their section break, and
 * a section's own records stay with the mark that ends it (rule 11). Every node is a copy with
 * the edit's ids: the paragraph it splits keeps its own nodes, and one part must not hold a node
 * id twice.
 */
export function copiedProps(
  ctx: Ctx,
  props: readonly OrderedPropertyRecord[]
): OrderedPropertyRecord[] {
  return props
    .filter((record) => record.target !== 'sectPr')
    .map((record) => {
      if (record.kind !== 'element') return record;
      const kept = record.element ? withoutSectionBreak(record.element) : null;
      const element = kept && kept.kind !== 'textValue' ? withNewIds(ctx, kept) : record.element;
      return {
        ...record,
        element,
        ...(record.fileRecord ? { fileRecord: withNewIds(ctx, record.fileRecord) } : {}),
      };
    });
}

/** The mark's format records with their faces copied under the edit's ids. */
export function copiedFmts(ctx: Ctx, fmts: readonly OrderedFormatRecord[]): OrderedFormatRecord[] {
  return fmts.map((record) => {
    if (record.kind !== 'face') return record;
    return {
      ...record,
      face: record.face ? withNewIds(ctx, record.face) : null,
      ...(record.fileRecord ? { fileRecord: withNewIds(ctx, record.fileRecord) } : {}),
    };
  });
}

/**
 * Enter at a view offset: a new mark that ends the first part. It copies the mark it splits as
 * it is now, with the pending records on it so far, but not a struck mark's deletion (a
 * proposed join stays on the original mark) and not a section break (rule 11). Its items in the
 * author's own earlier suggestion (a line of a paste) are noted for its undo.
 */
export function splitParagraph(ctx: Ctx, p: OrderedViewParagraph, offset: number): string {
  refuseLocked(ctx, p, offset, offset);
  const prev = offset > 0 ? p.chars[offset - 1] : undefined;
  const next = p.chars[offset];
  const prompt = promptAround(prev, next);
  const { site, split } =
    prompt !== null
      ? { site: promptBreakSite(ctx, p, prompt), split: [] }
      : breakSite(ctx, p, offset);
  for (const open of split)
    if (open.token === 'sdt' && !enterDivides(open))
      throw new OrderedRefusal('single-line-control', 'Enter in a control that holds one line');
  const shell = p.mark.node;
  if (!shell || shell.kind === 'textValue')
    throw new OrderedRefusal('malformed-story', 'a paragraph mark without its paragraph');
  const suggestion = suggestionFor(ctx);
  const paragraph = newParagraphIdentity(ctx);
  const mark = createOrderedItem({
    space: paragraph.space,
    kind: 'mark',
    node: newMarkShell(ctx, shell, paragraph),
    face: p.mark.face ? withNewIds(ctx, p.mark.face) : null,
    fmts: copiedFmts(ctx, p.mark.fmts),
    props: copiedProps(ctx, p.mark.props),
    by: suggestion.id,
    within: p.mark.within,
    ...(p.mark.home === undefined ? {} : { home: p.mark.home }),
  });
  const { closes, opens } = splitContainers(ctx, split, suggestion.id, p.mark);
  const items = [...closes, mark, ...opens];
  insertAtSite(ctx, site, items);
  noteOwnAddition(ctx, suggestion.id, unitsOf(ctx.story, items));
  return suggestion.id;
}

/**
 * The containers of the prompt Enter falls in, or null. Inside a prompt, prompt units lie on both
 * sides, and Enter falls in the containers they share: at the edge of a control inside the
 * prompt, beside it. At a block prompt's edge (the paragraph's start or end) Enter is in the
 * prompt too: Word selects a prompt whole, so Return replaces it wherever the caret was (policy
 * 4). At an inline prompt's edge Enter goes beside the control.
 */
function promptAround(
  prev: OrderedViewChar | undefined,
  next: OrderedViewChar | undefined
): string | null {
  if (prev?.item.prompt && next?.item.prompt) {
    return commonContainers(prev, next);
  }

  // At the paragraph's start or end: the one side there is
  let edge: OrderedViewChar | undefined;
  if (!prev) {
    edge = next;
  } else if (!next) {
    edge = prev;
  }

  if (!edge?.item.prompt) return null;
  const inline = isControlSegment(innermostSegment(edge.containers));
  return inline ? null : edge.containers;
}

/**
 * Enter inside links and controls splits them, as the engine does (task 3.5): before the new
 * mark, a close item for each (innermost first) ends the first part; after it, an open item
 * for each (outermost first), a copy of the container's element under new ids, starts the
 * rest, which the container's own close item still ends. They carry the Enter's suggestion: a
 * decision keeps the two containers (as the engine and a reopened file do), and only taking
 * the Enter back leaves one container again (`isSplitEdge`).
 */
export function splitContainers(
  ctx: Ctx,
  split: readonly OrderedItem[],
  suggestion: string,
  mark: OrderedItem
): { closes: OrderedItem[]; opens: OrderedItem[] } {
  const closes = [...split].reverse().map((open) =>
    createOrderedItem({
      space: ctx.ids.container(),
      offset: 1,
      kind: 'close',
      ...(open.token === undefined ? {} : { token: open.token }),
      ...(open.pair === undefined ? {} : { pair: open.pair }),
      by: suggestion,
      within: mark.within,
    })
  );
  const opens = split.map((open) => {
    const space = ctx.ids.container();
    const shell = open.node;

    // The copy's node id is the open item's space, as a new control's is (`controlEdges`)
    const copy =
      shell && shell.kind !== 'textValue'
        ? withoutPromptFlag(withId(withNewIds(ctx, shell), space))
        : null;

    // The copy reads as the container it came from (a control's content kind, its prompt)
    const original = shell ? ctx.doc.originals.get(shell.id) : undefined;
    if (copy && original) {
      ctx.doc.originals.set(copy.id, original);
      recordUndo(ctx.doc, () => ctx.doc.originals.delete(copy.id));
    }

    return createOrderedItem({
      space,
      offset: 0,
      kind: 'open',
      ...(open.token === undefined ? {} : { token: open.token }),
      ...(open.pair === undefined ? {} : { pair: open.pair }),
      node: copy,
      // Its own copy of the records: one part must not hold a node id twice
      props: copiedProps(ctx, open.props),
      // A control that showed its prompt: the piece's prompt hides as the control's does
      ...(open.prompt ? { prompt: true } : {}),
      by: suggestion,
      within: mark.within,
    });
  });
  return { closes, opens };
}

/**
 * A split piece's shell without `w:showingPlcHdr`: the piece holds no prompt of its own, so
 * its control's prompt rule (from the original, `ctx.doc.originals`) says when it shows one.
 */
function withoutPromptFlag(shell: OoxmlElement): OoxmlElement {
  return withChildren(
    shell,
    shell.children.map((child) =>
      child.kind !== 'textValue' && isWml(child, 'sdtPr')
        ? withChildren(
            child,
            child.children.filter((c) => !isWml(c, 'showingPlcHdr'))
          )
        : child
    )
  );
}

/** Whether a control holds more than one line: the engine's rule for line breaks in it. */
export function holdsLines(open: OrderedItem): boolean {
  const shell = open.node;
  if (!shell || shell.kind === 'textValue') return true;
  return propertiesHoldLineBreaks(contentControlPropertiesOf(shell));
}

/**
 * Whether Enter may divide an inline control: the engine's rule (`enterDividesControl`). Word
 * leaves a rich text control whole, as a control that holds one line.
 */
function enterDivides(open: OrderedItem): boolean {
  const shell = open.node;
  if (!shell || shell.kind === 'textValue') return true;
  return enterDividesControl(contentControlPropertiesOf(shell));
}

/**
 * The close items before a mark and the open items after it, in its two paragraphs, that the
 * same Enter put there (a split link or control): they go with the mark when its author takes it
 * back. Words typed at the first part's end go after its close items.
 */
function splitAround(ctx: Ctx, mark: OrderedItem): OrderedItem[] {
  const out: OrderedItem[] = [];
  const ofEnter = (item: OrderedItem, kind: 'open' | 'close'): boolean =>
    item.kind === kind && isSplitEdge(item) && mark.by !== null && item.by === mark.by;
  for (let at = ctx.story.tree.prev(mark); at && at.kind !== 'mark'; at = ctx.story.tree.prev(at)) {
    if (ofEnter(at, 'close')) {
      out.push(at);
    }
  }

  for (let at = ctx.story.tree.next(mark); at && at.kind !== 'mark'; at = ctx.story.tree.next(at)) {
    if (ofEnter(at, 'open')) {
      out.push(at);
    }
  }

  return out;
}

/** Strike a mark. Your own pending break is taken back; a struck mark keeps its first deletion. */
export function strikeMark(
  ctx: Ctx,
  mark: OrderedItem,
  into: OrderedSuggestion | null,
  options: { readonly evenStruck?: boolean } = {}
): string | null {
  const own = pendingOwn(ctx, mark.by);
  if (own && insertedPending(ctx.doc, mark)) {
    // Taking back one's own Enter leaves a link or control it split as one again
    for (const item of [mark, ...splitAround(ctx, mark)]) takeBack(ctx, ctx.story, item);
    return into?.id ?? own.id;
  }

  if (!options.evenStruck && struckAlready(ctx, mark.dels)) return into?.id ?? null;
  const suggestion = into ?? suggestionFor(ctx);
  strikeItems(ctx, ctx.story, [mark], suggestion.id);
  return suggestion.id;
}

/**
 * Whether the mark between two paragraphs is part of a block control's prompt: both are in the
 * control, and it shows its prompt. A prompt is no content (rule 7), and Word changes one only
 * in Design Mode, so a deletion leaves the mark as it leaves the prompt's text.
 */
export function markInPrompt(
  ctx: Ctx,
  previous: OrderedViewParagraph,
  p: OrderedViewParagraph
): boolean {
  const control = previous.blocks.at(-1);
  if (!control || control !== p.blocks.at(-1)) return false;
  return showsBlockPrompt(new ComposeContext(ctx.doc, ctx.story, {}), control);
}

/**
 * Backspace at the start of `p` (or Delete at the end of `previous`): strike the mark between
 * them. It does nothing at a cell's or a block control's edge (rule 6), or inside a block
 * control's prompt (rule 7). Inline prompts at the two paragraphs' edges are outside the mark,
 * so those paragraphs join as other text does.
 */
export function joinParagraphs(
  ctx: Ctx,
  previous: OrderedViewParagraph,
  p: OrderedViewParagraph,
  options: { readonly evenStruck?: boolean } = {}
): string | null {
  if (!siblings(ctx, previous, p) || markInPrompt(ctx, previous, p)) return null;
  refuseLocked(ctx, previous, 0, 0);
  refuseLocked(ctx, p, 0, 0);

  // A direct join rearranges both paragraphs, as the engine's `joinParagraphs` does: a locked
  // control in either refuses it. A suggested join marks only the mark between them
  if (ctx.direct) for (const q of [previous, p]) refuseContentLock(ctx, q, 0, q.chars.length);
  return strikeMark(ctx, previous.mark, null, options);
}

/** A format on a paragraph's mark (a format over several paragraphs reaches the marks between). */
export function formatMark(
  ctx: Ctx,
  p: OrderedViewParagraph,
  write: OrderedFormatWrite,
  into: OrderedSuggestion | null
): string | null {
  if (!formatPassesContentLock(ctx, write.properties)) refuseLockedBlocks(ctx, p.blocks);
  const mark = p.mark;
  if (formatKeeps(ctx, mark, currentFace(ctx, mark, false), write)) return into?.id ?? null;
  const own = pendingOwn(ctx, mark.by);
  const seq = nextSeq(ctx.doc);
  if (own && insertedPending(ctx.doc, mark)) {
    relabel(ctx.doc, ctx.story, mark, {
      fmts: [...mark.fmts, ownFormat(ctx, mark, own, seq, write)],
    });
    return into?.id ?? own.id;
  }

  const suggestion = into ?? suggestionFor(ctx);
  relabel(
    ctx.doc,
    ctx.story,
    mark,
    withPressFormat(ctx, mark, {
      kind: 'properties',
      id: `f:${suggestion.id}#${seq}`,
      suggestion: suggestion.id,
      seq,
      ...write,
    })
  );
  return suggestion.id;
}

/** The engine refuses the press on the paragraph as it shows now (`shown`). */
function checkParagraphOp(
  ctx: Ctx,
  shown: OoxmlElement | null,
  op: OrderedParagraphOp,
  mint: () => string
): void {
  const known = ctx.doc.numbering;

  // A list number the file defines, or one a change added that this author may use (D14)
  if (
    op.op === 'setListNumbering' &&
    op.numId !== null &&
    known !== null &&
    !known.has(op.numId) &&
    orderedListOf(ctx.doc, op.numId, ctx.attribution) === null
  )
    throw new OrderedRefusal('not-covered', `a new list definition (${op.numId.slice(0, 32)})`);
  const current = shown
    ? withChildren(
        shown,
        shown.children.filter((c) => !isWml(c, 'pPrChange'))
      )
    : null;
  const refusal = paragraphOpRefusal(ctx.story.part, current, op, mint);
  if (refusal !== null) throw new OrderedRefusal('engine-refused', `${op.op}: ${refusal}`);
}

/**
 * A paragraph property press, as a record on the paragraph's mark. On a paragraph its author
 * proposed, the press is part of that proposal and draws with no record. A press that leaves the
 * paragraph as it is (`paragraphOpKeeps`) makes no record. Returns the suggestion that holds the
 * press, if any.
 */
export function paragraphOp(
  ctx: Ctx,
  p: OrderedViewParagraph,
  op: OrderedParagraphOp,
  into: OrderedSuggestion | null
): string | null {
  refuseLockedBlocks(ctx, p.blocks);
  const mark = p.mark;
  let scratch = 0;
  const mint = () => `check#${(scratch += 1)}`;
  const shown = shownParagraphProperties(ctx, mark);
  checkParagraphOp(ctx, shown, op, mint);
  if (paragraphOpKeeps(ctx.story.part, shown, op, mint, ctx.attribution)) {
    return into?.id ?? null;
  }

  const own = pendingOwn(ctx, mark.by);
  if (own && insertedPending(ctx.doc, mark)) {
    relabel(ctx.doc, ctx.story, mark, {
      props: [
        ...mark.props,
        {
          kind: 'paragraphOp',
          suggestion: own.id,
          seq: nextSeq(ctx.doc),
          target: 'pPr',
          op,
          untracked: true,
        },
      ],
    });
    return own.id;
  }

  const suggestion = into ?? suggestionFor(ctx);
  relabel(ctx.doc, ctx.story, mark, {
    props: withPropertyRecord(ctx, mark, {
      kind: 'paragraphOp',
      suggestion: suggestion.id,
      seq: nextSeq(ctx.doc),
      target: 'pPr',
      op,
    }),
  });
  return suggestion.id;
}
