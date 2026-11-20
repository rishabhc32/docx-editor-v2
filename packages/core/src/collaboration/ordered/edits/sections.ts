// Section edits (task 9.8, `model/sections.ts`): page setup, furniture options and section
// breaks, as records on the items that hold the sections.

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { defaultSectionChildren } from '../../../store/store/tree-op-section.ts';
import {
  metricsOfSection,
  sectionBreakTypeOf,
} from '../../../store/store/tree-op-section-address.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import { applyTreeOp } from '../../../store/store/tree-op-apply.ts';
import { withStoryBlocks } from '../compose/paragraph-ops.ts';
import {
  bodySectionNode,
  isSectionItem,
  markSection,
  orderedSectionsOf,
  sectionOpKeeps,
  sectionRecordsOf,
  sectionWithoutPendingWrites,
  type OrderedSectionView,
} from '../compose/sections.ts';
import { insertedPending, itemAt, pendingByOther, statusOf } from '../model/document.ts';
import { nextSeq, relabel } from '../model/mutations.ts';
import type { OrderedMaker } from '../model/person.ts';
import { addressLabel, OrderedRefusal } from '../model/refusal.ts';
import type { OrderedSectionOp, OrderedSectionRecord } from '../model/sections.ts';
import type {
  OrderedDocument,
  OrderedStory,
  OrderedItem,
  OrderedItemAddress,
  OrderedPropertyRecord,
  OrderedSuggestion,
} from '../model/types.ts';
import { refuseLockedAt } from '../placement.ts';
import { pendingOwn, scratchCompose, suggestionFor, withNewIds, type Ctx } from './context.ts';
import { withTakeOver } from './text.ts';

/**
 * A story's sections as they show now, counted as the engine counts them, so a change that
 * names a section by its holder writes where the editor wrote.
 */
export function sectionsOfStory(doc: OrderedDocument, story: OrderedStory): OrderedSectionView[] {
  return orderedSectionsOf(scratchCompose(doc, story, 'section'));
}

/** The main story's sections as they show now, in the engine's order. */
export function sectionsNow(ctx: Ctx): OrderedSectionView[] {
  return sectionsOfStory(ctx.doc, ctx.story);
}

/** The section whose holder is at `address`, with its index; refused when none ends there. */
export function sectionAt(
  ctx: Ctx,
  address: OrderedItemAddress
): { readonly index: number; readonly section: OrderedSectionView } {
  const holder = itemAt(ctx.story, address);
  const sections = sectionsNow(ctx);
  const index = holder === null ? -1 : sections.findIndex((s) => s.holder === holder);
  const section = sections[index];
  if (section === undefined) throw new OrderedRefusal('not-a-section', addressLabel(address));
  return { index, section };
}

/**
 * The record list with `record` added. A pending record of another author on the same section
 * is taken over (policy 2: Word keeps one `w:sectPrChange` for each section); a direct edit
 * takes over none, and a section break's own record stays with its mark.
 */
function withSectionRecord(
  ctx: Ctx,
  props: readonly OrderedPropertyRecord[],
  record: OrderedSectionRecord
): OrderedPropertyRecord[] {
  if (ctx.direct || record.untracked === true) return [...props, record];
  return withTakeOver(ctx, props, record, (kept) => kept.target !== 'sectPr' || isOwnBreak(kept));
}

/** A section break's own record on its author's new mark: untracked, it is no change record. */
const isOwnBreak = (record: OrderedPropertyRecord): boolean =>
  record.kind === 'sectionOp' && record.untracked === true;

/** The type of the break that ends a section; a section with no properties starts a page. */
const breakTypeOf = (section: OrderedSectionView): string =>
  section.sectPr ? sectionBreakTypeOf(section.sectPr) : 'nextPage';

/** Add one write to a section's holder, for `suggestion`. */
export function addSectionWrite(
  ctx: Ctx,
  holder: OrderedItem,
  op: OrderedSectionOp,
  suggestion: OrderedSuggestion,
  flags: Pick<OrderedSectionRecord, 'untracked' | 'keptOnReject'> = {}
): void {
  const record: OrderedSectionRecord = {
    kind: 'sectionOp',
    suggestion: suggestion.id,
    seq: nextSeq(ctx.doc),
    target: 'sectPr',
    op,
    ...(flags.untracked === true ? { untracked: true } : {}),
    ...(flags.keptOnReject === true ? { keptOnReject: true } : {}),
  };
  relabel(ctx.doc, ctx.story, holder, { props: withSectionRecord(ctx, holder.props, record) });
}

/** The engine's own check of a page setup write on one section as it shows now; refused by name. */
function checkPage(
  ctx: Ctx,
  section: OrderedSectionView,
  op: Extract<OrderedSectionOp, { op: 'page' }>
) {
  const { op: _page, ...fields } = op;
  const scratch = withStoryBlocks(ctx.story.part, section.sectPr ? [section.sectPr] : []);
  const result = applyTreeOp(scratch, { op: 'setSectionProperties', ...fields });
  if (!result.ok)
    throw new OrderedRefusal('engine-refused', `setSectionProperties: ${result.reason}`);
}

/**
 * One write to each of `sections`, for one suggestion: `into` when given, else the edit's own. A
 * suggested write that leaves a section as it shows (`sectionOpKeeps`) makes no record, so a
 * press that changes nothing makes no suggestion. Returns the suggestion that holds the writes,
 * if any.
 */
function addSectionWrites(
  ctx: Ctx,
  sections: readonly OrderedSectionView[],
  op: OrderedSectionOp,
  into: OrderedSuggestion | null
): string | null {
  const shown = scratchCompose(ctx.doc, ctx.story, 'section');
  const written = ctx.direct
    ? sections
    : sections.filter((section) => !sectionOpKeeps(shown, section.sectPr, op, ctx.attribution));
  if (written.length === 0) {
    return into?.id ?? null;
  }

  const suggestion = into ?? suggestionFor(ctx);
  for (const section of written) {
    addSectionWrite(ctx, section.holder, op, suggestion);
  }

  return suggestion.id;
}

/**
 * Page size, orientation and margins (`setSectionProperties`) on the sections the document has
 * now: every one, the body's own too when the file wrote none (the engine mints it), or with
 * `pageAt` the one that governs that paragraph (a paragraph a join took is in the section of the
 * paragraph it joined). Each is checked as the engine checks it. Suggested, it records the
 * section before it (`w:sectPrChange`). Returns the suggestion that holds the writes, if any.
 */
export function setPageSetup(
  ctx: Ctx,
  pageAt: OrderedItemAddress | null,
  op: Extract<OrderedSectionOp, { op: 'page' }>
): string | null {
  const named = pageSetupSections(ctx, pageAt);
  for (const section of named) checkPage(ctx, section, op);
  return addSectionWrites(ctx, named, op, null);
}

/** A story and its document, as a page setup reads them. */
interface StoryOf {
  readonly doc: OrderedDocument;
  readonly story: OrderedStory;
}

/** The sections a page setup goes to now: every one, or with `pageAt` the one that governs it. */
function pageSetupSections(
  where: StoryOf,
  pageAt: OrderedItemAddress | null
): readonly OrderedSectionView[] {
  const sections = sectionsOfStory(where.doc, where.story);
  return pageAt === null ? withBodySection(where, sections) : [governing(where, sections, pageAt)];
}

/**
 * Another person's pending record on the section `holder` holds, or null. A section holds one
 * `w:sectPrChange`, so a suggested write there would take that record over and the change would
 * lose its author. A break on its author's own new mark has no record of its own and is not one.
 */
function othersRecordOn(
  doc: OrderedDocument,
  holder: OrderedItem,
  maker: OrderedMaker
): string | null {
  for (const record of sectionRecordsOf(holder)) {
    if (isOwnBreak(record)) {
      continue;
    }

    if (pendingByOther(doc, record.suggestion, maker)) {
      return record.suggestion;
    }
  }

  return null;
}

/** Another person's pending record on a section a page setup goes to, or null. */
export function othersSectionRecord(
  doc: OrderedDocument,
  story: OrderedStory,
  pageAt: OrderedItemAddress | null,
  maker: OrderedMaker
): string | null {
  for (const section of pageSetupSections({ doc, story }, pageAt)) {
    const other = othersRecordOn(doc, section.holder, maker);
    if (other !== null) return other;
  }

  return null;
}

/** Another person's pending record on a section the furniture options name, or null. */
export function othersFurnitureRecord(
  doc: OrderedDocument,
  story: OrderedStory,
  sections: readonly OrderedItemAddress[],
  maker: OrderedMaker
): string | null {
  for (const address of sections) {
    const holder = itemAt(story, address);
    const other = holder === null ? null : othersRecordOn(doc, holder, maker);
    if (other !== null) return other;
  }

  return null;
}

/** Every section, and the body's own when the engine does not count it (it ends no page). */
function withBodySection(
  where: StoryOf,
  sections: readonly OrderedSectionView[]
): readonly OrderedSectionView[] {
  if (sections.some((section) => section.body)) return sections;
  let holder: OrderedItem | null = null;
  for (const item of where.story.tree.items()) if (isSectionItem(item)) holder = item;
  if (holder === null) throw new OrderedRefusal('not-a-section', 'the document has no body');
  const sectPr = bodySectionNode(scratchCompose(where.doc, where.story, 'section'), holder);
  return [...sections, { holder, sectPr, body: true }];
}

/** The section that governs the paragraph whose mark is at `address`. */
function governing(
  where: StoryOf,
  sections: readonly OrderedSectionView[],
  address: OrderedItemAddress
): OrderedSectionView {
  const mark = itemAt(where.story, address);
  if (mark === null || mark.kind !== 'mark' || mark.offset !== address.offset)
    throw new OrderedRefusal('unknown-address', address.space.slice(0, 160));
  return governingSection(where.story, withBodySection(where, sections), mark);
}

/**
 * A different first page, and the header and footer distances, on the sections whose holders are
 * at `addresses`, for `into` when given. Suggested, it records the section before it
 * (`w:sectPrChange`), as page setup does. Returns the suggestion that holds the writes, if any.
 */
export function setFurnitureOptions(
  ctx: Ctx,
  addresses: readonly OrderedItemAddress[],
  op: Extract<OrderedSectionOp, { op: 'furniture' }>,
  into: OrderedSuggestion | null
): string | null {
  const sections = addresses.map((address) => sectionAt(ctx, address).section);
  return addSectionWrites(ctx, sections, op, into);
}

/**
 * A section break at a paragraph mark (`setSectionMark`): the mark takes a copy of the section
 * it is in, as the engine clones the governing section, and the section that starts after it
 * takes the break's type when it starts another way. The copy is the section as it shows
 * without its pending tracked writes, as the copy's own: another person's pending break there
 * gives the section it ends, and a decision on that break leaves the copy as it is, as in the
 * editor alone. The pending tracked writes go to the new section too, under their own
 * suggestions, as the engine's clone keeps their change record.
 *
 * Suggested, the mark must be its author's own new mark (Enter in the same change): the break
 * rides on it, and leaves with it. A mark that is not one has no tracked form. The type goes on
 * untracked, as Word writes it beside the proposed mark (no `w:sectPrChange`): a reject leaves
 * it, as Word's does, and a take-back removes it, as the editor's undo does.
 */
export function insertSectionBreak(
  ctx: Ctx,
  markAddress: OrderedItemAddress,
  type: 'nextPage' | 'continuous' | undefined
): string {
  const mark = itemAt(ctx.story, markAddress);
  if (mark === null || mark.kind !== 'mark' || mark.offset !== markAddress.offset)
    throw new OrderedRefusal('unknown-address', markAddress.space.slice(0, 160));
  refuseLockedAt(ctx, mark);
  const sections = sectionsNow(ctx);
  if (sections.some((s) => s.holder === mark))
    throw new OrderedRefusal('engine-refused', 'setSectionMark: the paragraph ends a section');
  const own = pendingOwn(ctx, mark.by);
  const onOwnMark = own !== null && insertedPending(ctx.doc, mark);
  if (!ctx.direct && !onOwnMark)
    throw new OrderedRefusal('invalid-edit', 'a section break rides on its own new paragraph mark');
  const governing = governingSection(ctx.story, sections, mark);
  const suggestion = onOwnMark && own ? own : suggestionFor(ctx);
  const shown = copiedSection(ctx, governing);
  const copy = shown ? withNewIds(ctx, shown) : null;
  addSectionWrite(ctx, mark, { op: 'mark', sectPr: copy ?? defaultSection(ctx) }, suggestion, {
    untracked: onOwnMark,
  });

  // The pending tracked writes on the section go with its copy, under their own suggestions
  // (an untracked one is in the copy already)
  for (const record of sectionRecordsOf(governing.holder)) {
    if (statusOf(ctx.doc, record.suggestion) === 'pending' && !isOwnBreak(record)) {
      relabel(ctx.doc, ctx.story, mark, {
        props: [...mark.props, { ...record, seq: nextSeq(ctx.doc) }],
      });
    }
  }

  if (type !== undefined && breakTypeOf(governing) !== type)
    addSectionWrite(
      ctx,
      governing.holder,
      { op: 'type', type },
      suggestion,
      onOwnMark ? { untracked: true, keptOnReject: true } : {}
    );
  return suggestion.id;
}

/**
 * Another person's pending record that a section break at the mark would copy onto its new
 * section, or null. The break copies the section it cuts with the pending writes on it, as the
 * engine's clone keeps the section's change record, so the copy is inside that person's
 * suggestion: rejecting the break would take their record back with it (design D9, Enter).
 */
export function othersCopiedRecord(
  doc: OrderedDocument,
  story: OrderedStory,
  markAddress: OrderedItemAddress,
  maker: OrderedMaker
): string | null {
  const mark = itemAt(story, markAddress);
  if (mark === null || mark.kind !== 'mark') return null;
  return othersRecordOn(
    doc,
    governingSection(story, sectionsOfStory(doc, story), mark).holder,
    maker
  );
}

/**
 * Whether a section break at the mark gives the section after it another type, which Word
 * writes untracked: a reject of the break leaves it. False when the mark does not resolve (the
 * change's targets refuse that by name).
 */
export function retypesNextSection(
  doc: OrderedDocument,
  story: OrderedStory,
  markAddress: OrderedItemAddress,
  type: 'nextPage' | 'continuous' | undefined
): boolean {
  if (type === undefined) return false;
  const mark = itemAt(story, markAddress);
  if (mark === null || mark.kind !== 'mark') return false;
  const governing = governingSection(story, sectionsOfStory(doc, story), mark);
  return breakTypeOf(governing) !== type;
}

// ---- helpers ----

/**
 * The section a paragraph mark is in: the first section that ends at or after it, outside a
 * table (`targetSectionNodes`), else the body's.
 */
export function governingSection(
  story: OrderedStory,
  sections: readonly OrderedSectionView[],
  mark: OrderedItem
): OrderedSectionView {
  const order = new Map<OrderedItem, number>();
  let k = 0;
  for (const item of story.tree.items()) order.set(item, (k += 1));
  const at = order.get(mark) ?? 0;
  const found = sections.find((s) => !s.body && (order.get(s.holder) ?? 0) >= at);
  const body = sections.find((s) => s.body);
  const section = found ?? body;
  if (!section) throw new OrderedRefusal('not-a-section', 'the document has no section');
  return section;
}

/** A section as it shows without its pending tracked writes: what a copy of it takes as its own. */
function copiedSection(ctx: Ctx, section: OrderedSectionView): OoxmlElement | null {
  const holder = section.holder;
  const base = holder.kind === 'mark' ? markSection(holder) : holder.node;
  return sectionWithoutPendingWrites(
    scratchCompose(ctx.doc, ctx.story, 'section'),
    holder,
    base && base.kind !== 'textValue' ? base : null
  );
}

/** A section a document that wrote none takes: the defaults the engine writes out. */
function defaultSection(ctx: Ctx): OoxmlElement {
  const mint = () => ctx.ids.node();
  return build(
    mint(),
    'generic',
    'sectPr',
    [],
    defaultSectionChildren(metricsOfSection(null), mint)
  );
}
