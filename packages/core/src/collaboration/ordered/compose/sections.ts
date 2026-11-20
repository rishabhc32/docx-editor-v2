// A section's properties under their records (task 9.8, `model/sections.ts`), each write the
// engine's own, run on this one section.

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { sectionFurnitureChildren } from '../../../store/package/hf-lifecycle.ts';
import { sectionChildrenWithNoteProperties } from '../../../store/package/note-lifecycle-props.ts';
import { cloneWithNewIds } from '../../../store/store/tree-op-nodes.ts';
import {
  sameSectionProperties,
  sectionRecordChildren,
  withSectionChangeRecord,
} from '../../../store/store/tree-op-section-record.ts';
import {
  applySetSectionProperties,
  defaultSectionChildren,
  withSectionBreakType,
} from '../../../store/store/tree-op-section.ts';
import { bodySectionOf, metricsOfSection } from '../../../store/store/tree-op-section-address.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import type { RevisionAttributionInput } from '../../../store/store/tree-op-types.ts';
import { isHidden, statusOf } from '../model/document.ts';
import { REVISION_ROLE } from '../model/revision-ids.ts';
import type { OrderedSectionOp, OrderedSectionRecord } from '../model/sections.ts';
import type {
  OrderedDocument,
  OrderedElementRecord,
  OrderedItem,
  OrderedStory,
} from '../model/types.ts';
import {
  asWml,
  attributeOf,
  childWhere,
  isParagraphProperties,
  isWml,
  relationshipIdAttribute,
  withChildren,
  wmlAttribute,
} from '../ooxml-nodes.ts';
import type { ComposeContext, DrawnTakeover } from './context.ts';
import { withStoryBlocks } from './paragraph-ops.ts';
import { acceptedAndPending, aroundPending, fileRecordAlone, noteTakeover } from './properties.ts';

type SectionRecord = OrderedElementRecord | OrderedSectionRecord;

const isReference = (node: OoxmlNode): boolean =>
  isWml(node, 'headerReference') || isWml(node, 'footerReference');

/** A section's records: its file's own change record and the writes changes made. */
export function sectionRecordsOf(item: OrderedItem): SectionRecord[] {
  return item.props.filter((record): record is SectionRecord => record.target === 'sectPr');
}

const newSection = (mint: () => string): OoxmlElement => build(mint(), 'generic', 'sectPr', [], []);

/** The references, then the properties a record gives (a file's `w:sectPrChange` record). */
function withRecordedElement(
  current: OoxmlElement | null,
  record: OrderedElementRecord,
  mint: () => string
) {
  const given = (record.element?.children ?? []).filter(
    (c) => !isReference(c) && !isWml(c, 'sectPrChange')
  );
  const base = current ?? newSection(mint);
  return withChildren(base, [...base.children.filter(isReference), ...given]);
}

/** A section with its reference of one kind and variant set to `id`, or removed. */
function withReference(
  section: OoxmlElement,
  op: Extract<OrderedSectionOp, { op: 'reference' }>,
  mint: () => string
): OoxmlElement {
  const localName = op.kind === 'header' ? 'headerReference' : 'footerReference';
  const kept = section.children.filter(
    (c) =>
      c.kind === 'textValue' || c.localName !== localName || attributeOf(c, 'type') !== op.variant
  );
  if (op.id === null) return withChildren(section, kept);
  const reference = build(
    mint(),
    'generic',
    localName,
    [wmlAttribute('type', op.variant), relationshipIdAttribute(op.id)],
    []
  );

  // As the engine writes it: at the head of the sequence, where references go (`CT_SectPr`)
  return withChildren(section, [reference, ...kept]);
}

/** Page size, orientation and margins, by the engine's own write on a scratch part. */
function withPage(
  ctx: ComposeContext,
  section: OoxmlElement | null,
  op: Extract<OrderedSectionOp, { op: 'page' }>
): OoxmlElement | null {
  const { op: _page, ...fields } = op;
  const scratch = withStoryBlocks(ctx.story.part, section ? [section] : []);
  const result = applySetSectionProperties(scratch, { op: 'setSectionProperties', ...fields });
  if (!result.ok) return section;
  const written = bodySectionOf(result.part);
  return written && written.kind !== 'textValue' ? (written as OoxmlElement) : section;
}

/** One write on one section. */
function applySectionOp(
  ctx: ComposeContext,
  section: OoxmlElement | null,
  op: OrderedSectionOp
): OoxmlElement | null {
  const mint = ctx.mint;
  switch (op.op) {
    case 'mark':
      return op.sectPr;
    case 'page':
      return withPage(ctx, section, op);
    case 'furniture': {
      const { op: _furniture, ...options } = op;
      const base = section ?? newSection(mint);
      return withChildren(base, sectionFurnitureChildren(base.children, options, mint));
    }
    case 'reference':
      return withReference(section ?? newSection(mint), op, mint);
    case 'type': {
      // A section that was never written starts from the defaults the engine writes out
      const base =
        section ??
        withChildren(newSection(mint), defaultSectionChildren(metricsOfSection(null), mint));
      return withChildren(base, withSectionBreakType(base.children, op.type, mint));
    }
    case 'setNoteProperties': {
      const base = section ?? newSection(mint);
      return withChildren(base, sectionChildrenWithNoteProperties(base.children, op, mint));
    }
  }
}

/**
 * Whether the engine's write for a press by `revision`'s author leaves the section as it shows
 * (`section`, with any change record): the section has every value the press sets, and the
 * engine's tracked write records nothing for it (`withSectionChangeRecord`). Word records none
 * either.
 */
export function sectionOpKeeps(
  ctx: ComposeContext,
  section: OoxmlElement | null,
  op: OrderedSectionOp,
  revision: RevisionAttributionInput
): boolean {
  const prior = section?.children ?? [];
  const after = withSectionChangeRecord({
    prior,
    next: applySectionOp(ctx, section, op)?.children ?? [],
    revision,
    mint: ctx.mint,
    nextRevisionId: ctx.mint,
  });
  const record = (children: readonly OoxmlNode[]) => children.find((c) => isWml(c, 'sectPrChange'));
  return sameSectionProperties(after, prior) && record(after) === record(prior);
}

function applyRecord(
  ctx: ComposeContext,
  section: OoxmlElement | null,
  record: SectionRecord
): OoxmlElement | null {
  return record.kind === 'element'
    ? withRecordedElement(section, record, ctx.mint)
    : applySectionOp(ctx, section, record.op);
}

/** Ids a composed section takes: a function of its holder, so a compose again keeps them. */
function holderIds(holder: OrderedItem): () => string {
  let n = 0;
  return () => `${holder.space}#sect${(n += 1)}`;
}

/** A section without its change record (`w:sectPrChange`): the records draw it again. */
function withoutSectionRecord(base: OoxmlElement | null): OoxmlElement | null {
  if (!base) return null;
  return withChildren(
    base,
    base.children.filter((c) => !isWml(c, 'sectPrChange'))
  );
}

/**
 * The section a holder shows: `base` (its own, as loaded) under its records, or null when it
 * has none. The accepted writes apply in the order changes made them, then the pending ones
 * with one `w:sectPrChange` holding the section before them, under the latest pending author,
 * as Word keeps one change record for each section (ECMA-376 Part 1, §17.13.5.32) and the
 * engine's tracked section writes record it (`sectionRecordChildren`). A file's own record alone
 * is written back as the file had it.
 */
export function sectionAfter(
  ctx: ComposeContext,
  holder: OrderedItem,
  base: OoxmlElement | null
): OoxmlElement | null {
  const records = sectionRecordsOf(holder);
  if (records.length === 0) return base;

  // A write Word makes untracked shows as a decided one, pending or rejected
  const kept = (r: SectionRecord) => r.kind === 'sectionOp' && r.keptOnReject === true;
  const shown = acceptedAndPending(
    ctx,
    records.filter((r) => !kept(r))
  );
  const accepted = [...records.filter(kept), ...shown.accepted].sort((a, b) => a.seq - b.seq);
  const pending = shown.pending;
  const [earlier, later] = aroundPending(accepted, pending);
  let section = withoutSectionRecord(base);
  for (const record of earlier) section = applyRecord(ctx, section, record);
  const prior = section?.children ?? [];
  const taken: DrawnTakeover[] = [];
  for (const record of pending) {
    const before = section?.children ?? [];
    section = applyRecord(ctx, section, record);
    noteTakeover(ctx, taken, record, before, section?.children ?? []);
  }

  for (const record of later) section = applyRecord(ctx, section, record);
  if (section === null) return null;
  const fileRecord = fileRecordAlone(accepted, pending);
  const latest = pending.filter((r) => !(r.kind === 'sectionOp' && r.untracked === true)).at(-1);
  const ids = holderIds(holder);
  let children = section.children.filter((c) => !isWml(c, 'sectPrChange'));
  if (fileRecord) children = [...children, fileRecord];
  else if (latest) {
    const record = build(
      ids(),
      'generic',
      'sectPrChange',
      ctx.revisionMark(latest.suggestion, REVISION_ROLE.sectionProperties),
      [build(ids(), 'generic', 'sectPr', [], sectionRecordChildren(prior, children, ids))]
    );
    children = [...children, record];
  }

  const drawn = cloneWithNewIds(withChildren(section, children), ids) as OoxmlElement;
  ctx.keepTakeovers(drawn.children, taken);
  return drawn;
}

/**
 * The section a holder shows without its pending tracked writes: `base` under its decided
 * records and the untracked ones still pending, in the order changes made them. An untracked
 * record has no change record to take back (a break's own section on its author's new mark, a
 * type Word writes untracked), so the section shows it as it is. What a section break's copy
 * takes as its own: the engine copies the section as it shows, and the pending tracked writes
 * go with the copy under their own suggestions.
 */
export function sectionWithoutPendingWrites(
  ctx: ComposeContext,
  holder: OrderedItem,
  base: OoxmlElement | null
): OoxmlElement | null {
  const shown = (r: SectionRecord): boolean => {
    const status = statusOf(ctx.doc, r.suggestion);
    if (r.kind === 'sectionOp' && r.keptOnReject === true) return true;
    if (r.kind === 'sectionOp' && r.untracked === true) return status === 'pending';
    return status === 'accepted';
  };
  let section = withoutSectionRecord(base);
  const records = sectionRecordsOf(holder)
    .filter(shown)
    .sort((a, b) => a.seq - b.seq);
  for (const record of records) section = applyRecord(ctx, section, record);
  return section;
}

/** A paragraph mark's own section, as its node holds it. */
export function markSection(mark: OrderedItem): OoxmlElement | null {
  const shell = mark.node;
  if (!shell || shell.kind === 'textValue') return null;
  const pPr = childWhere(shell, isParagraphProperties);
  return (pPr && childWhere(pPr, (c) => isWml(c, 'sectPr'))) ?? null;
}

/**
 * The `w:sectPr` an item holds: a paragraph mark's (the mark ends a section), or the body's last
 * section. Null for any other item. The notes' properties write itself is a section record
 * (`model/sections.ts`), which compose runs with the section's other writes.
 */
export function sectionPropertiesOf(item: OrderedItem): OoxmlElement | null {
  const node = item.node;
  if (!node || node.kind === 'textValue') return null;
  if (item.kind === 'zero') return item.block === true ? asWml(node, 'sectPr') : null;
  return item.kind === 'mark' ? markSection(item) : null;
}

/**
 * A mark's `w:pPr` with the section its records give in place of its own: last, before a
 * change record (`CT_PPr`). Unchanged when the mark has no section record.
 */
export function withMarkSection(
  ctx: ComposeContext,
  mark: OrderedItem,
  pPr: OoxmlElement | null
): OoxmlElement | null {
  if (sectionRecordsOf(mark).length === 0) return pPr;
  const own = pPr ? (childWhere(pPr, (c) => isWml(c, 'sectPr')) ?? null) : null;
  const section = sectionAfter(ctx, mark, own);
  const base = pPr ?? build(ctx.mint(), 'paragraphProperties', 'pPr', [], []);
  const children = base.children.filter((c) => !isWml(c, 'sectPr'));
  if (section) {
    const at = children.findIndex((c) => isWml(c, 'pPrChange'));
    children.splice(at < 0 ? children.length : at, 0, section);
  }

  return children.length === 0 && !pPr ? null : withChildren(base, children);
}

/** The body's own section item: its node under its records (null when none shows). */
export function bodySectionNode(ctx: ComposeContext, item: OrderedItem): OoxmlElement | null {
  const base = item.node && item.node.kind !== 'textValue' ? item.node : null;
  return sectionAfter(ctx, item, base);
}

// ---- the sections of a story, as the engine counts them ----

/** One section: the item that holds its properties, and its properties as they show. */
export interface OrderedSectionView {
  readonly holder: OrderedItem;
  /** The section's properties as compose writes them; null for one the file never wrote. */
  readonly sectPr: OoxmlElement | null;
  /** The body's own section (the last one), held by a block item. */
  readonly body: boolean;
}

/**
 * The block item that holds the body's own section: loaded, or the one load gives a file
 * without.
 */
export const isSectionItem = (item: OrderedItem): boolean =>
  item.kind === 'zero' && item.block === true && (item.node === null || isWml(item.node, 'sectPr'));

/**
 * The sections of the main story, in the engine's order (`collectSectionPropertyNodes`): the
 * paragraphs that end a section, in the body and in block controls but never in a table, then
 * the body's own section unless the last block ends a section and the body states none.
 */
export function orderedSectionsOf(ctx: ComposeContext): OrderedSectionView[] {
  const doc: OrderedDocument = ctx.doc;
  const story: OrderedStory = ctx.story;
  const out: OrderedSectionView[] = [];
  let tables = 0;
  let anyBlock = false;
  let lastEnds = false;
  let body: OrderedItem | null = null;
  for (const item of story.tree.items()) {
    if (item.token === 'tbl') {
      if (item.kind === 'open') {
        if (tables === 0 && !isHidden(doc, item)) {
          anyBlock = true;
          lastEnds = false;
        }

        tables += 1;
      } else tables -= 1;
      continue;
    }

    if (tables > 0) continue;
    if (isSectionItem(item)) {
      body = item;
      continue;
    }

    if (item.kind !== 'mark' || isHidden(doc, item)) continue;
    anyBlock = true;
    const sectPr = sectionAfter(ctx, item, markSection(item));
    lastEnds = sectPr !== null;
    if (sectPr) out.push({ holder: item, sectPr, body: false });
  }

  if (body === null) return out;
  const sectPr = bodySectionNode(ctx, body);
  if (!anyBlock || !lastEnds || sectPr !== null) out.push({ holder: body, sectPr, body: true });
  return out;
}
