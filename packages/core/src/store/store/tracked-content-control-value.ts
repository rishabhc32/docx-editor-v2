// Tracked content-control VALUES — a pick, a toggle or a date proposed in suggesting mode.
//
// Word records a value change as CONTENT: the old display text struck in a `w:del` and the new
// one in a `w:ins`, both inside `w:sdtContent`. A checkbox toggle is the same edit over its
// glyph. `CT_SdtPr` has no change record, so the type's own record of the value — `w:lastValue`,
// `w:date/@w:fullDate`, `w14:checked`, `w:showingPlcHdr` — moves to the new value untracked,
// exactly as the untracked write moves it.
//
// So Reject restores the text, and each property is read back from the content it describes
// once a decision changes what the control shows (`resolveKeepingControlState`). A checkbox's
// state IS the glyph it shows, so `w14:checked` follows the live glyph: without that, a rejected
// toggle left a box drawn unchecked that reported itself checked, and the next click wrote the
// glyph it already showed. `w:lastValue` names the item a list shows, and `w:fullDate` the date
// a date control's text shows through its format, as Word reads it back. An empty control shows
// its prompt, so a decision that empties one brings the prompt back under `w:showingPlcHdr`, as
// Word does, with no value record.
//
// A prompt is never struck. Word removes it without a revision, as the control's state rather
// than its text, and proposes only what takes its place (`proposedOverPrompt`): struck, it came
// back on Reject as text the control no longer marked as its prompt.
//
// The new content is never built here. The untracked write builds it — the display run, its
// face, the glyph — and this module strikes the control's text and inserts what that write made.
// The strike takes every character in the control, so a shape where the untracked write would
// have KEPT text beside the value is refused rather than half-proposed.

import {
  contentControlPropertiesOf as contentControlSummaryOf,
  contentControlTextOf,
  contentControlsIn,
} from '../package/content-control-nodes.ts';
import { parseContentControlDate, parseIsoDate } from './content-control-dates.ts';
import {
  createNodeIdAllocator,
  findNode,
  insertChildren,
  replaceChildren,
  replaceNode,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import type {
  OoxmlElement,
  OoxmlNode,
  OoxmlParagraphNode,
  OoxmlPart,
} from '../package/ooxml-tree.ts';
import {
  decodeCheckboxGlyph,
  isInertInline,
  isInlineControl,
  isRangeEnd,
  isWml,
  runContentOf,
} from './content-control-checkbox.ts';
import { controlHoldsContent, withRestoredPrompt } from './content-control-prompt-restore.ts';
import { equivalentNodes } from './ooxml-node-equality.ts';
import { applyTreeOp } from './tree-op-apply.ts';
import { contentControlEffect, editedProperties } from './tree-op-content-controls.ts';
import {
  checkboxPayloadOf,
  contentControlContentOf,
  contentControlPropertiesOf,
  contentControlValueTypeOf,
  findContentControl,
  fromEdit,
  isShowingPlaceholder,
  isTemporaryControl,
  ok,
  paragraphContainingNode,
} from './tree-op-nodes.ts';
import {
  invalidRevisionAttribution,
  type RevisionAttributionInput,
} from './tree-op-revision-attribution.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import { resolveRevisionOperation, type RevisionResolveResult } from './tree-op-revisions.ts';
import { paragraphLength, paragraphOffsetIndex } from './tree-op-segments.ts';
import { applyDeleteTracked } from './tree-op-tracked-delete.ts';
import { build, copy, revisionAttributes } from './tree-op-tracked.ts';
import { insideOwnInsertion } from './tree-op-tracked-properties.ts';
import type { TreeDocOp, TreeOpResult } from './tree-op-types.ts';

type SetValue = Extract<TreeDocOp, { op: 'setContentControlValue' }>;

/** Bound on the walk over a checkbox's content, which comes from the file. */
const MAX_DEPTH = 32;

/** Where a control's value lives: the paragraph whose offsets address it, and its holder. */
interface ValueSite {
  readonly paragraph: OoxmlParagraphNode;
  readonly start: number;
  readonly end: number;
  /** The node the inserted value goes into: `w:sdtContent` inline, the paragraph at block level. */
  readonly container: OoxmlElement;
}

/**
 * The site of a control's value, or null for a shape a value cannot be proposed in.
 *
 * Block level takes ONE paragraph: the value stays a run edit inside it. More would need
 * tracked paragraph marks for the paragraphs the value removes, and a cell- or row-level
 * control holds no paragraph of its own at all.
 */
function valueSiteOf(part: OoxmlPart, control: OoxmlElement): ValueSite | null {
  const content = contentControlContentOf(control);
  if (!content) return null;
  if (isInlineControl(part, control.id)) {
    const paragraph = paragraphContainingNode(part, control.id);
    const span = paragraph ? paragraphOffsetIndex(paragraph).spanOf(control) : null;
    if (!paragraph || !span) return null;
    return { paragraph, start: span.start, end: span.end, container: content };
  }
  const blocks = content.children.filter((child) => !isInertInline(child));
  const only = blocks.length === 1 ? blocks[0]! : null;
  if (!only || only.kind !== 'paragraph') return null;
  const paragraph = only as OoxmlParagraphNode;
  return { paragraph, start: 0, end: paragraphLength(paragraph), container: paragraph };
}

/** Why no value can be proposed in this control, whatever the value, or null. */
function proposalShapeRefusal(part: OoxmlPart, control: OoxmlElement): string | null {
  // The write removes a temporary control, and nothing records a removed `w:sdt`.
  if (isTemporaryControl(control)) {
    return 'suggesting mode cannot propose a value for a temporary content control';
  }
  if (valueSiteOf(part, control)) return null;
  return 'suggesting mode can propose a value only in a content control with one paragraph';
}

/**
 * Why suggesting mode refuses every value for this control, or null, so `can` gives the reason
 * the write would give. `author` is the one suggesting mode attributes to; without one the
 * value is not tracked. A refusal that depends on the value — one that clears the control to
 * its prompt, or keeps text beside it — is known only at the write.
 */
export function trackedValueRefusal(
  part: OoxmlPart,
  op: TreeDocOp,
  author: string | undefined
): string | null {
  const writer = author?.trim();
  if (op.op !== 'setContentControlValue' || !writer) return null;
  const control = findContentControl(part, op.controlId);
  if (!control || insideOwnInsertion(part, control.id, writer)) return null;
  return proposalShapeRefusal(part, control);
}

/** The runs a container shows, history skipped: its own and those inside an insertion. */
function liveRuns(container: OoxmlElement): readonly OoxmlElement[] {
  return container.children.flatMap((child): OoxmlElement[] => {
    if (child.kind === 'run') return [child];
    if (child.kind === 'revisionInsert' || child.kind === 'revisionMoveTo') {
      return child.children.flatMap((inner) => (inner.kind === 'run' ? [inner] : []));
    }
    return [];
  });
}

/**
 * The runs the untracked write produced, or null when it kept a run that carries text.
 *
 * A run the write left as it was is not part of the value, and the strike over the control's
 * span would take it anyway, so the proposal would delete text the value never replaced. An
 * empty run is not part of it either: clearing a value proposes the strike alone.
 */
function writtenValueRuns(before: ValueSite, after: ValueSite): readonly OoxmlElement[] | null {
  const prior = new Map(liveRuns(before.container).map((run) => [run.id, run]));
  const offsets = paragraphOffsetIndex(after.paragraph);
  const runs: OoxmlElement[] = [];
  for (const run of liveRuns(after.container)) {
    if (offsets.lengthOf(run) === 0) continue;
    const was = prior.get(run.id);
    if (was !== undefined && equivalentNodes([was], [run])) return null;
    runs.push(run);
  }
  return runs;
}

/** Content that the written value's `w:ins` goes in front of, so a range over the value covers it. */
function insertionIndex(container: OoxmlElement): number {
  let at = container.children.length;
  while (at > 0 && isRangeEnd(container.children[at - 1]!)) at--;
  return at;
}

/**
 * The checkbox state a control's LIVE glyph shows, or null when it shows neither state or both.
 *
 * Struck and moved-away runs are history, never the display, so a pending toggle reads as the
 * state it proposes — which is what `w14:checked` says while it is pending.
 */
function checkboxStateShown(control: OoxmlElement): boolean | null {
  const payload = checkboxPayloadOf(control);
  const content = contentControlContentOf(control);
  if (!payload || !content) return null;
  const checked = decodeCheckboxGlyph(payload.checkedGlyph);
  const unchecked = decodeCheckboxGlyph(payload.uncheckedGlyph);
  if (checked === null || unchecked === null || checked === unchecked) return null;
  const shown = new Set<boolean>();
  const walk = (node: OoxmlNode, depth: number): void => {
    if (node.kind === 'textValue' || depth > MAX_DEPTH) return;
    if (node.kind === 'revisionDelete' || node.kind === 'revisionMoveFrom') return;
    if (node.kind === 'run') {
      const glyph = glyphOf(node);
      if (glyph === checked) shown.add(true);
      else if (glyph === unchecked) shown.add(false);
      return;
    }
    if (isWml(node, 'sdtPr') || node.kind === 'paragraphProperties') return;
    for (const child of node.children) walk(child, depth + 1);
  };
  walk(content, 0);
  return shown.size === 1 ? shown.has(true) : null;
}

/** The one character a run shows, through `w:sym` or `w:t`, or null for anything else. */
function glyphOf(run: OoxmlElement): string | null {
  const content = runContentOf(run);
  const only = content.length === 1 ? content[0]! : null;
  if (!only || only.kind === 'textValue') return null;
  if (isWml(only, 'sym')) {
    return decodeCheckboxGlyph(only.attributes.find((a) => a.localName === 'char')?.value);
  }
  if (only.kind !== 'text') return null;
  return only.children.map((child) => (child.kind === 'textValue' ? child.value : '')).join('');
}

/** Whether a value write changes nothing a reader sees, so there is nothing to propose. */
function showsSameValue(before: OoxmlElement, after: OoxmlElement): boolean {
  if (contentControlValueTypeOf(before) === 'checkbox') {
    const state = checkboxStateShown(before);
    return state !== null && state === checkboxStateShown(after);
  }
  return (
    !isShowingPlaceholder(before) && contentControlTextOf(before) === contentControlTextOf(after)
  );
}

/**
 * `setContentControlValue` with a revision: the old content struck and the new value inserted,
 * inside the control, with the control's properties moved to the new value.
 *
 * Refused rather than approximated for what has no faithful tracked form: a control that no
 * value can be proposed in (`proposalShapeRefusal`), a value that clears the control back to
 * its prompt (`w:showingPlcHdr` cannot be pending), and content the value would replace only in
 * part. A value over a prompt is proposed alone (`proposedOverPrompt`).
 */
export function applyTrackedContentControlValue(
  part: OoxmlPart,
  op: SetValue,
  options?: EditOptions
): TreeOpResult {
  const { revision, ...untracked } = op;
  if (!revision || invalidRevisionAttribution(revision)) {
    return { ok: false, reason: 'invalid-property-value' };
  }
  const control = findContentControl(part, op.controlId);
  if (!control) return { ok: false, reason: 'unknown-content-control' };
  // The author's own proposed control: rejecting that insertion takes the control and its
  // value together, so the value is part of the proposal already — the rule formatting follows.
  if (insideOwnInsertion(part, control.id, revision.author)) {
    return applyTreeOp(part, untracked, options);
  }
  if (proposalShapeRefusal(part, control) !== null) return { ok: false, reason: 'unsupported' };
  const written = applyTreeOp(part, untracked, options);
  if (!written.ok) return written;
  const after = findContentControl(written.part, control.id);
  if (!after || isShowingPlaceholder(after)) return { ok: false, reason: 'unsupported' };
  if (showsSameValue(control, after))
    return ok(part, contentControlEffect(control.id, 'text-local'));
  if (isShowingPlaceholder(control)) {
    return proposedOverPrompt(part, control, written, revision, options);
  }
  const before = valueSiteOf(part, control);
  const value = valueSiteOf(written.part, after);
  const runs = before && value ? writtenValueRuns(before, value) : null;
  if (!before || !runs) return { ok: false, reason: 'unsupported' };

  const struck =
    before.start < before.end
      ? applyDeleteTracked(part, before.paragraph, before.start, before.end, revision, options)
      : ok(part, contentControlEffect(control.id, 'text-local'));
  if (!struck.ok) return struck;
  const container = findNode(struck.part, before.container.id);
  if (!container || container.kind === 'textValue') return { ok: false, reason: 'tree-invariant' };
  const mint = createNodeIdAllocator(struck.part);
  const children = [...container.children];
  if (runs.length > 0) {
    const id = nextRevisionId(struck.part)();
    const copies = runs.map((run) => copy(mint, run));
    const insertion = build(
      mint(),
      'revisionInsert',
      'ins',
      revisionAttributes(id, revision),
      copies
    );
    children.splice(insertionIndex(container), 0, insertion);
  }
  const inserted = replaceChildren(struck.part, container.id, children, options);
  if (!inserted.ok) return fromEdit(inserted, contentControlEffect(control.id, 'text-local'));

  // The properties exactly as the untracked write left them. Their ids are kept, not copied:
  // the write edits `w:sdtPr` in place, and a collaboration replica addresses its children by
  // those ids — re-minted, the journal refused them as unknown and the replica realigned the
  // pick away. Ids the write minted come from the same lineage as the strike's, so none
  // collides.
  const target = findContentControl(inserted.part, control.id);
  const properties = contentControlPropertiesOf(after);
  if (!target || !properties) return { ok: false, reason: 'tree-invariant' };
  const prior = contentControlPropertiesOf(target);
  const moved = prior
    ? replaceNode(inserted.part, prior.id, properties, options)
    : insertChildren(inserted.part, target.id, 0, [properties], options);
  return fromEdit(moved, contentControlEffect(before.paragraph.id, 'flow-structural'));
}

/**
 * A write over a control's prompt, proposed: the prompt goes without a revision, and the runs the
 * untracked write put in its place go into one insertion where they stand. Rejecting it empties
 * the control, and the decision shows the prompt again.
 *
 * `written` is that untracked write. A control in the author's own pending insertion takes it
 * directly, as a value does. A shape where the write left anything beside its new runs is
 * refused rather than half-proposed.
 */
export function proposedOverPrompt(
  part: OoxmlPart,
  control: OoxmlNode,
  written: TreeOpResult,
  revision: RevisionAttributionInput,
  options?: EditOptions
): TreeOpResult {
  if (!written.ok || insideOwnInsertion(part, control.id, revision.author)) return written;
  const after = findContentControl(written.part, control.id);
  const site = after ? valueSiteOf(written.part, after) : null;
  if (!site) return { ok: false, reason: 'unsupported' };
  const children: OoxmlNode[] = [...site.container.children];
  const offsets = paragraphOffsetIndex(site.paragraph);
  const shown = children.filter((child) => offsets.lengthOf(child) > 0);
  const first = shown.length > 0 ? children.indexOf(shown[0]!) : -1;
  if (
    first < 0 ||
    shown.some((child, index) => child.kind !== 'run' || children[first + index] !== child)
  ) {
    return { ok: false, reason: 'unsupported' };
  }
  const mint = createNodeIdAllocator(written.part);
  const id = options?.trackedRevisionIds?.mint() ?? nextRevisionId(written.part)();
  const insertion = build(mint(), 'revisionInsert', 'ins', revisionAttributes(id, revision), shown);
  children.splice(first, shown.length, insertion);
  return fromEdit(
    replaceChildren(written.part, site.container.id, children, options),
    written.effect
  );
}

/**
 * Resolve revisions, then let each control the decision reached read its state back from its
 * content: a checkbox whose live glyph changed says so in `w14:checked`, and a control the
 * decision emptied shows its prompt again.
 *
 * Only a changed control moves: one the decision did not reach keeps whatever its file says,
 * consistent or not.
 */
export function resolveKeepingControlState(
  part: OoxmlPart,
  action: Parameters<typeof resolveRevisionOperation>[1],
  operation: Parameters<typeof resolveRevisionOperation>[2],
  options?: EditOptions
): RevisionResolveResult {
  const resolved = resolveRevisionOperation(part, action, operation, options);
  if (!resolved.ok || !resolved.part) return resolved;
  let next = resolved.part;
  // Innermost first: an inner control's prompt is content its emptied outer control holds.
  for (const { node } of [...contentControlsIn(part.root)].reverse()) {
    const after = findContentControl(next, node.id);
    // The same node: the decision did not reach this control.
    if (!after || after === node) continue;
    if (controlHoldsContent(node) && !controlHoldsContent(after)) {
      next = withRestoredPrompt(next, node.id, options);
    }
    const current = findContentControl(next, node.id);
    const properties = current ? contentControlPropertiesOf(current) : undefined;
    if (!current || !properties) continue;
    const edits = recordEdits(node, current);
    if (Object.keys(edits).length === 0) continue;
    const edited = editedProperties(properties, edits, createNodeIdAllocator(next));
    const written = replaceNode(next, properties.id, edited, options);
    if (!written.ok) return { ok: false, reason: 'tree-invariant' };
    next = written.part;
  }
  return { ...resolved, part: next };
}

interface RecordEdits {
  checked?: boolean;
  lastValue?: string | null;
  fullDate?: string | null;
}

/** The edits that make a control's own record agree with what a decision left it showing. */
function recordEdits(before: OoxmlElement, after: OoxmlElement): RecordEdits {
  if (contentControlValueTypeOf(after) === 'checkbox') {
    const shown = checkboxStateShown(after);
    if (shown === null || shown === checkboxStateShown(before)) return {};
    return checkboxPayloadOf(after)?.checked === shown ? {} : { checked: shown };
  }
  const unchanged =
    contentControlTextOf(after) === contentControlTextOf(before) &&
    isShowingPlaceholder(after) === isShowingPlaceholder(before);
  return unchanged ? {} : valueRecordEdits(after);
}

/**
 * The value record a list or date control's content states: `w:lastValue` the value of the item
 * it shows, or a combo box's own text, and `w:fullDate` the date it shows, read back through its
 * format as Word reads it after a decision. A control showing its prompt holds no value. Text
 * that names no item and no date leaves the record as it is.
 */
function valueRecordEdits(control: OoxmlElement): RecordEdits {
  const summary = contentControlSummaryOf(control);
  const text = contentControlTextOf(control);
  if (summary.type === 'dropDownList' || summary.type === 'comboBox') {
    if (summary.showingPlaceholder)
      return summary.lastValue === undefined ? {} : { lastValue: null };
    const item = summary.listItems.find((candidate) => candidate.displayText === text);
    const value = item?.value ?? (summary.type === 'comboBox' ? text : undefined);
    return value === undefined || value === summary.lastValue ? {} : { lastValue: value };
  }
  if (summary.type !== 'date') return {};
  const current = summary.date?.fullDate;
  if (summary.showingPlaceholder) return current === undefined ? {} : { fullDate: null };
  // A two-digit year stays in the century the record already names.
  const previous = current === undefined ? null : parseIsoDate(current);
  const century = previous ? Math.floor(previous.year / 100) * 100 : undefined;
  const shown = parseContentControlDate(text, summary.date?.dateFormat, century);
  if (!shown) return {};
  const iso = `${String(shown.year).padStart(4, '0')}-${String(shown.month).padStart(2, '0')}-${String(shown.day).padStart(2, '0')}`;
  const fullDate = `${iso}T00:00:00Z`;
  return fullDate === current ? {} : { fullDate };
}
