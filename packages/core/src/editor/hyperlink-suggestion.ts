// Links proposed while suggesting: a new link over existing words, a new target, tooltip or
// words for an existing link, or the removal of a link.
//
// Each strikes the words and inserts a copy right after them, in a link unless the link is
// being removed: the struck-then-replacement order the review queue pairs into one card. A new
// target or a removal needs this because OOXML has no revision element for either; done in
// place, neither could be reviewed or rejected. A new link needs it because a bare wrap of
// words is no reviewable change. New words need it to keep the link's format: inserted into
// the link beside the struck words, they take none, and only the op that makes a link writes
// its character style.
//
// The copy is plain text given the words' own format, so only words a copy reproduces
// exactly are proposed: text characters in one format an op can write, held by the paragraph
// or the link (or by a pending insertion there, as another author's proposed link holds its
// words), and for an existing link no target frame or document location. Anything else is
// refused rather than copied with something missing, except new words alone, which the lane
// still writes into the link itself.

import type { TreeApplyResult, TreeDocxSessionView } from '@docx-editor.dev/core/binding';
import {
  authorableHyperlinkTarget,
  findNode,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { SemanticPosition } from '@docx-editor.dev/core/layout';
import { fromHyphenGlyphs } from '../store/package/hyphen-text.ts';
import {
  AUTHORABLE_RUN_PROPERTIES,
  authoredProperties,
  propertyContainer,
} from '../store/store/direct-properties.ts';
import { isTextRun } from '../store/store/tree-op-hyperlink-divisible.ts';
import { attributeValueOf, parentOf } from '../store/store/tree-op-nodes.ts';
import type { RevisionAddress } from '../store/store/tree-op-revision-attribution.ts';
import { segmentsOf } from '../store/store/tree-op-segments.ts';
import type { OoxmlProperty } from '../store/store/tree-op-types.ts';
import {
  hyperlinkStyleWrite,
  unlinkedStyleId,
  type HyperlinkStyleWrite,
} from './hyperlink-style.ts';
import type { HyperlinkOpsDeps, SurfaceHyperlink } from './surface-hyperlinks.ts';

/** Proofing marks: they change no glyph, so a copy may go without them. */
const PROOFING_PROPERTIES: ReadonlySet<string> = new Set(['lang', 'noProof']);

// Exported for the diagnostic codes, which match a refusal by its sentence.
export const LINK_SUGGESTION_REFUSAL =
  'only a link of plain text in one format can change its target or tooltip in suggesting mode';
export const NEW_LINK_REFUSAL =
  'only plain text in one format can become a link in suggesting mode';
export const UNLINK_REFUSAL =
  'only a link of plain text in one format can be removed in suggesting mode';
export const FRESH_STRIKE_REFUSAL =
  'this link cannot change while your deletion right beside it is under a minute old; wait a minute, or accept or reject that deletion first';

/** What a link's requested target, text and tooltip are. */
interface LinkInput {
  readonly url?: string;
  readonly anchor?: string;
  readonly text?: string;
  readonly tooltip?: string;
}

/** A run's own format: its character style and the properties an op can write. */
interface RunFormat {
  readonly key: string;
  readonly properties: readonly OoxmlProperty[];
  readonly styleId?: string;
}

/** What the hyperlink lane lends this proposal. */
export interface LinkSuggestionLane {
  readonly linkAtCaret: () => SurfaceHyperlink | null;
  readonly storyPart: () => OoxmlPart | null;
  readonly withinLink: (
    link: SurfaceHyperlink,
    range: { from: SemanticPosition; to: SemanticPosition }
  ) => boolean;
  readonly applyOps: (
    ops: Parameters<TreeDocxSessionView['applyTreeOps']>[0],
    before?: Parameters<TreeDocxSessionView['applyTreeOps']>[1],
    after?: Parameters<TreeDocxSessionView['applyTreeOps']>[2],
    options?: Parameters<TreeDocxSessionView['applyTreeOps']>[4]
  ) => TreeApplyResult;
}

/**
 * A proposal suggesting mode can make, ready to commit; why it refuses one; or null when it
 * makes none and the lane edits as before.
 *
 * Read before anything is written, so `can` reads the same refusal `exec` commits by.
 */
type Proposal = { readonly refusal: string } | { readonly commit: () => boolean } | null;

/**
 * Propose the link the lane was asked for as a tracked replacement, when suggesting.
 *
 * Null when this does not apply, and the lane edits as before. Otherwise whether the proposal
 * committed; a refusal commits nothing, mints no relationship, and says why.
 */
export function suggestLink(
  deps: HyperlinkOpsDeps,
  lane: LinkSuggestionLane,
  input: LinkInput
): boolean | null {
  return settle(deps, linkProposal(deps, lane, input, deps.orderedRange()));
}

/**
 * Why suggesting would refuse to propose `input` over `range`, or null. Writes nothing; the
 * caller passes a range it read without settling queued typing.
 */
export function linkSuggestionRefusal(
  deps: HyperlinkOpsDeps,
  lane: LinkSuggestionLane,
  input: LinkInput,
  range: { from: SemanticPosition; to: SemanticPosition }
): string | null {
  return refusalOf(linkProposal(deps, lane, input, range));
}

/**
 * Propose removing `link` when suggesting: its words stay struck in the link, and a plain copy
 * of them follows.
 *
 * Null when not suggesting, or when the link is this author's own pending insertion, which the
 * lane unwraps in place; when that link is their copy of words they struck for it, the whole
 * proposal is withdrawn instead. Otherwise whether the proposal committed; a refusal says why.
 */
export function suggestUnlink(
  deps: HyperlinkOpsDeps,
  lane: LinkSuggestionLane,
  link: SurfaceHyperlink
): boolean | null {
  return settle(deps, unlinkProposal(deps, lane, link));
}

/** Why suggesting would refuse to propose removing `link`, or null. Writes nothing. */
export function unlinkSuggestionRefusal(
  deps: HyperlinkOpsDeps,
  lane: LinkSuggestionLane,
  link: SurfaceHyperlink
): string | null {
  return refusalOf(unlinkProposal(deps, lane, link));
}

/** Commit `proposal`, or refuse it with its reason; null when there is none. */
function settle(deps: HyperlinkOpsDeps, proposal: Proposal): boolean | null {
  if (proposal === null) return null;
  return 'refusal' in proposal ? refuse(deps, proposal.refusal) : proposal.commit();
}

function refusalOf(proposal: Proposal): string | null {
  return proposal !== null && 'refusal' in proposal ? proposal.refusal : null;
}

/** The lane's own split: the caret in a link changes it, anything else makes a new one. */
function linkProposal(
  deps: HyperlinkOpsDeps,
  lane: LinkSuggestionLane,
  input: LinkInput,
  range: { from: SemanticPosition; to: SemanticPosition }
): Proposal {
  const link = lane.linkAtCaret();
  const collapsed =
    range.from.paragraphId === range.to.paragraphId && range.from.offset === range.to.offset;
  if (link && (collapsed || lane.withinLink(link, range)))
    return linkChangeProposal(deps, lane, link, input);
  return collapsed ? null : newLinkProposal(deps, lane, range, input);
}

/**
 * A new target, tooltip or words for `link`. Null when not suggesting, when nothing would
 * change, or when the link is this author's own pending insertion, which the lane changes in
 * place. Null too for new words alone that a copy cannot carry: the lane writes them into the
 * link itself, where they stay linked without the link's format.
 */
function linkChangeProposal(
  deps: HyperlinkOpsDeps,
  lane: LinkSuggestionLane,
  link: SurfaceHyperlink,
  input: LinkInput
): Proposal {
  if (landingAfterStrike(deps, link) === null) return null;
  const retargets = changesTarget(link, input);
  if (!retargets && !changesText(link, input)) return null;

  const text = input.text ?? fromHyphenGlyphs(link.text);
  const part = lane.storyPart();
  const format = part ? linkFormat(part, link.id) : null;
  if (!format || text.length === 0) return retargets ? { refusal: LINK_SUGGESTION_REFUSAL } : null;
  return {
    commit: () => {
      const relationshipId = mintFor(deps, input);
      if (relationshipId === null) return false;
      const tooltip = input.tooltip ?? link.tooltip;
      // The link's own character style, or Word's `Hyperlink` for a link that had none.
      const style =
        format.styleId === undefined
          ? hyperlinkStyleWrite(deps.session.documentStyles())
          : { styleId: format.styleId, packageEdits: [] };
      const { paragraphId } = link;
      const end = link.end + text.length;
      const target = { relationshipId, anchor: input.anchor, tooltip };
      const ops: TreeDocOp[] = [
        // The copy lands in the link, right after its struck words, then leaves it for its own.
        ...replacementOps(paragraphId, link.start, link.end, text, lastCharacterAim(deps, link)),
        { op: 'removeHyperlink', linkId: link.id, range: { start: link.end, end } },
        ...linkedCopyOps(paragraphId, link.end, end, format, target, style.styleId),
      ];
      return commitProposal(deps, lane, ops, style.packageEdits);
    },
  };
}

/** A new link over the selected words, which stay struck before the linked copy. */
function newLinkProposal(
  deps: HyperlinkOpsDeps,
  lane: LinkSuggestionLane,
  range: { from: SemanticPosition; to: SemanticPosition },
  input: LinkInput
): Proposal {
  const paragraphId = range.from.paragraphId;
  // A link cannot span paragraphs; the lane says so itself.
  if (range.to.paragraphId !== paragraphId) return null;
  const { offset: start } = range.from;
  const { offset: end } = range.to;
  // Non-null exactly when suggesting: where the copy lands once the words are struck.
  const landing = deps.replacementLanding?.(paragraphId, start, end);
  if (landing === null || landing === undefined) return null;
  const text = input.text ?? deps.textOf(paragraphId).slice(start, end);
  const part = lane.storyPart();
  const format = part ? selectionFormat(part, paragraphId, start, end) : null;
  if (!format || text.length === 0) return { refusal: NEW_LINK_REFUSAL };
  return {
    commit: () => {
      const relationshipId = mintFor(deps, input);
      if (relationshipId === null) return false;
      const style = hyperlinkStyleWrite(deps.session.documentStyles());
      const copyEnd = landing + text.length;
      const target = { relationshipId, anchor: input.anchor, tooltip: input.tooltip };
      const ops: TreeDocOp[] = [
        ...replacementOps(paragraphId, start, end, text),
        ...linkedCopyOps(paragraphId, landing, copyEnd, format, target, style.styleId),
      ];
      // The caret lands after the linked copy, where the next keystroke belongs.
      const after = { paragraphId, offset: copyEnd };
      return commitProposal(deps, lane, ops, style.packageEdits, after);
    },
  };
}

/** Removing `link`: null when not suggesting, or for this author's own pending link. */
function unlinkProposal(
  deps: HyperlinkOpsDeps,
  lane: LinkSuggestionLane,
  link: SurfaceHyperlink
): Proposal {
  const landing = landingAfterStrike(deps, link);
  if (landing === null) return withdrawalProposal(deps, lane, link);
  const part = lane.storyPart();
  if (!part || !linkFormat(part, link.id)) return { refusal: UNLINK_REFUSAL };
  return {
    commit: () => {
      const { paragraphId } = link;
      const text = deps.textOf(paragraphId).slice(link.start, link.end);
      const end = landing + text.length;
      const ops: TreeDocOp[] = [
        // The copy replaces the struck words, so it lands after them in the link and takes
        // their run's whole format.
        ...replacementOps(paragraphId, link.start, link.end, text),
        // Then the link lets go of the copy alone, and of the style it gave it, as an untracked
        // unlink does: the copy reads as the plain text it will be.
        {
          op: 'removeHyperlink',
          linkId: link.id,
          range: { start: landing, end },
          styleId: unlinkedStyleId(deps.session.documentStyles()),
        },
      ];
      // The caret lands after the copy, where the next keystroke belongs.
      return commitProposal(deps, lane, ops, [], { paragraphId, offset: end });
    },
  };
}

/**
 * Withdrawing this author's own proposed link, when taking it off would leave nothing to review:
 * the link holds only their copy of the words they struck right before it, so the copy, unlinked,
 * reads as those words again and the card as "Docs" replaced with "Docs". Both halves are
 * rejected and the words stand as they were. Null when not suggesting, and for any other own
 * link, which the lane unwraps in place: new or changed words stay proposed.
 */
function withdrawalProposal(
  deps: HyperlinkOpsDeps,
  lane: LinkSuggestionLane,
  link: SurfaceHyperlink
): Proposal {
  const landing = deps.replacementLanding?.(link.paragraphId, link.start, link.end);
  const part = lane.storyPart();
  const pair = landing === null || landing === undefined || !part ? null : copiedWords(part, link);
  if (!pair) return null;
  const ops: TreeDocOp[] = pair.map((node) => ({
    op: 'rejectRevision',
    revision: revisionAddressOf(node),
    localName: node.kind === 'revisionInsert' ? 'ins' : 'del',
    siteNodeIds: [node.id],
  }));
  // The struck words come back right before where the link began, and the caret follows them.
  return {
    commit: () =>
      commitProposal(deps, lane, ops, [], { paragraphId: link.paragraphId, offset: link.start }),
  };
}

/**
 * The link's own insertion and the deletion right before it, when the one copies the other: one
 * author, one edit, the same text, and nothing in either but text runs.
 */
function copiedWords(part: OoxmlPart, link: SurfaceHyperlink): readonly OoxmlNode[] | null {
  const node = findNode(part, link.id);
  const paragraph = parentOf(part, link.id);
  if (!node || node.kind !== 'hyperlink' || paragraph?.kind !== 'paragraph') return null;
  const insertion = node.children.length === 1 ? node.children[0]! : null;
  const index = paragraph.children.findIndex((child) => child.id === link.id);
  const deletion = index > 0 ? paragraph.children[index - 1]! : null;
  if (insertion?.kind !== 'revisionInsert' || deletion?.kind !== 'revisionDelete') return null;
  if (
    !insertion.children.every((run) => isTextRun(run, 't')) ||
    !deletion.children.every((run) => isTextRun(run, 'delText'))
  ) {
    return null;
  }
  const sameEdit = ['author', 'date'].every(
    (name) => attributeValueOf(insertion, name) === attributeValueOf(deletion, name)
  );
  return sameEdit && textUnder(insertion) === textUnder(deletion) ? [insertion, deletion] : null;
}

/** The `w:id`, author and date that address a revision element. */
function revisionAddressOf(node: OoxmlNode): RevisionAddress {
  const id = attributeValueOf(node, 'id') ?? '';
  const author = attributeValueOf(node, 'author') ?? '';
  const date = attributeValueOf(node, 'date');
  return date === undefined ? { id, author } : { id, author, date };
}

/** The characters under a node, `w:t` or `w:delText` alike. */
function textUnder(node: OoxmlNode): string {
  if (node.kind === 'textValue') return node.value;
  return node.children.map(textUnder).join('');
}

/**
 * Where a copy of `link`'s words lands once they are struck. Null when not suggesting, or when
 * the strike would retract the whole link as this author's own insertion: nothing to review.
 */
function landingAfterStrike(deps: HyperlinkOpsDeps, link: SurfaceHyperlink): number | null {
  const landing = deps.replacementLanding?.(link.paragraphId, link.start, link.end);
  return landing === null || landing === undefined || landing <= link.start ? null : landing;
}

/**
 * Strike `[start, end)` and insert `text` right after the struck words.
 *
 * The insertion aims at `aim`, the strike's front edge unless a caller names a point inside the
 * struck words. The store moves an insertion at the front edge past the words the same
 * transaction struck, into a link that holds them, and gives it their run's format. An aim at
 * the far edge sits on any bookmark end there, and text inserted at a bookmark's end stays
 * outside the bookmark: the copy left a bookmark around the words, and accepting emptied it.
 */
function replacementOps(
  paragraphId: string,
  start: number,
  end: number,
  text: string,
  aim = start
): TreeDocOp[] {
  return [
    { op: 'deleteText', paragraphId, start, end },
    { op: 'insertText', paragraphId, offset: aim, text },
  ];
}

/**
 * An aim before `link`'s last character, or at its start for a link of one character.
 *
 * Inside the struck words, the store puts the insertion right after the deletion that holds the
 * aim, in the link. An aim at an edge of the strike follows it instead, and a strike that joins
 * one this author made just outside the link a moment ago ends outside the link. The copy takes
 * no format there; the caller gives it one.
 */
function lastCharacterAim(deps: HyperlinkOpsDeps, link: SurfaceHyperlink): number {
  const last = deps.textOf(link.paragraphId).charCodeAt(link.end - 1);
  // A low surrogate ends a pair, which an offset must not split.
  const aim = link.end - (last >= 0xdc00 && last <= 0xdfff ? 2 : 1);
  return Math.max(link.start, aim);
}

/**
 * Format the copy at `[start, end)` and wrap it in its link.
 *
 * The format is given outright: which run the copy joins depends on what sits around where it
 * lands, which need not be the original words. On this author's own insertion the write is
 * direct, not a format revision.
 */
function linkedCopyOps(
  paragraphId: string,
  start: number,
  end: number,
  format: RunFormat,
  target: { readonly relationshipId?: string; readonly anchor?: string; readonly tooltip?: string },
  styleId: string
): TreeDocOp[] {
  return [
    { op: 'setRunProperties', paragraphId, start, end, properties: format.properties },
    {
      op: 'insertHyperlink',
      paragraphId,
      start,
      end,
      ...(target.relationshipId !== undefined ? { relationshipId: target.relationshipId } : {}),
      ...(target.anchor ? { anchor: target.anchor } : {}),
      ...(target.tooltip !== undefined ? { tooltip: target.tooltip } : {}),
      styleId,
    },
  ];
}

/** Refuse with `reason`, as the lane's other refusals do, and write nothing. */
function refuse(deps: HyperlinkOpsDeps, reason: string): false {
  deps.commit(() => ({ committed: false, rejected: true, opCount: 0, reason }));
  return false;
}

/**
 * The relationship for an external target, minted only once the proposal is allowed, so a
 * refusal leaves none behind. Undefined for a bookmark target, null when the URL is refused.
 */
function mintFor(deps: HyperlinkOpsDeps, input: LinkInput): string | undefined | null {
  if (input.url === undefined || input.url.length === 0) return undefined;
  const url = input.url;
  const mint = (): string | null =>
    deps.session.ensureHyperlinkRelationship(url, deps.storyScope());
  return deps.withMintActor ? deps.withMintActor(mint) : mint();
}

/** Commit the proposal as one undo step, the caret moving to `after` when given. */
function commitProposal(
  deps: HyperlinkOpsDeps,
  lane: LinkSuggestionLane,
  ops: readonly TreeDocOp[],
  packageEdits: HyperlinkStyleWrite['packageEdits'],
  after?: SemanticPosition
): boolean {
  let committed = false;
  deps.commit(
    () => {
      const result = lane.applyOps(
        ops,
        deps.selectionMark(),
        after
          ? { paragraphId: after.paragraphId, start: after.offset, end: after.offset }
          : undefined,
        { packageEdits }
      );
      committed = result.committed;
      // A deletion of the author's own from the last minute merges with the strike and carries
      // the copy out of the link, so the store refuses the range; say why instead of its code.
      return result.committed || result.reason !== 'invalid-range'
        ? result
        : { ...result, reason: FRESH_STRIKE_REFUSAL };
    },
    () => (after ? { anchor: after, head: after } : null)
  );
  return committed;
}

/** Whether the requested target or tooltip differs from the link's, compared as written. */
function changesTarget(link: SurfaceHyperlink, input: LinkInput): boolean {
  if (input.tooltip !== undefined && input.tooltip !== link.tooltip) return true;
  if (input.anchor) return link.kind !== 'internal' || link.anchor !== input.anchor;
  return link.kind !== 'external' || authorableHyperlinkTarget(input.url ?? '') !== link.authored;
}

/** Whether the requested words differ from the link's live words. */
function changesText(link: SurfaceHyperlink, input: LinkInput): boolean {
  return input.text !== undefined && input.text !== fromHyphenGlyphs(link.text);
}

/**
 * The format a copy of the link must carry, or null when a copy would lose something: the link
 * sits in a wrapper, holds anything but text runs, mixes formats, or has a target frame or
 * document location the new link cannot take. Runs in a pending insertion count: another
 * author's proposed link holds its words in their insertion.
 */
function linkFormat(part: OoxmlPart, linkId: string): RunFormat | null {
  const link = findNode(part, linkId);
  if (!link || link.kind !== 'hyperlink' || parentOf(part, linkId)?.kind !== 'paragraph') {
    return null;
  }
  if (link.attributes.some((a) => a.localName === 'tgtFrame' || a.localName === 'docLocation')) {
    return null;
  }
  const runs = link.children.flatMap((child): readonly OoxmlNode[] =>
    child.kind === 'revisionInsert' ? child.children : [child]
  );
  return runs.every((run) => isTextRun(run, 't')) ? commonFormat(runs) : null;
}

/**
 * The format of the selected words, or null when a copy would lose something: they hold
 * anything but text characters (a tab, a break, a picture, a field), sit in a link, a control
 * or a deletion, or mix formats. Text in a pending insertion is still the paragraph's text.
 */
function selectionFormat(
  part: OoxmlPart,
  paragraphId: string,
  start: number,
  end: number
): RunFormat | null {
  const paragraph = findNode(part, paragraphId);
  if (!paragraph || paragraph.kind !== 'paragraph') return null;
  const runs: OoxmlNode[] = [];
  for (const segment of segmentsOf(paragraph)) {
    if (segment.end <= start || segment.start >= end) continue;
    if (segment.removeNodeIds || segment.node.kind !== 'textValue') return null;
    if (parentOf(part, segment.node.id)?.kind !== 'text') return null;
    const run = findNode(part, segment.runId);
    const holder = parentOf(part, segment.runId);
    const held =
      holder?.kind === 'paragraph' ||
      (holder?.kind === 'revisionInsert' && parentOf(part, holder.id)?.kind === 'paragraph');
    if (!run || !held) return null;
    runs.push(run);
  }
  return commonFormat(runs);
}

/** The one format every run shares, or null when they differ or one cannot be written. */
function commonFormat(runs: readonly OoxmlNode[]): RunFormat | null {
  let format: RunFormat | null = null;
  for (const run of runs) {
    const next = runFormat(run);
    if (!next || (format && format.key !== next.key)) return null;
    format = next;
  }
  return format;
}

/** A run's own format, or null when it carries a visible property no op can write. */
function runFormat(run: OoxmlNode): RunFormat | null {
  const container = propertyContainer(run, 'runProperties', 'rPr');
  const children = container && container.kind !== 'textValue' ? container.children : [];
  let styleId: string | undefined;
  for (const child of children) {
    if (child.kind === 'textValue' || child.namespaceUri !== WML_NAMESPACE_URI) return null;
    if (child.localName === 'rStyle') {
      styleId = child.attributes.find((a) => a.localName === 'val')?.value;
    } else if (
      !AUTHORABLE_RUN_PROPERTIES.has(child.localName) &&
      !PROOFING_PROPERTIES.has(child.localName)
    ) {
      return null;
    }
  }
  const properties = authoredProperties(container, AUTHORABLE_RUN_PROPERTIES);
  return {
    key: JSON.stringify([styleId ?? null, properties]),
    properties,
    ...(styleId !== undefined ? { styleId } : {}),
  };
}
