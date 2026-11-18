import { equationPlainText } from './surface-equations.ts';
import { selectionContainsClipboardObject } from './clipboard-object-selection.ts';
import { partOfNodeId } from './surface-scope.ts';
// Copy, cut and paste for the paginated surface (paginated-surface seam).
//
// Thin glue, on purpose. The flavour payload lives in `clipboard-copy-payload.ts`, the
// fidelity ladder in `clipboard-paste-router.ts`, and control-character normalization in
// `clipboard-plain-text.ts`. What is left here is what needs the mount closure: which range
// a paste replaces, where the replacement lands once a tracked strike has had its say, and
// the ONE commit each gesture becomes.
//
// The force-plain arm lives here too, with the paste that consumes it. It is a deadline, not
// a flag: when no paste follows the chord — a denied clipboard, a focus loss — a stale arm
// must not silently downgrade a LATER ordinary Cmd+V.
//
// Mutable mount state arrives as a getter and is read at the ORIGINAL read point, never
// hoisted: `orderedRange()` and `flushPendingInputAndLayout()` both republish the layout, so
// a `layout()` lifted above one would hand the caller the pre-flush revision.

// The VIEW, not the concrete binding session: `paginated-surface.ts` is the only file in
// this lane allowed to name that one (`store/__tests__/prosemirror-isolation.test.ts`).
import type { TreeApplyResult, TreeDocxSessionView } from '@docx-editor.dev/core/binding';
import type { CollaborationGateContext } from '../collaboration/index.ts';
import type { StoryScope, TreeDocOp } from '@docx-editor.dev/core/store';
import type { RevisionAttributionInput } from '../store/store/tree-op-types.ts';
import { findNode } from '../store/package/ooxml-edit.ts';
import { smartSpacedText, spacingNeighbours } from '../store/store/fragment-smart-spacing.ts';
import { placeholderControlForInsertion } from '../store/store/tree-op-content-controls.ts';
import {
  cellSelectionText,
  paragraphsInCells,
  type CellSelection,
  type SemanticLayout,
  type SemanticPosition,
  type SemanticSelection,
} from '@docx-editor.dev/core/layout';
import { buildCopyFlavours, type CopyFlavours } from './clipboard-copy-payload.ts';
import { routePaste } from './clipboard-paste-router.ts';
import { insertableText } from './clipboard-plain-text.ts';
import {
  paragraphTextMarkingOptionalHyphens,
  withoutOptionalHyphens,
} from '../layout/optional-hyphen-copy-text.ts';
import { fieldResultSpeller } from '../layout/field-copy-text.ts';
import { clampTextFormPaste } from './text-form-field-paste.ts';
import { promptInsertionLanding } from './content-control-prompt-landing.ts';
import {
  collapsedAt,
  fragmentCoverageOf,
  selectedTextIn,
  type RangeDeletionPlan,
} from './surface-selection-ops.ts';
import type { SurfaceEditingMode } from './paginated-surface-contract.ts';

/** A history mark: one paragraph and an offset range within it. */
type HistoryMark = { paragraphId: string; start: number; end: number };

/** What this lane borrows from the mount closure. Mutable state arrives as a getter. */

/**
 * The plain-text clipboard flavour: an optional hyphen element is a hint for line breaking,
 * not a character of the text, so plain text leaves it out. A literal U+00AD in run text is
 * text and stays. The HTML flavour keeps both.
 */
const plainCopyText = withoutOptionalHyphens;

export interface SurfaceClipboardDeps {
  session: TreeDocxSessionView;
  textFormFieldId?(): string | null;
  /**
   * False shuts the rich lane: a text box is in edit, or the surface refuses a write here (a
   * table of contents, a blocked or read-only document). The plain lane then reports why, as
   * every other write lane does.
   */
  richPasteAllowed?(): boolean;
  paragraphIds?(): readonly string[];
  layout(): SemanticLayout;
  cellSelection(): CellSelection | null;
  editingMode(): SurfaceEditingMode;
  storyScope(): StoryScope;
  paragraphOrder(): readonly string[];
  /** The collaboration actor, or undefined when no session is attached. */
  actorId(): string | undefined;
  /**
   * The collaboration write gate the typing lane asks: the refusal code, or null when the
   * gate admits — always null with no session attached, so the non-collaborative paste
   * path stays exactly as it was.
   */
  collaborationGate(
    ops: readonly TreeDocOp[],
    scope: StoryScope,
    context?: CollaborationGateContext
  ): string | null;
  /** The suggesting author, or undefined when an edit lands untracked: edit mode, no author. */
  trackedAuthor(): string | undefined;
  /** The date a tracked edit records, the one typing writes. */
  trackedDate(): string;
  /** Told after a paste commits as a tracked change, as typing tells it. */
  onTrackedChange(): void;
  /** Monotonic clock for the force-plain deadline. */
  now(): number;
  flushPendingInputAndLayout(): void;
  /** Range-edit lane: the selection in document order. */
  orderedRange(): { from: SemanticPosition; to: SemanticPosition };
  /** Range-edit lane: the current selection as a history mark. */
  selectionMark(): HistoryMark | null;
  /** Range-edit lane: what removing the selection takes, and where a replacement lands. */
  deleteSelectionPlan(): RangeDeletionPlan;
  /** The armed caret format, as ops over the pre-split offsets, consumed by this insert. */
  consumePendingFormatOps(
    paragraphId: string,
    offset: number,
    length: number,
    replacing?: { readonly start: number; readonly end: number }
  ): TreeDocOp[];
  /** Apply with the armed format, and — if the store refuses the pair — without it. */
  withoutPendingOnRejection(
    withFormat: readonly TreeDocOp[],
    withoutFormat: readonly TreeDocOp[],
    mark: HistoryMark | null,
    redoMark?: HistoryMark
  ): TreeApplyResult;
  caretMark(position: { paragraphId: string; offset: number }): HistoryMark;
  commit(
    run: () => TreeApplyResult | boolean,
    selectionAfter?: () => SemanticSelection | null
  ): void;
}

export interface SurfaceClipboardOps {
  /** Insert text, turning newlines into real paragraph splits rather than literal characters. */
  insertPlainText(text: string): void;
  /** Every clipboard flavour for the current selection — see clipboard-copy-payload.ts. */
  copyFlavoursNow(): CopyFlavours;
  /** The paste router entry — fidelity order with continuous degrade to plain. */
  pasteRichNow(text: string, html: string | null): boolean;
  /** Arm Cmd+Shift+V: the next paste inside the deadline routes plain. */
  armForcePlainPaste(): void;
}

const FORCE_PLAIN_PASTE_WINDOW_MS = 2000;

/**
 * Store refusals of a rich paste that Word does not refuse and that plain text would only
 * approximate: reported by name, with no plain-text fallback.
 */
const NAMED_PASTE_REFUSALS: ReadonlySet<string> = new Set(['fragment-paragraphs-in-link']);

export function createSurfaceClipboardOps(deps: SurfaceClipboardDeps): SurfaceClipboardOps {
  const { session } = deps;

  /**
   * Armed by Cmd+Shift+V; the next paste routes plain. Deadline-bound: when no paste
   * event follows the chord (denied clipboard, focus loss), a stale arm must not silently
   * downgrade a LATER ordinary Cmd+V.
   */
  let forcePlainPasteArmedAt: number | null = null;

  /**
   * The visible characters around a place, for Word's paste spacing; null at a control's
   * prompt, whose text a paste replaces.
   */
  function neighboursAt(at: SemanticPosition) {
    const part = partOfNodeId(session, at.paragraphId) ?? session.part();
    const paragraph = findNode(part, at.paragraphId);
    if (!paragraph || paragraph.kind !== 'paragraph') return null;
    if (placeholderControlForInsertion(part, at.paragraphId, at.offset)) return null;
    return spacingNeighbours(paragraph, at.offset);
  }

  /**
   * Insert text, turning newlines into real paragraph splits rather than literal characters.
   * `smartSpacing`: a paste, which gets Word's space on each side between two words (Smart cut
   * and paste), except in a text form field.
   */
  function insertPlainText(text: string, options?: { readonly smartSpacing?: boolean }): void {
    // Normalized first: a Windows clipboard carries CRLF, a page break arrives as a form
    // feed, and either one left in run text is a control character the store refuses —
    // which vetoes the whole transaction and makes the paste do nothing at all.
    const range = deps.orderedRange();
    const spaced = options?.smartSpacing === true && !deps.textFormFieldId?.();
    const normalized = spaced
      ? smartSpacedText(
          insertableText(text),
          neighboursAt(range.from)?.before ?? null,
          neighboursAt(range.to)?.after ?? null
        )
      : insertableText(text);
    const lines = clampTextFormPaste(
      session,
      range.from,
      range.to,
      normalized,
      deps.textFormFieldId?.() ?? null
    ).split('\n');

    // ONE COMMIT, TWO OPS, whatever the clipboard holds.
    //
    // A newline in pasted plain text is a paragraph boundary — a new `w:p`, never a
    // character in run text. Committing once per line laid out and repainted the whole
    // document per pasted paragraph, so a four-page paste cost two hundred layouts of a
    // growing document: quadratic in document size, and the reason paste lagged long
    // before typing did. The whole paste is one op list instead: the joined text lands in
    // the caret's paragraph with a single insert, and one `splitParagraphMany` cuts that
    // paragraph at every newline offset in a single pass — one rebuild of the body's child
    // sequence, however many paragraphs the clipboard carried.
    const plan = deps.deleteSelectionPlan();
    // WHERE THE REPLACEMENT GOES, the same question `type()` answers: in suggesting mode a
    // deletion keeps the characters it strikes, so the replacement belongs AFTER them.
    // Pasting at the range start put the new text in front of the struck words and left
    // the caret inside it, so the next keystroke landed mid-word.
    const target = plan.replaceAt ?? plan.collapseTo;
    const joined = lines.join('');
    const ops: TreeDocOp[] = [...plan.ops];
    // Plain text pasted at a caret takes the armed typing format, like typed text — Word
    // formats a plain paste as if you had typed it. Written over the PRE-SPLIT offsets, so
    // the op runs before `splitParagraphMany` cuts the paragraph up.
    const pendingOps = deps.consumePendingFormatOps(
      target.paragraphId,
      target.offset,
      joined.length
    );
    if (joined.length > 0) {
      ops.push({
        op: 'insertText',
        paragraphId: target.paragraphId,
        offset: target.offset,
        text: joined,
      });
      ops.push(...pendingOps);
    }
    // Text pasted at a prompt's edge replaces the prompt, as typing does: it lands where the
    // prompt began, so the cuts and the caret count from there. Counted from the pressed
    // offset, every cut fell inside the next line.
    const start =
      joined.length > 0
        ? promptInsertionLanding(
            partOfNodeId(session, target.paragraphId) ?? session.part(),
            target.paragraphId,
            target.offset,
            0
          )
        : target.offset;
    const boundaries: number[] = [];
    let consumed = 0;
    for (let index = 0; index < lines.length - 1; index += 1) {
      consumed += lines[index]!.length;
      boundaries.push(start + consumed);
    }
    if (boundaries.length > 0) {
      ops.push({ op: 'splitParagraphMany', paragraphId: target.paragraphId, offsets: boundaries });
    }
    if (ops.length === 0) return;

    const before = new Set(deps.paragraphIds?.() ?? session.paragraphIdsIn(deps.storyScope()));
    const lastLine = lines[lines.length - 1]!;
    // A paste that stays in ONE paragraph knows exactly where it ends, so redo can put the
    // caret there. A multi-line paste mints its paragraphs inside the transaction, so the
    // landing id does not exist yet and the mark stays undefined — redo then falls back to
    // the clamp, which is where this lane started.
    const redoMark =
      boundaries.length === 0
        ? deps.caretMark({ paragraphId: target.paragraphId, offset: start + lastLine.length })
        : undefined;
    const withoutFormat = ops.filter((op) => !pendingOps.includes(op));
    deps.commit(
      () => deps.withoutPendingOnRejection(ops, withoutFormat, deps.selectionMark(), redoMark),
      () => {
        if (boundaries.length === 0) {
          return collapsedAt({ paragraphId: target.paragraphId, offset: start + lastLine.length });
        }
        // The caret lands at the end of the pasted text: in the LAST minted paragraph, right
        // after the final line. Scoped story ids are in document order, so the last unfamiliar
        // id is the tail that carries the final line and whatever followed the caret.
        const minted = (deps.paragraphIds?.() ?? session.paragraphIdsIn(deps.storyScope())).filter(
          (id) => !before.has(id)
        );
        const landing = minted[minted.length - 1];
        return landing ? collapsedAt({ paragraphId: landing, offset: lastLine.length }) : null;
      }
    );
  }

  /** Every clipboard flavour for the current selection — see clipboard-copy-payload.ts. */
  function copyFlavoursNow(): CopyFlavours {
    deps.flushPendingInputAndLayout();
    const rectangle = deps.cellSelection();
    const range = deps.orderedRange();
    const order = deps.paragraphOrder();
    const selected = rectangle
      ? paragraphsInCells(deps.layout(), rectangle.cellIds)
      : order.slice(order.indexOf(range.from.paragraphId), order.indexOf(range.to.paragraphId) + 1);
    const ranges = new Map(
      selected.map((id) => [
        id,
        {
          start: !rectangle && id === range.from.paragraphId ? range.from.offset : 0,
          end: !rectangle && id === range.to.paragraphId ? range.to.offset : Infinity,
        },
      ])
    );
    const copyScope = deps.storyScope();
    // Clipboard preflight reads the selected part without retaining a story-store slot.
    const copyPart =
      copyScope.kind === 'body'
        ? session.part()
        : partOfNodeId(session, selected[0] ?? range.from.paragraphId);
    if (copyPart && selectionContainsClipboardObject(copyPart, ranges)) {
      deps.commit(
        () =>
          ({
            committed: false,
            rejected: true,
            opCount: 0,
            reason: 'unsupported-content',
          }) as TreeApplyResult
      );
      return { text: '', html: null, reason: 'unsupported-content' };
    }
    // Plain text spells each computed field with its painted result, not its one model unit.
    const fields = fieldResultSpeller(deps.layout(), paragraphTextMarkingOptionalHyphens);
    if (rectangle) {
      return buildCopyFlavours({
        text: plainCopyText(cellSelectionText(deps.layout(), rectangle, fields.spelledText)),
        cellRectangle: true,
        coverage: null,
        pkg: null,
      });
    }
    const { from, to } = deps.orderedRange();
    const textPart = session.part();
    // Plain text spells each equation in its linear form instead of U+FFFC.
    const text = plainCopyText(
      selectedTextIn(
        deps.layout(),
        from,
        to,
        deps.paragraphOrder(),
        (id, offset) =>
          fields.resultAt(id, offset) ??
          (textPart ? equationPlainText(textPart, id, offset) : null),
        fields.markedText
      )
    );
    const scope = deps.storyScope();
    const collapsed = from.paragraphId === to.paragraphId && from.offset === to.offset;
    if (collapsed || scope.kind !== 'body') return { text, html: null };
    // A copy is a pure READ: `session.part()` is the body part, and the body-only guard
    // above keeps body extraction on the package reader. Non-body preflight uses its active story.
    const part = session.part();
    const coverage = fragmentCoverageOf(deps.layout(), part, from, to, deps.paragraphOrder());
    return buildCopyFlavours({
      text,
      cellRectangle: false,
      coverage,
      pkg: session.currentPackage(),
      // Word copies the accepted final view from a document that tracks changes
      acceptChanges: deps.editingMode() === 'suggest',
    });
  }

  /**
   * The attribution a rich paste writes, or null when the rich lane is shut: edit mode lands
   * the fragment plainly, and suggesting mode with an author lands it as one tracked insertion,
   * as Word pastes with Track Changes on. Suggesting without an author keeps to the plain lane,
   * whose refusal names the missing author.
   */
  function richPasteRevision(): { readonly revision?: RevisionAttributionInput } | null {
    const mode = deps.editingMode();
    if (mode === 'edit') return {};
    const author = mode === 'suggest' ? deps.trackedAuthor() : undefined;
    return author === undefined ? null : { revision: { author, date: deps.trackedDate() } };
  }

  /**
   * Land a fragment package at the selection, ONE commit: the selection-clearing ops plus
   * the resource merge plus `insertFragment`, promoted to a package undo unit in the
   * session. Unsupported content and a collaboration refusal stop routing; other refusals
   * permit the next flavour.
   */
  function pasteFragmentBytes(
    bytes: Uint8Array,
    lastMarkCovered: boolean
  ): boolean | 'unsupported-content' | 'refused' {
    const attribution = richPasteRevision();
    if (attribution === null) return false;
    if (deps.storyScope().kind !== 'body') return false;
    if (deps.cellSelection()) return false;
    deps.flushPendingInputAndLayout();
    const plan = deps.deleteSelectionPlan();
    const target = plan.replaceAt ?? plan.collapseTo;
    let landed = false;
    let unsupported = false;
    let refusedByName = false;
    let collaborationRefusal: string | null = null;
    deps.commit(
      () => {
        const actorId = deps.actorId();
        const result = session.applyFragmentPaste(
          { kind: 'body' },
          {
            paragraphId: target.paragraphId,
            offset: target.offset,
            fragmentBytes: bytes,
            lastMarkCovered,
            // Word's Smart cut and paste: a space on each side between two words
            smartSpacing: true,
            priorOps: plan.ops as unknown as TreeDocOp[],
            ...(actorId !== undefined ? { actorId } : {}),
            ...attribution,
            // The gate the typing lane asks, asked INSIDE the store transaction: only there
            // does the write exist whole: the selection-clearing ops, the exact
            // `insertFragment`, and the package with the styles, numbering, media and
            // relationships the merge brought in. Asked before the commit instead, a session
            // saw only the cleared selection and the paste landed past it.
            gate: (ops, pkg) =>
              (collaborationRefusal = deps.collaborationGate(
                ops,
                { kind: 'body' },
                { package: pkg }
              )),
          }
        );
        landed = result.ok;
        unsupported = !result.ok && result.detail === 'fragment-merge:unsupported-content';
        refusedByName = !result.ok && NAMED_PASTE_REFUSALS.has(result.reason);
        return {
          committed: result.ok,
          rejected: !result.ok,
          opCount: result.ok ? 1 : 0,
          ...(result.ok ? {} : { reason: collaborationRefusal || result.detail || result.reason }),
        } as unknown as TreeApplyResult;
      },
      () => collapsedAt(target)
    );
    if (landed && attribution.revision) deps.onTrackedChange();
    if (unsupported) return 'unsupported-content';
    // A session that refused the write refuses it in every flavour: a plain-text fallback
    // would put the same words past the same gate, or past a gate that never saw them. So
    // does a paste Word lands in a form the engine cannot write yet: plain text would
    // approximate it.
    return collaborationRefusal || refusedByName ? 'refused' : landed;
  }

  /** The paste router entry — fidelity order with continuous degrade to plain. */
  function pasteRichNow(text: string, html: string | null): boolean {
    const forcePlain =
      forcePlainPasteArmedAt !== null &&
      deps.now() - forcePlainPasteArmedAt < FORCE_PLAIN_PASTE_WINDOW_MS;
    forcePlainPasteArmedAt = null;
    const lane = routePaste(
      {
        richLaneOpen:
          (deps.richPasteAllowed?.() ?? true) &&
          richPasteRevision() !== null &&
          deps.storyScope().kind === 'body' &&
          deps.cellSelection() === null,
        pasteFragment: (fragmentBytes, lastMarkCovered) =>
          pasteFragmentBytes(fragmentBytes, lastMarkCovered),
        insertPlainText: (pasted) => insertPlainText(pasted, { smartSpacing: true }),
      },
      { html, text, forcePlain }
    );
    // The commit that refused already reported its reason.
    if (lane === 'refused') return false;
    if (lane === 'unsupported-content') {
      deps.commit(
        () =>
          ({
            committed: false,
            rejected: true,
            opCount: 0,
            reason: 'unsupported-content',
          }) as TreeApplyResult
      );
      return false;
    }
    return lane !== 'none';
  }

  return {
    insertPlainText,
    copyFlavoursNow,
    pasteRichNow,
    armForcePlainPaste: () => {
      forcePlainPasteArmedAt = deps.now();
    },
  };
}
