import type {
  CollaborationEditorPosition,
  CollaborationEditorSelection,
  CollaborationLocalSelection,
  CollaborationSelectionMove,
} from '../../index.ts';
import type { CollaborationDocumentPort } from '../../replication.ts';
import { mapOffsetAcrossText } from '../../offset-mapping.ts';
import { placeOfChange } from '../change/change-place.ts';
import { orderedKindSpecOf } from '../change/kinds/registry.ts';
import {
  isOrderedEditChange,
  type OrderedChange,
  type OrderedChangePlace,
} from '../change/types.ts';
import { notesStoryOf } from '../edits/notes.ts';
import { resolveOrderedPlace } from '../edits/places.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import { paragraphOfNode, placeAt, viewText, type ListParagraph } from './paragraphs.ts';
import { OrderedClientRefusal } from './refusal.ts';

/**
 * A place of the list and its story. With `after`, the point is right after the unit at `place`
 * (a caret after typed text keeps to that text); else right before it.
 */
export interface OrderedPlacedPoint {
  readonly story: string;
  readonly place: OrderedChangePlace;
  readonly after?: boolean;
}

/**
 * An editor point as a place of the list. `editorText` is the paragraph's text in the editor:
 * where it differs from the list's (the engine placed typed text elsewhere), the offset is
 * carried across the difference. Null when the list holds no such paragraph.
 */
export function placeOfPoint(
  doc: OrderedDocument,
  point: CollaborationEditorPosition,
  editorText: string | null,
  keep: 'before' | 'after' = 'before'
): OrderedPlacedPoint | null {
  try {
    const found = listPointOf(doc, point, editorText);
    if (!found) return null;
    const { paragraph, offset: clamped } = found;
    const story = paragraph.story.name;

    // Keep to the unit before the point: text another person types at the point goes after it
    if (keep === 'after' && clamped > 0)
      return { story, place: placeAt(paragraph, clamped - 1), after: true };
    return { story, place: placeAt(paragraph, clamped) };
  } catch (error) {
    if (error instanceof OrderedClientRefusal) return null;
    throw error;
  }
}

/** An editor point as the list's paragraph and an offset in its view, or null. */
function listPointOf(
  doc: OrderedDocument,
  point: CollaborationEditorPosition,
  editorText: string | null
): { readonly paragraph: ListParagraph; readonly offset: number } | null {
  try {
    const paragraph = paragraphOfNode(doc, point.nodeId);
    const text = viewText(paragraph.view);
    const offset =
      editorText === null || editorText === text
        ? point.offset
        : mapOffsetAcrossText(point.offset, editorText, text);
    return { paragraph, offset: Math.max(0, Math.min(offset, paragraph.view.chars.length)) };
  } catch (error) {
    if (error instanceof OrderedClientRefusal) return null;
    throw error;
  }
}

/** A range of the list in one story: from the unit at `from` up to the unit at `to`. */
export interface OrderedPlacedRange {
  readonly story: string;
  readonly from: OrderedChangePlace;
  readonly to: OrderedChangePlace;
}

/**
 * Two editor points as a range of the list, from the first in the story's order to the last;
 * null when they are one point, in two stories, or in a paragraph the list does not hold.
 */
export function rangeOfPoints(
  doc: OrderedDocument,
  a: CollaborationEditorPosition,
  b: CollaborationEditorPosition,
  texts: { readonly a: string | null; readonly b: string | null } = { a: null, b: null }
): OrderedPlacedRange | null {
  const first = listPointOf(doc, a, texts.a);
  const last = listPointOf(doc, b, texts.b);
  if (!first || !last || first.paragraph.story !== last.paragraph.story) return null;
  const order = first.paragraph.index - last.paragraph.index || first.offset - last.offset;
  if (order === 0) return null;
  const [from, to] = order < 0 ? [first, last] : [last, first];
  return {
    story: from.paragraph.story.name,
    from: placeAt(from.paragraph, from.offset),
    to: placeAt(to.paragraph, to.offset),
  };
}

/** A place of the list as an editor point now, or null when it does not resolve. */
export function pointOfPlace(
  doc: OrderedDocument,
  placed: OrderedPlacedPoint
): CollaborationEditorPosition | null {
  try {
    const address = placeOfChange(placed.place);
    const resolved = resolveOrderedPlace(doc, placed.story, address);
    const char = resolved.paragraph.chars[resolved.offset];

    // After a unit that still shows; a unit that went leaves the point where it was
    const shown =
      placed.after === true &&
      char !== undefined &&
      char.address.space === address.space &&
      char.address.offset === address.offset;
    return { nodeId: resolved.paragraph.id, offset: resolved.offset + (shown ? 1 : 0) };
  } catch {
    return null;
  }
}

/**
 * The end of the paragraph of the note a change inserted, read from the list after the change:
 * the note's nodes take ids from the change (`n:<change>.<k>`), and it is the last of its story.
 */
function newNoteCaret(
  doc: OrderedDocument,
  change: Extract<OrderedChange, { kind: 'insert-note' }>
): OrderedPlacedPoint | null {
  const notes = notesStoryOf(doc, change.note);
  if (!notes) return null;
  const minted = `n:${change.id}.`;
  let mark: OrderedItem | null = null;
  for (let item = notes.tree.last(); item; item = notes.tree.prev(item)) {
    if (item.kind === 'mark') mark = item;
    const node = item.node;
    if (item.kind === 'open' && item.token === 'note' && node?.id.startsWith(minted) === true)
      return mark ? { story: notes.name, place: { mark: { space: mark.space, offset: 0 } } } : null;
  }

  return null;
}

/** Where a change leaves the caret, as a place; null when it leaves the caret where it was. */
export function caretOfChange(
  change: OrderedChange,
  doc: OrderedDocument
): OrderedPlacedPoint | null {
  switch (change.kind) {
    case 'insert-note':
      // In the new note, as the editor alone and Word leave it
      return newNoteCaret(doc, change) ?? { story: change.part, place: change.place };
    case 'insert':
    case 'insert-tab':
    case 'insert-break':
    case 'split':
    case 'insert-table':
      return { story: change.part, place: change.place };
    case 'delete':
      return { story: change.part, place: change.from };
    case 'join':
      return { story: change.part, place: { mark: change.mark } };
    default: {
      // A kind a group added says where it leaves the caret in its spec
      if (!isOrderedEditChange(change)) return null;
      const place = orderedKindSpecOf(change.kind)?.caret?.(change) ?? null;
      return place === null ? null : { story: change.part, place };
    }
  }
}

/**
 * Where a batch of changes leaves the caret: after the text it typed (a paste's breaks fall
 * inside that text), else where its last change leaves it.
 */
export function caretOfBatch(
  changes: readonly OrderedChange[],
  doc: OrderedDocument
): OrderedPlacedPoint | null {
  let typed: OrderedPlacedPoint | null = null;
  let last: OrderedPlacedPoint | null = null;
  for (const change of changes) {
    const caret = caretOfChange(change, doc);
    if (caret === null) continue;
    last = caret;
    if (
      change.kind === 'insert' ||
      change.kind === 'insert-tab' ||
      change.kind === 'insert-break' ||
      orderedKindSpecOf(change.kind)?.types === true
    )
      typed = caret;
  }

  return typed ?? last;
}

/** A selection endpoint in node ids, with the paragraph's text in the editor when it is read. */
interface EditorPoint {
  readonly nodeId: string;
  readonly offset: number;
  /** The editor's text of the paragraph; null when the point is already in the list's terms. */
  readonly text: string | null;
}

/**
 * The local selection across publishes: `from` is what the editor holds, `anchor` and `head`
 * where it stands now. They differ while the editor has not yet taken a carried move: a second
 * publish before the editor's next layout carries the first one's result on.
 */
interface EditorSelection {
  readonly from: CollaborationEditorSelection;
  readonly anchor: EditorPoint;
  readonly head: EditorPoint;
}

/** The local selection and its endpoints as places of the list, held across a write. */
export interface HeldSelection {
  readonly before: EditorSelection;
  readonly places: { readonly anchor: OrderedPlacedPoint; readonly head: OrderedPlacedPoint };
}

/**
 * The local selection as the editor last published it, and the move a publish carried it by.
 * Across a publish, each endpoint is a place of the list: an address, so the caret stays beside
 * the same characters whatever moved around it.
 */
export class OrderedSelectionCarry {
  #last: CollaborationLocalSelection | null = null;
  #move: CollaborationSelectionMove | null = null;

  /** The selection the editor last published. */
  get last(): CollaborationLocalSelection | null {
    return this.#last;
  }

  /** Where the last publish carried the selection, until the editor publishes a new one. */
  get move(): CollaborationSelectionMove | null {
    return this.#move;
  }

  /** The editor published its selection: it took the carried move, or the person moved it. */
  track(selection: CollaborationLocalSelection | null): void {
    this.#last = selection;
    this.#move = null;
  }

  /**
   * The selection as places of the list, read before a publish and read back in the editor's node
   * ids after it. `caret` stands for an endpoint in a paragraph the list does not hold (one the
   * engine made under another id): the caret the edit itself leaves.
   */
  hold(
    doc: OrderedDocument,
    port: CollaborationDocumentPort,
    caret: OrderedPlacedPoint | null
  ): HeldSelection | null {
    const before = this.#current(port);
    if (!before) return null;

    // A caret keeps to the text before it; a range keeps what it covers and no more
    const { anchor: a, head: h } = before;
    const collapsed = a.nodeId === h.nodeId && a.offset === h.offset;
    const forward = a.nodeId !== h.nodeId || a.offset <= h.offset;
    const place = (point: EditorPoint, keep: 'before' | 'after') =>
      placeOfPoint(doc, point, point.text, keep) ?? caret;
    const anchor = place(a, collapsed || !forward ? 'after' : 'before');
    const head = place(h, collapsed || forward ? 'after' : 'before');
    return anchor && head ? { before, places: { anchor, head } } : null;
  }

  /**
   * The selection now as a range of the list, from its first point to its last; null for a
   * caret, a range over two stories, or a point in a paragraph the list does not hold.
   */
  range(doc: OrderedDocument, port: CollaborationDocumentPort): OrderedPlacedRange | null {
    const now = this.#current(port);
    return now
      ? rangeOfPoints(doc, now.anchor, now.head, { a: now.anchor.text, b: now.head.text })
      : null;
  }

  /** After a publish: the held places as the editor's points now. */
  carry(doc: OrderedDocument, held: HeldSelection | null): void {
    this.#move = null;
    if (!held) return;
    const anchor = pointOfPlace(doc, held.places.anchor);
    const head = pointOfPlace(doc, held.places.head);
    if (anchor && head) this.#move = { from: held.before.from, to: { anchor, head } };
  }

  /** The selection now: a move the editor has not taken yet, or the one it published. */
  #current(port: CollaborationDocumentPort): EditorSelection | null {
    const move = this.#move;
    if (move)
      return {
        from: move.from,
        anchor: { ...move.to.anchor, text: null },
        head: { ...move.to.head, text: null },
      };
    const selection = this.#last;
    if (!selection) return null;
    const point = (address: CollaborationLocalSelection['anchor']): EditorPoint | null => {
      const paragraph = port.paragraphByStableId(address.paragraphId);
      return paragraph
        ? { nodeId: paragraph.nodeId, offset: address.offset, text: paragraph.text }
        : null;
    };
    const anchor = point(selection.anchor);
    const head = point(selection.head);
    if (!anchor || !head) return null;
    const from = {
      anchor: { nodeId: anchor.nodeId, offset: anchor.offset },
      head: { nodeId: head.nodeId, offset: head.offset },
    };
    return { from, anchor, head };
  }
}
