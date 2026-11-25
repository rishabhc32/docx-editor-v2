// Notes converted to the other kind (group 9.9): Word's Convert, for one note or every note of a
// kind, as the engine's `applyConvertNote` converts them.

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { isNotesStory } from '../compose/note-references.ts';
import { createOrderedItem, isHidden } from '../model/document.ts';
import { insertItems } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedItem, OrderedStory } from '../model/types.ts';
import { attributeOf, isWml, renamed, withAttributes, withChildren } from '../ooxml-nodes.ts';
import { closeOf, itemsInside } from '../placement.ts';
import { refuseLockedAt } from '../locked-content.ts';
import { holdWithin, suggestionFor, withNewIds, type Ctx } from './context.ts';
import { settle } from './direct.ts';
import { copiedFmts, copiedProps, newMarkShell } from './marks.ts';
import { addNotesStory, noteIdsFor, notesStoryOf, type OrderedNoteKind } from './notes.ts';
import { newParagraphIdentity } from './para-ids.ts';

/** Word's built-in character style of each kind's marks (`w:footnoteRef`, a reference). */
const REFERENCE_STYLE: Readonly<Record<OrderedNoteKind, string>> = {
  footnote: 'FootnoteReference',
  endnote: 'EndnoteReference',
};

/** One conversion: the kinds, and the note id the copy takes. */
interface Conversion {
  readonly from: OrderedNoteKind;
  readonly to: OrderedNoteKind;
  readonly id: number;
}

const otherKind = (kind: OrderedNoteKind): OrderedNoteKind =>
  kind === 'footnote' ? 'endnote' : 'footnote';

/** The kind of notes a story holds, or null for any other story. */
function kindOfStory(story: OrderedStory): OrderedNoteKind | null {
  if (!isNotesStory(story)) return null;
  return story.part.root.localName === 'footnotes' ? 'footnote' : 'endnote';
}

/** An element's node, as an element, or null. */
const elementOf = (item: OrderedItem): OoxmlElement | null =>
  item.node && item.node.kind !== 'textValue' ? item.node : null;

/**
 * A face with `from`'s built-in reference style as `to`'s; any other style stays, as the engine
 * restyles a converted note's reference run.
 */
function restyled(face: OoxmlElement | null, c: Conversion): OoxmlElement | null {
  if (!face) return null;
  let changed = false;
  const children = face.children.map((child) => {
    if (child.kind === 'textValue' || !isWml(child, 'rStyle')) return child;
    if (attributeOf(child, 'val') !== REFERENCE_STYLE[c.from]) return child;
    changed = true;
    return withAttributes(
      child,
      child.attributes.map((a) =>
        a.localName === 'val' ? { ...a, value: REFERENCE_STYLE[c.to] } : a
      )
    );
  });
  return changed ? withChildren(face, children) : face;
}

/** An element with its `w:id` set (added when it has none). */
function withNoteId(element: OoxmlElement, id: number): OoxmlElement {
  const has = element.attributes.some((a) => a.localName === 'id');
  const attributes = has
    ? element.attributes.map((a) => (a.localName === 'id' ? { ...a, value: String(id) } : a))
    : [
        ...element.attributes,
        {
          kind: 'genericExtension' as const,
          namespaceUri: element.namespaceUri,
          localName: 'id',
          prefix: element.prefix,
          value: String(id),
        },
      ];
  return withAttributes(element, attributes);
}

/** Whether an item is a reference to a note of a kind (any id). */
const isReferenceOf = (item: OrderedItem, kind: OrderedNoteKind): boolean =>
  item.kind === 'atom' && elementOf(item)?.localName === `${kind}Reference`;

/**
 * Every reference to the notes of a kind outside the notes stories, shown or not, by note id:
 * a reference that a revert shows again must be converted too.
 */
function referencesOf(ctx: Ctx, kind: OrderedNoteKind): Map<string, [OrderedStory, OrderedItem][]> {
  const out = new Map<string, [OrderedStory, OrderedItem][]>();
  for (const story of ctx.doc.stories.values()) {
    if (isNotesStory(story)) continue;
    for (const item of story.tree.items()) {
      const node = elementOf(item);
      const id = node ? attributeOf(node, 'id') : undefined;
      if (!isReferenceOf(item, kind) || id === undefined) continue;
      const list = out.get(id) ?? [];
      list.push([story, item]);
      out.set(id, list);
    }
  }

  return out;
}

/** The labels of an item for its copy: its own, and held by the conversion. */
function copiedLabels(ctx: Ctx, item: OrderedItem, suggestion: string) {
  return {
    by: item.by,
    dels: item.dels,
    fmts: copiedFmts(ctx, item.fmts),
    props: copiedProps(ctx, item.props),
    move: item.move,
    retracted: item.retracted,
    taken: item.taken ?? null,
    gone: item.gone,
    quiet: item.quiet,
    within: { by: [...item.within.by, suggestion], dels: item.within.dels },
  };
}

/**
 * Copies of a note's items for the other kind's story: every space gets a new one of its kind
 * (a paragraph mark's is its new paragraph's identity), offsets and labels stay, and every node
 * is a copy under the change's ids.
 */
function copiedNote(
  ctx: Ctx,
  target: Ctx,
  items: readonly OrderedItem[],
  c: Conversion,
  suggestion: string
): OrderedItem[] {
  const copy: NoteCopy = { ctx, target, c, suggestion, ...newSpaces(ctx, target, items) };
  return items.map((item) => copiedItem(copy, item));
}

/** A new paragraph's identity: its mark's space and its node id. */
type ParagraphIdentity = ReturnType<typeof newParagraphIdentity>;

/** What every item of one note's copy shares: its new spaces and paragraph identities. */
interface NoteCopy {
  readonly ctx: Ctx;
  readonly target: Ctx;
  readonly c: Conversion;
  readonly suggestion: string;
  readonly spaces: ReadonlyMap<string, string>;
  readonly identities: ReadonlyMap<OrderedItem, ParagraphIdentity>;
}

/** Every item's new space, minted first and in order: a field's marker may name a later field. */
function newSpaces(
  ctx: Ctx,
  target: Ctx,
  items: readonly OrderedItem[]
): {
  spaces: Map<string, string>;
  identities: Map<OrderedItem, ParagraphIdentity>;
} {
  const spaces = new Map<string, string>();
  const identities = new Map<OrderedItem, ParagraphIdentity>();
  for (const item of items) {
    if (item.kind === 'mark') {
      const identity = newParagraphIdentity(target);
      identities.set(item, identity);
      spaces.set(item.space, identity.space);
    } else if (!spaces.has(item.space)) {
      spaces.set(item.space, newSpaceFor(ctx, item.kind));
    }
  }

  return { spaces, identities };
}

/** A new space of an item's kind (a mark's comes with its paragraph's identity). */
function newSpaceFor(ctx: Ctx, kind: OrderedItem['kind']): string {
  switch (kind) {
    case 'open':
    case 'close':
      return ctx.ids.container();
    case 'zero':
      return ctx.ids.marker();
    default:
      return ctx.ids.units();
  }
}

/** The item's node under new ids: a note's and its own mark's names changed. */
function copiedNode(
  copy: NoteCopy,
  item: OrderedItem,
  node: OoxmlElement | null
): OoxmlElement | null {
  const { ctx, c } = copy;
  if (!node) {
    return null;
  }

  const identity = copy.identities.get(item);
  if (identity) {
    return newMarkShell(copy.target, node, identity);
  }

  const fresh = withNewIds(ctx, node);
  if (item.kind === 'open' && item.token === 'note') {
    return renamed(withNoteId(fresh, c.id), 'note', c.to);
  }

  if (item.kind === 'atom' && fresh.localName === `${c.from}Ref`) {
    return renamed(fresh, 'noteRef', `${c.to}Ref`);
  }

  return fresh;
}

/** One item's copy: new space and nodes, its offsets and labels kept. */
function copiedItem(copy: NoteCopy, item: OrderedItem): OrderedItem {
  const { ctx, c } = copy;
  const spaceOf = (space: string): string => copy.spaces.get(space) ?? space;
  const node = elementOf(item);
  if (item.kind === 'mark' && !node)
    throw new OrderedRefusal('malformed-story', 'a note paragraph without its mark');
  const slot = item.fieldSlot;
  const [fieldSpace, fieldOffset] = slot ? slot.field.split('|') : [];
  const face = item.face ? withNewIds(ctx, item.face) : null;
  return createOrderedItem({
    space: spaceOf(item.space),
    offset: item.offset,
    length: item.length,
    kind: item.kind,
    text: item.text,
    node: copiedNode(copy, item, node),
    face: item.kind === 'atom' ? restyled(face, c) : face,
    run: item.run,
    ...(item.textShell ? { textShell: withNewIds(ctx, item.textShell) } : {}),
    ...(item.pieces
      ? {
          pieces: item.pieces.map((piece) => ({
            ...piece,
            node: piece.node.kind === 'textValue' ? piece.node : withNewIds(ctx, piece.node),
          })),
        }
      : {}),
    ...(item.inRun === undefined ? {} : { inRun: item.inRun }),
    ...(item.bare === undefined ? {} : { bare: item.bare }),
    ...(item.block === undefined ? {} : { block: item.block }),
    ...(item.instr === undefined ? {} : { instr: item.instr }),
    ...(item.prompt === undefined ? {} : { prompt: item.prompt }),
    ...(item.token === undefined ? {} : { token: item.token }),
    ...(item.pair === undefined ? {} : { pair: spaceOf(item.pair) }),
    ...(item.moveRange === undefined ? {} : { moveRange: item.moveRange }),
    ...(slot && fieldSpace !== undefined
      ? { fieldSlot: { ...slot, field: `${spaceOf(fieldSpace)}|${fieldOffset}` } }
      : {}),
    ...copiedLabels(ctx, item, copy.suggestion),
  });
}

/** A reference's copy, right after it: the other kind's reference to the copy. */
function copiedReference(ctx: Ctx, item: OrderedItem, c: Conversion, suggestion: string) {
  const node = elementOf(item);
  if (!node) throw new OrderedRefusal('malformed-story', 'a note reference without its node');
  return createOrderedItem({
    space: ctx.ids.units(),
    kind: 'atom',
    text: item.text,
    node: renamed(withNoteId(withNewIds(ctx, node), c.id), 'noteReference', `${c.to}Reference`),
    face: restyled(item.face ? withNewIds(ctx, item.face) : null, c),
    run: item.run,
    ...copiedLabels(ctx, item, suggestion),
  });
}

/**
 * Convert the notes whose open items `opens` names, in order, from the kind of `ctx.story` to
 * the other kind, as one direct edit (Word converts a note directly, never as a suggestion).
 * Returns the edit's suggestion.
 *
 * Nothing moves (design D1), so a converted note is a copy at the end of the other kind's story,
 * and every reference gets a copy right after it. The suggestion holds the copies through
 * `within.by` and the originals through `within.dels`, so `revert` gives the originals back. A
 * suggestion inside the note labels the copy too: a decision on it gives one result.
 */
export function convertNotes(ctx: Ctx, opens: readonly OrderedItem[]): string {
  if (!ctx.direct)
    throw new OrderedRefusal(
      'invalid-edit',
      'Word converts a note directly, never as a suggestion'
    );
  const from = kindOfStory(ctx.story);
  if (from === null) throw new OrderedRefusal('invalid-edit', 'not a notes story');
  if (new Set(opens).size !== opens.length)
    throw new OrderedRefusal('invalid-edit', 'a note named twice');
  const to = otherKind(from);
  const references = referencesOf(ctx, from);
  const suggestion = suggestionFor(ctx).id;
  const target = notesStoryOf(ctx.doc, to) ?? addNotesStory(ctx, to, suggestion);
  const inTarget: Ctx = { ...ctx, story: target };
  const nextId = noteIdsFor(target);
  for (const open of opens) {
    const node = elementOf(open);
    if (!node || open.kind !== 'open' || open.token !== 'note')
      throw new OrderedRefusal('invalid-edit', 'not a note');
    const cites = references.get(attributeOf(node, 'id') ?? '') ?? [];

    // A note shows while one of its references shows (or it has none, as the file had it)
    if (
      isHidden(ctx.doc, open) ||
      (cites.length > 0 && cites.every(([, r]) => isHidden(ctx.doc, r)))
    )
      throw new OrderedRefusal('hidden-target', 'the note does not show');

    // A reference in a locked control's content keeps its kind
    for (const [story, reference] of cites) refuseLockedAt({ ...ctx, story }, reference);
    const c: Conversion = { from, to, id: nextId() };
    const items = [open, ...itemsInside(ctx.story, open), closeOf(ctx.story, open)];
    insertItems(
      ctx.doc,
      target,
      target.tree.last(),
      copiedNote(ctx, inTarget, items, c, suggestion)
    );
    for (const item of items) {
      holdWithin(ctx, ctx.story, item, suggestion);
    }

    for (const [story, reference] of cites) {
      insertItems(ctx.doc, story, reference, [copiedReference(ctx, reference, c, suggestion)]);
      holdWithin(ctx, story, reference, suggestion);
    }
  }

  settle(ctx, suggestion);
  return suggestion;
}
