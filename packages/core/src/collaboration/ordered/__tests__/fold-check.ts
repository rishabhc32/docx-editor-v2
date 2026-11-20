// What a fold keeps (design D14): a room writes its document as a checkpoint and the next opening
// reads it back under another file number. Everything the room and its pages read must come back
// the same. Ids and addresses are new after a fold, so suggestions compare by who made them and
// when, and items by what they show.
//
// Left out, as the same file and the same decisions follow from them:
//  - internal shapes: format records come back as `face` records with the file's record, paragraph
//    and section ops as `element` records, a control's prompt as the glossary's placeholder;
//  - a suggestion's `version`, which counts the writes since it was loaded;
//  - a suggestion its author took back (`retracted`): nothing shows it, and only its redo needs it,
//    as the undo history, which a fold does not keep;
//  - a suggested note reference that a direct deletion removed: the room keeps its record only for
//    that deletion's undo; the note shows nowhere and has no card;
//  - a record its author made on their own new words or paragraph (`untracked`): it shows with no
//    change record, so a fold writes it as part of them, and a decision keeps or takes it with
//    them; only the undo of the change that made it reads it.

import { expect } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { readOoxmlPackage, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart, type OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { isCardUnit, isOrderedItemHidden } from '../card-of.ts';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { openOrderedCheckpoint, type OrderedWordFile } from '../client/open.ts';
import { currentFace } from '../client/properties.ts';
import { orderedPackageMedia, writeOrderedCheckpoint } from '../compose/checkpoint.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { orderedReplacedSuggestions } from '../compose/replaced.ts';
import { orderedCards } from '../decide.ts';
import { MOVE_RANGE } from '../load/moves.ts';
import { setStatus } from '../model/mutations.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import { orderedUnitTyper } from '../model/typers.ts';
import type {
  OrderedDocument,
  OrderedItem,
  OrderedStory,
  OrderedSuggestion,
} from '../model/types.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import { normalizedBody, semanticBody } from './matrix/compare.ts';

/** The media bytes of a room's Word file, by SHA-256: the room's storage. */
export function mediaStore(base: OoxmlPackage): (sha256: string) => Uint8Array | undefined {
  const bytes = new Map(orderedPackageMedia(base).map((m) => [m.sha256, m.bytes]));
  return (sha256) => bytes.get(sha256);
}

/** The document folded: its checkpoint written, and opened as file `file`. */
export async function folded(
  doc: OrderedDocument,
  base: OoxmlPackage,
  file: number,
  blobs: (sha256: string) => Uint8Array | undefined
): Promise<OrderedWordFile> {
  const chunks: Uint8Array[] = [];
  await writeOrderedCheckpoint(doc, base, {
    media: async (sha256) => {
      const bytes = blobs(sha256);
      if (bytes === undefined) {
        throw new Error(`no media ${sha256}`);
      }

      return bytes;
    },
    write: async (chunk) => {
      chunks.push(chunk);
    },
  });
  return openOrderedCheckpoint(new Uint8Array(Buffer.concat(chunks)), { file });
}

/** A suggestion by who made it and when: a fold keeps these, not its id. */
const makerKey = ({ attribution }: OrderedSuggestion): string =>
  JSON.stringify([attribution.person, attribution.author, attribution.date ?? null]);

/** An element as XML: none and an empty one (a face with no properties) read alike. */
const xmlOf = (face: OoxmlElement | null): string =>
  face === null || face.children.length === 0
    ? ''
    : serializeOoxmlPart({ id: 'x', name: '/x', contentType: 'application/xml', root: face });

/**
 * Whether an item shows: a prompt's units that content took the place of are the glossary's
 * placeholder after a fold, and show nowhere either way; a link's edges are no unit, as the room
 * labels them and a reopened link has no label.
 */
const showsOn = (doc: OrderedDocument, item: OrderedItem): boolean =>
  item.kind === 'zero'
    ? !isOrderedItemHidden(doc, item) && !isMoveRange(item)
    : isCardUnit(doc, item);

/** A move's range marker: a fold writes the file's own, which the room draws from the move. */
const isMoveRange = (item: OrderedItem): boolean =>
  item.node?.kind !== 'textValue' && MOVE_RANGE.test(item.node?.localName ?? '');

/**
 * Pending suggestions that show something, by the room's read: one taken back by its author, or
 * one only on a prompt that content took the place of, shows nothing.
 */
function shownPending(doc: OrderedDocument): OrderedSuggestion[] {
  return [...doc.suggestions.values()].filter(
    (s) =>
      s.status === 'pending' &&
      (doc.partRevisions.has(s.id) ||
        [...(doc.labelled.get(s.id) ?? [])].some((item) => !isOrderedItemHidden(doc, item)))
  );
}

/** Everything a fold must keep, as plain data with no ids or addresses. */
export function foldPrint(doc: OrderedDocument, base: OoxmlPackage, blobs: FoldBlobs): unknown {
  const key = (id: string | null): string | null => {
    const suggestion = id === null ? undefined : doc.suggestions.get(id);
    return suggestion?.status === 'pending' ? makerKey(suggestion) : null;
  };
  const keys = (ids: readonly string[]) => ids.map(key).filter((k) => k !== null);

  // Each pending suggestion's own units, so two suggestions of one maker cannot trade them
  const owned = new Map<string, string[]>();
  const own = (roles: Record<string, readonly (string | null)[]>, unit: string) => {
    for (const [role, ids] of Object.entries(roles)) {
      for (const id of ids) {
        if (id === null || key(id) === null) {
          continue;
        }

        const list = owned.get(id) ?? [];
        list.push(`${role}:${unit}`);
        owned.set(id, list);
      }
    }
  };

  // Each unit of an item on its own: a load joins units of one face into one item
  const units = (story: OrderedStory, item: OrderedItem): string[] => {
    const node = item.node?.kind === 'textValue' ? '' : (item.node?.localName ?? '');

    // A unit in a new row or note keeps its label in the room, where a load reads it from the
    // container
    const by = item.quiet ? null : item.by;
    const recorded = (records: readonly { suggestion: string; untracked?: boolean }[]) =>
      records
        .filter((r) => r.untracked !== true || r.suggestion !== item.by)
        .map((r) => r.suggestion);
    const roles = {
      by: [by],
      del: item.dels,
      fmt: recorded(item.fmts),
      prop: recorded(item.props),
      within: [...item.within.by, ...item.within.dels],
    };
    const labels = {
      kind: item.kind,
      node,
      by: key(by),
      dels: keys(item.dels),
      fmts: keys(roles.fmt),
      replaced: keys(item.fmts.flatMap((f) => (f.replaced === undefined ? [] : [f.replaced]))),
      // Records apply in their order of creation, whatever their place in the list
      props: keys(roles.prop).sort(),
      move: item.move !== null && (key(item.by) !== null || keys(item.dels).length > 0),
      within: [keys(item.within.by), keys(item.within.dels)],
    };
    return Array.from({ length: Math.max(item.text.length, 1) }, (_, k) => {
      own(roles, `${storyLabel(story.name)}:${node}:${item.text[k] ?? ''}`);
      return JSON.stringify({
        ...labels,
        text: item.text[k] ?? '',
        face: xmlOf(currentFace(doc, story, item)),
        typer: orderedUnitTyper(doc, item, item.offset + k),
      });
    });
  };
  const pending = shownPending(doc);
  return {
    stories: [...doc.stories]
      .map(([name, story]) => {
        const shown = story.tree.toArray().filter((item) => showsOn(doc, item));
        const listed = shown.flatMap((item) => units(story, item));
        return [storyLabel(name), ...listed].join('\n');
      })
      .sort(),
    suggestions: pending.map(makerKey).sort(),
    owned: [...owned].map(([id, list]) => [key(id), ...list.sort()].join('|')).sort(),
    cards: orderedCards(doc)
      .map((card) => keys(card.suggestions).join('|'))
      .sort(),
    seen: pending
      .map((s) =>
        (orderedCardSeen(doc, s.id) ?? [])
          .map((seen) => key(seen.suggestion))
          .sort()
          .join('|')
      )
      .sort(),
    replaced: pending
      .flatMap((s) =>
        orderedReplacedSuggestions(doc, s.id).map((r) =>
          JSON.stringify([makerKey(s), key(r.suggestion), r.formattingKind, r.formattingChanges])
        )
      )
      .sort(),
    comments: [...doc.comments.values()].map(({ element, ...comment }) =>
      JSON.stringify({ ...comment, element: xmlOf(element) })
    ),
    partRevisions: [...doc.partRevisions]
      .filter(([id]) => key(id) !== null)
      .map(([id, r]) => JSON.stringify([key(id), r.part, r.element, r.styleId ?? null])),
    view: viewOf(doc),
    file: partsOf(orderedPackageBytes(doc, base, { blobs })),
  };
}

type FoldBlobs = (sha256: string) => Uint8Array | undefined;

/** A story by its name: a text box's names its object, whose id a fold mints again. */
const storyLabel = (name: string): string => (isTextBoxStoryName(name) ? 'text box' : name);

/**
 * What each story's page shows: every character and its strike. A character's own face is left
 * out: a fold writes the decided formats into it, which the room keeps as records.
 */
function viewOf(doc: OrderedDocument): string[] {
  const out: string[] = [];
  for (const [name, story] of doc.stories) {
    for (let k = 0; k < story.tree.summary.marks; k += 1) {
      const view = orderedViewParagraph(doc, name, k);
      if (view !== null) {
        out.push(view.chars.map((c) => `${c.ch}${c.deleted ? '-' : ''}`).join(''));
      }
    }
  }

  return out.sort();
}

/** A Word file's XML parts by name, and the size of each other part. */
function partsOf(bytes: Uint8Array): Record<string, string> {
  return Object.fromEntries(
    Object.entries(unzipSync(bytes)).map(([name, data]) => [
      name,
      /\.(xml|rels)$/.test(name) ? strFromU8(data) : `${data.length} bytes`,
    ])
  );
}

/**
 * The Word file after `ids` are decided as `status`, each XML part by what it says (as a reopen
 * compares decisions: `semanticBody`); the document is restored. A decision on a folded copy may
 * split a run where the room keeps one, or keep an empty `w:rPr`: the same meaning.
 */
function decidedFile(
  doc: OrderedDocument,
  base: OoxmlPackage,
  blobs: FoldBlobs,
  ids: readonly string[],
  status: 'accepted' | 'rejected'
): Record<string, string> {
  const { record } = withUndoRecord(doc, () => {
    for (const id of ids) {
      setStatus(doc, id, status);
    }
  });
  const read = readOoxmlPackage(orderedPackageBytes(doc, base, { blobs }), {});
  takeBack(doc, record);
  if (!read.ok) {
    throw new Error(`the decided file does not read: ${read.reason}`);
  }

  return Object.fromEntries(
    [...read.package.parts].map(([name, part]) => [
      name,
      semanticBody(normalizedBody(part, (id) => id)),
    ])
  );
}

/** A link's relationship in a relationships part. */
const LINK_RELATIONSHIP = /<Relationship [^>]*Type="[^"]*\/hyperlink"[^>]*\/>/g;

/**
 * A folded copy's decided file without the link relationships that the room's own decided file
 * does not hold. A known difference: a reject of a suggested link in a folded copy keeps the
 * link's relationship, which nothing names.
 */
function withoutFoldedLinks(
  file: Record<string, string>,
  live: Record<string, string>
): Record<string, string> {
  return Object.fromEntries(
    Object.entries(file).map(([name, xml]) => [
      name,
      xml.replace(LINK_RELATIONSHIP, (link) => (live[name]?.includes(link) ? link : '')),
    ])
  );
}

/** The ids of each maker's pending suggestions that show something. */
function byMaker(doc: OrderedDocument): Map<string, string[]> {
  const out = new Map<string, string[]>();
  for (const s of shownPending(doc)) {
    out.set(makerKey(s), [...(out.get(makerKey(s)) ?? []), s.id]);
  }

  return out;
}

const DECISIONS = ['accepted', 'rejected'] as const;

/**
 * Fold `doc`, and fold the folded copy again: check that each copy gives back everything
 * `foldPrint` reads, and the same Word file after an accept and after a reject of each maker's
 * pending suggestions. `then`, when given, is an edit made on every copy after the folds.
 */
export async function checkFold(
  doc: OrderedDocument,
  base: OoxmlPackage,
  blobs: FoldBlobs,
  label: string,
  then?: (doc: OrderedDocument) => void
): Promise<void> {
  const once = await folded(doc, base, doc.file + 1, blobs);
  const twice = await folded(once.document, once.base, doc.file + 2, blobs);
  then?.(doc);
  const print = foldPrint(doc, base, blobs);
  const live = [...byMaker(doc)].flatMap(([maker, ids]) =>
    DECISIONS.map((status) => ({ maker, status, file: decidedFile(doc, base, blobs, ids, status) }))
  );
  for (const [fold, copy] of [
    ['once', once],
    ['twice', twice],
  ] as const) {
    then?.(copy.document);
    expect({ label, fold, print: foldPrint(copy.document, copy.base, blobs) }).toEqual({
      label,
      fold,
      print,
    });
    const reopened = byMaker(copy.document);
    for (const { maker, status, file } of live) {
      const ids = reopened.get(maker) ?? [];
      const decided = decidedFile(copy.document, copy.base, blobs, ids, status);
      expect({ label, fold, maker, status, file: withoutFoldedLinks(decided, file) }).toEqual({
        label,
        fold,
        maker,
        status,
        file,
      });
    }
  }
}
