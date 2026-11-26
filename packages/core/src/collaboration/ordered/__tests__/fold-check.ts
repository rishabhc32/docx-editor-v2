// What a fold keeps (design D14): a room writes its document as a checkpoint and the next opening
// reads it back under another file number. Everything the room and its pages read must come back
// the same. A suggestion keeps its id, and the print names each by it; addresses are new after a
// fold, so items compare by what they show.
// `checkFold` and `checkEveryCard` also compare the room cleaned as a close cleans it
// (`withCleanup`), live, with the same checks.
//
// Left out, as the same file and the same decisions follow from them:
//  - internal shapes: format records come back as `face` records with the file's record, paragraph
//    and section ops as `element` records, a control's prompt as the glossary's placeholder;
//  - a suggestion's `version`, which counts the writes since it was loaded;
//  - the record of a suggestion its author took back, with the units it inserted (`retracted`):
//    only its redo reads them, as the undo history, which a fold does not keep. What the take-back
//    leaves is compared: the file, the cards, the shown units and their labels;
//  - the record of a suggestion that shows nowhere now: a direct or an accepted deletion took every
//    unit it is on, or the note or text box it is in shows nowhere (`CardRead.isLeftOut`). The room
//    keeps the record only for that change's undo; it has no card, and the file holds none of it;
//  - a record its author made on their own new words or paragraph (`untracked`): it shows with no
//    change record, so a fold writes it as part of them, and a decision keeps or takes it with
//    them; only the undo of the change that made it reads it. So does the type a suggested section
//    break gives the next section (`keptOnReject`): Word writes it untracked, a fold writes it as
//    the section's own, and every decision keeps it.
//  - how many format records of one suggestion are on a unit: a format that took over another's
//    keeps the record it took under its own suggestion, and once the one it took from is decided,
//    a fold reads the one change record of the run as one record. The face shows what they give.

import { expect } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { readOoxmlPackage, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { canonicalOoxmlFingerprint } from '../../../store/package/ooxml-serialize.ts';
import { serializeOoxmlPart, type OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { CardRead, isCardUnit, isOrderedItemHidden } from '../card-of.ts';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { compactOrdered } from '../change/compact.ts';
import { orderedChangeId, orderedIdMinter } from '../change/ids.ts';
import { settleDecidedTables } from '../change/settle-columns.ts';
import { settleEmptiedControls } from '../change/settle-emptied.ts';
import { settleNamedPrompts } from '../change/settle-named-prompts.ts';
import { settlePromptMarks } from '../change/settle-prompts.ts';
import type { OrderedSeenSuggestion } from '../change/types.ts';
import { openOrderedCheckpoint, type OrderedWordFile } from '../client/open.ts';
import { orderedSuggestionsOfRevision } from '../client/ops/review.ts';
import { currentFace } from '../client/properties.ts';
import { isBlockToken } from '../compose/blocks.ts';
import { orderedPackageMedia, writeOrderedCheckpoint } from '../compose/checkpoint.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import { ComposeContext } from '../compose/context.ts';
import { showsPrompt } from '../compose/paragraph.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { orderedReplacedSuggestions } from '../compose/replaced.ts';
import { isSectionItem } from '../compose/sections.ts';
import { orderedCards } from '../decide.ts';
import { MOVE_RANGE } from '../load/moves.ts';
import { isHidden, movedAwayBy } from '../model/document.ts';
import { setStatus } from '../model/mutations.ts';
import { snapshotDifference } from '../model/snapshot-compare.ts';
import { takenFrom } from '../model/takeovers.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import { orderedUnitTyper } from '../model/typers.ts';
import type {
  OrderedDocument,
  OrderedItem,
  OrderedStory,
  OrderedSuggestion,
} from '../model/types.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import { attributeOf, eachElement } from '../ooxml-nodes.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { normalizedBody, semanticBody } from './matrix/compare.ts';
import { testPerson } from './test-persons.ts';

/** The media bytes of a room's Word file, by SHA-256: the room's storage. */
export function mediaStore(base: OoxmlPackage): (sha256: string) => Uint8Array | undefined {
  const bytes = new Map(orderedPackageMedia(base).map((m) => [m.sha256, m.bytes]));
  return (sha256) => bytes.get(sha256);
}

/**
 * The document folded: its checkpoint written, and opened as file `file`. The opened copy holds
 * every stored field as `doc` does, and saves to the same bytes.
 */
export async function folded(
  doc: OrderedDocument,
  base: OoxmlPackage,
  file: number,
  blobs: FoldBlobs
): Promise<OrderedWordFile> {
  const bytes = await checkpointBytes(doc, base, blobs);
  const opened = openOrderedCheckpoint(bytes, { file });
  expect({ file, difference: snapshotDifference(doc, opened.document) }).toEqual({
    file,
    difference: null,
  });

  // Before any compose: each one moves the counter of the node ids it mints
  expect(await checkpointBytes(opened.document, opened.base, blobs)).toEqual(bytes);
  return opened;
}

/** The bytes of the document's checkpoint. */
export async function checkpointBytes(
  doc: OrderedDocument,
  base: OoxmlPackage,
  blobs: FoldBlobs
): Promise<Uint8Array> {
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
  return new Uint8Array(Buffer.concat(chunks));
}

/** The entries a checkpoint holds, with those of the Word package version 6 holds inside it. */
export function heldEntries(checkpoint: Uint8Array): Record<string, Uint8Array> {
  const { 'ordered/base.docx': base, ...entries } = unzipSync(checkpoint);
  return base === undefined ? entries : { ...entries, ...unzipSync(base) };
}

/**
 * A Word file by what it says: each XML part by the engine's semantic oracle (prefixes and where
 * namespaces are declared are the writer's choice), each other part by its bytes.
 */
export function wordMeaning(bytes: Uint8Array): Record<string, string> {
  const read = readOoxmlPackage(bytes, {});
  if (!read.ok) {
    throw new Error(read.reason);
  }

  const { parts, partBytes } = read.package;
  return Object.fromEntries(
    [...partBytes].map(([name, data]) => {
      const part = parts.get(name);
      return [name, part === undefined ? orderedBlobHash(data) : canonicalOoxmlFingerprint(part)];
    })
  );
}

/** Who made a suggestion, and when. */
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
 * labels them and a reopened link has no label; a note or a text box compose leaves out shows
 * nowhere.
 */
const showsOn = (read: CardRead, story: OrderedStory, item: OrderedItem): boolean =>
  !read.isLeftOut(item, story) &&
  (item.kind === 'zero'
    ? !isOrderedItemHidden(read.doc, item) && !isMoveRange(item)
    : isCardUnit(read.doc, item));

/**
 * An item's element by name. The body's own section reads alike with its properties or none: a
 * load gives a file that states none an item with none, and a fold loads the section the room drew.
 */
function nodeName(item: OrderedItem): string {
  if (isSectionItem(item)) {
    return 'sectPr';
  }

  return item.node?.kind === 'textValue' ? '' : (item.node?.localName ?? '');
}

/** A move's range marker: a fold writes the file's own, which the room draws from the move. */
const isMoveRange = (item: OrderedItem): boolean =>
  item.node?.kind !== 'textValue' && MOVE_RANGE.test(item.node?.localName ?? '');

/**
 * The items a story's drawing leaves out: the prompts content took the place of, with the controls
 * those prompts held (a fold keeps them in the glossary only), and the edges of a table whose every
 * row is gone (compose leaves it out, as Word removes it).
 */
function drawnNowhere(doc: OrderedDocument, story: OrderedStory): ReadonlySet<OrderedItem> {
  const ctx = new ComposeContext(doc, story, {});
  const out = new Set<OrderedItem>();
  const tables: { open: OrderedItem; rows: boolean }[] = [];
  for (const item of story.tree.items()) {
    if (item.kind === 'open' && item.prompt === true) {
      showsPrompt(ctx, item);
    }

    if (item.token === 'tbl' && item.kind === 'open') {
      tables.push({ open: item, rows: false });
    } else if (item.token === 'tbl' && item.kind === 'close') {
      const table = tables.pop();
      if (table && !table.rows) {
        out.add(table.open).add(item);
      }
    } else if (item.token === 'tr' && item.kind === 'open' && !isOrderedItemHidden(doc, item)) {
      const table = tables.at(-1);
      if (table) {
        table.rows = true;
      }
    }
  }

  for (const item of ctx.hiddenPrompts) {
    out.add(item);
  }

  return out;
}

/**
 * The units that show in no paragraph: words a mark that shows does not end before a block starts
 * or ends, or before the story ends. Live and after a fold, every unit is in a paragraph. A note
 * or a text box that compose leaves out (`CardRead.isLeftOut`) shows nowhere.
 */
function unitsOutsideParagraphs(doc: OrderedDocument): string[] {
  const out: string[] = [];
  const read = new CardRead(doc);
  for (const story of doc.stories.values()) {
    const nowhere = drawnNowhere(doc, story);
    let loose: OrderedItem[] = [];
    const flush = () => {
      out.push(...loose.map((item) => `${story.name}: ${item.kind} ${JSON.stringify(item.text)}`));
      loose = [];
    };
    for (const item of story.tree.items()) {
      if (isBlockToken(item)) {
        flush();
      } else if (item.kind === 'mark') {
        if (!isHidden(doc, item)) {
          loose = [];
        }
      } else if (
        !(item.kind === 'zero' && item.block) &&
        !isOrderedItemHidden(doc, item) &&
        !nowhere.has(item) &&
        !read.isLeftOut(item, story)
      ) {
        loose.push(item);
      }
    }

    flush();
  }

  return out;
}

/**
 * Pending suggestions that show something: a fold keeps none whose every item shows nowhere
 * (`isOrderedItemHidden`: taken back, taken away by a decision or a direct deletion, a prompt that
 * content took the place of; or in a note or a text box compose leaves out). A move's range
 * markers alone show nothing: compose draws no range around no words.
 */
function keptPending(read: CardRead): OrderedSuggestion[] {
  const { doc } = read;
  const kept = (item: OrderedItem) =>
    item.moveRange !== true && !isOrderedItemHidden(doc, item) && !read.isLeftOut(item);
  return [...doc.suggestions.values()].filter(
    (s) =>
      s.status === 'pending' &&
      (doc.partRevisions.has(s.id) || [...(doc.labelled.get(s.id) ?? [])].some(kept))
  );
}

/** Everything a fold must keep, as plain data with no ids or addresses. */
export function foldPrint(doc: OrderedDocument, base: OoxmlPackage, blobs: FoldBlobs): unknown {
  const key = (id: string | null): string | null => {
    const suggestion = id === null ? undefined : doc.suggestions.get(id);
    return suggestion?.status === 'pending' ? suggestion.id : null;
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
    const node = nodeName(item);

    // A unit in a new row or note keeps its label in the room, where a load reads it from the
    // container
    const by = item.quiet ? null : item.by;
    const recorded = (
      records: readonly { suggestion: string; untracked?: boolean; keptOnReject?: boolean }[]
    ) =>
      records
        .filter(
          (r) => r.untracked !== true || (r.suggestion !== item.by && r.keptOnReject !== true)
        )
        .map((r) => r.suggestion);
    const roles = {
      by: [by],
      del: item.dels,
      // A run holds one change record: a format names its suggestion once on a unit
      fmt: [...new Set(recorded(item.fmts))],
      prop: recorded(item.props),
      within: [...item.within.by, ...item.within.dels],
    };
    const labels = {
      kind: item.kind,
      node,
      by: key(by),
      dels: keys(item.dels),
      fmts: keys(roles.fmt),
      // A decided record takes over nothing: a fold resolves it. Records apply in their order of
      // creation, and a fold puts the records taken in front of their holder's. Each pending
      // suggestion a record was taken from counts, oldest first: a take-back gives it back
      replaced: [...item.fmts, ...item.props]
        .flatMap((r) => (key(r.suggestion) === null ? [] : [keys(takenFrom(r)).join('<')]))
        .filter((chain) => chain !== '')
        .sort(),
      // Records apply in their order of creation, whatever their place in the list
      props: keys(roles.prop).sort(),
      // A move's name stays on its words after its decision: they show as a pending move's
      // destination, or as its source the move struck
      move:
        item.move !== null &&
        (key(item.by) !== null ||
          item.dels.some((d) => key(d) !== null && movedAwayBy(doc, item, d))),
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
  const read = new CardRead(doc);
  const pending = keptPending(read);
  const file = orderedPackageBytes(doc, base, { blobs });
  return {
    // A story that shows nothing (a text box whose object went) is in no file
    stories: [...doc.stories]
      .flatMap(([name, story]) => {
        const hidden = drawnNowhere(doc, story);
        const shown = story.tree.toArray().filter((i) => showsOn(read, story, i) && !hidden.has(i));
        const listed = shown.flatMap((item) => units(story, item));
        return listed.length === 0 ? [] : [[storyLabel(name), ...listed].join('\n')];
      })
      .sort(),
    suggestions: pending.map((s) => `${s.id} ${makerKey(s)}`).sort(),
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
          JSON.stringify([s.id, r.suggestion, r.formattingKind, r.formattingChanges])
        )
      )
      .sort(),
    comments: [...doc.comments.values()].map(({ element, ...comment }) =>
      JSON.stringify({ ...comment, element: xmlOf(element) })
    ),
    partRevisions: [...doc.partRevisions]
      .filter(([id]) => key(id) !== null)
      .map(([id, r]) => JSON.stringify([key(id), r.part, r.element, r.styleId ?? null])),
    view: viewOf(read),
    revisions: revisionReads(doc, file),
    file: partsOf(file),
  };
}

/** The suggestions `orderedSuggestionsOfRevision` names for each revision mark of the file. */
function revisionReads(doc: OrderedDocument, file: Uint8Array): string[] {
  const read = readOoxmlPackage(file, {});
  if (!read.ok) {
    throw new Error(`the file does not read: ${read.reason}`);
  }

  const out: string[] = [];
  for (const [part, { root }] of read.package.parts) {
    eachElement(root, (element) => {
      const id = attributeOf(element, 'id');
      const author = attributeOf(element, 'author');
      if (id === undefined || author === undefined) {
        return;
      }

      const date = attributeOf(element, 'date');
      const name = { part, id, author, element: element.localName };
      const found = orderedSuggestionsOfRevision(
        doc,
        date === undefined ? name : { ...name, date }
      );
      out.push(`${part} ${element.localName} ${id}: ${found.join(' ')}`);
    });
  }

  return out;
}

type FoldBlobs = (sha256: string) => Uint8Array | undefined;

/** A story by its name: a text box's names its object, whose id a fold mints again. */
const storyLabel = (name: string): string => (isTextBoxStoryName(name) ? 'text box' : name);

/**
 * What each story's page shows: every character and its strike, but for a note or a text box
 * compose leaves out. A character's own face is left out: a fold writes the decided formats into
 * it, which the room keeps as records.
 */
function viewOf(read: CardRead): string[] {
  const out: string[] = [];
  for (const [name, story] of read.doc.stories) {
    for (let k = 0; k < story.tree.summary.marks; k += 1) {
      const view = orderedViewParagraph(read.doc, name, k);
      if (view !== null && !read.isLeftOut(view.mark, story)) {
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

/** Who decides in `checkFold` and `checkEveryCard`, and when. */
const DECIDER = { person: testPerson('Olive'), date: '2026-11-25T12:00:00Z' };

/**
 * The Word file after `ids` are decided as `status`, with the controls' prompts settled as a
 * decision settles them (`applyOrderedChange`), each XML part by what it says (as a reopen
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
  const apply = { ...DECIDER, ids: orderedIdMinter(orderedChangeId('check', 1)) };
  const { record } = withUndoRecord(doc, () => {
    for (const id of ids) {
      setStatus(doc, id, status);
    }

    settleDecidedTables(doc, ids, apply.ids);
    settleEmptiedControls(doc, apply);
    settleNamedPrompts(doc, apply);
    settlePromptMarks(doc);
  });
  const bytes = orderedPackageBytes(doc, base, { blobs });
  takeBack(doc, record);
  return semanticParts(bytes);
}

/** A decided Word file's XML parts, each by what it says (as a reopen compares decisions). */
function semanticParts(bytes: Uint8Array): Record<string, string> {
  const read = readOoxmlPackage(bytes, {});
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

/** The ids of each maker's pending suggestions that a fold keeps. */
function byMaker(doc: OrderedDocument): Map<string, string[]> {
  const out = new Map<string, string[]>();
  for (const s of keptPending(new CardRead(doc))) {
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
  expect({ label, outside: unitsOutsideParagraphs(doc) }).toEqual({ label, outside: [] });
  const print = foldPrint(doc, base, blobs);
  const live = [...byMaker(doc)].flatMap(([maker, ids]) =>
    DECISIONS.map((status) => ({ maker, status, file: decidedFile(doc, base, blobs, ids, status) }))
  );
  withCleanup(doc, () => {
    const fold = 'cleaned';
    const outside = unitsOutsideParagraphs(doc);
    expect({ label, fold, outside }).toEqual({ label, fold, outside: [] });
    expect({ label, fold, print: foldPrint(doc, base, blobs) }).toEqual({ label, fold, print });
    const cleaned = byMaker(doc);
    for (const { maker, status, file } of live) {
      const decided = decidedFile(doc, base, blobs, cleaned.get(maker) ?? [], status);
      expect({ label, fold, maker, status, file: decided }).toEqual({
        label,
        fold,
        maker,
        status,
        file,
      });
    }
  });
  for (const [fold, copy] of [
    ['once', once],
    ['twice', twice],
  ] as const) {
    then?.(copy.document);
    const outside = unitsOutsideParagraphs(copy.document);
    expect({ label, fold, outside }).toEqual({ label, fold, outside: [] });
    expect({ label, fold, print: foldPrint(copy.document, copy.base, blobs) }).toEqual({
      label,
      fold,
      print,
    });
    const reopened = byMaker(copy.document);
    for (const { maker, status, file } of live) {
      const ids = reopened.get(maker) ?? [];
      const decided = decidedFile(copy.document, copy.base, blobs, ids, status);
      expect({ label, fold, maker, status, file: decided }).toEqual({
        label,
        fold,
        maker,
        status,
        file,
      });
    }
  }
}

/**
 * One card decided as the organiser's decision change decides it, from the card as the organiser
 * saw it: the suggestions decided, or the refusal's reason, and the Word file; the document is
 * restored.
 */
function cardDecided(
  doc: OrderedDocument,
  base: OoxmlPackage,
  blobs: FoldBlobs,
  seen: readonly OrderedSeenSuggestion[],
  status: 'accepted' | 'rejected'
): {
  readonly decided: readonly string[] | string;
  readonly file: Record<string, string>;
  readonly outside: readonly string[];
} {
  const id = orderedChangeId('check', 1);
  const change = { v: 1, id, kind: 'decide', cards: [{ suggestions: seen }], status } as const;
  try {
    const { value, record } = withUndoRecord(doc, () =>
      applyOrderedChange(doc, change, { ...DECIDER, ids: orderedIdMinter(id) })
    );
    const bytes = orderedPackageBytes(doc, base, { blobs });
    const outside = unitsOutsideParagraphs(doc);
    takeBack(doc, record);
    return { decided: value.suggestions, file: semanticParts(bytes), outside };
  } catch (error) {
    if (!(error instanceof OrderedRefusal)) {
      throw error;
    }

    return {
      decided: error.reason,
      file: semanticParts(orderedPackageBytes(doc, base, { blobs })),
      outside: unitsOutsideParagraphs(doc),
    };
  }
}

/**
 * Fold `doc`, and fold the folded copy again: each card the room shows, decided on its own as the
 * organiser saw it, accepted and rejected, decides the same suggestions and gives the same Word
 * file on each copy as live. A suggestion keeps its id and its version across a fold, so the card
 * the organiser saw names the same suggestions on every copy. Returns the number of cards.
 */
export async function checkEveryCard(
  doc: OrderedDocument,
  base: OoxmlPackage,
  blobs: FoldBlobs,
  label: string
): Promise<number> {
  const once = await folded(doc, base, doc.file + 1, blobs);
  const twice = await folded(once.document, once.base, doc.file + 2, blobs);
  const cards = orderedCards(doc);
  for (const card of cards) {
    const seen = orderedCardSeen(doc, card.id) ?? [];
    for (const status of DECISIONS) {
      const live = cardDecided(doc, base, blobs, seen, status);
      expect({ label, card: card.suggestions, status, outside: live.outside }).toEqual({
        label,
        card: card.suggestions,
        status,
        outside: [],
      });
      const cleaned = withCleanup(doc, () => cardDecided(doc, base, blobs, seen, status));
      expect({ label, fold: 'cleaned', card: card.suggestions, status, ...cleaned }).toEqual({
        label,
        fold: 'cleaned',
        card: card.suggestions,
        status,
        ...live,
      });
      for (const [fold, copy] of [
        ['once', once],
        ['twice', twice],
      ] as const) {
        const shown = cardDecided(copy.document, copy.base, blobs, seen, status);
        expect({ label, fold, card: card.suggestions, status, ...shown }).toEqual({
          label,
          fold,
          card: card.suggestions,
          status,
          ...live,
        });
      }
    }
  }

  return cards.length;
}

/**
 * Run `read` on the document cleaned as a close cleans it (`compactOrdered`), then take the
 * cleanup back. Each item it removed throws on any read of its fields meanwhile, so a reader that
 * still holds one fails by name.
 */
export function withCleanup<T>(doc: OrderedDocument, read: () => T): T {
  const items = [...doc.stories.values()].flatMap((story) => story.tree.toArray());
  const { record } = withUndoRecord(doc, () => compactOrdered(doc));
  const restore = items.filter((item) => item.leaf === undefined).map(poison);
  try {
    return read();
  } finally {
    for (const undo of restore) {
      undo();
    }

    takeBack(doc, record);
  }
}

/** Make each read of an item's fields but its tree leaf throw; returns how to undo it. */
function poison(item: OrderedItem): () => void {
  const fields = Object.entries(Object.getOwnPropertyDescriptors(item)).filter(
    ([key]) => key !== 'leaf'
  );
  for (const [key] of fields) {
    Object.defineProperty(item, key, {
      configurable: true,
      enumerable: true,
      get: () => {
        throw new Error(`a read of a removed item's ${key}`);
      },
    });
  }

  return () => {
    for (const [key, descriptor] of fields) {
      Object.defineProperty(item, key, descriptor);
    }
  };
}
