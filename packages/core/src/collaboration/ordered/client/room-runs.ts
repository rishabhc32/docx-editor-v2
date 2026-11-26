// A checkpoint's pieces of one room run joined again (design D14).

import { ooxmlTreesEqual } from '../../../store/package/ooxml-serialize.ts';
import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { createOrderedItem } from '../model/document.ts';
import { replaceItems } from '../model/mutations.ts';
import { sameTakenFrom } from '../model/takeovers.ts';
import type { OrderedDocument, OrderedFormatRecord, OrderedItem } from '../model/types.ts';

/** A room run read so far: its first piece, and the format records its pieces hold. */
interface RoomRun {
  readonly first: Pick<OrderedItem, 'run' | 'face' | 'home'>;
  /** The first record read of each record of the room. */
  readonly fmts: OrderedFormatRecord[];
}

/** Whether two elements say the same: both none, or equal but for their node ids. */
const same = (a: OoxmlElement | null | undefined, b: OoxmlElement | null | undefined): boolean =>
  a === b ||
  (a !== null && a !== undefined && b !== null && b !== undefined && ooxmlTreesEqual(a, b));

/** Whether two format records are one record of the room, read from two pieces of its run. */
const sameRecord = (a: OrderedFormatRecord, b: OrderedFormatRecord): boolean =>
  a.kind === 'face' &&
  b.kind === 'face' &&
  a.suggestion === b.suggestion &&
  a.seq === b.seq &&
  sameTakenFrom(a, b) &&
  same(a.face, b.face) &&
  same(a.fileRecord, b.fileRecord);

/** The first value a map holds for `key`, set to `value` when it holds none. */
function firstOf<T>(firsts: Map<string, T>, key: string, value: T): T {
  const first = firsts.get(key);
  if (first !== undefined) {
    return first;
  }

  firsts.set(key, value);
  return value;
}

/**
 * A piece's format records, each as the first record read in its room run that says the same: a
 * pending suggestion can split the run before the record's first piece, so its pieces join each
 * other, not only the run's first piece.
 */
function joinedFormats(
  known: OrderedFormatRecord[],
  fmts: readonly OrderedFormatRecord[]
): readonly OrderedFormatRecord[] {
  const joined = fmts.map((record) => {
    const earlier = known.find((f) => sameRecord(f, record));
    if (earlier !== undefined) {
      return earlier;
    }

    known.push(record);
    return record;
  });
  return joined.every((record, i) => record === fmts[i]) ? fmts : joined;
}

/**
 * Join again the pieces of each room run a checkpoint numbers (`oc:run`): `numbers` holds the
 * number on each run and text element, by its node id, which is the `run` and the text shell of
 * the units the load read from it. A pending suggestion split the room's run, and the file holds
 * each piece as a run of its own. The units of each later piece take the first piece's run, the
 * paragraph it was loaded in (a piece after a paragraph break is drawn with new ids, as in the
 * room), and its face where theirs says the same; each format record becomes the first one read in
 * the run that says the same (`joinedFormats`); a text element takes the text shell of the first
 * one of its number. So a decision that takes the suggestion away draws the run as the room does.
 * Runs after the takeovers: a record keeps what it took.
 */
export function withCheckpointRuns(
  doc: OrderedDocument,
  numbers: ReadonlyMap<string, string>
): void {
  if (numbers.size === 0) {
    return;
  }

  const runs = new Map<string, RoomRun>();
  const texts = new Map<string, OoxmlElement>();
  for (const story of doc.stories.values()) {
    for (let item = story.tree.first(); item !== null; item = story.tree.next(item)) {
      const runNumber = numbers.get(item.run);
      const room =
        runNumber === undefined ? undefined : firstOf(runs, runNumber, { first: item, fmts: [] });
      const first = room?.first ?? item;
      const fmts = room === undefined ? item.fmts : joinedFormats(room.fmts, item.fmts);
      const { textShell } = item;
      const textNumber = textShell === undefined ? undefined : numbers.get(textShell.id);
      const shell =
        textShell === undefined || textNumber === undefined
          ? textShell
          : firstOf(texts, textNumber, textShell);
      if (first.run === item.run && shell === textShell && fmts === item.fmts) {
        continue;
      }

      const joined = joinedItem(item, first, fmts, same(shell, textShell) ? shell : textShell);
      if (joined.run !== item.run) {
        doc.runShells.delete(item.run);
      }

      replaceItems(doc, story, [item], [joined]);
      item = joined;
    }
  }
}

/** A later piece's unit in the first piece's run, with `fmts` and `textShell`. */
function joinedItem(
  item: OrderedItem,
  first: RoomRun['first'],
  fmts: readonly OrderedFormatRecord[],
  textShell: OoxmlElement | undefined
): OrderedItem {
  return createOrderedItem({
    ...item,
    run: first.run,
    home: first.home,
    face: same(item.face, first.face) ? first.face : item.face,
    fmts,
    ...(textShell === undefined ? {} : { textShell }),
  });
}
