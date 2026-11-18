// Every write to an ordered document is taken back exactly by its undo record: the same items
// (the same objects), the same labels, indexes, sums, suggestions, counters and resources.

import { addOrderedResource } from '../model/resources.ts';
import { describe, expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint } from '../../../store/package/ooxml-tree.ts';
import {
  composeOrderedPart,
  createOrderedItem,
  insertOrderedItems,
  isolateOrderedUnits,
  loadOrderedPackage,
  newOrderedSuggestion,
  orderedItemAt,
  OrderedRefusal,
  relabelOrderedItem,
  setOrderedStatus,
  splitOrderedItem,
  takeBackOrdered,
  withOrderedUndoRecord,
  type OrderedDocument,
  type OrderedStory,
} from '../index.ts';
import { snapshot } from './snapshot.ts';
import { docx, p, packageOf, t } from './word-fixtures.ts';
import { testAttribution, testPerson } from './test-persons.ts';

const BODY = '/word/document.xml';
const REV = 'w:author="Cara" w:date="2026-01-01T00:00:00Z"';

function load(): OrderedDocument {
  return loadOrderedPackage(
    packageOf(
      docx(
        p(t('The fee is 100 USD.')) +
          p(t('Due in ') + `<w:ins w:id="7" ${REV}>${t('30')}</w:ins>` + t(' days.')) +
          p(t('Doc \u{1F4C4} here'))
      )
    )
  );
}

function bodyOf(doc: OrderedDocument): OrderedStory {
  const story = doc.stories.get(BODY);
  if (!story) throw new Error('no body');
  return story;
}

/** The space of the n-th loaded paragraph's units. */
function paragraphSpace(doc: OrderedDocument, n: number): string {
  const marks = bodyOf(doc)
    .tree.toArray()
    .filter((item) => item.kind === 'mark');
  const mark = marks[n];
  if (!mark) throw new Error(`no paragraph ${n}`);
  return `b:${mark.space.slice(2)}`;
}

/** Ana types "45 " at offset 7 of paragraph 0. */
function anaTypes(doc: OrderedDocument): void {
  const story = bodyOf(doc);
  const space = paragraphSpace(doc, 0);
  const [before] = isolateOrderedUnits(doc, story, space, 0, 7).slice(-1);
  const ana = newOrderedSuggestion(doc, testAttribution('Ana', '2026-02-01T00:00:00Z'));
  doc.counters.insertion += 1;
  const typed = createOrderedItem({
    space: `i:${ana.id}`,
    kind: 'text',
    length: 3,
    text: '45 ',
    run: `i:${ana.id}`,
    by: ana.id,
  });
  insertOrderedItems(doc, story, before ?? null, [typed]);
}

/** Ben strikes "Due" in paragraph 1. */
function benStrikes(doc: OrderedDocument): void {
  const story = bodyOf(doc);
  const ben = newOrderedSuggestion(doc, testAttribution('Ben'));
  for (const item of isolateOrderedUnits(doc, story, paragraphSpace(doc, 1), 0, 3))
    relabelOrderedItem(doc, story, item, { dels: [...item.dels, ben.id] });
}

/** The server accepts Cara's insertion, and a link's address is added. */
function serverDecides(doc: OrderedDocument): void {
  setOrderedStatus(doc, 'S1', 'accepted');
  addOrderedResource(doc, {
    kind: 'relationship',
    key: 'p:olive-1:1.1',
    suggestion: 'S1',
    owner: '/word/document.xml',
    id: 'rOolive-1_1_1',
    type: 'hyperlink',
    target: { external: 'https://example.com/new' },
  });
}

describe('the undo record', () => {
  const edits: Record<string, (doc: OrderedDocument) => void> = {
    'a split': (doc) => {
      const story = bodyOf(doc);
      const item = orderedItemAt(story, { space: paragraphSpace(doc, 0), offset: 4 });
      if (item) splitOrderedItem(doc, story, item, 4);
    },
    'an isolation in three pieces': (doc) => {
      isolateOrderedUnits(doc, bodyOf(doc), paragraphSpace(doc, 0), 4, 7);
    },
    'every label': (doc) => {
      const story = bodyOf(doc);
      const s = newOrderedSuggestion(doc, testAttribution('Ana'));
      const [item] = isolateOrderedUnits(doc, story, paragraphSpace(doc, 0), 0, 3);
      if (!item) return;
      relabelOrderedItem(doc, story, item, {
        by: s.id,
        dels: [s.id],
        fmts: [
          {
            kind: 'properties',
            id: 'f1',
            suggestion: s.id,
            seq: s.seq,
            properties: [{ localName: 'b' }],
          },
        ],
        move: 'mv',
        retracted: true,
        gone: true,
        quiet: true,
      });
    },
    'an insertion': anaTypes,
    'a strike': benStrikes,
    'a status and a relationship': serverDecides,
    'a status set twice': (doc) => {
      setOrderedStatus(doc, 'S1', 'rejected');
      setOrderedStatus(doc, 'S1', 'accepted');
    },
  };

  for (const [name, edit] of Object.entries(edits)) {
    test(`takes back ${name} exactly`, () => {
      const doc = load();
      const before = snapshot(doc);
      const part = composeOrderedPart(doc, BODY).part;
      const { record } = withOrderedUndoRecord(doc, () => edit(doc));
      expect(snapshot(doc)).not.toEqual(before);
      takeBackOrdered(doc, record);
      expect(snapshot(doc)).toEqual(before);
      expect(canonicalOoxmlFingerprint(composeOrderedPart(doc, BODY).part)).toBe(
        canonicalOoxmlFingerprint(part)
      );
    });
  }

  test('an edit that throws midway is taken back, and the error goes on', () => {
    const doc = load();
    const before = snapshot(doc);
    const story = bodyOf(doc);
    const space = paragraphSpace(doc, 2);
    expect(() =>
      withOrderedUndoRecord(doc, () => {
        anaTypes(doc);
        // "Doc " then the two units of one emoji: a place between them is refused
        isolateOrderedUnits(doc, story, space, 0, 5);
      })
    ).toThrow(OrderedRefusal);
    expect(snapshot(doc)).toEqual(before);
    expect(doc.undo).toBeNull();
  });

  test('records do not nest, and nothing is taken back while one is open', () => {
    const doc = load();
    withOrderedUndoRecord(doc, () => {
      expect(() => withOrderedUndoRecord(doc, () => undefined)).toThrow();
      expect(() => takeBackOrdered(doc, { steps: [] })).toThrow();
    });
  });

  test('records taken back last first restore the document', () => {
    const doc = load();
    const before = snapshot(doc);
    const a = withOrderedUndoRecord(doc, () => anaTypes(doc)).record;
    const b = withOrderedUndoRecord(doc, () => benStrikes(doc)).record;
    takeBackOrdered(doc, b);
    takeBackOrdered(doc, a);
    expect(snapshot(doc)).toEqual(before);
  });

  test('taking back two pending changes, applying a server change, and applying them again equals the server change first', () => {
    const replica = load();
    const a = withOrderedUndoRecord(replica, () => anaTypes(replica)).record;
    const b = withOrderedUndoRecord(replica, () => benStrikes(replica)).record;
    takeBackOrdered(replica, b);
    takeBackOrdered(replica, a);
    withOrderedUndoRecord(replica, () => serverDecides(replica));
    withOrderedUndoRecord(replica, () => anaTypes(replica));
    withOrderedUndoRecord(replica, () => benStrikes(replica));

    const server = load();
    withOrderedUndoRecord(server, () => serverDecides(server));
    anaTypes(server);
    benStrikes(server);

    expect(snapshot(replica, false)).toEqual(snapshot(server, false));
    expect(canonicalOoxmlFingerprint(composeOrderedPart(replica, BODY).part)).toBe(
      canonicalOoxmlFingerprint(composeOrderedPart(server, BODY).part)
    );
  });
});

// Found by the random sessions (four people, seed 89): a party's control value over the
// destination of another person's move. The value's template (the control's words, read with
// the loader) holds the move's range markers, and the loader noted the move as one whose markers
// are items (`movesWithRanges`). The take-back kept that note, so the save no longer wrote the
// move's range markers on the copy that took the value back
describe('a control value over a move', () => {
  test('keeps the move range markers, and its take-back gives the save back', async () => {
    const {
      apply,
      at,
      opens,
      addressOf,
      load: loadApply,
      EDIT,
    } = await import('./change-apply-fixture.ts');
    const doc = loadApply();
    // Ana moves "fee" into the control's words: "Party: Ac[fee]me"
    apply(doc, 'ana:1', {
      kind: 'move',
      ...EDIT,
      from: at(doc, 0, 4),
      to: at(doc, 0, 7),
      place: at(doc, 2, 9),
    });
    const saved = () => composeOrderedPart(doc, BODY).part;
    const before = canonicalOoxmlFingerprint(saved());
    expect(before).toContain('moveToRangeStart');
    const [control] = opens(doc, 'sdt');
    if (!control) throw new Error('fixture');
    const { applyOrderedChange } = await import('../change/apply.ts');
    const { orderedIdMinter } = await import('../change/ids.ts');
    const change = {
      v: 1,
      id: 'ben:1',
      kind: 'control-value',
      ...EDIT,
      control: addressOf(control),
      value: { kind: 'text', text: 'Beta' },
    } as const;
    const ctx = {
      person: testPerson('Ben'),
      date: '2026-10-28T09:30:00Z',
      ids: orderedIdMinter(change.id),
    };
    const { record } = withOrderedUndoRecord(doc, () => applyOrderedChange(doc, change, ctx));
    // While the value stands, the move keeps its range markers too: the edit keeps none of the
    // template's
    expect(canonicalOoxmlFingerprint(saved()).match(/move(From|To)RangeStart/g)).toHaveLength(2);
    takeBackOrdered(doc, record);
    expect(canonicalOoxmlFingerprint(saved())).toBe(before);
  });
});
