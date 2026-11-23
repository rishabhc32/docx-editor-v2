// The room knows who wrote each item by person ID (design D9), with two real editors: two people
// of one company with one name own only their own suggestions, one person in two tabs owns the
// same ones, a namespace stays with one person, saved files carry names only (and no record the
// organiser took from a party), a checkpoint keeps every person into the next room, and presence
// names the person.

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { writeOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import { openOrderedCheckpoint, openOrderedRoomFile } from '../../client/open.ts';
import { applyOrderedChange } from '../../change/apply.ts';
import { writeOrderedCheckpoint } from '../../compose/checkpoint.ts';
import {
  composeOrderedPackage,
  orderedPackageBytes,
  writeOrderedPackage,
} from '../../compose/package.ts';
import type { OrderedPerson } from '../../model/person.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { OrderedRoom, OrderedRoomOpenError } from '../../room/room.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { markOf } from '../change-apply-fixture.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
  join,
  localRoom,
  ORGANISER,
  PARTY,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';

afterEach(closeAll);

const ANA: OrderedPerson = { id: 'id-ana-7f3c', name: 'Ana', org: 'Acme' };
const OTHER_ANA: OrderedPerson = { id: 'id-ana-91d2', name: 'Ana', org: 'Acme' };
const BEN: OrderedPerson = { id: 'id-ben-5e08', name: 'Ben', org: '' };
const OLGA: OrderedPerson = { id: 'id-olga-2b9e', name: 'Olga', org: 'Lexcorp' };

/** A file suggestion under Ana's Word name, which no person of the room made. */
const SOURCE = docx(
  p(
    t('The fee is ') +
      `<w:ins w:id="1" w:author="Ana (Acme)" w:date="2026-10-01T00:00:00Z">${t('net ')}</w:ins>` +
      t('100 USD.')
  ) + p(t('Due in 30 days.'))
);

const MANIFEST = 'ordered-checkpoint.json';

/** A room with the suggestions-only policy, and a party who joins it as a person. */
const partyRoom = (file = 1) => localRoom(SOURCE, { policy: suggestionsOnlyPolicy(), file });
const party = (room: OrderedLocalRoom, person: OrderedPerson, wordFile = SOURCE) =>
  join(room, person.name, { person, wordFile });

/** The person's refusals as `code/reason`, then none. */
function refused(person: Person): string[] {
  const reasons = person.refusals.map((r) => `${r.code}/${r.reason ?? ''}`);
  person.refusals.length = 0;
  return reasons;
}

/** Every copy is the room's, nothing is refused, and every editor shows its list. */
function converged(room: OrderedLocalRoom, ...people: Person[]): void {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
  }
}

/** The pending suggestions as `[author, person ID]`, sorted: null for a file's suggestion. */
const makers = (doc: OrderedDocument): [string, string | null][] =>
  [...doc.suggestions.values()]
    .filter((s) => s.status === 'pending')
    .map((s): [string, string | null] => [s.attribution.author, s.attribution.person?.id ?? null])
    .sort();

/** The id of the pending suggestion a person made. */
function suggestionOf(doc: OrderedDocument, person: OrderedPerson): string {
  const found = [...doc.suggestions.values()].find(
    (s) => s.status === 'pending' && s.attribution.person?.id === person.id
  );
  if (!found) throw new Error(`no suggestion of ${person.id}`);
  return found.id;
}

/** The bytes a streaming writer writes. */
async function collected(
  run: (write: (chunk: Uint8Array) => Promise<void>) => Promise<void>
): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  await run(async (chunk) => {
    chunks.push(chunk);
  });
  return new Uint8Array(Buffer.concat(chunks));
}

const noMedia = async (sha256: string): Promise<Uint8Array> => {
  throw new Error(`no media is read: ${sha256}`);
};

/** The room's document as a checkpoint. */
function checkpointOf(room: OrderedLocalRoom, file: number): Promise<Uint8Array> {
  const { base } = openOrderedRoomFile(SOURCE, { file });
  return collected((write) =>
    writeOrderedCheckpoint(room.room.document, base, { media: noMedia, write })
  );
}

/** The entries of a zip that hold `text`. */
const entriesWith = (bytes: Uint8Array, text: string): string[] =>
  Object.entries(unzipSync(bytes))
    .filter(([, data]) => strFromU8(data).includes(text))
    .map(([name]) => name);

describe('a room knows each person by ID', () => {
  test('two people of one company with one name own only their own suggestions', async () => {
    const room = partyRoom();
    const first = await party(room, ANA);
    const second = await party(room, OTHER_ANA);
    caretAt(first, 1, 0);
    surface(first).insertPlainText('Rent ');
    await settle();

    // The other Ana's words inside them are another person's change
    caretAt(second, 1, 2);
    surface(second).insertPlainText('x');
    await settle();
    expect(refused(second)).toEqual(['ordered-refused/nested']);

    // The first Ana adds to her own words
    caretAt(first, 1, 2);
    surface(first).insertPlainText('e');
    await settle();
    expect(editorTexts(second)[1]).toBe('Reent Due in 30 days.');
    expect(makers(room.room.document)).toEqual([
      ['Ana (Acme)', null],
      ['Ana (Acme)', 'id-ana-7f3c'],
    ]);

    // The other Ana may not take them back
    const doc = room.room.document;
    const actor = { person: OTHER_ANA, ...PARTY, namespace: second.handle.session.namespace };
    const retract = { v: 1, id: `${actor.namespace}:900`, kind: 'retract' };
    const sent = room.room.submit(actor, [{ ...retract, suggestions: [suggestionOf(doc, ANA)] }]);
    expect(sent.refused.map((r) => `${r.code}/${r.reason ?? ''}`)).toEqual([
      'policy-not-own/others_suggestions',
    ]);

    // The first Ana takes them back
    surface(first).undo();
    surface(first).undo();
    await settle();
    expect(editorTexts(second)[1]).toBe('Due in 30 days.');
    converged(room, first, second);
  });

  test('one person in two tabs owns the same suggestions; a namespace stays theirs', async () => {
    const room = partyRoom();
    const tab = await party(room, ANA);
    const other = await party(room, ANA);
    const namespace = tab.handle.session.namespace;
    expect(other.handle.session.namespace).not.toBe(namespace);
    caretAt(tab, 1, 0);
    surface(tab).insertPlainText('Rent ');
    await settle();
    caretAt(other, 1, 2);
    surface(other).insertPlainText('e');
    await settle();
    expect(editorTexts(tab)[1]).toBe('Reent Due in 30 days.');

    // One suggestion: the second tab's words joined the first tab's
    expect(makers(room.room.document)).toEqual([
      ['Ana (Acme)', null],
      ['Ana (Acme)', 'id-ana-7f3c'],
    ]);
    surface(other).undo();
    await settle();
    expect(editorTexts(tab)[1]).toBe('Rent Due in 30 days.');
    caretAt(tab, 0, 23);
    surface(tab).insertPlainText(' Net.');
    await settle();
    converged(room, tab, other);

    // Another person on Ana's namespace is refused by name
    const insert = {
      v: 1,
      id: `${namespace}:900`,
      kind: 'insert',
      part: BODY,
      mode: 'suggest',
      place: markOf(room.room.document, 1),
      text: 'x',
    };
    const sent = room.room.submit({ person: BEN, ...PARTY, namespace }, [insert]);
    expect(sent.refused.map((r) => r.code)).toEqual(['stamp']);

    // A log that writes one of Ana's changes under Ben's ID does not open
    const log = room.room.joinSnapshot(0);
    const last = log.filter((change) => change.id.startsWith(`${namespace}:`)).at(-1);
    if (!last) throw new Error('no change of the first tab');
    const reopen = (entries: typeof log) => {
      try {
        new OrderedRoom({
          document: openOrderedRoomFile(SOURCE, { file: 1 }).document,
          applyChange: applyOrderedChange,
          policy: suggestionsOnlyPolicy(),
          now: () => 0,
          log: entries,
        });
      } catch (error) {
        return error instanceof OrderedRoomOpenError ? `${error.code}@${error.at}` : 'other';
      }

      return 'opened';
    };
    expect(reopen(log)).toBe('opened');
    const forged = log.map((change) => (change === last ? { ...change, by: BEN } : change));
    expect(reopen(forged)).toBe(`log-stamp@${last.n}`);
  });

  test('saved files write the Word name, and no person ID or takeover', async () => {
    const room = partyRoom();
    const ana = await party(room, ANA);
    const ben = await party(room, BEN);
    const olga = await join(room, OLGA.name, { person: OLGA, wordFile: SOURCE, login: ORGANISER });
    caretAt(ana, 1, 0);
    surface(ana).insertPlainText('Rent ');
    caretAt(ben, 0, 23);
    surface(ben).insertPlainText(' Net.');
    await settle();
    caretAt(ben, 1, 0, 4);
    expect(ben.handle.session.addSelectionComment('Why rent?')).not.toBeNull();
    await settle();

    // Olga's italic takes Ana's bold on "Due" (policy 2)
    caretAt(ana, 1, 5, 8);
    expect(ana.editor.exec({ type: 'toggleMark', mark: 'bold' }).ok).toBe(true);
    await settle();
    caretAt(olga, 1, 5, 8);
    expect(olga.editor.exec({ type: 'toggleMark', mark: 'italic' }).ok).toBe(true);
    await settle();
    converged(room, ana, ben, olga);
    const doc = room.room.document;
    const { base } = openOrderedRoomFile(SOURCE, { file: 1 });
    const exported = {
      bytes: orderedPackageBytes(doc, base),
      streamed: await collected((write) =>
        writeOrderedPackage(doc, base, { media: noMedia, write })
      ),
      composed: writeOoxmlPackage(composeOrderedPackage(doc, base)),
      editor: new Uint8Array(await ana.editor.save()),
    };
    for (const [form, bytes] of Object.entries(exported)) {
      const body = strFromU8(unzipSync(bytes)['word/document.xml'] ?? new Uint8Array());
      const authors = new Set([...body.matchAll(/w:author="([^"]*)"/g)].map((m) => m[1]));
      expect({ form, authors: [...authors].sort() }).toEqual({
        form,
        authors: ['Ana (Acme)', 'Ben', 'Olga (Lexcorp)'],
      });
      expect({ form, records: body.match(/<w:rPrChange /g)?.length }).toEqual({ form, records: 1 });
      expect({ form, manifest: MANIFEST in unzipSync(bytes) }).toEqual({ form, manifest: false });
      const comments = strFromU8(unzipSync(bytes)['word/comments.xml'] ?? new Uint8Array());
      expect({ form, comment: comments.includes('w:author="Ben"') }).toEqual({
        form,
        comment: true,
      });
      for (const person of [ANA, BEN, OLGA]) {
        expect({ form, id: person.id, entries: entriesWith(bytes, person.id) }).toEqual({
          form,
          id: person.id,
          entries: [],
        });
      }
    }

    // The checkpoint keeps them and the takeover in its manifest; its parts name only who typed
    // a run, by the namespace of the change (`oc:typer`)
    const checkpoint = await checkpointOf(room, 1);
    const parts = Object.fromEntries(
      Object.entries(unzipSync(checkpoint)).map(([name, data]) => [
        name,
        strToU8(strFromU8(data).replace(/ oc:typer="[^"]*"/g, '')),
      ])
    );
    for (const person of [ANA, BEN, OLGA]) {
      expect(entriesWith(zipSync(parts), person.id)).toEqual([MANIFEST]);
    }

    const manifest = JSON.parse(strFromU8(unzipSync(checkpoint)[MANIFEST] ?? new Uint8Array()));
    expect(manifest.takeovers).toHaveLength(1);
  });

  test('presence names the person by their Word name', async () => {
    const room = partyRoom();
    const ana = await party(room, ANA);
    const ben = await party(room, BEN);
    caretAt(ana, 1, 4);
    caretAt(ben, 1, 0);
    await settle();
    expect(ben.handle.session.participants().map((x) => [x.name, x.actorId, x.isLocal])).toEqual([
      ['Ben', BEN.id, true],
      ['Ana (Acme)', ANA.id, false],
    ]);
    expect(ben.handle.session.remoteSelections()).toEqual([
      expect.objectContaining({ actorId: ANA.id, name: 'Ana (Acme)' }),
    ]);
    expect(ana.handle.session.remoteSelections()).toEqual([
      expect.objectContaining({ actorId: BEN.id, name: 'Ben' }),
    ]);
  });

  test('a login ID with spaces at its ends is the trimmed person everywhere', async () => {
    const room = partyRoom();
    const ana = await party(room, { ...ANA, id: ` ${ANA.id} ` });
    const ben = await party(room, BEN);
    const namespace = ana.handle.session.namespace;
    caretAt(ana, 1, 0);
    surface(ana).insertPlainText('Rent ');
    await settle();
    expect(room.room.namespaceOwner(namespace)).toBe(ANA.id);

    // Her presence in her own namespace reaches Ben
    caretAt(ana, 1, 4);
    await settle();
    const ids = surface(ben).session.paragraphIds();
    const seen = ben.handle.session.remoteSelections();
    expect(seen).toHaveLength(1);
    expect(seen[0]).toMatchObject({ actorId: ANA.id, head: { nodeId: ids[1], offset: 4 } });

    // A reconnect joins her namespace again
    room.disconnect(namespace);
    await settle();
    ana.handle.session.reconnect();
    await settle();
    expect(ana.handle.session.status()).toBe('ready');
    caretAt(ana, 0, 23);
    surface(ana).insertPlainText(' Net.');
    await settle();
    expect(editorTexts(ben)[0]).toBe('The fee is net 100 USD. Net.');
    converged(room, ana, ben);
  });
});

describe('a checkpoint keeps every person', () => {
  /** Ana, the other Ana and Ben each suggest, in a room of file 3. */
  async function suggested() {
    const room = partyRoom(3);
    const ana = await party(room, ANA);
    const other = await party(room, OTHER_ANA);
    const ben = await party(room, BEN);
    caretAt(ana, 1, 0);
    surface(ana).insertPlainText('Rent ');
    caretAt(other, 1, 15);
    surface(other).insertPlainText(' net');
    caretAt(ben, 0, 23);
    surface(ben).insertPlainText(' Paid.');
    await settle();
    converged(room, ana, other, ben);
    return { room, checkpoint: await checkpointOf(room, 3) };
  }

  const EXPECTED: [string, string | null][] = [
    ['Ana (Acme)', null],
    ['Ana (Acme)', 'id-ana-7f3c'],
    ['Ana (Acme)', 'id-ana-91d2'],
    ['Ben', 'id-ben-5e08'],
  ];

  test('under its own file number, and as file 1 of the next room', async () => {
    const { room, checkpoint } = await suggested();
    expect(makers(room.room.document)).toEqual(EXPECTED);
    expect(makers(openOrderedCheckpoint(checkpoint, { file: 3 }).document)).toEqual(EXPECTED);
    expect(makers(openOrderedCheckpoint(checkpoint, { file: 1 }).document)).toEqual(EXPECTED);
  });

  test("the next room knows whose each suggestion is, and the file one is nobody's", async () => {
    const { checkpoint } = await suggested();
    const next = localRoom(checkpoint, { policy: suggestionsOnlyPolicy() });
    expect(makers(next.room.document)).toEqual(EXPECTED);
    const ana = await party(next, ANA, checkpoint);
    const other = await party(next, OTHER_ANA, checkpoint);

    // Ana adds to her own words of the last room; the other Ana may not
    caretAt(ana, 1, 2);
    surface(ana).insertPlainText('e');
    await settle();
    caretAt(other, 1, 3);
    surface(other).insertPlainText('x');
    await settle();
    expect(refused(other)).toEqual(['ordered-refused/nested']);

    // The file's suggestion under her Word name is not hers
    caretAt(ana, 0, 13);
    surface(ana).insertPlainText('x');
    await settle();
    expect(refused(ana)).toEqual(['ordered-refused/nested']);
    expect(editorTexts(other)[1]).toBe('Reent Due in 30 days. net');
    converged(next, ana, other);
  });

  test('persons that do not fit the file are refused by name', async () => {
    const { checkpoint } = await suggested();
    const entries = unzipSync(checkpoint);
    const manifest = strFromU8(entries[MANIFEST] ?? new Uint8Array());
    entries[MANIFEST] = new TextEncoder().encode(manifest.replace('"org":"Acme"', '"org":"Other"'));
    expect(() => openOrderedCheckpoint(zipSync(entries), { file: 1 })).toThrow(
      'checkpoint-manifest'
    );
  });
});
