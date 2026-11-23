// A suggestion that took a party's records (policy 2) keeps them across a close, with two real
// editors and one room: the checkpoint names each taken record, and a reopen in the same room or
// as file 1 of the next room holds them again. The organiser's take-back gives the party's
// change back, every decision gives the live document, and `orderedReplacedSuggestions` names the
// party's suggestion on every copy. The room takes over formats, paragraph properties and
// section properties; a table, row or cell change over another author's is refused (the engine's
// `other-authors-row-change`), so no table record is taken over.

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync, zipSync } from 'fflate';
import { openOrderedCheckpoint, openOrderedRoomFile } from '../../client/open.ts';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import { writeOrderedCheckpoint } from '../../compose/checkpoint.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import { orderedReplacedSuggestions } from '../../compose/replaced.ts';
import { decideOrdered } from '../../decide.ts';
import { retractSuggestion } from '../../edits/retract.ts';
import { isPending } from '../../model/document.ts';
import { makerOf, type OrderedPerson } from '../../model/person.ts';
import { OrderedRefusal } from '../../model/refusal.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { OrderedRoom } from '../../room/room.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { realApply } from '../room-harness.ts';
import {
  closeAll,
  join,
  localRoom,
  ORGANISER,
  settle,
  storiesOf,
  surface,
  type Person,
} from './editor-harness.ts';
import { exec, select } from './format-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { converged } from './tables-harness.ts';
import { docx, p, t } from '../word-fixtures.ts';

afterEach(closeAll);

const review = reviewTestModule();
const ANA: OrderedPerson = { id: 'id-ana-7f3c', name: 'Ana', org: 'Acme' };
const OLGA: OrderedPerson = { id: 'id-olga-2b9e', name: 'Olga', org: 'Lexcorp' };
const MANIFEST = 'ordered-checkpoint.json';

/** Two paragraphs and the body's own section, as Word writes a new document's. */
const SOURCE = docx(
  p(t('The fees are below.')) +
    p(t('Signed here.')) +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720"' +
    ' w:footer="720" w:gutter="0"/></w:sectPr>'
);

/** The room's Word file number: a reopen in the same room uses it, the next room 1. */
const FILE = 3;

type Step = (person: Person) => void;

/** A toggle of `mark` on `[start, end)` of the first paragraph: the words "fees" by default. */
const toggle =
  (mark: 'bold' | 'italic' | 'underline', start = 4, end = 8): Step =>
  (person) => {
    select(person.editor, start, end);
    exec(person.editor, { type: 'toggleMark', mark });
  };

/** An alignment of the first paragraph. */
const align =
  (value: 'center' | 'right'): Step =>
  (person) => {
    select(person.editor, 0, 0);
    exec(person.editor, { type: 'setAlignment', align: value });
  };

/** Page margins of every section. */
const margins =
  (margin: 'marginLeftTwips' | 'marginTopTwips'): Step =>
  (person) => {
    expect(surface(person).setSectionProperties({ [margin]: 1080 })).toBe(true);
  };

/** The party's change, the organiser's over it, and what the read names. */
const KINDS = [
  {
    name: 'a format',
    party: toggle('bold'),
    organiser: toggle('italic'),
    kind: 'rPrChange',
    changes: [{ property: 'bold', value: 'true' }],
  },
  {
    name: 'a paragraph property',
    party: align('center'),
    organiser: align('right'),
    kind: 'pPrChange',
    changes: [{ property: 'alignment', value: 'center' }],
  },
  {
    name: 'a section property',
    party: margins('marginLeftTwips'),
    organiser: margins('marginTopTwips'),
    kind: 'sectPrChange',
    changes: [],
  },
] as const;

/** A room of the suggestions-only policy, Ana as a party and Olga as the organiser. */
async function people(file = SOURCE) {
  const room = localRoom(file, { policy: suggestionsOnlyPolicy(), file: FILE });
  const ana = await join(room, 'Ana', { person: ANA, wordFile: file, review });
  const olga = await join(room, 'Olga', { person: OLGA, wordFile: file, login: ORGANISER, review });
  return { room, ana, olga };
}

/** Each person's steps in turn, every copy converged after each. */
async function steps(room: OrderedLocalRoom, ...turns: [Person, Step][]): Promise<void> {
  const everyone = [...new Set(turns.map(([person]) => person))];
  for (const [person, step] of turns) {
    step(person);
    await settle();
    converged(room, ...everyone);
  }
}

/** The room's document as a checkpoint. */
async function checkpointOf(room: OrderedLocalRoom, file = SOURCE): Promise<Uint8Array> {
  const { base } = openOrderedRoomFile(file, { file: FILE });
  const chunks: Uint8Array[] = [];
  await writeOrderedCheckpoint(room.room.document, base, {
    media: async (sha256) => {
      throw new Error(`no media is read: ${sha256}`);
    },
    write: async (chunk) => {
      chunks.push(chunk);
    },
  });
  return new Uint8Array(Buffer.concat(chunks));
}

/** The id of the pending suggestion a person made last. */
function suggestionOf(doc: OrderedDocument, person: OrderedPerson): string {
  const made = [...doc.suggestions.values()].filter(
    (s) => s.status === 'pending' && s.attribution.person?.id === person.id
  );
  const last = made.at(-1);
  if (!last) {
    throw new Error(`no suggestion of ${person.id}`);
  }

  return last.id;
}

/** What the read names for the organiser's suggestion, without ids. */
const replacedBy = (doc: OrderedDocument) =>
  orderedReplacedSuggestions(doc, suggestionOf(doc, OLGA)).map(
    ({ suggestion: _, ...named }) => named
  );

/**
 * A copy's pending suggestions by maker, and the records they hold, with the suggestion each
 * took them from: no ids and no addresses (those hold the file number).
 */
function held(doc: OrderedDocument) {
  const maker = (id: string) => {
    const attribution = doc.suggestions.get(id)?.attribution;
    return `${attribution?.author}/${attribution?.person?.id ?? null}`;
  };
  const records = new Set<string>();
  for (const story of doc.stories.values()) {
    for (const item of story.tree.toArray()) {
      const listed = [
        ...item.fmts.map((r) => ['rPr', r] as const),
        ...item.props.map((r) => [r.target, r] as const),
      ];
      for (const [target, record] of listed) {
        if (isPending(doc, record.suggestion)) {
          const from = record.replaced === undefined ? '-' : maker(record.replaced);
          records.add(`${target} ${maker(record.suggestion)} < ${from}`);
        }
      }
    }
  }

  const suggestions = [...doc.suggestions.values()]
    .filter((s) => s.status === 'pending')
    .map((s) => maker(s.id))
    .sort();
  return { suggestions, records: [...records].sort() };
}

/**
 * Every story of a copy, its revision ids left out: a new suggestion's own differ. An empty
 * `w:rPr` or `w:pPr` reads as none: a reopened file's change record keeps the empty properties it
 * held, which a reject gives back, with or without a takeover.
 */
const shown = (doc: OrderedDocument): string[] =>
  storiesOf(doc).map((xml) => xml.replace(/ w:id="[^"]*"/g, '').replace(/<w:(rPr|pPr)\/>/g, ''));

describe('a suggestion that took a party’s records keeps them across a close', () => {
  for (const { name, party, organiser, kind, changes } of KINDS) {
    test(`${name}: the read names the party’s suggestion live, and after a reopen`, async () => {
      const { room, ana, olga } = await people();
      await steps(room, [ana, party], [olga, organiser]);
      const live = room.room.document;
      const expected = [
        {
          author: 'Ana (Acme)',
          person: ANA,
          date: expect.any(String),
          formattingKind: kind,
          formattingChanges: changes,
        },
      ];
      expect(replacedBy(live)).toEqual(expected);
      for (const person of [ana, olga]) {
        expect(replacedBy(person.handle.session.orderedDocument())).toEqual(replacedBy(live));
      }

      // The same room, and the next room's file 1
      const checkpoint = await checkpointOf(room);
      for (const file of [FILE, 1]) {
        const reopened = openOrderedCheckpoint(checkpoint, { file }).document;
        expect(held(reopened)).toEqual(held(live));
        expect(replacedBy(reopened)).toEqual(replacedBy(live));
        expect(shown(reopened)).toEqual(shown(live));
      }
    });

    test(`${name}: the organiser's take-back gives the party's change back, as live`, async () => {
      const { room, ana, olga } = await people();
      await steps(room, [ana, party], [olga, organiser]);
      const checkpoint = await checkpointOf(room);
      surface(olga).undo();
      await settle();
      converged(room, ana, olga);
      expect(held(room.room.document).records.filter((r) => r.includes('Olga'))).toEqual([]);

      // The next room's pages read the same; the organiser takes it back there too
      const next = localRoom(checkpoint, { policy: suggestionsOnlyPolicy() });
      const page = await join(next, 'Olga', {
        person: OLGA,
        wordFile: checkpoint,
        login: ORGANISER,
        review,
      });
      const doc = next.room.document;
      expect(replacedBy(page.handle.session.orderedDocument())).toEqual(replacedBy(doc));
      const actor = { person: OLGA, ...ORGANISER, namespace: page.handle.session.namespace };
      const retract = { v: 1, id: `${actor.namespace}:900`, kind: 'retract' };
      const sent = next.room.submit(actor, [
        { ...retract, suggestions: [suggestionOf(doc, OLGA)] },
      ]);
      expect(sent.refused).toEqual([]);
      expect(held(doc)).toEqual(held(room.room.document));
      expect(shown(doc)).toEqual(shown(room.room.document));
    });

    test(`${name}: each decision in each order gives the live document`, async () => {
      const { room, ana, olga } = await people();
      await steps(room, [ana, party], [olga, organiser]);
      const checkpoint = await checkpointOf(room);

      // A copy of the live room, from its log
      const log = room.room.joinSnapshot(0);
      const liveCopy = () =>
        new OrderedRoom({
          document: openOrderedRoomFile(SOURCE, { file: FILE }).document,
          applyChange: realApply,
          policy: suggestionsOnlyPolicy(),
          now: () => 0,
          log,
        }).document;
      const orders = [
        ['accept', 'accept'],
        ['accept', 'reject'],
        ['reject', 'accept'],
        ['reject', 'reject'],
      ] as const;
      for (const first of [ANA, OLGA]) {
        for (const decisions of orders) {
          const order = first === ANA ? [ANA, OLGA] : [OLGA, ANA];
          const decide = (doc: OrderedDocument) => {
            const ids = order.map((person) => suggestionOf(doc, person));
            const outcomes = ids.map((id, k) => {
              try {
                return decideOrdered(doc, id, decisions[k] ?? 'accept').length;
              } catch (error) {
                if (error instanceof OrderedRefusal) {
                  return error.reason;
                }

                throw error;
              }
            });
            return { outcomes, held: held(doc), shown: shown(doc) };
          };
          const expected = decide(liveCopy());
          const reopened = openOrderedCheckpoint(checkpoint, { file: FILE }).document;
          expect({ first: first.name, decisions, ...decide(reopened) }).toEqual({
            first: first.name,
            decisions,
            ...expected,
          });
        }
      }
    });
  }
});

describe('the records of a takeover', () => {
  test('one change record id in several places names the party’s suggestion at each', async () => {
    const { room, ana, olga } = await people();

    // Ana bolds "The", then underlines "are"; Olga's italic over the whole sentence takes both
    await steps(
      room,
      [ana, toggle('bold', 0, 3)],
      [ana, toggle('underline', 9, 12)],
      [olga, toggle('italic', 0, 19)]
    );
    const live = room.room.document;
    expect(replacedBy(live).map((r) => r.formattingChanges)).toEqual([
      [{ property: 'bold', value: 'true' }],
      [{ property: 'underline', value: 'single' }],
    ]);
    const checkpoint = await checkpointOf(room);
    const manifest = JSON.parse(strFromU8(unzipSync(checkpoint)[MANIFEST] ?? new Uint8Array()));

    // Olga's one `w:id` marks the paragraph mark first, then "The", " fees ", "are" and " below.":
    // the second and the fourth record hold a takeover, each by its label
    const body = strFromU8(unzipSync(checkpoint)['word/document.xml'] ?? new Uint8Array());
    const records = [...body.matchAll(/<w:[a-zA-Z]+Change [^>]*>/g)].map(([record]) => record);
    expect(records.flatMap((record, k) => (record.includes(' oc:label=') ? [k] : []))).toEqual([
      1, 3,
    ]);
    expect(manifest.takeovers).toHaveLength(2);
    const reopened = openOrderedCheckpoint(checkpoint, { file: 1 }).document;
    expect(held(reopened)).toEqual(held(live));
    expect(replacedBy(reopened)).toEqual(replacedBy(live));
  });

  test('a reopen names the party’s suggestions in the order the live read does', async () => {
    const { room, ana, olga } = await people();

    // Ana's underline runs on to "Signed": only that one keeps a change record of its own
    const underline: Step = (person) => {
      const [first = '', second = ''] = surface(person).session.paragraphIds();
      surface(person).setSelection({
        anchor: { paragraphId: first, offset: 9 },
        head: { paragraphId: second, offset: 6 },
      });
      exec(person.editor, { type: 'toggleMark', mark: 'underline' });
    };
    await steps(room, [ana, toggle('bold')], [ana, underline], [olga, toggle('italic', 0, 19)]);
    const live = room.room.document;
    expect(replacedBy(live).map((r) => r.formattingChanges)).toEqual([
      [{ property: 'bold', value: 'true' }],
      [{ property: 'underline', value: 'single' }],
    ]);
    const checkpoint = await checkpointOf(room);
    for (const file of [FILE, 1]) {
      expect(replacedBy(openOrderedCheckpoint(checkpoint, { file }).document)).toEqual(
        replacedBy(live)
      );
    }
  });

  test('a unit whose taken records add nothing new leaves the read the same', async () => {
    const { room, ana, olga } = await people();

    // Ana's underline takes her bold on "fees": "The " and "fees" add a change, " are" nothing
    await steps(
      room,
      [ana, toggle('bold', 0, 8)],
      [ana, toggle('underline', 4, 12)],
      [olga, toggle('italic', 0, 19)]
    );
    const live = room.room.document;
    expect(replacedBy(live).map((r) => r.formattingChanges)).toEqual([
      [{ property: 'bold', value: 'true' }],
      [
        { property: 'bold', value: 'true' },
        { property: 'underline', value: 'single' },
      ],
    ]);
    const checkpoint = await checkpointOf(room);
    expect(replacedBy(openOrderedCheckpoint(checkpoint, { file: 1 }).document)).toEqual(
      replacedBy(live)
    );
  });

  test('a taken record that writes the italic already there names it once', async () => {
    const { room, ana, olga } = await people();

    // Ana's underline writes "fees" italic again, over her italic there that it takes
    await steps(
      room,
      [ana, toggle('italic', 4, 8)],
      [ana, toggle('underline', 0, 12)],
      [olga, toggle('bold', 0, 12)]
    );
    const live = room.room.document;
    expect(replacedBy(live).map((r) => r.formattingChanges)).toEqual([
      [
        { property: 'underline', value: 'single' },
        { property: 'italic', value: 'true' },
      ],
    ]);
    const checkpoint = await checkpointOf(room);
    expect(replacedBy(openOrderedCheckpoint(checkpoint, { file: 1 }).document)).toEqual(
      replacedBy(live)
    );
  });

  test('each run of a field gives the party’s format back after a reopen', async () => {
    const field =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>7</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const file = docx(`<w:p>${'<w:r><w:t xml:space="preserve">Page </w:t></w:r>'}${field}</w:p>`);
    const { room, ana, olga } = await people(file);

    // "Page " and the field's one unit
    await steps(room, [ana, toggle('bold', 0, 6)], [olga, toggle('italic', 0, 6)]);
    const checkpoint = await checkpointOf(room, file);
    const reopened = openOrderedCheckpoint(checkpoint, { file: FILE }).document;
    expect(held(reopened)).toEqual(held(room.room.document));
    expect(shown(reopened)).toEqual(shown(room.room.document));

    // Olga takes hers back: Ana's bold shows again on "Page ", the field's result and the mark
    surface(olga).undo();
    await settle();
    converged(room, ana, olga);
    retractSuggestion(reopened, suggestionOf(reopened, OLGA), makerOf(OLGA));
    expect(shown(reopened)).toEqual(shown(room.room.document));
    expect(shown(reopened)[0]?.match(/<w:b\/>/g)).toHaveLength(3);
  });

  test('a party’s suggestion with marks of its own in the file is one suggestion again', async () => {
    const { room, ana, olga } = await people();

    // Olga's italic takes "The" of Ana's bold; " fees" keeps Ana's own change record
    await steps(room, [ana, toggle('bold', 0, 8)], [olga, toggle('italic', 0, 3)]);
    const live = room.room.document;
    expect(held(live).suggestions).toEqual([
      'Ana (Acme)/id-ana-7f3c',
      'Olga (Lexcorp)/id-olga-2b9e',
    ]);
    const checkpoint = await checkpointOf(room);
    for (const file of [FILE, 1]) {
      const reopened = openOrderedCheckpoint(checkpoint, { file }).document;
      expect(held(reopened)).toEqual(held(live));
      expect(replacedBy(reopened)).toEqual(replacedBy(live));
    }
  });

  test('a manifest whose takeovers do not read or do not fit the file is refused', async () => {
    const { room, ana, olga } = await people();
    await steps(room, [ana, toggle('bold')], [olga, toggle('italic')]);
    const checkpoint = await checkpointOf(room);
    const entries = unzipSync(checkpoint);
    const manifest = JSON.parse(strFromU8(entries[MANIFEST] ?? new Uint8Array()));
    const [takeover] = manifest.takeovers;
    const cases = [
      { ...manifest, takeovers: [{ ...takeover, label: takeover.label + 1 }] },
      { ...manifest, takeovers: [{ ...takeover, label: -1 }] },
      { ...manifest, takeovers: [{ ...takeover, label: String(takeover.label) }] },
      { ...manifest, takeovers: [takeover, takeover] },
      { ...manifest, takeovers: [{ ...takeover, taken: [] }] },
      {
        ...manifest,
        takeovers: [{ ...takeover, taken: [{ ...takeover.taken[0], suggestion: 1 }] }],
      },
      {
        ...manifest,
        takeovers: [{ ...takeover, taken: [{ ...takeover.taken[0], properties: '<w:pPr/>' }] }],
      },
      {
        ...manifest,
        takeovers: [{ ...takeover, taken: [{ ...takeover.taken[0], properties: '<!DOCTYPE x>' }] }],
      },
      { ...manifest, suggestions: [{ ...manifest.suggestions[0], person: { id: 'x' } }] },
      { ...manifest, suggestions: [{ ...manifest.suggestions[0], author: 'Other' }] },
      {
        ...manifest,
        suggestions: [
          {
            ...manifest.suggestions[0],
            attributes: [{ namespace: 'urn:x', name: 'a b', value: '1' }],
          },
        ],
      },
      { ...manifest, suggestions: [] },
    ];
    for (const changed of cases) {
      entries[MANIFEST] = new TextEncoder().encode(JSON.stringify(changed));
      expect(() => openOrderedCheckpoint(zipSync(entries), { file: 1 })).toThrow(
        'checkpoint-manifest'
      );
    }
  });

  test('exports write the Word file only: the checkpoint’s Word parts are the same but for its marks', async () => {
    const { room, ana, olga } = await people();
    await steps(room, [ana, toggle('bold')], [olga, toggle('italic')]);
    const { base } = openOrderedRoomFile(SOURCE, { file: FILE });
    const exported = unzipSync(orderedPackageBytes(room.room.document, base));
    const checkpoint = unzipSync(await checkpointOf(room));
    expect(Object.keys(checkpoint)).toEqual([...Object.keys(exported), MANIFEST]);

    // The checkpoint's own attributes (`oc:label`, `oc:typer`) and their declaration taken off
    const unmarked = (bytes: Uint8Array) =>
      strFromU8(bytes).replace(/ (?:xmlns:)?oc(?::[a-z]+)?="[^"]*"/g, '');
    for (const [entry, bytes] of Object.entries(exported)) {
      const same = unmarked(checkpoint[entry] ?? new Uint8Array()) === strFromU8(bytes);
      expect({ entry, same }).toEqual({ entry, same: true });
    }
  });
});
