// What one key press costs the editors' layout in an ordered room (task 6.8), as counts, not
// time: a multi-page document, a party typing in suggesting mode and the organiser typing
// directly, each watched on their own editor and on the other one. Each press has four
// moments: the editor's own commit, the publish of the list's version after it, the room's
// confirmation (with the room's date), and the other editor hearing of it. Frames from the
// room wait per person, so each moment's work is counted alone.
//
// Counted: what the editor's store told layout (impact, and the paragraphs it names), the
// paragraphs broken into lines again, the blocks prepared again, and the passes of the live
// layout. A press reaches the paragraphs it touches and no other, and lays the typist's pages
// out once: the publish of the list's version lands before the commit's layout, not in a
// second pass after it, and a host that reads the editor on each change, inside the commit,
// reads the pages on screen until the list's version has landed. The confirmation publishes nothing when the room's date is the one the
// copy chose, and one paragraph when the room's second turned in between.

import { afterEach, describe, expect, test } from 'bun:test';
import { liveLayoutPassTestRecorder } from '../../../../editor/surface-layout-passes.ts';
import { paragraphBreakTestRecorder } from '../../../../layout/paragraph-break-request.ts';
import { preparedBlockTestRecorder } from '../../../../layout/section-preparation.ts';
import type { TreeModelChange } from '../../../../store/store/tree-store.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  HeldTransport,
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
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

const LONG =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
  'ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ' +
  'ullamco laboris nisi ut aliquip ex ea commodo consequat.';
const PARAGRAPHS = 150;
const FILE = docx(Array.from({ length: PARAGRAPHS }, (_, k) => p(t(`${k} ${LONG}`))).join(''));

/** The room's clock, which the test moves; each person's runs hours behind it. */
let clock = Date.UTC(2026, 9, 28, 9, 0, 0, 100);
const SKEW = (2 * 3600 + 17 * 60 + 3) * 1000 + 400;

interface Seat {
  readonly person: Person;
  readonly link: HeldTransport;
}

async function seat(
  room: OrderedLocalRoom,
  name: string,
  login: typeof ORGANISER,
  mode: 'edit' | 'suggesting'
): Promise<Seat> {
  const link = new HeldTransport(room.transport({ person: testPerson(name), ...login }));
  const person = await join(room, name, {
    mode,
    login,
    wordFile: FILE,
    transport: link,
    now: () => clock - SKEW,
  });
  link.held = true;
  // A host reading the editor on each change, as a React store hook does
  person.editor.on('change', () => person.editor.snapshot());
  return { person, link };
}

/** One moment's work on one editor. */
interface Work {
  /** The changes the editor's store published: impact, and how many paragraphs each names. */
  readonly changes: readonly { readonly impact: string; readonly paragraphs: number }[];
  readonly breaks: number;
  readonly prepared: number;
  /** Passes of the live layout, on either editor. */
  readonly passes: number;
}

const breaks = paragraphBreakTestRecorder();
const prepared = preparedBlockTestRecorder();
const passes = liveLayoutPassTestRecorder();

/** Start counting what `target`'s editor does; the returned function stops and reads. */
function counting(target: Seat): () => Work {
  const changes: { impact: string; paragraphs: number }[] = [];
  const stop = surface(target.person).session.subscribe((change: TreeModelChange) => {
    changes.push({
      impact: change.impact,
      paragraphs: change.dirty.length + change.created.length + change.deleted.length,
    });
  });
  breaks.reset();
  prepared.reset();
  passes.reset();
  return () => {
    stop();
    return { changes, breaks: breaks.breaks, prepared: prepared.prepared, passes: passes.passes };
  };
}

/** Count what `run` makes `target`'s editor do, every queued task included. */
async function watch(target: Seat, run: () => void | Promise<void>): Promise<Work> {
  const done = counting(target);
  await run();
  return done();
}

/** One press, counted at its four moments. */
interface Press {
  readonly commit: Work;
  readonly publish: Work;
  readonly confirm: Work;
  readonly remote: Work;
}

async function press(typist: Seat, other: Seat, act: (person: Person) => void): Promise<Press> {
  // The commit lays itself out at once; the publish runs in a microtask after it
  const read = counting(typist);
  act(typist.person);
  const commit = read();
  const publish = await watch(typist, () => settle(2));
  const confirm = await watch(typist, async () => {
    typist.link.release();
    await settle(2);
  });
  const remote = await watch(other, async () => {
    other.link.release();
    await settle(2);
  });
  return { commit, publish, confirm, remote };
}

/** Work that names only the paragraphs a press touched: never the document. */
function local(work: Work, most: number): void {
  for (const change of work.changes) {
    expect(change.impact).not.toBe('global');
    expect(change.paragraphs).toBeGreaterThan(0);
    expect(change.paragraphs).toBeLessThanOrEqual(most);
  }
  expect(work.breaks).toBeLessThanOrEqual(most);
  expect(work.prepared).toBeLessThanOrEqual(most);
}

const nothing = (work: Work): void =>
  expect({ changes: work.changes.length, breaks: work.breaks, passes: work.passes }).toEqual({
    changes: 0,
    breaks: 0,
    passes: 0,
  });

function converged(room: OrderedLocalRoom, ...seats: Seat[]): void {
  for (const { person } of seats) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
  }
}

/** The presses: typing, more typing, Backspace, Enter, bold, and Backspace at a start. */
const PRESSES: readonly { readonly name: string; readonly most: number; act(p: Person): void }[] = [
  {
    name: 'type',
    most: 1,
    act: (person) => {
      caretAt(person, 60, 10);
      surface(person).insertPlainText('x');
    },
  },
  { name: 'type again', most: 1, act: (person) => surface(person).insertPlainText('y') },
  // A key in a browser: queued, and committed by the editor's own timer
  { name: 'type a queued key', most: 1, act: (person) => surface(person).enqueueType('q') },
  {
    name: 'Backspace',
    most: 1,
    act: (person) => {
      caretAt(person, 70, 10);
      surface(person).deleteBackward();
    },
  },
  {
    name: 'Enter',
    most: 3,
    act: (person) => {
      caretAt(person, 80, 10);
      surface(person).splitParagraph();
    },
  },
  {
    name: 'bold',
    most: 1,
    act: (person) => {
      caretAt(person, 90, 2, 8);
      surface(person).toggleRunProperty('b');
    },
  },
  {
    name: 'Backspace at a start',
    most: 3,
    act: (person) => {
      caretAt(person, 100, 0);
      surface(person).deleteBackward();
    },
  },
];

describe('one key press lays out only the paragraphs it touches', () => {
  for (const typistIsParty of [true, false]) {
    const who = typistIsParty ? 'a party suggesting' : 'the organiser editing';
    test(`${who}, on their editor and on the other one`, async () => {
      clock = Date.UTC(2026, 9, 28, 9, 0, 0, 100);
      const room = localRoom(FILE, { now: () => clock });
      const party = await seat(room, 'Ana', PARTY, 'suggesting');
      const organiser = await seat(room, 'Org', ORGANISER, 'edit');
      const [typist, other] = typistIsParty ? [party, organiser] : [organiser, party];
      expect(surface(typist.person).state().pageCount).toBeGreaterThan(10);
      // A first press teaches the copy the room's clock
      await press(typist, other, (person) => {
        caretAt(person, 5, 3);
        surface(person).insertPlainText('w');
      });
      for (const step of PRESSES) {
        clock += 150;
        const counted = await press(typist, other, step.act);
        expect({ step: step.name, refusals: typist.person.refusals }).toEqual({
          step: step.name,
          refusals: [],
        });
        local(counted.commit, step.most);
        local(counted.publish, step.most);
        // The editor's commit and the list's version of it, at most
        expect(counted.commit.changes.length + counted.publish.changes.length).toBeLessThanOrEqual(
          2
        );
        // One pass for the key: the commit's, with the list's version already in it
        expect({ step: step.name, passes: counted.commit.passes + counted.publish.passes }).toEqual(
          { step: step.name, passes: 1 }
        );
        // The room's date is the one the copy chose: its confirmation draws nothing
        nothing(counted.confirm);
        local(counted.remote, step.most);
        expect(counted.remote.changes).toHaveLength(1);
        expect({ step: step.name, passes: counted.remote.passes }).toEqual({
          step: step.name,
          passes: 1,
        });
      }
      converged(room, party, organiser);
    }, 60_000);
  }

  test('a confirmation with a date the copy did not foresee draws one paragraph again', async () => {
    clock = Date.UTC(2026, 9, 28, 9, 0, 0, 100);
    const room = localRoom(FILE, { now: () => clock });
    const party = await seat(room, 'Ana', PARTY, 'suggesting');
    const organiser = await seat(room, 'Org', ORGANISER, 'edit');
    await press(party, organiser, (person) => {
      caretAt(person, 5, 3);
      surface(person).insertPlainText('w');
    });
    const read = counting(party);
    caretAt(party.person, 40, 4);
    surface(party.person).insertPlainText('z');
    const commit = read();
    // A slow trip: the room stamps the change a second and a half later
    clock += 1500;
    const publish = await watch(party, () => settle(2));
    const confirm = await watch(party, async () => {
      party.link.release();
      await settle(2);
    });
    organiser.link.release();
    await settle(2);
    local(commit, 1);
    local(publish, 1);
    expect(commit.passes + publish.passes).toBe(1);
    expect(confirm.changes).toEqual([{ impact: 'text-local', paragraphs: 1 }]);
    expect(confirm.breaks).toBe(1);
    expect(confirm.passes).toBe(1);
    converged(room, party, organiser);
  }, 60_000);
});
