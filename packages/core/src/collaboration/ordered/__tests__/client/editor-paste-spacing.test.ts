// Word's Smart cut and paste spacing in an ordered room, beside a sentence end, a comma and
// inside Thai (Word 16.114 for Mac, W4 evidence `~$paste-out-w4-<case>.docx`): "fee" after
// "Stop." gets a space on each side, before a comma one before only; inside Thai, "fee" gets
// a space on each side and a Thai word one before only. The spaces travel inside the paste's
// change. The organiser's paste is direct; a party's is one insertion that holds the spaces.
// Every copy equals the room, and the room equals the editor alone.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import type { OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { grade, normalizedBody } from '../matrix/compare.ts';
import { checkSaveAndLoad } from '../sessions/session.ts';
import { docx, p, packageOf, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { copiedBetween, type Copied } from './paste-harness.ts';
import { reviewTestModule } from './review-module.ts';

const review = reviewTestModule();

afterEach(closeAll);

type Room = ReturnType<typeof localRoom>;

/** The targets of W4 samples a, b and d (and g). */
const FILE = docx(p(t('Stop.Go now.')) + p(t('Go, then stop.')) + p(t('สวัสดีครับ')));

const BEFORE = ['Stop.Go now.', 'Go, then stop.', 'สวัสดีครับ'];

/** Word's sources: bold "fee", and bold Thai "ราคา", each followed by more text. */
const SOURCE = docx(p(t('fee', '<w:b/>') + t(' more')) + p(t('ราคา', '<w:b/>') + t(' more')));

const fee = (): Copied => copiedBetween(0, 0, 'fee'.length, SOURCE);
const price = (): Copied => copiedBetween(1, 1, 'ราคา'.length, SOURCE);

const kinds = (room: Room): string[] => room.room.joinSnapshot(0).map((change) => change.kind);

const checked = (room: Room, people: readonly Person[]) => {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
  }
};

const comparable = (part: OoxmlPart): string => normalizedBody(part, () => null);

/** What the editor alone writes for the same paste at the same place. */
function alone(flavours: Copied, party: boolean, paragraph: number, offset: number): string {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: FILE,
    author: party ? 'Ana' : 'Olive',
    mode: party ? 'suggesting' : 'edit',
    modules: [review],
  });
  try {
    const s = editor.surface;
    if (!s) throw new Error('no surface');
    const id = s.session.paragraphIds()[paragraph] ?? '';
    s.setSelection({ anchor: { paragraphId: id, offset }, head: { paragraphId: id, offset } });
    expect(s.pasteRich(flavours.text, flavours.html)).toBe(true);
    return comparable(s.session.part());
  } finally {
    editor.destroy();
    container.remove();
  }
}

const agrees = (room: Room, engine: string, party: boolean): string => {
  const g = grade(comparable(composeOrderedPart(room.room.document, BODY).part), engine);
  return g === 'exact' || g === 'same' || (party && g === 'ids') ? 'agrees' : g;
};

async function people(room: Room) {
  const olive = await join(room, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    review,
    wordFile: FILE,
  });
  const ana = await join(room, 'Ana', { review, wordFile: FILE });
  return { olive, ana };
}

/** The text of each tracked insertion of Ana's in a copy. */
const anasInsertions = (person: Person): string[] =>
  [...editorXml(person).matchAll(/<w:ins [^>]*w:author="Ana"[^>]*>(.*?)<\/w:ins>/gs)].map((m) =>
    [...m[1]!.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((x) => x[1]).join('')
  );

interface Case {
  readonly sample: string;
  readonly copy: () => Copied;
  readonly paragraph: number;
  readonly offset: number;
  /** The paragraph's text after the paste, as Word writes it. */
  readonly text: string;
  /** The pasted text with Word's spaces. */
  readonly inserted: string;
}

const CASES: readonly Case[] = [
  {
    sample: 'a, after a sentence end',
    copy: fee,
    paragraph: 0,
    offset: 'Stop.'.length,
    text: 'Stop. fee Go now.',
    inserted: ' fee ',
  },
  {
    sample: 'b, before a comma',
    copy: fee,
    paragraph: 1,
    offset: 'Go'.length,
    text: 'Go fee, then stop.',
    inserted: ' fee',
  },
  {
    sample: 'd, Latin inside Thai',
    copy: fee,
    paragraph: 2,
    offset: 'สวัสดี'.length,
    text: 'สวัสดี fee ครับ',
    inserted: ' fee ',
  },
  {
    sample: 'g, Thai inside Thai',
    copy: price,
    paragraph: 2,
    offset: 'สวัสดี'.length,
    text: 'สวัสดี ราคาครับ',
    inserted: ' ราคา',
  },
];

const after = (item: Case): string[] =>
  BEFORE.map((line, k) => (k === item.paragraph ? item.text : line));

for (const item of CASES)
  describe(`smart paste spacing in an ordered room, W4 ${item.sample}`, () => {
    test("the organiser's: direct, as the editor alone pastes it; undo and redo", async () => {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const { olive, ana } = await people(room);
      const flavours = item.copy();
      caretAt(olive, item.paragraph, item.offset);
      expect(surface(olive).pasteRich(flavours.text, flavours.html)).toBe(true);
      await settle(16);
      expect(kinds(room)).toEqual(['insert-fragment']);
      checked(room, [olive, ana]);
      expect(editorTexts(ana)).toEqual(after(item));
      expect(editorXml(ana)).not.toContain('<w:ins ');
      expect(agrees(room, alone(flavours, false, item.paragraph, item.offset), false)).toBe(
        'agrees'
      );
      checkSaveAndLoad(room.room.document, item.sample, packageOf(FILE));
      expect(olive.handle.session.undo()).toBe(true);
      await settle(16);
      checked(room, [olive, ana]);
      expect(editorTexts(ana)).toEqual(BEFORE);
      expect(olive.handle.session.redo()).toBe(true);
      await settle(16);
      checked(room, [olive, ana]);
      expect(editorTexts(ana)).toEqual(after(item));
    });

    test("a party's: one insertion that holds the spaces, as the editor alone; a reject", async () => {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const { olive, ana } = await people(room);
      const flavours = item.copy();
      caretAt(ana, item.paragraph, item.offset);
      expect(surface(ana).pasteRich(flavours.text, flavours.html)).toBe(true);
      await settle(16);
      expect(kinds(room)).toEqual(['insert-fragment']);
      checked(room, [olive, ana]);
      expect(editorTexts(olive)).toEqual(after(item));
      // Word's tracked paste holds both spaces in its one insertion (W4 sample h)
      expect(anasInsertions(olive)).toEqual([item.inserted]);
      expect(agrees(room, alone(flavours, true, item.paragraph, item.offset), true)).toBe('agrees');
      checkSaveAndLoad(room.room.document, `party ${item.sample}`, packageOf(FILE));
      const [card] = olive.editor
        .getReviewItems()
        .filter((placement) => placement.kind === 'revision')
        .map((placement) => placement.key);
      if (card === undefined) throw new Error('no card');
      expect(olive.editor.rejectReviewItem(card).ok).toBe(true);
      await settle(16);
      checked(room, [olive, ana]);
      expect(editorTexts(olive)).toEqual(BEFORE);
    });
  });
