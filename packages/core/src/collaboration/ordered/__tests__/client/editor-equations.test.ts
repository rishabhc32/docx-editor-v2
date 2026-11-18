// Equations through real editors and one room (group 9.9): the organiser rewrites an equation
// from its linear form and removes one, with undo and redo; the other editor follows; every copy
// ends as the room, and the screen follows the list. A file's equation that Word struck whole is
// one card the organiser decides; Word's marks on some runs only are refused by name, and the
// editor's Accept All and Reject All decide every other card beside them.

import { afterEach, describe, expect, test } from 'bun:test';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import { docx, M, p, t } from '../word-fixtures.ts';
import {
  closeAll,
  editorXml,
  join,
  listXml,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const FILE = docx(
  p(t('Let ') + `<m:oMath xmlns:m="${M}"><m:r><m:t>x=1</m:t></m:r></m:oMath>` + t(' hold.')) +
    p(t('Signed here.'))
);

function checked(room: OrderedLocalRoom, ...people: Person[]): void {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(editorXml(person)).toBe(listXml(person));
    expect(sameAsRoom(person, room)).toBe(true);
  }
}

/** The equation in the first paragraph, as the editor sees it. */
function equationOf(person: Person) {
  const s = surface(person);
  const [paragraphId = ''] = s.session.paragraphIds();
  s.setSelection({ anchor: { paragraphId, offset: 4 }, head: { paragraphId, offset: 4 } });
  const equation = s.equations.equationAtCaret();
  if (!equation) throw new Error('no equation at the caret');
  return equation;
}

async function two() {
  const room = localRoom(FILE);
  const olive = await join(room, 'Olive', { wordFile: FILE, mode: 'edit', login: ORGANISER });
  const ana = await join(room, 'Ana', { wordFile: FILE });
  return { room, olive, ana };
}

describe('equations through the editor (group 9.9)', () => {
  test("the organiser's new form, its undo and its redo", async () => {
    const { room, olive, ana } = await two();
    const before = listXml(olive);
    const equation = equationOf(olive);
    expect(surface(olive).equations.applyEquation(equation.id, '{a+b}/{2}')).toBe(true);
    await settle();
    expect(room.room.joinSnapshot(0)[0]?.kind).toBe('equation');
    expect(editorXml(ana)).toContain('<m:f>');
    expect(editorXml(ana)).not.toContain('x=1');
    // The editor still names the equation it rewrote
    expect(surface(olive).equations.equationById(equation.id)?.linear).toBe('{a+b}/{2}');
    checked(room, olive, ana);
    olive.editor.exec({ type: 'undo' });
    await settle();
    expect(listXml(ana)).toBe(before);
    checked(room, olive, ana);
    olive.editor.exec({ type: 'redo' });
    await settle();
    expect(editorXml(ana)).toContain('<m:f>');
    checked(room, olive, ana);
  });

  test('the organiser removes an equation, and undo brings it back', async () => {
    const { room, olive, ana } = await two();
    const before = listXml(olive);
    const equation = equationOf(olive);
    expect(surface(olive).equations.removeEquation(equation.id)).toBe(true);
    await settle();
    expect(room.room.joinSnapshot(0)[0]?.kind).toBe('delete');
    expect(editorXml(ana)).not.toContain('m:oMath');
    checked(room, olive, ana);
    olive.editor.exec({ type: 'undo' });
    await settle();
    expect(listXml(ana)).toBe(before);
    checked(room, olive, ana);
  });

  test("Word's marks in an equation: a whole one is one card; some runs only are refused by name", async () => {
    const review = reviewTestModule();
    const stamp = 'w:author="Ana" w:date="2026-10-29T23:18:00Z"';
    const marked = (id: number, text: string, mark = 'del') =>
      `<m:r><w:${mark} w:id="${id}" ${stamp}><m:t>${text}</m:t></w:${mark}></m:r>`;
    const file = (inside: string) =>
      docx(p(t('Let ') + `<m:oMath xmlns:m="${M}">${inside}</m:oMath>` + t(' hold.')));
    const open = async (bytes: Uint8Array) => {
      const room = localRoom(bytes);
      const olive = await join(room, 'Olive', {
        wordFile: bytes,
        mode: 'edit',
        login: ORGANISER,
        review,
      });
      const cards = () => olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
      return { room, olive, cards };
    };
    // Struck whole, as Word writes it: one card, and the organiser's accept removes it
    const whole = await open(file(marked(0, 'x') + marked(1, '=1')));
    expect(whole.cards()).toHaveLength(1);
    expect(whole.olive.editor.acceptReviewItem(whole.cards()[0]!.key).ok).toBe(true);
    await settle();
    checked(whole.room, whole.olive);
    expect(editorXml(whole.olive)).not.toContain('<m:oMath');
    // "x=1" made "x=2": the list keeps the marks as the file has them, and refuses their decision
    const part = await open(
      file('<m:r><m:t>x=</m:t></m:r>' + marked(3, '1') + marked(4, '2', 'ins'))
    );
    expect(part.cards()).toHaveLength(2);
    const before = listXml(part.olive);
    part.olive.editor.acceptReviewItem(part.cards()[0]!.key);
    await settle();
    expect(part.olive.refusals.map((r) => `${r.code}:${r.reason}`)).toEqual([
      'ordered-unsupported-op:acceptRevision.equation-part',
    ]);
    expect(listXml(part.olive)).toBe(before);
    expect(editorXml(part.olive)).toBe(listXml(part.olive));
  });

  // Audit R1 row 8: one card inside part of an equation refused the editor's whole Accept All
  // and Reject All, so every other card stayed pending. Now they decide every other card in one
  // change and name the equation's marks, which stay as the file has them
  for (const action of ['accept', 'reject'] as const)
    test(`the editor's own ${action} all decides the rest beside marks in part of an equation`, async () => {
      const review = reviewTestModule();
      const stamp = 'w:author="Ana" w:date="2026-10-29T23:18:00Z"';
      const bytes = docx(
        p(
          t('Let ') +
            `<m:oMath xmlns:m="${M}"><m:r><m:t>x=</m:t></m:r>` +
            `<m:r><w:del w:id="3" ${stamp}><m:t>1</m:t></w:del></m:r>` +
            `<m:r><w:ins w:id="4" ${stamp}><m:t>2</m:t></w:ins></m:r></m:oMath>` +
            t(' hold.')
        ) + p(t('Signed ') + `<w:ins w:id="5" ${stamp}>${t('here')}</w:ins>` + t('.'))
      );
      const room = localRoom(bytes);
      const olive = await join(room, 'Olive', {
        wordFile: bytes,
        mode: 'edit',
        login: ORGANISER,
        review,
      });
      const done = olive.editor.exec({
        type: 'resolveAllReviewChanges',
        action,
        scope: 'document',
      });
      expect(done.ok).toBe(true);
      await settle();
      expect(room.room.joinSnapshot(0).map((change) => change.kind)).toEqual(['decide']);
      expect(olive.refusals.map((r) => `${r.code}:${r.reason}`)).toEqual([
        'ordered-unsupported-op:acceptRevision.equation-part',
      ]);
      expect(sameAsRoom(olive, room)).toBe(true);
      expect(editorXml(olive)).toBe(listXml(olive));
      // The ordinary card is decided; the equation's marks stay
      expect(editorXml(olive)).not.toContain('w:id="5"');
      expect(editorXml(olive)).toContain(action === 'accept' ? '>here<' : 'Signed ');
      if (action === 'reject') expect(editorXml(olive)).not.toContain('>here<');
      expect(editorXml(olive)).toContain('<w:del ');
      expect(editorXml(olive)).toContain('<w:ins ');
    });
});
