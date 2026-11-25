// Content control edits with real editors and one ordered room (task 9.6): the organiser's new
// controls (around words, at a caret, a custom node), removed controls (keeping their content or
// not) and changed properties, each the same on every copy and on the screen as the editor alone
// writes it, and each taken back and made again exactly by the organiser's undo and redo. A
// party's are refused by the party rule the host's earlier check applies (a direct edit), except a
// new control around words: from the toolbar in suggesting mode, as from the editing API, it is a
// suggestion (the words struck and a copy in the new control) the organiser decides.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { stubReviewModule } from '../../../../editor/__tests__/review-test-module.ts';
import { serializeOoxmlPart, type OoxmlNode } from '../../../../store/package/ooxml-tree.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { checkSaveAndLoad } from '../sessions/session.ts';
import { docx, p, packageOf, t } from '../word-fixtures.ts';
import { reviewTestModule } from './review-module.ts';
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

afterEach(closeAll);

const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

/** A paragraph of words, one with a plain text control, one more. */
const FILE = docx(
  p(t('Alpha beta gamma.')) +
    p(
      t('Name: ') +
        sdt('<w:alias w:val="Name"/><w:tag w:val="name"/><w:id w:val="7"/><w:text/>', t('Acme')) +
        t('.')
    ) +
    p(t('Signed here.'))
);

/** The editor's body XML without paragraph ids and revision dates. */
const bodyOf = (xml: string): string =>
  xml
    .slice(xml.indexOf('<w:body>'), xml.indexOf('</w:body>'))
    .replace(/ w14:(paraId|textId)="[^"]*"/g, '')
    .replace(/<w:id w:val="\d+"\/>/g, '<w:id/>');

function converged(room: ReturnType<typeof localRoom>, ...people: Person[]): void {
  for (const person of people) {
    expect({ person: person.name, refusals: person.refusals }).toEqual({
      person: person.name,
      refusals: [],
    });
    expect({ person: person.name, room: sameAsRoom(person, room) }).toEqual({
      person: person.name,
      room: true,
    });
    expect({ person: person.name, screen: screenFollowsList(person) }).toEqual({
      person: person.name,
      screen: true,
    });
  }
}

/** The surface's gated write, as the editing API's ops take it. */
function apply(person: Person, op: TreeDocOp): boolean {
  const s = surface(person) as unknown as {
    applyGatedTreeOps(ops: readonly TreeDocOp[]): { committed: boolean; rejected?: boolean };
  };
  const result = s.applyGatedTreeOps([op]);
  return result.committed && result.rejected !== true;
}

function controlId(person: Person, index = 0): string {
  const part = surface(person).session.part();
  const ids: string[] = [];
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'contentControl') ids.push(node.id);
    for (const child of node.children) visit(child);
  };
  visit(part.root);
  const id = ids[index];
  if (id === undefined) throw new Error(`no control ${index}`);
  return id;
}

/**
 * The same edit on the editor alone (no room), in edit mode: its body, as `bodyOf` reads it.
 * `edit` gets the editor's surface.
 */
function alone(edit: (person: Person) => void): string {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: FILE,
    author: 'Olive',
    mode: 'edit',
    modules: [stubReviewModule()],
  });
  try {
    const person = { name: 'alone', editor, container } as unknown as Person;
    edit(person);
    return bodyOf(editorXml(person));
  } finally {
    editor.destroy();
    container.remove();
  }
}

async function organiserRoom() {
  const room = localRoom(FILE);
  const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
  const ben = await join(room, 'Ben', { wordFile: FILE });
  return { room, olive, ben };
}

/** The organiser's undo, then redo: back to `before`, then to `after`, on every copy. */
async function undoRedo(
  room: ReturnType<typeof localRoom>,
  olive: Person,
  ben: Person,
  before: string,
  after: string
): Promise<void> {
  surface(olive).undo();
  await settle();
  converged(room, olive, ben);
  expect(bodyOf(editorXml(olive))).toBe(before);
  surface(olive).redo();
  await settle();
  converged(room, olive, ben);
  expect(bodyOf(editorXml(olive))).toBe(after);
}

describe("the organiser's control edits", () => {
  test('a new control around words, then undo and redo', async () => {
    const { room, olive, ben } = await organiserRoom();
    const before = bodyOf(editorXml(olive));
    caretAt(olive, 0, 6, 10);
    expect(
      olive.editor.exec({ type: 'insertContentControl', subtype: 'richText' } as never)
    ).toMatchObject({ ok: true });
    await settle();
    converged(room, olive, ben);
    const after = bodyOf(editorXml(olive));
    expect(after).toContain('<w:sdtContent><w:r><w:t>beta</w:t></w:r></w:sdtContent>');
    expect(after).toBe(
      alone((person) => {
        caretAt(person, 0, 6, 10);
        person.editor.exec({ type: 'insertContentControl', subtype: 'richText' } as never);
      })
    );
    expect(bodyOf(editorXml(ben))).toBe(after);
    await undoRedo(room, olive, ben, before, after);
  });

  test('a new control at a caret shows its prompt, then undo and redo', async () => {
    const { room, olive, ben } = await organiserRoom();
    const before = bodyOf(editorXml(olive));
    caretAt(olive, 0, 6);
    expect(
      olive.editor.exec({ type: 'insertContentControl', subtype: 'plainText' } as never)
    ).toMatchObject({ ok: true });
    await settle();
    converged(room, olive, ben);
    const after = bodyOf(editorXml(olive));
    expect(after).toContain('<w:showingPlcHdr/>');
    expect(after).toBe(
      alone((person) => {
        caretAt(person, 0, 6);
        person.editor.exec({ type: 'insertContentControl', subtype: 'plainText' } as never);
      })
    );
    expect(editorTexts(ben)[0]).toBe('Alpha Click or tap here to enter text.beta gamma.');
    await undoRedo(room, olive, ben, before, after);
  });

  test('a custom node holds its text', async () => {
    const { room, olive, ben } = await organiserRoom();
    const [first = ''] = surface(olive).session.paragraphIds();
    expect(
      apply(olive, {
        op: 'insertInlineContentControl',
        paragraphId: first,
        offset: 6,
        tag: 'cite:1',
        text: '[1]',
        lock: 'sdtContentLocked',
      })
    ).toBe(true);
    await settle();
    converged(room, olive, ben);
    expect(editorTexts(ben)[0]).toBe('Alpha [1]beta gamma.');
    expect(bodyOf(editorXml(ben))).toContain('<w:tag w:val="cite:1"/>');
    expect(bodyOf(editorXml(ben))).toBe(
      alone((person) => {
        const [id = ''] = surface(person).session.paragraphIds();
        apply(person, {
          op: 'insertInlineContentControl',
          paragraphId: id,
          offset: 6,
          tag: 'cite:1',
          text: '[1]',
          lock: 'sdtContentLocked',
        });
      })
    );
  });

  test('a removed control keeps its words, then undo and redo', async () => {
    const { room, olive, ben } = await organiserRoom();
    const before = bodyOf(editorXml(olive));
    caretAt(olive, 1, 7);
    expect(olive.editor.exec({ type: 'removeContentControl' } as never)).toMatchObject({
      ok: true,
    });
    await settle();
    converged(room, olive, ben);
    const after = bodyOf(editorXml(olive));
    expect(after).not.toContain('<w:sdt>');
    expect(editorTexts(ben)[1]).toBe('Name: Acme.');
    expect(after).toBe(
      alone((person) => {
        caretAt(person, 1, 7);
        person.editor.exec({ type: 'removeContentControl' } as never);
      })
    );
    await undoRedo(room, olive, ben, before, after);
  });

  test('a control removed with its words, then undo and redo', async () => {
    const { room, olive, ben } = await organiserRoom();
    const before = bodyOf(editorXml(olive));
    expect(
      apply(olive, { op: 'removeContentControl', controlId: controlId(olive), keepContent: false })
    ).toBe(true);
    await settle();
    converged(room, olive, ben);
    expect(editorTexts(ben)[1]).toBe('Name: .');
    expect(bodyOf(editorXml(olive))).toBe(
      alone((person) =>
        apply(person, {
          op: 'removeContentControl',
          controlId: controlId(person),
          keepContent: false,
        })
      )
    );
    await undoRedo(room, olive, ben, before, bodyOf(editorXml(olive)));
  });

  test("a control's tag, title and lock, then undo and redo", async () => {
    const { room, olive, ben } = await organiserRoom();
    const before = bodyOf(editorXml(olive));
    expect(
      apply(olive, {
        op: 'setContentControlProperties',
        controlId: controlId(olive),
        tag: 'party',
        alias: null,
        lock: 'sdtLocked',
      })
    ).toBe(true);
    await settle();
    converged(room, olive, ben);
    const after = bodyOf(editorXml(olive));
    expect(after).toContain('<w:tag w:val="party"/>');
    expect(after).toContain('<w:lock w:val="sdtLocked"/>');
    expect(after).not.toContain('w:alias');
    expect(after).toBe(
      alone((person) =>
        apply(person, {
          op: 'setContentControlProperties',
          controlId: controlId(person),
          tag: 'party',
          alias: null,
          lock: 'sdtLocked',
        })
      )
    );
    await undoRedo(room, olive, ben, before, after);
  });

  test("the editors' save reopens as the room's document", async () => {
    const { room, olive } = await organiserRoom();
    caretAt(olive, 0, 6, 10);
    olive.editor.exec({ type: 'insertContentControl', subtype: 'richText' } as never);
    await settle();
    apply(olive, { op: 'setContentControlProperties', controlId: controlId(olive, 1), tag: 'x' });
    await settle();
    const expected = serializeOoxmlPart(composeOrderedPart(room.room.document, BODY).part);
    const saved = new Uint8Array(await olive.editor.save());
    expect(readOoxmlPackage(saved).ok).toBe(true);
    const again = openOrderedWordFile(saved, { file: 2 }).document;
    const reopened = serializeOoxmlPart(composeOrderedPart(again, BODY).part);
    expect(bodyOf(reopened)).toBe(bodyOf(expected));
  });
});

describe("a party's control edits", () => {
  async function partyRoom() {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    const ana = await join(room, 'Ana', { wordFile: FILE });
    return { room, olive, ana };
  }

  /** A party's new control around "beta" from the toolbar, in suggesting mode. */
  async function partyWrap(review = reviewTestModule()) {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      review,
      wordFile: FILE,
    });
    const ana = await join(room, 'Ana', { review, wordFile: FILE });
    const before = bodyOf(editorXml(olive));
    caretAt(ana, 0, 6, 10);
    expect(
      ana.editor.exec({ type: 'insertContentControl', subtype: 'plainText', tag: 'term' } as never)
    ).toEqual({ ok: true, changed: true });
    await settle();
    converged(room, olive, ana);
    return { room, olive, ana, before, after: bodyOf(editorXml(olive)) };
  }

  const revisionFree = (body: string) => body.replace(/ w:(date|id)="[^"]*"/g, '');

  test("a party's control from the toolbar is a suggestion, as the editor alone writes it", async () => {
    const { room, olive, ana, before, after } = await partyWrap();
    expect(room.room.order).toBe(1);
    // The words struck and a copy of them in the new control, the editing API's form
    expect(after).toContain('<w:delText>beta</w:delText>');
    expect(after).toMatch(/<w:ins [^>]*w:author="Ana"[^>]*><w:sdt>/);
    expect(bodyOf(editorXml(ana))).toBe(after);
    // The editor alone, suggesting as Ana, writes the same
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({
      container,
      document: FILE,
      author: 'Ana',
      mode: 'suggesting',
      modules: [stubReviewModule()],
    });
    try {
      const person = { name: 'alone', editor, container } as unknown as Person;
      caretAt(person, 0, 6, 10);
      editor.exec({ type: 'insertContentControl', subtype: 'plainText', tag: 'term' });
      expect(revisionFree(bodyOf(editorXml(person)))).toBe(revisionFree(after));
    } finally {
      editor.destroy();
      container.remove();
    }
    // Her undo takes it back on every copy, and redo makes it again
    surface(ana).undo();
    await settle();
    converged(room, olive, ana);
    expect(bodyOf(editorXml(olive))).toBe(before);
    surface(ana).redo();
    await settle();
    converged(room, olive, ana);
    expect(revisionFree(bodyOf(editorXml(olive)))).toBe(revisionFree(after));
  });

  for (const decision of ['accept', 'reject'] as const)
    test(`the organiser's ${decision} of a party's toolbar control holds on every copy and in the file`, async () => {
      const { room, olive, ana, before } = await partyWrap();
      const keys = olive.editor
        .getReviewItems()
        .filter((placement) => placement.kind === 'revision')
        .map((placement) => placement.key);
      expect(keys.length).toBeGreaterThan(0);
      for (const key of keys)
        expect(
          (decision === 'accept'
            ? olive.editor.acceptReviewItem(key)
            : olive.editor.rejectReviewItem(key)
          ).ok
        ).toBe(true);
      await settle();
      converged(room, olive, ana);
      const decided = bodyOf(editorXml(olive));
      expect(decided).not.toMatch(/<w:(ins|del) /);
      expect(editorTexts(ana)[0]).toBe('Alpha beta gamma.');
      if (decision === 'reject') expect(decided).toBe(before);
      else expect(decided).toContain('<w:tag w:val="term"/>');
      checkSaveAndLoad(room.room.document, `toolbar control ${decision}`, packageOf(FILE));
    });

  test('a tracked control from the editing API is a suggestion; its author takes it back', async () => {
    const { room, olive, ana } = await partyRoom();
    const before = bodyOf(editorXml(olive));
    const [first = ''] = surface(ana).session.paragraphIds();
    expect(
      apply(ana, {
        op: 'insertContentControl',
        paragraphId: first,
        start: 6,
        end: 10,
        type: 'richText',
        tag: 'term',
        revision: { author: 'Ana', date: '2026-10-28T08:00:00Z' },
      })
    ).toBe(true);
    await settle();
    converged(room, olive, ana);
    const xml = bodyOf(editorXml(olive));
    expect(xml).toContain('<w:delText>beta</w:delText>');
    expect(xml).toMatch(/<w:ins [^>]*><w:sdt>/);
    // The party's undo takes the suggestion back on every copy, and redo makes it again (a new
    // change: the room's later date and new revision ids)
    const revisionFree = (body: string) => body.replace(/ w:(date|id)="[^"]*"/g, '');
    surface(ana).undo();
    await settle();
    converged(room, olive, ana);
    expect(bodyOf(editorXml(olive))).toBe(before);
    surface(ana).redo();
    await settle();
    converged(room, olive, ana);
    expect(revisionFree(bodyOf(editorXml(olive)))).toBe(revisionFree(xml));
  });
});

describe('control edits at the same time as other edits', () => {
  test('the organiser removes a control with its words while a party types in it', async () => {
    const room = localRoom(FILE);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    const ana = await join(room, 'Ana', { wordFile: FILE });
    caretAt(ana, 1, 8);
    // Both go out before either copy hears of the other
    surface(ana).insertPlainText('ZZ');
    apply(olive, { op: 'removeContentControl', controlId: controlId(olive), keepContent: false });
    await settle();
    converged(room, olive, ana);
    // A direct removal takes what the control holds when the room applies it, the party's
    // pending words too, as the engine's removal does (a direct deletion, task 6.7)
    expect(editorTexts(olive)[1]).toBe('Name: .');
    expect(bodyOf(editorXml(olive))).not.toContain('<w:sdt>');
    // The organiser's undo shows them again as they were: the party's insertion, pending
    surface(olive).undo();
    await settle();
    converged(room, olive, ana);
    expect(editorTexts(olive)[1]).toBe('Name: AcZZme.');
    expect(bodyOf(editorXml(olive))).toMatch(/<w:ins w:author="Ana"[^>]*><w:r><w:t>ZZ<\/w:t>/);
  });

  test('the organiser removes a control, keeping its words, while a party types in it', async () => {
    const room = localRoom(FILE);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    const ana = await join(room, 'Ana', { wordFile: FILE });
    caretAt(ana, 1, 8);
    surface(ana).insertPlainText('ZZ');
    caretAt(olive, 1, 7);
    olive.editor.exec({ type: 'removeContentControl' } as never);
    await settle();
    converged(room, olive, ana);
    expect(editorTexts(ana)[1]).toBe('Name: AcZZme.');
    expect(bodyOf(editorXml(ana))).not.toContain('<w:sdt>');
  });

  test('two organisers change one control: each property holds, the last tag wins', async () => {
    const room = localRoom(FILE);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    const oscar = await join(room, 'Oscar', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    apply(olive, { op: 'setContentControlProperties', controlId: controlId(olive), tag: 'olive' });
    apply(oscar, {
      op: 'setContentControlProperties',
      controlId: controlId(oscar),
      alias: 'Oscar',
    });
    await settle();
    converged(room, olive, oscar);
    const both = bodyOf(editorXml(olive));
    expect(both).toContain('<w:tag w:val="olive"/>');
    expect(both).toContain('<w:alias w:val="Oscar"/>');
    // Olive's undo takes back her tag only
    surface(olive).undo();
    await settle();
    converged(room, olive, oscar);
    expect(bodyOf(editorXml(oscar))).toContain('<w:tag w:val="name"/>');
    expect(bodyOf(editorXml(oscar))).toContain('<w:alias w:val="Oscar"/>');
    apply(olive, { op: 'setContentControlProperties', controlId: controlId(olive), tag: 'a' });
    apply(oscar, { op: 'setContentControlProperties', controlId: controlId(oscar), tag: 'b' });
    await settle();
    converged(room, olive, oscar);
    expect(bodyOf(editorXml(olive))).toMatch(/<w:tag w:val="[ab]"\/>/);
  });
});

describe("a party's control edits, by the host's party rules", () => {
  test("each direct control edit is refused with the reason today's check gives", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: FILE });
    const [first = ''] = surface(ana).session.paragraphIds();
    const cases: readonly [TreeDocOp, string][] = [
      [{ op: 'removeContentControl', controlId: controlId(ana) }, 'package'],
      [{ op: 'removeContentControl', controlId: controlId(ana), keepContent: false }, 'text'],
      [{ op: 'setContentControlProperties', controlId: controlId(ana), tag: 'x' }, 'package'],
      [
        { op: 'insertInlineContentControl', paragraphId: first, offset: 0, tag: 't', text: 'x' },
        'text',
      ],
      [
        {
          op: 'insertContentControl',
          paragraphId: first,
          start: 0,
          end: 0,
          type: 'richText',
        },
        'text',
      ],
    ];
    for (const [op, reason] of cases) {
      ana.refusals.length = 0;
      apply(ana, op);
      await settle();
      expect([op.op, ana.refusals.map((r) => r.reason)]).toEqual([op.op, [reason]]);
      expect(sameAsRoom(ana, room)).toBe(true);
      expect(screenFollowsList(ana)).toBe(true);
    }
    expect(room.room.order).toBe(0);
  });
});

describe("a control's value and its properties at once", () => {
  const DROP = docx(
    p(
      t('Pick: ') +
        sdt(
          '<w:tag w:val="pick"/><w:id w:val="9"/><w:dropDownList w:lastValue="1">' +
            '<w:listItem w:displayText="One" w:value="1"/><w:listItem w:displayText="Two" w:value="2"/>' +
            '</w:dropDownList>',
          t('One')
        ) +
        t('.')
    )
  );

  test("a party's pending value and the organiser's tag both hold; a rejected value leaves the tag", async () => {
    const review = reviewTestModule();
    const room = localRoom(DROP);
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      wordFile: DROP,
      review,
    });
    const ana = await join(room, 'Ana', { wordFile: DROP, review });
    caretAt(ana, 0, 7);
    expect(ana.editor.exec({ type: 'setContentControlValue', value: '2' } as never)).toMatchObject({
      ok: true,
    });
    await settle();
    apply(olive, { op: 'setContentControlProperties', controlId: controlId(olive), tag: 'chosen' });
    await settle();
    converged(room, olive, ana);
    const both = bodyOf(editorXml(olive));
    expect(both).toContain('<w:tag w:val="chosen"/>');
    expect(both).toContain('w:lastValue="2"');
    const [card] = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
    if (card === undefined) throw new Error('no card');
    expect(olive.editor.rejectReviewItem(card.key).ok).toBe(true);
    await settle();
    converged(room, olive, ana);
    const after = bodyOf(editorXml(ana));
    expect(after).toContain('<w:tag w:val="chosen"/>');
    expect(after).toContain('w:lastValue="1"');
    expect(editorTexts(ana)[0]).toBe('Pick: One.');
  });
});

// The engine writes and reads Gregorian dates only: a date pick in a control in another calendar
// is refused by name, for a party and for the organiser, and nothing reaches the room
describe('a date control in another calendar', () => {
  const HIJRI = docx(
    p(
      t('Due: ') +
        sdt(
          '<w:id w:val="11"/><w:date><w:dateFormat w:val="dd/MM/yyyy"/>' +
            '<w:calendar w:val="hijri"/></w:date>',
          t('Soon')
        ) +
        t('.')
    )
  );

  test('refuses a date pick in the editor and in the room', async () => {
    const room = localRoom(HIJRI);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: HIJRI });
    const ana = await join(room, 'Ana', { wordFile: HIJRI });
    for (const person of [ana, olive]) {
      caretAt(person, 0, 6);
      const picked = person.editor.exec({
        type: 'setContentControlValue',
        value: '2026-09-15',
      } as never);
      expect({ person: person.name, picked }).toMatchObject({
        person: person.name,
        picked: { ok: false },
      });
      await settle();
      expect(person.refusals.splice(0)).toEqual([
        {
          code: 'ordered-refused',
          reason: 'engine-refused',
          detail: 'setContentControlValue: unsupported',
        },
      ]);
      converged(room, olive, ana);
      expect(editorTexts(person)[0]).toBe('Due: Soon.');
      expect(bodyOf(editorXml(person))).not.toContain('w:fullDate');
    }
  });
});
