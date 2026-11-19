// Fields in an ordered room (group 9.5), with real editors and one room: a page field typed by the
// organiser and proposed by a party, a field's new instruction and result through the editing
// API's lane, and a format of a field's result. Each saves as the editor alone saves the same
// edit; the organiser's undo and redo and a party's take-back are exact; decisions and a saved
// file's reopening work; the party rules hold; and edits at the same time converge.

import { afterEach, describe, expect, test } from 'bun:test';
import { writeOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { findNode } from '../../../../store/package/ooxml-edit.ts';
import { locateFieldResults } from '../../../../store/store/tree-op-field-results.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedCards } from '../../decide.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, R, t, W, WML_CT } from '../word-fixtures.ts';
import {
  closeAll,
  FILE,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  type Editor,
} from './editor-harness.ts';
import {
  act,
  converged,
  roomFile,
  sameAsAlone,
  select,
  storyAsRead,
  type Act,
} from './links-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const fieldRuns = (code: string, result: string, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  t(result, rPr) +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

/** "Page 1 of 3": a complex PAGE field and a simple NUMPAGES field, with their results. */
const FIELDS = docx(
  p(
    t('Page ') +
      fieldRuns('PAGE', '1') +
      t(' of ') +
      '<w:fldSimple w:instr=" NUMPAGES ">' +
      t('3') +
      '</w:fldSimple>'
  ) +
    p(t('Due in 30 days.')) +
    p(t('Signed here.'))
);

type PageField = 'PAGE' | 'NUMPAGES' | 'SECTIONPAGES' | 'PAGE_X_OF_Y';

const HEADER_PART = '/word/header1.xml';

/** Two paragraphs and a header that the file has already, so opening it creates nothing. */
const WITH_HEADER = docx(
  p(t('The fee is 100 USD.')) +
    p(t('Due in 30 days.')) +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr>',
  {
    parts: {
      'word/header1.xml': `<w:hdr xmlns:w="${W}" xmlns:r="${R}">${p(t('Acme Ltd, page '))}</w:hdr>`,
    },
    contentTypes: `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>`,
    relationships: `<Relationship Id="rIdHeader" Type="${R}/header" Target="header1.xml"/>`,
  }
);

/** The organiser opens the header and types a page field at its end, with the editor's command. */
const inHeader =
  (field: PageField): Act =>
  (e) => {
    expect(e.exec({ type: 'editHeaderFooter', position: 'header' }).ok).toBe(true);
    return e.exec({ type: 'insertPageField', field });
  };

/** A page field in the body, through the editing API's lane. */
const inBody = (index: number, offset: number, field: PageField): Act =>
  write((editor) => {
    const id = editor.surface?.session.paragraphIds()[index] ?? '';
    return [{ op: 'insertPageField', paragraphId: id, offset, field }];
  });

/** The fields of the index-th body paragraph, as the editor reads them. */
function fieldsOf(editor: Editor, index: number) {
  const s = editor.surface;
  if (!s) throw new Error('no surface');
  const id = s.session.paragraphIds()[index] ?? '';
  const paragraph = findNode(s.session.part(), id);
  if (!paragraph || paragraph.kind !== 'paragraph') throw new Error(`no paragraph ${index}`);
  return { id, fields: locateFieldResults(paragraph) };
}

/** Ops through the editing API's lane, which asks the same gate as every other edit. */
const write =
  (ops: (editor: Editor) => readonly TreeDocOp[]): Act =>
  (editor) => {
    const s = editor.surface;
    if (!s) throw new Error('no surface');
    const result = s.applyAutomationOps(() => ops(editor));
    return { ...result, ok: result.committed };
  };

/** A field's new instruction, named by the editor's own node id. */
const code = (k: number, value: string) =>
  write((editor) => {
    const { id, fields } = fieldsOf(editor, 0);
    const field = fields[k];
    if (!field) throw new Error(`no field ${k}`);
    return [{ op: 'setFieldCode', paragraphId: id, fieldNodeId: field.fieldNodeId, code: value }];
  });

/** Fields' new results, as the editing API or a save-time refresh writes them. */
const results = (texts: readonly (string | null)[]) =>
  write((editor) => {
    const { id, fields } = fieldsOf(editor, 0);
    const updates = fields.flatMap((field, k) => {
      const text = texts[k];
      return text === null || text === undefined
        ? []
        : [{ paragraphId: id, fieldNodeId: field.fieldNodeId, text }];
    });
    return [{ op: 'refreshFieldResults', updates }];
  });

const bold: Act = (e) => e.exec({ type: 'toggleMark', mark: 'bold' });

describe('page fields', () => {
  test('the organiser types one in a header, as the editor alone types it', async () => {
    await sameAsAlone(WITH_HEADER, 'edit', [inHeader('PAGE')], HEADER_PART);
    await sameAsAlone(WITH_HEADER, 'edit', [inHeader('PAGE_X_OF_Y')], HEADER_PART);
  });

  test('in the body: direct for the organiser, one insertion for a party, as the editor alone', async () => {
    await sameAsAlone(FILE, 'edit', [inBody(1, 4, 'PAGE'), inBody(2, 0, 'PAGE_X_OF_Y')]);
    const { room } = await sameAsAlone(FILE, 'suggesting', [
      inBody(1, 4, 'SECTIONPAGES'),
      inBody(2, 6, 'PAGE_X_OF_Y'),
    ]);
    const story = storyAsRead(roomFile(room.room.document, FILE));
    expect(story).toMatch(/<w:ins[ >]/);
    expect(story).toContain('SECTIONPAGES');
  });

  test("the party rules: a party's page field in a header is refused", async () => {
    const room = localRoom(WITH_HEADER, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: WITH_HEADER });
    const before = storyAsRead(roomFile(room.room.document, WITH_HEADER), HEADER_PART);
    await act(ana, inHeader('PAGE'));
    expect(ana.refusals.map((r) => `${r.code}:${r.reason}`)).toEqual(['ordered-refused:furniture']);
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(storyAsRead(roomFile(room.room.document, WITH_HEADER), HEADER_PART)).toBe(before);
  });

  test("a party's proposal: accepted it stays, rejected nothing is left", async () => {
    const review = reviewTestModule();
    const original = storyAsRead(roomFile(localRoom().room.document, FILE));
    for (const decision of ['accept', 'reject'] as const) {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const ana = await join(room, 'Ana', { review });
      const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
      await act(ana, inBody(1, 4, 'PAGE_X_OF_Y'));
      const [card] = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
      if (!card) throw new Error('no card');
      const decided =
        decision === 'accept'
          ? olive.editor.acceptReviewItem(card.key)
          : olive.editor.rejectReviewItem(card.key);
      expect(decided.ok).toBe(true);
      await settle();
      converged(room, [ana, olive]);
      const story = storyAsRead(roomFile(room.room.document, FILE));
      expect(story).not.toMatch(/<w:ins[ >]/);
      if (decision === 'accept') expect(story).toContain('NUMPAGES');
      else expect(story).toBe(original);
    }
  });

  test("a party's undo takes the proposed field back, and redo proposes it again", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana');
    const original = storyAsRead(roomFile(room.room.document, FILE));
    await act(ana, inBody(1, 4, 'PAGE_X_OF_Y'));
    const proposed = storyAsRead(roomFile(room.room.document, FILE));
    await act(ana, (e) => e.surface?.undo());
    expect(storyAsRead(roomFile(room.room.document, FILE))).toBe(original);
    await act(ana, (e) => e.surface?.redo());
    expect(storyAsRead(roomFile(room.room.document, FILE))).toBe(proposed);
    converged(room, [ana]);
  });

  test("the organiser's undo and redo of a page field are exact", async () => {
    const room = localRoom(WITH_HEADER);
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      wordFile: WITH_HEADER,
    });
    const header = () => storyAsRead(roomFile(room.room.document, WITH_HEADER), HEADER_PART);
    const original = header();
    await act(olive, inHeader('PAGE_X_OF_Y'));
    const typed = header();
    expect(typed).not.toBe(original);
    await act(olive, (e) => e.surface?.undo());
    expect(header()).toBe(original);
    await act(olive, (e) => e.surface?.redo());
    expect(header()).toBe(typed);
    converged(room, [olive]);
  });

  test('a saved proposal reopens with its card still to decide', async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana');
    await act(ana, inBody(1, 4, 'PAGE'));
    const bytes = writeOoxmlPackage(roomFile(room.room.document, FILE));
    const again = openOrderedWordFile(bytes, { file: 2 }).document;
    expect(orderedCards(again).length).toBe(1);
  });
});

describe("a field's instruction and result (the editing API's lane)", () => {
  test('a new instruction and new results save as the editor alone writes them', async () => {
    const { room } = await sameAsAlone(FIELDS, 'edit', [code(0, 'NUMPAGES')]);
    expect(storyAsRead(roomFile(room.room.document, FIELDS))).not.toContain(' PAGE ');
    await sameAsAlone(FIELDS, 'edit', [code(1, 'page')]);
    await sameAsAlone(FIELDS, 'edit', [results(['7', '12'])]);
    // A result that is the text already, and one the field does not have, change nothing
    await sameAsAlone(FIELDS, 'edit', [results(['1', '4\t5'])]);
  });

  test("the organiser's undo of a new result shows the old one, and redo the new", async () => {
    const room = localRoom(FIELDS);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FIELDS });
    const original = storyAsRead(roomFile(room.room.document, FIELDS));
    await act(olive, results(['9', null]));
    await act(olive, code(1, 'PAGE'));
    const changed = storyAsRead(roomFile(room.room.document, FIELDS));
    expect(changed).toContain('>9<');
    await act(olive, (e) => e.surface?.undo());
    await act(olive, (e) => e.surface?.undo());
    expect(storyAsRead(roomFile(room.room.document, FIELDS))).toBe(original);
    await act(olive, (e) => e.surface?.redo());
    await act(olive, (e) => e.surface?.redo());
    expect(storyAsRead(roomFile(room.room.document, FIELDS))).toBe(changed);
    converged(room, [olive]);
  });

  test("the party rules: a party's field instruction or result is refused as untracked", async () => {
    const room = localRoom(FIELDS, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: FIELDS });
    const original = storyAsRead(roomFile(room.room.document, FIELDS));
    await act(ana, code(0, 'NUMPAGES'));
    await act(ana, results(['5', null]));
    expect(ana.refusals.map((r) => `${r.code}:${r.reason}`)).toEqual([
      'ordered-refused:package',
      'ordered-refused:text',
    ]);
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(storyAsRead(roomFile(room.room.document, FIELDS))).toBe(original);
  });

  test('a refresh and typing beside the field at the same time converge', async () => {
    const room = localRoom(FIELDS, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: FIELDS });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FIELDS });
    select(ana.editor, 0, 5);
    ana.editor.surface?.insertPlainText('No. ');
    results(['2', '8'])(olive.editor);
    await settle();
    converged(room, [ana, olive]);
    const story = storyAsRead(roomFile(room.room.document, FIELDS));
    expect(story).toContain('No. ');
    expect(story).toContain('>8<');
  });
});

describe("a format of a field's result", () => {
  const BOLD_FIELD = docx(p(t('Page ') + fieldRuns('PAGE', '1', '<w:i/>') + t(' here.')));

  test('bold over a field reaches its result, as the editor alone writes it', async () => {
    const { room } = await sameAsAlone(BOLD_FIELD, 'edit', [(e) => (select(e, 0, 0, 11), bold(e))]);
    expect(storyAsRead(roomFile(room.room.document, BOLD_FIELD))).toContain(
      '<w:rPr><w:b/><w:i/></w:rPr><w:t>1</w:t>'
    );
    await sameAsAlone(BOLD_FIELD, 'suggesting', [(e) => (select(e, 0, 5, 6), bold(e))]);
  });

  // `{ IF { MERGEFIELD Name } = "" "Dear" "" }`: the outer field's result takes the format, else
  // its own separator's run, never a run of the field in its code
  const ifField = (inner: string, result: string) =>
    docx(
      p(
        t('Page ') +
          '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
          '<w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>' +
          inner +
          '<w:r><w:instrText xml:space="preserve"> = "" "Dear" "" </w:instrText></w:r>' +
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
          result +
          '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
          t(' here.')
      )
    );
  const mergeWithoutResult =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> MERGEFIELD Name </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const boldField: Act = (e) => (select(e, 0, 5, 6), bold(e));

  test('bold over a field with a field in its code reaches the outer field only', async () => {
    await sameAsAlone(ifField(mergeWithoutResult, ''), 'suggesting', [boldField]);
    const withResults = ifField(fieldRuns('MERGEFIELD Name', '«Name»'), t('Dear'));
    await sameAsAlone(withResults, 'suggesting', [boldField]);
  });
});
