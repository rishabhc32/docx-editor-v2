// Two controls side by side whose node ids read as a prefix of each other (`...#0.0.0.1` and
// `...#0.0.0.10`), with two real editors and one ordered room. The caret between them is outside
// both: typed words go after the first control's end (rule 5), and a new control may be made
// there, as in the editor alone.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { stubReviewModule } from '../../../../editor/__tests__/review-test-module.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorXml,
  join,
  listXml,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';

afterEach(closeAll);

const sdt = (id: number, text: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="${id}"/></w:sdtPr><w:sdtContent>${t(text)}</w:sdtContent></w:sdt>`;
/** Eight bookmark markers, so control B is child 10 of the paragraph and control A child 1. */
const MARKERS = [0, 2, 4, 6]
  .map((id) => `<w:bookmarkStart w:id="${id}" w:name="b${id}"/><w:bookmarkEnd w:id="${id}"/>`)
  .join('');
const FILE = docx(p(t('x') + sdt(1, 'AAA') + MARKERS + sdt(2, 'BBB')));
/** Between `AAA` and `BBB`. */
const BETWEEN = 4;

/** The body without paragraph ids and control ids. */
const bodyOf = (xml: string): string =>
  xml
    .slice(xml.indexOf('<w:body>'), xml.indexOf('</w:body>'))
    .replace(/ w14:(paraId|textId)="[^"]*"/g, '')
    .replace(/<w:id w:val="\d+"\/>/g, '<w:id/>');

function converged(room: ReturnType<typeof localRoom>, ...people: Person[]): void {
  for (const person of people) {
    expect({
      person: person.name,
      refusals: person.refusals,
      room: sameAsRoom(person, room),
      screen: screenFollowsList(person),
    }).toEqual({ person: person.name, refusals: [], room: true, screen: true });
  }
}

/** The same edit on the editor alone, in edit mode: its body. */
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

describe('two controls whose ids read as a prefix of each other', () => {
  test('a word typed between them goes after the first control, outside both', async () => {
    const room = localRoom(FILE);
    const ana = await join(room, 'Ana', { wordFile: FILE });
    const ben = await join(room, 'Ben', { wordFile: FILE });
    caretAt(ana, 0, BETWEEN);
    surface(ana).insertPlainText('Z');
    await settle();
    converged(room, ana, ben);
    // Z is after the first control's end and past the bookmarks, inside neither control
    expect(listXml(ana)).toMatch(
      /AAA<\/w:t><\/w:r><\/w:sdtContent><\/w:sdt>(<w:bookmark[^>]*\/>)+<w:ins [^>]*><w:r><w:t>Z<\/w:t><\/w:r><\/w:ins><w:sdt>/
    );
    expect(bodyOf(editorXml(ben))).toBe(bodyOf(editorXml(ana)));
  });

  test('a new control between them is made, as in the editor alone', async () => {
    const room = localRoom(FILE);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    const ben = await join(room, 'Ben', { wordFile: FILE });
    const insert = (person: Person) => {
      caretAt(person, 0, BETWEEN);
      return person.editor.exec({ type: 'insertContentControl', subtype: 'plainText' });
    };
    expect(insert(olive)).toMatchObject({ ok: true });
    await settle();
    converged(room, olive, ben);
    expect(bodyOf(editorXml(olive))).toBe(alone(insert));
    expect(bodyOf(editorXml(ben))).toBe(bodyOf(editorXml(olive)));
  });
});
