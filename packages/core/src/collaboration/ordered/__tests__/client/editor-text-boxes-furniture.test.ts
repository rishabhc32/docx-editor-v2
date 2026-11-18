// Text boxes in a header and in a footer, in an ordered room: each is a story of its own, as a
// main document's text box is, so the editing API's edit of one reaches the other copy and the
// room's Word file as the editor alone writes it. A party is refused there, as everywhere in a
// header or a footer (the host's party rule). Before, the room refused every such edit: only the
// main document's text boxes were stories. (A note's text box: the editor alone edits none either.)

import { afterEach, expect, test } from 'bun:test';
import { createBrowserAutomationHost } from '../../../../editor/automation-host.ts';
import { writeOoxmlPackage, type OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { docx, MATRIX_ITEMS, p, packageOf, R, t, W, WML_CT } from '../word-fixtures.ts';
import { closeAll, join, localRoom, ORGANISER, PARTY, type Editor } from './editor-harness.ts';
import { act, alone, converged, roomFile, storyAsRead } from './links-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';

afterEach(closeAll);

const BOX = MATRIX_ITEMS.find((item) => item.key === 'textbox')?.xml ?? '';
const DRAWING_NS =
  `xmlns:w="${W}" xmlns:r="${R}" ` +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

/** A body, a default header and a default footer, each of them with a text box ("Box text"). */
const FILE = docx(
  p(t('Body')) +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/>' +
    '<w:footerReference w:type="default" r:id="rIdFooter1"/></w:sectPr>',
  {
    parts: {
      'word/header1.xml': `<w:hdr ${DRAWING_NS}><w:p>${t('Head ')}${BOX}</w:p></w:hdr>`,
      'word/footer1.xml': `<w:ftr ${DRAWING_NS}><w:p>${t('Foot ')}${BOX}</w:p></w:ftr>`,
    },
    relationships:
      `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>` +
      `<Relationship Id="rIdFooter1" Type="${R}/footer" Target="footer1.xml"/>`,
    contentTypes:
      `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>` +
      `<Override PartName="/word/footer1.xml" ContentType="${WML_CT}.footer+xml"/>`,
  }
);

type Host = ReturnType<typeof createBrowserAutomationHost>;
/** The value of one read. */
function read<T>(host: Host, operation: object): T {
  const [result] = host.execute({ operations: [operation] } as never).results;
  if (result?.status !== 'ok') throw new Error(`refused: ${JSON.stringify(result)}`);
  return (result as unknown as { value: T }).value;
}
type Handle = { readonly handle: unknown };
type Handles = { readonly handles: readonly unknown[] };

/** The body of the text box in the default header or footer of the first section. */
function boxBody(host: Host, kind: Furniture): unknown {
  const document = read<Handle>(host, { op: 'getDocument' }).handle;
  const [section] = read<Handles>(host, { op: 'getSections', document }).handles;
  const body = read<Handle>(host, { op: 'getFurniture', section, kind, variant: 'default' }).handle;
  const [shape] = read<Handles>(host, { op: 'getShapes', span: { body } }).handles;
  return read<Handle>(host, { op: 'getShapeBody', shape }).handle;
}

/** The editing API replaces "Box" with "Crate" in a text box (tracked in suggesting mode). */
const replaceIn =
  (where: Furniture, tracked = false) =>
  (editor: Editor) => {
    const host = createBrowserAutomationHost(editor);
    if (tracked) read(host, { op: 'setChangeTrackingMode', mode: 'TrackMineOnly', author: 'Ana' });
    const body = boxBody(host, where);
    const [span] = read<{ spans: unknown[] }>(host, {
      op: 'search',
      scope: { body },
      text: 'Box',
    }).spans;
    return host.execute({ operations: [{ op: 'replaceSpan', span, text: 'Crate' }] } as never);
  };

const PARTS = { header: '/word/header1.xml', footer: '/word/footer1.xml' } as const;
type Furniture = keyof typeof PARTS;

/** A part of the saved file as Word reads it: the package writer syncs a text box's VML copy. */
const partAsRead = (pkg: OoxmlPackage, where: Furniture) =>
  storyAsRead(packageOf(writeOoxmlPackage(pkg)), PARTS[where]);

for (const where of ['header', 'footer'] as const)
  test(`the organiser's edit in a ${where}'s text box reaches the other copy and the file`, async () => {
    const room = localRoom(FILE);
    const review = reviewTestModule();
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      wordFile: FILE,
      review,
    });
    const ana = await join(room, 'Ana', { wordFile: FILE, review });
    expect(await act(olive, replaceIn(where))).toMatchObject({ ok: true });
    converged(room, [olive, ana]);
    const editorAlone = await alone(FILE, 'Olive', 'edit', replaceIn(where), review);
    const saved = roomFile(room.room.document, FILE);
    expect(partAsRead(saved, where)).toContain('Crate text');
    expect(partAsRead(saved, where)).toBe(partAsRead(editorAlone, where));
  });

test("a party is refused in a header's text box, as in the rest of the header", async () => {
  const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
  const review = reviewTestModule();
  const olive = await join(room, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    wordFile: FILE,
    review,
  });
  const ana = await join(room, 'Ana', { login: PARTY, wordFile: FILE, review });
  await act(ana, replaceIn('header', true));
  expect(ana.refusals).toMatchObject([{ code: 'ordered-refused', reason: 'furniture' }]);
  ana.refusals.length = 0;
  converged(room, [olive, ana]);
  expect(partAsRead(roomFile(room.room.document, FILE), 'header')).toContain('Box text');
});
