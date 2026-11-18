// The screen follows the list (spec "The screen follows the list", scenario "Placement
// difference"). Typed between a comment range and a bookmark, the engine moves the bookmark's
// start in front of the comment's end (a known engine difference, matrix `comment+bookmark |
// type between`); the list keeps its placement. After the commit the client publishes the
// list's paragraph, so the editor shows the list's version, and the saved file is the room's.

import { afterEach, expect, test } from 'bun:test';
import { mountPaginatedSurface } from '../../../../editor/paginated-surface.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { kindNamed } from '../matrix/actions.ts';
import { bodyOf, extrasOf, type MatrixKind } from '../matrix/catalogue.ts';
import { docx } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  join,
  listXml,
  localRoom,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
} from './editor-harness.ts';

afterEach(closeAll);

const x = kindNamed('comment');
const y = kindNamed('bookmark');
const BOTH: MatrixKind = {
  ...x,
  key: 'comment+bookmark',
  name: `${x.name}, then ${y.name}`,
  group: 'text',
  xml: (x.xml ?? '') + (y.xml ?? ''),
  text: x.text + y.text,
};
const FILE = docx(bodyOf(BOTH), extrasOf(BOTH));
/** Right after the comment's text: between the comment range and the bookmark. */
const CARET = 'Alpha '.length + x.text.length;

/** The engine alone, in suggesting mode, makes the same edit. */
function engineAlone(): string {
  const container = document.createElement('div');
  document.body.append(container);
  const opened = mountPaginatedSurface(container, FILE, {
    scale: 1,
    editingMode: 'suggest',
    author: 'Ana',
    revisionDisplayMode: 'all-markup',
  });
  if (!opened.ok) throw new Error(opened.reason);
  const s = opened.surface;
  const id = s.session.paragraphIds()[1];
  if (id === undefined) throw new Error('no paragraph');
  s.setSelection({
    anchor: { paragraphId: id, offset: CARET },
    head: { paragraphId: id, offset: CARET },
  });
  s.insertPlainText('mid');
  const xml = serializeOoxmlPart(s.session.part());
  s.destroy();
  container.remove();
  return xml;
}

/** The order of the comment range's end and the bookmark's start in a paragraph's XML. */
function markerOrder(xml: string): string {
  const end = xml.indexOf('<w:commentRangeEnd');
  const start = xml.indexOf('<w:bookmarkStart');
  const typed = xml.indexOf('>mid<');
  return [
    ['comment end', end],
    ['bookmark start', start],
    ['typed', typed],
  ]
    .filter(([, at]) => (at as number) >= 0)
    .sort((a, b) => (a[1] as number) - (b[1] as number))
    .map(([name]) => name)
    .join(', ');
}

test('the editor shows the list where the engine placed typed text elsewhere', async () => {
  const engine = engineAlone();
  const room = localRoom(FILE);
  const ana = await join(room, 'Ana', { wordFile: FILE });
  caretAt(ana, 1, CARET);
  surface(ana).insertPlainText('mid');
  await settle();
  expect(ana.refusals).toEqual([]);
  expect(screenFollowsList(ana)).toBe(true);
  expect(sameAsRoom(ana, room)).toBe(true);
  expect(editorTexts(ana)[1]).toContain('mid');
  // The engine alone moved the bookmark start; the editor in the room shows the list's order
  expect(markerOrder(engine)).not.toBe(markerOrder(listXml(ana)));
  expect(markerOrder(editorXml(ana))).toBe(markerOrder(listXml(ana)));
});
