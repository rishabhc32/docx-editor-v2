// A header change from the room repaints only the header band of each sheet, on every editor.
//
// The organiser's typing in a header shown on every page reaches the other editors as a remote
// change, and a party's typing there is refused by the room and taken back. Each of these used
// to build every sheet again, body included. Here each one keeps the sheets, their body layer
// and its blocks, and the screen is the one a full paint of the same state shows.

import { afterEach, describe, expect, test } from 'bun:test';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  join,
  listTexts,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { HEADER, inHeader } from './sections-fixtures.ts';

afterEach(closeAll);

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WML_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const PAGES = 10;

/** Ten pages, each with two paragraphs and the default header. */
const FILE = docx(
  Array.from(
    { length: PAGES },
    (_, page) =>
      p(t(`Page ${page + 1} text.`)) +
      p(t('More text.')) +
      (page < PAGES - 1 ? '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' : '')
  ).join('') +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>',
  {
    parts: {
      'word/header1.xml': `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>`,
    },
    relationships: `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes: `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>`,
  }
);

/** The painted nodes a band repaint must keep: the sheets, their body layer and its blocks. */
function bodyNodes(person: Person): Element[] {
  const { container } = person;
  return [
    ...container.querySelectorAll('.docx-page'),
    ...container.querySelectorAll('.docx-page > .docx-page-content'),
    ...container.querySelectorAll('.docx-page > .docx-page-content > *'),
  ];
}

/** How many of each person's body nodes `run` replaced. */
async function replacedBy(people: Person[], run: () => Promise<void>): Promise<number[]> {
  const before = people.map(bodyNodes);
  await run();
  return people.map((person, k) => {
    const after = bodyNodes(person);
    expect(after).toHaveLength(before[k]!.length);
    return after.filter((node, index) => node !== before[k]![index]).length;
  });
}

const headerTexts = (person: Person) =>
  [...person.container.querySelectorAll('.docx-page > [data-docx-hf="header"]')].map(
    (band) => band.textContent
  );

const sheets = (person: Person) =>
  [...person.container.querySelectorAll('.docx-page')].map((sheet) => sheet.outerHTML);

describe('a header change in a room repaints only the header bands', () => {
  test("the organiser's typing, the remote change and a party's take-back", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: FILE });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    expect(headerTexts(ana)).toHaveLength(PAGES);
    inHeader(olive, 6);

    // The organiser's keystroke, then the room's answer and the remote change on Ana's editor.
    for (const text of [' new', ' more']) {
      expect(
        await replacedBy([olive, ana], async () => surface(olive).insertPlainText(text))
      ).toEqual([0, 0]);
      expect(await replacedBy([olive, ana], () => settle())).toEqual([0, 0]);
    }
    expect(olive.refusals).toEqual([]);
    for (const person of [ana, olive]) {
      expect(sameAsRoom(person, room)).toBe(true);
      expect(listTexts(person, HEADER)).toEqual(['Header new more text']);
      expect(headerTexts(person)).toEqual(Array(PAGES).fill('Header new more text'));
    }

    // A party's typing: the room refuses it, and the editor takes it back. The party already
    // holds a suggestion in the body, as a party does: a person's first suggestion, and the
    // take-back of their last, add or remove a person in the paint's author map, which
    // repaints every sheet once.
    caretAt(ana, 0, 4);
    surface(ana).insertPlainText('new ');
    await settle();
    inHeader(ana, 6);
    expect(await replacedBy([ana, olive], async () => surface(ana).insertPlainText(' x'))).toEqual([
      0, 0,
    ]);
    expect(await replacedBy([ana, olive], () => settle())).toEqual([0, 0]);
    expect(ana.refusals.map((r) => `${r.code}:${r.reason}`)).toEqual(['ordered-refused:furniture']);
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(headerTexts(ana)).toEqual(Array(PAGES).fill('Header new more text'));

    // Each screen is the one a full paint of the same state shows. A zoom change paints every
    // sheet again, with the editor's own author colours.
    for (const person of [ana, olive]) {
      const shown = sheets(person);
      const zoom = person.editor.getZoom();
      expect(person.editor.setZoom(zoom * 2).ok).toBe(true);
      expect(person.editor.setZoom(zoom).ok).toBe(true);
      expect(sheets(person)).toEqual(shown);
    }
  });
});
