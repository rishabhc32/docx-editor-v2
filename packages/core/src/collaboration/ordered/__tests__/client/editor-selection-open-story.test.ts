// A host selection in a header part that no section shows, through real editors and one room.
// The surface refuses it by name and keeps the caret, so the next key types in
// the body for the organiser and for a party. Before, the organiser's key changed the hidden part
// directly, and a party's key went to the room and came back refused as `unknown-paragraph`.

import { afterEach, describe, expect, test } from 'bun:test';
import { serializeOoxmlPart, type OoxmlNode } from '../../../../store/package/ooxml-tree.ts';
import { OUTSIDE_STORY_REFUSAL } from '../../../../editor/surface-scope.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  closeAll,
  editorTexts,
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

afterEach(closeAll);

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WML_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const HIDDEN = '/word/header2.xml';

const FILE = docx(
  p(t('The fee is 100 USD.')) +
    p(t('Due in 30 days.')) +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/>' +
    '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
  {
    parts: {
      'word/header1.xml': `<w:hdr xmlns:w="${W}">${p(t('Shown header.'))}</w:hdr>`,
      'word/header2.xml': `<w:hdr xmlns:w="${W}">${p(t('Hidden header.'))}</w:hdr>`,
    },
    contentTypes:
      `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>` +
      `<Override PartName="/word/header2.xml" ContentType="${WML_CT}.header+xml"/>`,
    relationships:
      `<Relationship Id="rIdH1" Type="${R}/header" Target="header1.xml"/>` +
      `<Relationship Id="rIdH2" Type="${R}/header" Target="header2.xml"/>`,
  }
);

function firstParagraph(node: OoxmlNode): string | null {
  if (node.kind === 'textValue') return null;
  if (node.kind === 'paragraph') return node.id;
  for (const child of node.children) {
    const found = firstParagraph(child);
    if (found !== null) return found;
  }
  return null;
}

function hiddenParagraph(person: Person): string {
  const part = surface(person).session.currentPackage().parts.get(HIDDEN);
  const found = part ? firstParagraph(part.root) : null;
  if (found === null) throw new Error('no hidden header paragraph');
  return found;
}

const hiddenXml = (person: Person): string =>
  serializeOoxmlPart(surface(person).session.currentPackage().parts.get(HIDDEN)!);

describe('a selection in a header part no section shows is refused in a room', () => {
  for (const who of ['Olive', 'Ana'] as const) {
    test(`${who}: refused by name; the key types in the body on every copy`, async () => {
      const room = localRoom(FILE);
      const ana = await join(room, 'Ana', { wordFile: FILE });
      const olive = await join(room, 'Olive', { wordFile: FILE, mode: 'edit', login: ORGANISER });
      const person = who === 'Olive' ? olive : ana;
      const s = surface(person);
      const body = s.session.paragraphIds()[0]!;
      s.setSelection({
        anchor: { paragraphId: body, offset: 3 },
        head: { paragraphId: body, offset: 3 },
      });
      const hidden = hiddenParagraph(person);
      s.setSelection({
        anchor: { paragraphId: hidden, offset: 2 },
        head: { paragraphId: hidden, offset: 2 },
      });
      expect(s.state().lastRejection).toBe(OUTSIDE_STORY_REFUSAL);
      s.type('X');
      await settle();
      for (const each of [ana, olive]) {
        expect(each.refusals).toEqual([]);
        expect(sameAsRoom(each, room)).toBe(true);
        expect(editorXml(each)).toBe(listXml(each));
        expect(editorTexts(each)[0]).toBe('TheX fee is 100 USD.');
        expect(hiddenXml(each)).toContain('Hidden header.');
        expect(hiddenXml(each)).not.toContain('X');
      }
    });
  }
});
