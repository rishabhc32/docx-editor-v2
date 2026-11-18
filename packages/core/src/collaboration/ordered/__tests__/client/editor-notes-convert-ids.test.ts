// A converted note's paragraph ids, through real editors and one room. The
// footnote's paragraphs carry `w14:paraId` values the endnotes part already uses; after the
// organiser converts it, every copy and the saved file hold each id once, as Word writes them.

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage, type OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { docx, p, t } from '../word-fixtures.ts';
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

afterEach(closeAll);

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WML_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const ENDNOTES = '/word/endnotes.xml';

const para = (paraId: string, runs: string) =>
  `<w:p w14:paraId="${paraId}" w14:textId="77777777">${runs}</w:p>`;
const separators = (kind: 'footnote' | 'endnote') =>
  `<w:${kind} w:type="separator" w:id="-1">${para('00000001', '<w:r><w:separator/></w:r>')}</w:${kind}>` +
  `<w:${kind} w:type="continuationSeparator" w:id="0">` +
  `${para('00000002', '<w:r><w:continuationSeparator/></w:r>')}</w:${kind}>`;

const FILE = docx(
  p(t('The fee') + '<w:r><w:footnoteReference w:id="1"/></w:r>' + t(' is due.')) +
    p(t('Paid in 30 days.') + '<w:r><w:endnoteReference w:id="1"/></w:r>') +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
  {
    parts: {
      'word/footnotes.xml':
        `<w:footnotes xmlns:w="${W}" xmlns:w14="${W14}">${separators('footnote')}` +
        `<w:footnote w:id="1">${para('1A2B3C4D', '<w:r><w:footnoteRef/></w:r>' + t(' Fees.'))}` +
        `${para('2B3C4D5E', t('Net of tax.'))}</w:footnote></w:footnotes>`,
      'word/endnotes.xml':
        `<w:endnotes xmlns:w="${W}" xmlns:w14="${W14}">${separators('endnote')}` +
        `<w:endnote w:id="1">${para('1A2B3C4D', '<w:r><w:endnoteRef/></w:r>' + t(' Annex.'))}` +
        `${para('2B3C4D5E', t('Second.'))}</w:endnote></w:endnotes>`,
    },
    contentTypes:
      `<Override PartName="/word/footnotes.xml" ContentType="${WML_CT}.footnotes+xml"/>` +
      `<Override PartName="/word/endnotes.xml" ContentType="${WML_CT}.endnotes+xml"/>`,
    relationships:
      `<Relationship Id="rIdNotes" Type="${R}/footnotes" Target="footnotes.xml"/>` +
      `<Relationship Id="rIdEndnotes" Type="${R}/endnotes" Target="endnotes.xml"/>`,
  }
);

const paraIds = (xml: string): string[] =>
  [...xml.matchAll(/<w:p [^>]*w14:paraId="([0-9A-Fa-f]{8})"/g)].map((m) => m[1] ?? '');

const endnotesOf = (pkg: OoxmlPackage): string => {
  const part = pkg.parts.get(ENDNOTES);
  return part ? serializeOoxmlPart(part) : '';
};

async function saved(person: Person): Promise<OoxmlPackage> {
  const read = readOoxmlPackage(new Uint8Array(await person.editor.save()), {});
  if (!read.ok) throw new Error(read.reason);
  return read.package;
}

describe("a converted note's paragraphs take ids of their own", () => {
  test("the organiser's conversion, on every copy and in the saved file", async () => {
    const room = localRoom(FILE);
    const ana = await join(room, 'Ana', { wordFile: FILE });
    const olive = await join(room, 'Olive', { wordFile: FILE, mode: 'edit', login: ORGANISER });
    const result = olive.editor.exec({ type: 'convertNote', fromKind: 'footnote', noteId: 1 });
    expect(result.ok, result.ok ? '' : result.reason).toBe(true);
    await settle();
    for (const person of [ana, olive]) {
      expect(person.refusals).toEqual([]);
      expect(sameAsRoom(person, room)).toBe(true);
      expect(editorXml(person)).toBe(listXml(person));
      const shown = endnotesOf(surface(person).session.currentPackage());
      expect(shown).toContain('Net of tax.');
      const ids = paraIds(shown);
      // Two separators, the endnote's two paragraphs and the converted note's two
      expect(ids).toHaveLength(6);
      expect(new Set(ids).size).toBe(ids.length);
      const file = paraIds(endnotesOf(await saved(person)));
      expect(file).toHaveLength(6);
      expect(new Set(file).size).toBe(file.length);
    }
  });
});
