// Ink (`w14:contentPart`) on copy and paste. Word copies ink with its InkML part; so does the
// editor: the part travels as bytes (checked as XML text, never read), and the paste writes it
// under a new name with a new relationship, never one the target already uses.

import { expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { serializeOoxmlPart } from '../package/ooxml-tree.ts';
import { resolveInternalTarget } from '../package/opc-names.ts';
import { extractFragmentPackage } from '../store/clipboard-fragment-extract.ts';
import {
  bodyOf,
  buildPackage,
  CT,
  fullBodyCoverage,
  loadPackage,
  openStore,
  paragraphIdsUnder,
  R,
  REL,
  W,
} from './clipboard-fragment-fixtures.ts';

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

const INKED = loadPackage(
  zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/ink/ink1.xml" ContentType="application/inkml+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId5" ` +
        'Type="http://schemas.microsoft.com/office/2011/relationships/ink" Target="ink/ink1.xml"/></Relationships>'
    ),
    'word/ink/ink1.xml': strToU8(
      '<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:trace>0 0, 10 10</inkml:trace></inkml:ink>'
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="${W14}" xmlns:mc="${MC}" mc:Ignorable="w14">` +
        '<w:body><w:p><w:r><w:t>Ink</w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires="w14">' +
        '<w14:contentPart r:id="rId5"/></mc:Choice><mc:Fallback><w:t>x</w:t></mc:Fallback>' +
        '</mc:AlternateContent></w:r></w:p></w:body></w:document>'
    ),
  })
);

test('copied ink travels with its part; the paste names a new relationship to it', () => {
  const extracted = extractFragmentPackage(INKED, fullBodyCoverage(INKED));
  if (!extracted.ok) throw new Error(extracted.reason);
  // A target whose own rId5 is a link: a pasted rId5 would name it
  const target = buildPackage('<w:p/>', {
    'word/_rels/document.xml.rels':
      `<Relationships xmlns="${REL}"><Relationship Id="rId5" Type="${R}/hyperlink" ` +
      'Target="https://example.com/" TargetMode="External"/></Relationships>',
  });
  const store = openStore(target);
  const host = paragraphIdsUnder(
    bodyOf(store.currentPackage().parts.get('/word/document.xml')!)
  )[0]!;
  const result = store.applyFragmentPaste(
    { kind: 'body' },
    { paragraphId: host, offset: 0, fragmentBytes: extracted.bytes, lastMarkCovered: true }
  );
  expect(result.ok).toBe(true);
  const pkg = store.currentPackage();
  const xml = serializeOoxmlPart(pkg.parts.get('/word/document.xml')!);
  expect(xml).toContain('<w:t>Ink</w:t>');
  const id = /<w14:contentPart r:id="([^"]+)"/.exec(xml)?.[1];
  expect(id).toBeDefined();
  expect(id).not.toBe('rId5');
  const record = pkg.relationships.get('/word/document.xml')?.find((r) => r.id === id);
  expect(record?.type).toBe('http://schemas.microsoft.com/office/2011/relationships/ink');
  const name = resolveInternalTarget('/word/document.xml', record!.rawTarget);
  expect(name.ok && name.partName.startsWith('/word/ink/')).toBe(true);
  expect(
    new TextDecoder().decode(name.ok ? pkg.partBytes.get(name.partName) : undefined)
  ).toContain('<inkml:trace>0 0, 10 10</inkml:trace>');
});
