// A VML picture whose image older Word names by `o:relid` (the Office relationship id of
// ECMA-376 Part 4) instead of `r:id`, on copy and paste. Word copies the image with it; so does
// the editor: the copy carries the image, and the paste names it by the target's own
// relationship id.

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import type { OoxmlPackage } from '../package/ooxml-package.ts';
import { serializeOoxmlPart } from '../package/ooxml-tree.ts';
import { resolveInternalTarget } from '../package/opc-names.ts';
import { extractFragmentPackage } from '../store/clipboard-fragment-extract.ts';
import { PNG } from './canonical-primitive-journal-coverage-support.ts';
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

const V = 'urn:schemas-microsoft-com:vml';
const O = 'urn:schemas-microsoft-com:office:office';

/** A line with a VML picture whose image `rIdImg` names by `o:relid`. */
const SOURCE = loadPackage(
  zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Default Extension="png" ContentType="image/png"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rIdImg" Type="${R}/image" Target="media/image1.png"/></Relationships>`
    ),
    'word/media/image1.png': PNG,
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:v="${V}" xmlns:o="${O}"><w:body><w:p>` +
        '<w:r><w:t xml:space="preserve">Seal </w:t></w:r><w:r><w:pict>' +
        '<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:36pt;height:36pt">' +
        '<v:imagedata o:relid="rIdImg" o:title=""/></v:shape></w:pict></w:r>' +
        '<w:r><w:t xml:space="preserve"> here.</w:t></w:r></w:p></w:body></w:document>'
    ),
  })
);

/** Copy the source whole and paste it into a document whose own `rIdImg` is a link. */
function pasted(): OoxmlPackage {
  const extracted = extractFragmentPackage(SOURCE, fullBodyCoverage(SOURCE));
  if (!extracted.ok) throw new Error(extracted.reason);
  const store = openStore(
    buildPackage('<w:p/>', {
      'word/_rels/document.xml.rels':
        `<Relationships xmlns="${REL}"><Relationship Id="rIdImg" Type="${R}/hyperlink" ` +
        'Target="https://example.com/" TargetMode="External"/></Relationships>',
    })
  );
  const host = paragraphIdsUnder(
    bodyOf(store.currentPackage().parts.get('/word/document.xml')!)
  )[0]!;
  const result = store.applyFragmentPaste(
    { kind: 'body' },
    { paragraphId: host, offset: 0, fragmentBytes: extracted.bytes, lastMarkCovered: true }
  );
  if (!result.ok) throw new Error(result.reason);
  return store.currentPackage();
}

describe('a VML picture named by o:relid on copy and paste', () => {
  test('the image travels under a new relationship that o:relid names', () => {
    const pkg = pasted();
    const xml = serializeOoxmlPart(pkg.parts.get('/word/document.xml')!);
    const id = /<v:imagedata [^>]*o:relid="([^"]+)"/.exec(xml)?.[1];
    expect(id).toBeDefined();
    expect(id).not.toBe('rIdImg');
    const record = pkg.relationships.get('/word/document.xml')?.find((r) => r.id === id);
    expect(record?.type).toBe(`${R}/image`);
    const target = resolveInternalTarget('/word/document.xml', record?.rawTarget ?? '');
    expect(target.ok ? pkg.partBytes.get(target.partName) : undefined).toEqual(PNG);
  });
});
