// An embedded object (`w:object`: an OLE part or an embedded Office file, and its preview
// picture) on copy and paste. Word copies it with its parts; so does the editor. The parts are
// bytes, checked by signature and never opened. A linked object is never loaded: it stays out.

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
/** An OLE compound file's first bytes: the signature the copy checks. Nothing reads further. */
const OLE = new Uint8Array([
  0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0, 0, 0, 0, 0,
]);
const WORKBOOK = zipSync({ 'xl/workbook.xml': strToU8('<workbook/>') });
const OLE_TYPE = 'application/vnd.openxmlformats-officedocument.oleObject';
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** A line with an embedded object whose relationship `rIdOle` is `ole`, its part `bytes`. */
function source(ole: string, part?: readonly [string, Uint8Array]): OoxmlPackage {
  return loadPackage(
    zipSync({
      '[Content_Types].xml': strToU8(
        `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          '<Default Extension="png" ContentType="image/png"/>' +
          `<Default Extension="bin" ContentType="${OLE_TYPE}"/><Default Extension="xlsx" ContentType="${XLSX_TYPE}"/>` +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
      ),
      '_rels/.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
      ),
      'word/_rels/document.xml.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rIdPrev" Type="${R}/image" Target="media/image1.png"/>${ole}</Relationships>`
      ),
      'word/media/image1.png': PNG,
      ...(part ? { [part[0]]: part[1] } : {}),
      'word/document.xml': strToU8(
        `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:v="${V}" xmlns:o="${O}"><w:body><w:p>` +
          '<w:r><w:t xml:space="preserve">Sheet </w:t></w:r><w:r><w:object w:dxaOrig="1440" w:dyaOrig="720">' +
          '<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:72pt;height:36pt" o:ole="">' +
          '<v:imagedata r:id="rIdPrev" o:title=""/></v:shape>' +
          '<o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_i1025" DrawAspect="Content" ' +
          'ObjectID="_1234567890" r:id="rIdOle"/></w:object></w:r><w:r><w:t xml:space="preserve"> here.</w:t></w:r>' +
          '</w:p></w:body></w:document>'
      ),
    })
  );
}

/** Copy the source whole and paste it into a document whose own `rIdOle` is a link. */
function pasted(from: OoxmlPackage): OoxmlPackage {
  const extracted = extractFragmentPackage(from, fullBodyCoverage(from));
  if (!extracted.ok) throw new Error(extracted.reason);
  const store = openStore(
    buildPackage('<w:p/>', {
      'word/_rels/document.xml.rels':
        `<Relationships xmlns="${REL}"><Relationship Id="rIdOle" Type="${R}/hyperlink" ` +
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

/** The part a relationship of the body names, by the attribute that names it. */
function named(pkg: OoxmlPackage, pattern: RegExp): { type: string; bytes?: Uint8Array } | null {
  const xml = serializeOoxmlPart(pkg.parts.get('/word/document.xml')!);
  const id = pattern.exec(xml)?.[1];
  const record = pkg.relationships.get('/word/document.xml')?.find((r) => r.id === id);
  if (!record) return null;
  const target = resolveInternalTarget('/word/document.xml', record.rawTarget);
  return { type: record.type, ...(target.ok ? { bytes: pkg.partBytes.get(target.partName) } : {}) };
}

const OLE_ID = /<o:OLEObject [^>]*r:id="([^"]+)"/;
const PREVIEW_ID = /<v:imagedata [^>]*r:id="([^"]+)"/;

describe('an embedded object on copy and paste', () => {
  test('an OLE part travels whole, with its preview, under a new relationship', () => {
    const pkg = pasted(
      source(
        `<Relationship Id="rIdOle" Type="${R}/oleObject" Target="embeddings/oleObject1.bin"/>`,
        ['word/embeddings/oleObject1.bin', OLE]
      )
    );
    expect(named(pkg, OLE_ID)).toEqual({ type: `${R}/oleObject`, bytes: OLE });
    expect(named(pkg, PREVIEW_ID)).toEqual({ type: `${R}/image`, bytes: PNG });
  });

  test('an embedded workbook travels whole', () => {
    const pkg = pasted(
      source(
        `<Relationship Id="rIdOle" Type="${R}/package" Target="embeddings/Microsoft_Excel_Worksheet.xlsx"/>`,
        ['word/embeddings/Microsoft_Excel_Worksheet.xlsx', WORKBOOK]
      )
    );
    expect(named(pkg, OLE_ID)).toEqual({ type: `${R}/package`, bytes: WORKBOOK });
  });

  test('a part that is not what its type says, or a linked object, stays out; the text stays', () => {
    for (const from of [
      source(
        `<Relationship Id="rIdOle" Type="${R}/oleObject" Target="embeddings/oleObject1.bin"/>`,
        ['word/embeddings/oleObject1.bin', PNG]
      ),
      source(
        `<Relationship Id="rIdOle" Type="${R}/oleObject" Target="file:///C:/data.xlsx" TargetMode="External"/>`
      ),
    ]) {
      const xml = serializeOoxmlPart(pasted(from).parts.get('/word/document.xml')!);
      expect(xml).not.toContain('OLEObject');
      expect(xml).toContain('Sheet ');
      expect(xml).toContain(' here.');
    }
  });
});
