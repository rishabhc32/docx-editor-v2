// A picture linked by address (`a:blip/@r:link`, an external image relationship) on copy and
// paste. Word keeps the link and its relationship target as they are, and keeps an embedded
// copy only when the source had one (paste evidence 3, `~$paste-out-e3-*.docx`). The editor
// does the same: the address goes through the same gate as a link's (`sanitizeHref`), and
// nothing loads or fetches it.

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { serializeOoxmlPart } from '../package/ooxml-tree.ts';
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

const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const SEAL = 'https://images.example.com/seal.png';

/** A source with one picture whose blip has `blip`'s attributes, and `rels` beside it. */
function source(blip: string, rels: string) {
  return loadPackage(
    zipSync({
      '[Content_Types].xml': strToU8(
        `<Types xmlns="${CT}">` +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="png" ContentType="image/png"/>' +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
          '</Types>'
      ),
      '_rels/.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
      ),
      'word/_rels/document.xml.rels': strToU8(
        `<Relationships xmlns="${REL}">${rels}</Relationships>`
      ),
      'word/document.xml': strToU8(
        `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}">` +
          '<w:body><w:p><w:r><w:t>Seal</w:t></w:r>' +
          '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
          '<wp:extent cx="914400" cy="914400"/><wp:docPr id="11" name="LinkOnly"/>' +
          '<wp:cNvGraphicFramePr/>' +
          `<a:graphic><a:graphicData uri="${PIC}"><pic:pic>` +
          '<pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr>' +
          `<pic:blipFill><a:blip ${blip}/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
          '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
          '</pic:pic></a:graphicData></a:graphic>' +
          '</wp:inline></w:drawing></w:r></w:p></w:body></w:document>'
      ),
      'word/media/image1.png': PNG,
    })
  );
}

const linkTo = (target: string) =>
  `<Relationship Id="rIdSeal" Type="${R}/image" Target="${target}" TargetMode="External"/>`;
const EMBED = `<Relationship Id="rIdImg" Type="${R}/image" Target="media/image1.png"/>`;

/** Copy the source whole and paste it into an empty document. */
function pasted(blip: string, rels: string, tracked = false) {
  const from = source(blip, rels);
  const extracted = extractFragmentPackage(from, fullBodyCoverage(from));
  if (!extracted.ok) throw new Error(extracted.reason);
  const store = openStore(buildPackage('<w:p/>'));
  const host = paragraphIdsUnder(
    bodyOf(store.currentPackage().parts.get('/word/document.xml')!)
  )[0]!;
  const result = store.applyFragmentPaste(
    { kind: 'body' },
    {
      paragraphId: host,
      offset: 0,
      fragmentBytes: extracted.bytes,
      lastMarkCovered: true,
      ...(tracked ? { revision: { author: 'Ada', date: '2026-10-30T00:00:00Z' } } : {}),
    }
  );
  expect(result.ok).toBe(true);
  const pkg = store.currentPackage();
  const xml = serializeOoxmlPart(pkg.parts.get('/word/document.xml')!);
  const blipAttributes = /<a:blip ([^>]*?)\/?>/.exec(xml)?.[1] ?? '';
  const linkId = /r:link="([^"]+)"/.exec(blipAttributes)?.[1];
  const linked = pkg.externalTargets.find(
    (target) => target.id === linkId && target.type === `${R}/image`
  );
  const media = [...pkg.partBytes.keys()].filter((name) => name.includes('/media/'));
  return { xml, blipAttributes, linked, media };
}

describe('a picture linked by address, copied and pasted', () => {
  test('keeps its link and the address as it is, with no embedded copy (`-linkonly`)', () => {
    for (const tracked of [false, true]) {
      const { xml, blipAttributes, linked, media } = pasted(
        'r:link="rIdSeal"',
        linkTo(SEAL),
        tracked
      );
      expect(xml).toContain('<w:drawing>');
      expect(blipAttributes).not.toContain('r:embed');
      expect(linked?.rawTarget).toBe(SEAL);
      expect(media).toEqual([]);
      expect(xml.includes('<w:ins ')).toBe(tracked);
    }
  });

  test('keeps both its embedded copy and its link (`-linkembed`)', () => {
    const { blipAttributes, linked, media } = pasted(
      'r:embed="rIdImg" r:link="rIdSeal"',
      EMBED + linkTo(SEAL)
    );
    expect(blipAttributes).toMatch(/r:embed="[^"]+"/);
    expect(linked?.rawTarget).toBe(SEAL);
    expect(media).toHaveLength(1);
  });

  test('a link address the link gate refuses goes with the link, and the picture with it', () => {
    const { xml, linked } = pasted('r:link="rIdSeal"', linkTo('javascript:alert(1)'));
    expect(linked).toBeUndefined();
    expect(xml).not.toContain('javascript');
    expect(xml).not.toContain('<w:drawing>');
  });
});
