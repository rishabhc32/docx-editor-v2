// A list's picture bullet (`w:numPicBullet`, §17.9.20) on copy and paste. Word copies the
// bullet and its image with the list; the editor does too: the copy carries the bullet and its
// image as a relationship of the numbering part, and the paste imports them under fresh ids,
// reusing a bullet of the same look that the target holds.

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import type { OoxmlPackage } from '../package/ooxml-package.ts';
import { serializeOoxmlPart } from '../package/ooxml-tree.ts';
import { resolveInternalTarget } from '../package/opc-names.ts';
import { extractFragmentPackage } from '../store/clipboard-fragment-extract.ts';
import { PNG } from './canonical-primitive-journal-coverage-support.ts';
import {
  bodyOf,
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
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
/** A one-pixel red PNG: another image than `PNG`. */
const RED = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGO4IycHAALyARlzRAvLAAAAAElFTkSuQmCC'
  ),
  (c) => c.charCodeAt(0)
);

const bullet = (id: number, rel: string) =>
  `<w:numPicBullet w:numPicBulletId="${id}"><w:pict>` +
  '<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:9pt;height:9pt" o:bullet="t">' +
  `<v:imagedata r:id="${rel}" o:title="dot"/></v:shape></w:pict></w:numPicBullet>`;

const pictureList = (abstractId: number, bulletId: number) =>
  `<w:abstractNum w:abstractNumId="${abstractId}"><w:multiLevelType w:val="singleLevel"/>` +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>' +
  `<w:lvlPicBulletId w:val="${bulletId}"/><w:lvlJc w:val="left"/>` +
  '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>';

/** A document with one list paragraph whose level draws the bullet image `image`. */
function documentWith(image: Uint8Array, text: string): OoxmlPackage {
  return loadPackage(
    zipSync({
      '[Content_Types].xml': strToU8(
        `<Types xmlns="${CT}">` +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="png" ContentType="image/png"/>' +
          `<Override PartName="/word/document.xml" ContentType="${WML}.document.main+xml"/>` +
          `<Override PartName="/word/numbering.xml" ContentType="${WML}.numbering+xml"/></Types>`
      ),
      '_rels/.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
      ),
      'word/_rels/document.xml.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId9" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`
      ),
      'word/_rels/numbering.xml.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/image" Target="media/image1.png"/></Relationships>`
      ),
      'word/media/image1.png': image,
      'word/numbering.xml': strToU8(
        `<w:numbering xmlns:w="${W}" xmlns:r="${R}" xmlns:v="${V}" xmlns:o="${O}">` +
          `${bullet(0, 'rId1')}${pictureList(0, 0)}` +
          '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
      ),
      'word/document.xml': strToU8(
        `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>` +
          `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>` +
          '</w:body></w:document>'
      ),
    })
  );
}

/** The image bytes of a numbering part's relationship. */
function imageOf(pkg: OoxmlPackage, id: string): Uint8Array | undefined {
  const record = pkg.relationships.get('/word/numbering.xml')?.find((r) => r.id === id);
  const target = record ? resolveInternalTarget(record.ownerPart, record.rawTarget) : null;
  return target?.ok ? pkg.partBytes.get(target.partName) : undefined;
}

const numberingXml = (pkg: OoxmlPackage) =>
  serializeOoxmlPart(pkg.parts.get('/word/numbering.xml')!);

/** Paste the source's whole body at the start of the target's first paragraph. */
function paste(target: OoxmlPackage, fragment: Uint8Array): OoxmlPackage {
  const store = openStore(target);
  const host = paragraphIdsUnder(
    bodyOf(store.currentPackage().parts.get('/word/document.xml')!)
  )[0]!;
  const result = store.applyFragmentPaste(
    { kind: 'body' },
    { paragraphId: host, offset: 0, fragmentBytes: fragment, lastMarkCovered: true }
  );
  if (!result.ok) throw new Error(result.reason);
  return store.currentPackage();
}

const SOURCE = documentWith(PNG, 'Pasted');

function copied(options?: { omitMedia: boolean }): Uint8Array {
  const extracted = extractFragmentPackage(SOURCE, fullBodyCoverage(SOURCE), options);
  if (!extracted.ok) throw new Error(extracted.reason);
  return extracted.bytes;
}

/** `numId` of the first pasted paragraph, and the picture bullet id its level names. */
function pastedBullet(pkg: OoxmlPackage): { numId: string; bulletId: string | undefined } {
  const doc = serializeOoxmlPart(pkg.parts.get('/word/document.xml')!);
  const numId = /<w:numId w:val="(\d+)"\/>/.exec(doc)![1]!;
  const numbering = numberingXml(pkg);
  const abstractId = new RegExp(`<w:num w:numId="${numId}"><w:abstractNumId w:val="(\\d+)"/>`).exec(
    numbering
  )![1]!;
  const abstract = new RegExp(
    `<w:abstractNum w:abstractNumId="${abstractId}">(.*?)</w:abstractNum>`
  ).exec(numbering)![1]!;
  return { numId, bulletId: /<w:lvlPicBulletId w:val="(\d+)"\/>/.exec(abstract)?.[1] };
}

describe('picture bullets on copy and paste', () => {
  test('the copy carries the bullet and its image, as a relationship of the numbering part', () => {
    const extracted = extractFragmentPackage(SOURCE, fullBodyCoverage(SOURCE));
    if (!extracted.ok) throw new Error(extracted.reason);
    const numbering = numberingXml(extracted.package);
    expect(numbering).toContain('<w:numPicBullet w:numPicBulletId="0">');
    expect(numbering.indexOf('numPicBullet')).toBeLessThan(numbering.indexOf('abstractNum'));
    expect(imageOf(extracted.package, 'rId1')).toEqual(PNG);
    // The bytes on the clipboard read back the same
    expect(imageOf(loadPackage(extracted.bytes), 'rId1')).toEqual(PNG);
  });

  test('a paste imports the bullet under a fresh id with its own image', () => {
    const target = documentWith(RED, 'Target');
    const after = paste(target, copied());
    const { bulletId } = pastedBullet(after);
    expect(bulletId).toBeDefined();
    expect(bulletId).not.toBe('0');
    const numbering = numberingXml(after);
    const rel = new RegExp(
      `<w:numPicBullet w:numPicBulletId="${bulletId}">.*?<v:imagedata r:id="([^"]+)"`
    ).exec(numbering)![1]!;
    expect(imageOf(after, rel)).toEqual(PNG);
    // The target's own bullet keeps its own image, and every bullet precedes every definition
    expect(imageOf(after, 'rId1')).toEqual(RED);
    expect(numbering.lastIndexOf('<w:numPicBullet ')).toBeLessThan(
      numbering.indexOf('<w:abstractNum ')
    );
  });

  test('a second paste reuses the bullet of the first', () => {
    const once = paste(documentWith(RED, 'Target'), copied());
    const twice = paste(once, copied());
    expect(pastedBullet(twice).bulletId).toBe(pastedBullet(once).bulletId!);
    expect(numberingXml(twice).match(/<w:numPicBullet /g)).toHaveLength(2);
  });

  test('a paste into the document it came from keeps its list', () => {
    const after = paste(SOURCE, copied());
    expect(pastedBullet(after)).toEqual({ numId: '1', bulletId: '0' });
    expect(numberingXml(after).match(/<w:numPicBullet /g)).toHaveLength(1);
  });

  test('a copy that leaves the media behind leaves the bullet out, and the level keeps its text', () => {
    const after = paste(documentWith(RED, 'Target'), copied({ omitMedia: true }));
    const { numId, bulletId } = pastedBullet(after);
    expect(numId).not.toBe('1');
    expect(bulletId).toBeUndefined();
    expect(numberingXml(after).match(/<w:numPicBullet /g)).toHaveLength(1);
  });
});
