// A tracked page setup: `setSectionProperties` with a revision records the section before it
// in a `w:sectPrChange` (ECMA-376 Part 1, §17.13.5.32), as Word does. Word for Mac, with track
// changes on, wrote this for an orientation change (2026-10-29): the complete earlier section,
// `w:orient="portrait"` stated though the section had not stated it.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart, type OoxmlPart } from '../package/ooxml-tree.ts';
import { applyTreeOp, type TreeDocOp } from '../store/tree-ops.ts';
import { bodySectionOf } from '../store/tree-op-section-address.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const SECTION =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr>';

function load(body: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'app/xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function apply(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`${result.reason}: ${result.detail ?? ''}`);
  return result.part;
}

/** The body's section as XML, without the namespace declarations and the change's date. */
function section(part: OoxmlPart): string {
  const xml = serializeOoxmlPart(part);
  return (/<w:sectPr>[\s\S]*<\/w:sectPr>/.exec(xml)?.[0] ?? '').replace(/ w:date="[^"]*"/g, '');
}

const ana = { author: 'Ana', date: '2026-10-29T10:00:00Z' };
const landscape: TreeDocOp = {
  op: 'setSectionProperties',
  orientation: 'landscape',
  revision: ana,
};

describe('a tracked page setup', () => {
  const part = load('<w:p><w:r><w:t>Text.</w:t></w:r></w:p>' + SECTION);

  test('records the complete earlier section, with the attribute it adds stated', () => {
    const xml = section(apply(part, landscape));
    expect(xml).toContain('<w:pgSz w:h="12240" w:orient="landscape" w:w="15840"/>');
    const record = /<w:sectPrChange [^>]*w:author="Ana"[^>]*><w:sectPr>([\s\S]*?)<\/w:sectPr>/.exec(
      xml
    );
    expect(record?.[1]).toBe(
      '<w:pgSz w:h="15840" w:orient="portrait" w:w="12240"/>' +
        '<w:pgMar w:bottom="1440" w:footer="720" w:gutter="0" w:header="720" w:left="1440" w:right="1440" w:top="1440"/>'
    );
  });

  test('accepted, the record goes; rejected, the section is what it was', () => {
    const tracked = apply(part, landscape);
    const accepted = section(apply(tracked, { op: 'acceptAllRevisions' }));
    expect(accepted).not.toContain('sectPrChange');
    expect(accepted).toContain('w:orient="landscape"');
    const rejected = apply(tracked, { op: 'rejectAllRevisions' });
    expect(section(rejected)).not.toContain('sectPrChange');
    expect(section(rejected)).not.toContain('landscape');
    // Word's reject leaves the default orientation out (page2-rejected.docx, 2026-10-29).
    expect(section(rejected)).toContain('<w:pgSz w:h="15840" w:w="12240"/>');
    const size = bodySectionOf(rejected);
    expect(
      size && size.kind !== 'textValue' ? serializeOoxmlPart({ ...rejected, root: size }) : ''
    ).toContain('w:h="15840"');
  });

  test('a second write keeps the original record; a write back to it leaves none', () => {
    const twice = apply(apply(part, landscape), {
      op: 'setSectionProperties',
      marginLeftTwips: 720,
      revision: ana,
    });
    expect(section(twice).match(/<w:sectPrChange /g)).toHaveLength(1);
    expect(section(twice)).toContain('<w:pgSz w:h="15840" w:orient="portrait" w:w="12240"/>');
    const back = apply(apply(part, landscape), {
      op: 'setSectionProperties',
      orientation: 'portrait',
      revision: ana,
    });
    expect(section(back)).not.toContain('sectPrChange');
  });

  test('rejected, a recorded property the section lost goes back at its place in the sequence', () => {
    // A file's own record of a different first page taken off: `w:titlePg` comes before
    // `w:docGrid` (`CT_SectPrBase`, ECMA-376 Part 1, §17.6.17)
    const recorded = load(
      '<w:p><w:r><w:t>Text.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
        '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/>' +
        '<w:sectPrChange w:id="1" w:author="Cara" w:date="2026-01-01T00:00:00Z"><w:sectPr>' +
        '<w:pgSz w:w="12240" w:h="15840"/><w:titlePg/><w:docGrid w:linePitch="360"/>' +
        '</w:sectPr></w:sectPrChange></w:sectPr>'
    );
    expect(section(apply(recorded, { op: 'rejectAllRevisions' }))).toBe(
      '<w:sectPr><w:pgSz w:h="15840" w:w="12240"/><w:cols w:space="720"/><w:titlePg/>' +
        '<w:docGrid w:linePitch="360"/></w:sectPr>'
    );
  });

  test('without a revision it is written as before, untracked', () => {
    const { revision: _revision, ...untracked } = landscape as Extract<
      TreeDocOp,
      { op: 'setSectionProperties' }
    >;
    expect(section(apply(part, untracked))).not.toContain('sectPrChange');
  });
});
