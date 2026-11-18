// A tracked different first page and header and footer distances: `setSectionFurnitureOptions`
// with a revision records the section before it in a `w:sectPrChange`, as Word does. Word for
// Mac, with track changes on (2026-10-30, `/tmp/word-evidence-sections/`): a header distance is
// one "revision section property" ("Formatted: Header distance from edge"), its record the old
// `w:pgMar` (`~$hdrdist-tracked.docx`); a different first page is one too ("Formatted: Different
// first page header"), its record `<w:titlePg w:val="0"/>` (`~$titlepg-tracked.docx`).

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { applyHeaderFooterLifecycleOp } from '../package/hf-lifecycle.ts';
import { readOoxmlPackage, type OoxmlPackage } from '../package/ooxml-package.ts';
import { serializeOoxmlPart, type OoxmlPart } from '../package/ooxml-tree.ts';
import { applyTreeOp } from '../store/tree-ops.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MARGINS =
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';

function load(section: string, record = ''): OoxmlPackage {
  const bytes = zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Text.</w:t></w:r></w:p>` +
        `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>${MARGINS}${section}` +
        `<w:docGrid w:linePitch="360"/>${record}</w:sectPr></w:body></w:document>`
    ),
  });
  const read = readOoxmlPackage(bytes);
  if (!read.ok) throw new Error(read.reason);
  return read.package;
}

const ana = { author: 'Ana', date: '2026-10-30T10:00:00Z' };

/** The main part after a furniture write. */
function written(
  pkg: OoxmlPackage,
  options: {
    readonly titlePage?: boolean;
    readonly headerDistanceTwips?: number;
    readonly footerDistanceTwips?: number;
    readonly evenAndOddHeaders?: boolean;
  },
  tracked = true
): { readonly pkg: OoxmlPackage; readonly main: OoxmlPart } {
  const result = applyHeaderFooterLifecycleOp(pkg, {
    op: 'setSectionFurnitureOptions',
    ...(options.evenAndOddHeaders === undefined ? { sectionIndex: 0 } : {}),
    ...options,
    ...(tracked ? { revision: ana } : {}),
  });
  if (!result.ok) throw new Error(`${result.reason}: ${result.detail ?? ''}`);
  const main = result.package.parts.get(result.package.mainDocumentPart);
  if (!main) throw new Error('no main part');
  return { pkg: result.package, main };
}

/** The body's section as XML, without the change's date. */
const section = (part: OoxmlPart): string =>
  (/<w:sectPr>[\s\S]*<\/w:sectPr>/.exec(serializeOoxmlPart(part))?.[0] ?? '').replace(
    / w:date="[^"]*"/g,
    ''
  );

/** The record a tracked write leaves: what `w:sectPrChange` holds. */
const recordOf = (part: OoxmlPart): string | undefined =>
  /<w:sectPrChange [^>]*w:author="Ana"[^>]*><w:sectPr>([\s\S]*?)<\/w:sectPr>/.exec(
    section(part)
  )?.[1];

function decided(part: OoxmlPart, op: 'acceptAllRevisions' | 'rejectAllRevisions'): string {
  const result = applyTreeOp(part, { op });
  if (!result.ok) throw new Error(result.reason);
  return section(result.part);
}

describe('a tracked different first page', () => {
  test("turned on: Word's record states it off; accepted it stays, rejected the section is as before", () => {
    const before = load('');
    const { main } = written(before, { titlePage: true });
    expect(recordOf(main)).toContain('<w:titlePg w:val="0"/>');
    expect(section(main)).toContain('<w:titlePg/><w:docGrid');
    expect(decided(main, 'acceptAllRevisions')).toContain('<w:titlePg/>');
    expect(decided(main, 'acceptAllRevisions')).not.toContain('sectPrChange');
    const original = before.parts.get(before.mainDocumentPart);
    if (!original) throw new Error('no main part');
    expect(decided(main, 'rejectAllRevisions')).toBe(section(original));
  });

  test('turned off: the record states it on, and a reject puts it back in its place', () => {
    const { main } = written(load('<w:titlePg/>'), { titlePage: false });
    expect(recordOf(main)).toContain('<w:titlePg/>');
    expect(section(main).split('<w:sectPrChange')[0]).not.toContain('titlePg');
    expect(decided(main, 'rejectAllRevisions')).toContain('<w:titlePg/><w:docGrid');
  });

  test("Word's own form of turned off: a reject turns it on again, as Word does", () => {
    // Word for Mac (2026-10-30, `~$titlepgoff-tracked.docx` -> `~$titlepgoff-rejected.docx`):
    // the section states `<w:titlePg w:val="0"/>`, the record `<w:titlePg/>`; Reject All gives
    // `<w:titlePg/>` back.
    const record =
      '<w:sectPrChange w:id="1" w:author="Ana" w:date="2026-10-30T10:00:00Z"><w:sectPr>' +
      `${MARGINS}<w:titlePg/></w:sectPr></w:sectPrChange>`;
    const pkg = load('<w:titlePg w:val="0"/>', record);
    const main = pkg.parts.get(pkg.mainDocumentPart);
    if (!main) throw new Error('no main part');
    const rejected = decided(main, 'rejectAllRevisions');
    expect(rejected).toContain('<w:titlePg/><w:docGrid');
    expect(rejected).not.toContain('w:val="0"');
    expect(decided(main, 'acceptAllRevisions')).not.toContain('<w:titlePg/>');
  });

  test('back to the original by the same author leaves no record', () => {
    const on = written(load(''), { titlePage: true });
    const off = written(on.pkg, { titlePage: false });
    expect(section(off.main)).not.toContain('sectPrChange');
    expect(section(off.main)).not.toContain('titlePg');
  });
});

describe('tracked header and footer distances', () => {
  for (const [name, options, now, was] of [
    ['header', { headerDistanceTwips: 1080 }, 'w:header="1080"', 'w:header="720"'],
    ['footer', { footerDistanceTwips: 540 }, 'w:footer="540"', 'w:footer="720"'],
  ] as const)
    test(`the ${name} distance records the earlier margins; a reject gives them back`, () => {
      const { main } = written(load(''), options);
      expect(section(main).split('<w:sectPrChange')[0]).toContain(now);
      expect(recordOf(main)).toContain(was);
      expect(decided(main, 'acceptAllRevisions')).toContain(now);
      expect(decided(main, 'rejectAllRevisions')).toContain(was);
      expect(decided(main, 'rejectAllRevisions')).not.toContain(now);
    });
});

describe('untracked writes', () => {
  test('without a revision the section is written as before', () => {
    const { main } = written(load(''), { titlePage: true, headerDistanceTwips: 1080 }, false);
    expect(section(main)).not.toContain('sectPrChange');
    expect(section(main)).toContain('<w:titlePg/>');
  });

  test('odd and even pages is a setting: written, with no record in any part', () => {
    const { pkg, main } = written(load(''), { evenAndOddHeaders: true });
    expect(section(main)).not.toContain('sectPrChange');
    const settings = pkg.parts.get('/word/settings.xml');
    expect(settings ? serializeOoxmlPart(settings) : '').toContain('evenAndOddHeaders');
  });
});
