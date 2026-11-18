// A tracked `insertFragment`: a paste with Track Changes on, as Word writes one.
//
// Every pasted run sits in `w:ins` with its own `w:rPr` plain; every pasted paragraph mark is
// `w:pPr/w:rPr/w:ins`, with the paragraph's own `w:pPr` plain; every pasted row is
// `w:trPr/w:ins`, with `w:tblPr`, `w:tblGrid` and `w:tcPr` plain. A hyperlink or a simple
// field takes the wrapper inside it. The whole paste is one `w:id`. Accepting it gives the
// untracked paste; rejecting it gives the document from before the paste.

import { describe, expect, test } from 'bun:test';
import { writeOoxmlPackage } from '../package/ooxml-package.ts';
import { serializeOoxmlPart, type OoxmlNode } from '../package/ooxml-tree.ts';
import { extractFragmentPackage } from '../store/clipboard-fragment-extract.ts';
import { revisionItemsOf } from '../store/review-reads.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';
import type { TreePackageStore } from '../store/tree-package-store.ts';
import {
  bodyOf,
  buildPackage,
  fullBodyCoverage,
  loadPackage,
  openStore,
  paragraphIdsUnder,
  R,
  W,
} from './clipboard-fragment-fixtures.ts';
import {
  captureOneJournal,
  openStore as openBytes,
  replayAndCompare,
} from './canonical-primitive-journal-coverage-support.ts';

const DATE = '2026-10-29T00:00:00Z';
const REVISION = { author: 'Ada', date: DATE };
const INS = `w:author="Ada" w:date="${DATE}" w:id="0"`;
const HOST = '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>HostText</w:t></w:r></w:p>';

const NUMBERING =
  `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">` +
  '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
  '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
const STYLES =
  `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Heading1">` +
  '<w:name w:val="heading 1"/></w:style></w:styles>';

/** The body element of the main part, serialized. */
function body(store: TreePackageStore): string {
  const pkg = store.currentPackage();
  const xml = serializeOoxmlPart(pkg.parts.get(pkg.mainDocumentPart)!);
  return xml.slice(xml.indexOf('<w:body>'), xml.indexOf('</w:body>') + '</w:body>'.length);
}

function hostParagraphId(store: TreePackageStore): string {
  return paragraphIdsUnder(bodyOf(store.currentPackage().parts.get('/word/document.xml')!))[0]!;
}

/** Copy the whole of `fragmentBody` the way the clipboard does. */
function clipboard(fragmentBody: string, extra?: Record<string, string>): Uint8Array {
  const source = buildPackage(fragmentBody, extra);
  const extracted = extractFragmentPackage(source, fullBodyCoverage(source));
  if (!extracted.ok) throw new Error('extract failed');
  return extracted.bytes;
}

interface Pasted {
  readonly store: TreePackageStore;
  readonly before: string;
}

function paste(
  fragment: Uint8Array,
  offset: number,
  lastMarkCovered: boolean,
  revision: typeof REVISION | null = REVISION,
  host = HOST
): Pasted {
  const store = openStore(buildPackage(host));
  const before = body(store);
  const result = store.applyFragmentPaste(
    { kind: 'body' },
    {
      paragraphId: hostParagraphId(store),
      offset,
      fragmentBytes: fragment,
      lastMarkCovered,
      ...(revision ? { revision } : {}),
    }
  );
  if (!result.ok) throw new Error(`paste refused: ${result.reason} ${result.detail ?? ''}`);
  return { store, before };
}

function resolveAll(store: TreePackageStore, op: 'acceptAllRevisions' | 'rejectAllRevisions') {
  const result = store.transact({ kind: 'body' }, (ctx) => {
    ctx.apply({ op });
  });
  expect(result.ok).toBe(true);
}

/**
 * Plain runs side by side read as one. Rejecting an insertion leaves the run it split in two
 * halves, which is the same text in the same formatting.
 */
const joinedRuns = (xml: string): string => xml.replaceAll('</w:t></w:r><w:r><w:t>', '');

/** Accept-all equals the untracked paste; reject-all equals the document before it. */
function expectResolves(
  fragment: Uint8Array,
  offset: number,
  lastMarkCovered: boolean,
  host = HOST
): void {
  const tracked = paste(fragment, offset, lastMarkCovered, REVISION, host);
  const untracked = paste(fragment, offset, lastMarkCovered, null, host);
  resolveAll(tracked.store, 'acceptAllRevisions');
  expect(body(tracked.store)).toBe(body(untracked.store));
  tracked.store.undo();
  resolveAll(tracked.store, 'rejectAllRevisions');
  expect(joinedRuns(body(tracked.store))).toBe(tracked.before);
}

/** Every `@w:id` on a `w:ins` in the body. */
function insertionIds(store: TreePackageStore): string[] {
  return [...body(store).matchAll(/<w:ins [^>]*w:id="([^"]*)"/g)].map((match) => match[1]!);
}

describe('tracked insertFragment markup', () => {
  test('inline: the runs go into one w:ins and the host mark is untouched', () => {
    const fragment = clipboard('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:p>');
    const { store } = paste(fragment, 4, false);
    expect(body(store)).toBe(
      '<w:body><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Host</w:t></w:r>' +
        `<w:ins ${INS}><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:ins>` +
        '<w:r><w:t>Text</w:t></w:r></w:p></w:body>'
    );
    expect(revisionItemsOf(store.currentPackage().parts.get('/word/document.xml')!)).toHaveLength(
      1
    );
    expectResolves(fragment, 4, false);
  });

  // Word's own sample (paste evidence 1b): Cara's insertion, deletion and format change, copied
  // from a document that did not track changes, pasted with Track Changes on.
  const CARA = (minute: string) => `w:author="Cara" w:date="2026-01-02T10:${minute}:00Z"`;
  const OTHERS =
    '<w:p><w:r><w:t xml:space="preserve">Start </w:t></w:r>' +
    `<w:ins w:id="1" ${CARA('00')}><w:r><w:t xml:space="preserve">inserted </w:t></w:r></w:ins>` +
    `<w:del w:id="2" ${CARA('01')}><w:r><w:delText xml:space="preserve">deleted </w:delText></w:r></w:del>` +
    `<w:r><w:rPr><w:b/><w:rPrChange w:id="3" ${CARA('02')}><w:rPr/></w:rPrChange></w:rPr>` +
    '<w:t>formatted</w:t></w:r><w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>';

  test("other people's changes: their insertion absorbed, their deletion and format change nested", () => {
    const { store } = paste(clipboard(OTHERS), 0, true);
    const xml = body(store);
    expect(xml).toBe(
      '<w:body>' +
        `<w:p><w:pPr><w:rPr><w:ins ${INS}/></w:rPr></w:pPr><w:ins ${INS}>` +
        '<w:r><w:t xml:space="preserve">Start </w:t></w:r>' +
        '<w:r><w:t xml:space="preserve">inserted </w:t></w:r>' +
        `<w:del ${CARA('01')} w:id="2">` +
        '<w:r><w:delText xml:space="preserve">deleted </w:delText></w:r></w:del>' +
        `<w:r><w:rPr><w:b/><w:rPrChange ${CARA('02')} w:id="3"><w:rPr/></w:rPrChange></w:rPr>` +
        '<w:t>formatted</w:t></w:r><w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:ins></w:p>' +
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>HostText</w:t></w:r></w:p>' +
        '</w:body>'
    );
    // The ids are Word's too (2 and 3); the paste is one change, and Cara's two stay hers
    const items = revisionItemsOf(store.currentPackage().parts.get('/word/document.xml')!);
    expect(items.map((item) => item.author).sort()).toEqual(['Ada', 'Cara', 'Cara']);
  });

  test('a copy from a document that tracks changes is the accepted final view, as Word copies', () => {
    const source = buildPackage(OTHERS);
    const extracted = extractFragmentPackage(source, fullBodyCoverage(source), {
      acceptChanges: true,
    });
    if (!extracted.ok) throw new Error('extract failed');
    // Target not tracking: `~$paste-out-e1e.docx` (Word joins the two plain runs into one)
    const plain = body(paste(extracted.bytes, 0, true, null).store);
    expect(plain).toBe(
      '<w:body><w:p><w:r><w:t xml:space="preserve">Start </w:t></w:r>' +
        '<w:r><w:t xml:space="preserve">inserted </w:t></w:r>' +
        '<w:r><w:rPr><w:b/></w:rPr><w:t>formatted</w:t></w:r>' +
        '<w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>HostText</w:t></w:r></w:p></w:body>'
    );
    // Target tracking: one insertion by the paster, its mark too (`~$paste-out-e1g.docx`)
    const tracked = body(paste(extracted.bytes, 0, true).store);
    expect(tracked).not.toContain('Cara');
    expect(tracked).toContain(`<w:p><w:pPr><w:rPr><w:ins ${INS}/></w:rPr></w:pPr><w:ins ${INS}>`);
  });

  const LIST =
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
    '<w:r><w:t>Item</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>Last</w:t></w:r></w:p>';
  const listFragment = () =>
    clipboard(LIST, { 'word/styles.xml': STYLES, 'word/numbering.xml': NUMBERING });

  test('structural, last mark not covered: pasted marks inserted, the host mark ends the tail', () => {
    const fragment = listFragment();
    const { store } = paste(fragment, 4, false);
    const xml = body(store);
    const numId = /<w:numId w:val="(\d+)"\/>/.exec(xml)?.[1];
    expect(numId).toBeDefined();
    expect(xml).toBe(
      '<w:body>' +
        // The head ends on the fragment's first mark: its properties, inserted.
        `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:rPr><w:ins ${INS}/></w:rPr></w:pPr>` +
        `<w:r><w:t>Host</w:t></w:r><w:ins ${INS}><w:r><w:t>Title</w:t></w:r></w:ins></w:p>` +
        // A whole pasted paragraph: numbering plain, its mark inserted.
        `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>` +
        `<w:rPr><w:ins ${INS}/></w:rPr></w:pPr><w:ins ${INS}><w:r><w:t>Item</w:t></w:r></w:ins></w:p>` +
        // The tail keeps the host's own mark and properties.
        `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:ins ${INS}><w:r><w:t>Last</w:t></w:r></w:ins>` +
        '<w:r><w:t>Text</w:t></w:r></w:p>' +
        '</w:body>'
    );
    expect(xml.includes('pPrChange')).toBe(false);
    expect(new Set(insertionIds(store))).toEqual(new Set(['0']));
    expectResolves(fragment, 4, false);
  });

  test('structural, last mark covered: every pasted mark inserted, the host text after it', () => {
    const fragment = listFragment();
    const { store } = paste(fragment, 4, true);
    const xml = body(store);
    const numId = /<w:numId w:val="(\d+)"\/>/.exec(xml)?.[1];
    expect(xml).toBe(
      '<w:body>' +
        `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:rPr><w:ins ${INS}/></w:rPr></w:pPr>` +
        `<w:r><w:t>Host</w:t></w:r><w:ins ${INS}><w:r><w:t>Title</w:t></w:r></w:ins></w:p>` +
        `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>` +
        `<w:rPr><w:ins ${INS}/></w:rPr></w:pPr><w:ins ${INS}><w:r><w:t>Item</w:t></w:r></w:ins></w:p>` +
        `<w:p><w:pPr><w:rPr><w:ins ${INS}/></w:rPr></w:pPr>` +
        `<w:ins ${INS}><w:r><w:t>Last</w:t></w:r></w:ins></w:p>` +
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>' +
        '</w:body>'
    );
    expectResolves(fragment, 4, true);
  });

  test('a table: rows w:trPr/w:ins, cell marks and runs inserted, table and cell properties plain', () => {
    const fragment = clipboard(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
        '<w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>' +
        '<w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
        '<w:p><w:r><w:t>after</w:t></w:r></w:p>'
    );
    const { store } = paste(fragment, 4, false);
    expect(body(store)).toBe(
      '<w:body>' +
        // The head ends on the mark the split wrote, which the paste inserts.
        `<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:ins ${INS}/></w:rPr></w:pPr>` +
        '<w:r><w:t>Host</w:t></w:r></w:p>' +
        '<w:tbl><w:tblPr><w:tblW w:type="auto" w:w="0"/></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
        `<w:tr><w:trPr><w:cantSplit/><w:ins ${INS}/></w:trPr>` +
        '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2000"/></w:tcPr>' +
        `<w:p><w:pPr><w:rPr><w:ins ${INS}/></w:rPr></w:pPr>` +
        `<w:ins ${INS}><w:r><w:t>cell</w:t></w:r></w:ins></w:p></w:tc></w:tr></w:tbl>` +
        `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:ins ${INS}><w:r><w:t>after</w:t></w:r></w:ins>` +
        '<w:r><w:t>Text</w:t></w:r></w:p>' +
        '</w:body>'
    );
    expect(body(store).includes('cellIns')).toBe(false);
    expectResolves(fragment, 4, false);
  });

  test('a table at the start of the paragraph, or holding a nested table, resolves both ways', () => {
    const table =
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p>' +
      '<w:tbl><w:tblGrid><w:gridCol w:w="900"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
      '<w:p/></w:tc></w:tr></w:tbl>';
    const fragment = clipboard(table + '<w:p><w:r><w:t>after</w:t></w:r></w:p>');
    for (const [offset, covered] of [
      [0, false],
      [8, true],
    ] as const) {
      const { store } = paste(fragment, offset, covered);
      const xml = body(store);
      // Every row the paste wrote is inserted, the nested one too.
      expect(xml.match(/<w:tr>/g)).toHaveLength(2);
      expect(xml.match(/<w:trPr><w:ins /g)).toHaveLength(2);
      expect(new Set(insertionIds(store))).toEqual(new Set(['0']));
      expectResolves(fragment, offset, covered);
    }
  });

  test('a hyperlink takes the w:ins inside it, around its runs', () => {
    const fragment = clipboard(
      `<w:p><w:hyperlink w:anchor="target"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>`
    );
    const { store } = paste(fragment, 4, false);
    expect(body(store)).toBe(
      '<w:body><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Host</w:t></w:r>' +
        `<w:hyperlink w:anchor="target"><w:ins ${INS}><w:r><w:t>link</w:t></w:r></w:ins></w:hyperlink>` +
        '<w:r><w:t>Text</w:t></w:r></w:p></w:body>'
    );
    expectResolves(fragment, 4, false);
  });

  test('an external hyperlink keeps its relationship untracked', () => {
    const fragment = clipboard(
      `<w:p><w:hyperlink xmlns:r="${R}" r:id="rIdLink"><w:r><w:t>web</w:t></w:r></w:hyperlink></w:p>`,
      {
        'word/_rels/document.xml.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          `<Relationship Id="rIdLink" Type="${R}/hyperlink" Target="https://example.com/" TargetMode="External"/>` +
          '</Relationships>',
      }
    );
    const { store } = paste(fragment, 4, false);
    const xml = body(store);
    expect(xml).toMatch(
      new RegExp(`<w:hyperlink [^>]*r:id="[^"]+"><w:ins ${INS}><w:r><w:t>web</w:t></w:r></w:ins>`)
    );
    expectResolves(fragment, 4, false);
  });

  test('a simple field takes the w:ins inside it; a complex field is all inside one', () => {
    const simple = clipboard(
      '<w:p><w:r><w:t>A</w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r>' +
        '</w:fldSimple></w:p>'
    );
    const pastedSimple = paste(simple, 4, false);
    expect(body(pastedSimple.store)).toBe(
      '<w:body><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Host</w:t></w:r>' +
        `<w:ins ${INS}><w:r><w:t>A</w:t></w:r></w:ins>` +
        `<w:fldSimple w:instr=" PAGE "><w:ins ${INS}><w:r><w:t>1</w:t></w:r></w:ins></w:fldSimple>` +
        '<w:r><w:t>Text</w:t></w:r></w:p></w:body>'
    );
    expectResolves(simple, 4, false);

    const complex = clipboard(
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    );
    const pastedComplex = paste(complex, 4, false);
    expect(body(pastedComplex.store)).toBe(
      '<w:body><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Host</w:t></w:r>' +
        `<w:ins ${INS}><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:ins>' +
        '<w:r><w:t>Text</w:t></w:r></w:p></w:body>'
    );
    expectResolves(complex, 4, false);
  });

  test('the paste takes one id past the revisions the document already holds', () => {
    const host =
      '<w:p><w:r><w:t>HostText</w:t></w:r>' +
      '<w:del w:id="7" w:author="Ben"><w:r><w:delText>old</w:delText></w:r></w:del></w:p>';
    const fragment = clipboard(
      '<w:p><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p>'
    );
    const { store } = paste(fragment, 4, true, REVISION, host);
    expect(new Set(insertionIds(store))).toEqual(new Set(['8']));
  });

  test('save and reopen keeps the revisions, and reject still restores the text', () => {
    const fragment = listFragment();
    const { store, before } = paste(fragment, 4, false);
    const saved = body(store);
    const reopened = openStore(loadPackage(writeOoxmlPackage(store.currentPackage())));
    expect(body(reopened)).toBe(saved);
    const cards = (s: TreePackageStore) =>
      revisionItemsOf(s.currentPackage().parts.get('/word/document.xml')!).length;
    expect(cards(reopened)).toBe(cards(store));
    resolveAll(reopened, 'rejectAllRevisions');
    expect(joinedRuns(body(reopened))).toBe(before);
  });
});

describe('tracked insertFragment replication', () => {
  test('a tracked paste is one journal a replica replays to the same package', () => {
    const fragment = clipboard(
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
        '<w:r><w:t>Item</w:t></w:r></w:p>' +
        '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
        '<w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
        '<w:p><w:r><w:t>Last</w:t></w:r></w:p>',
      { 'word/numbering.xml': NUMBERING }
    );
    const bytes = writeOoxmlPackage(buildPackage(HOST));
    const source = openBytes(bytes);
    const replica = openBytes(bytes);
    const captured = captureOneJournal(source, () => {
      const result = source.applyFragmentPaste(
        { kind: 'body' },
        {
          paragraphId: hostParagraphId(source),
          offset: 4,
          fragmentBytes: fragment,
          lastMarkCovered: false,
          revision: REVISION,
        }
      );
      return result.ok ? { ok: true } : { ok: false, reason: result.reason };
    });
    expect(captured.result.ok).toBe(true);
    expect(captured.journal).not.toBeNull();
    replayAndCompare(replica, source, captured.journal!);
  });
});

describe('tracked insertFragment refusals', () => {
  function blocksOf(xml: string): readonly OoxmlNode[] {
    const pkg = buildPackage(xml);
    return bodyOf(pkg.parts.get(pkg.mainDocumentPart)!).children;
  }

  function insert(xml: string, revision = true): { ok: boolean; reason?: string } {
    const store = openStore(buildPackage(HOST));
    const op: TreeDocOp = {
      op: 'insertFragment',
      paragraphId: hostParagraphId(store),
      offset: 4,
      blocks: blocksOf(xml),
      lastMarkCovered: false,
      ...(revision ? { revision: REVISION } : {}),
    };
    const result = store.transact({ kind: 'body' }, (ctx) => {
      ctx.apply(op);
    });
    return result.ok ? { ok: true } : { ok: false, reason: result.reason };
  }

  /** The body after a tracked paste of `xml`, or the refusal. */
  function trackedBody(xml: string): string {
    const store = openStore(buildPackage(HOST));
    const result = store.transact({ kind: 'body' }, (ctx) => {
      ctx.apply({
        op: 'insertFragment',
        paragraphId: hostParagraphId(store),
        offset: 4,
        blocks: blocksOf(xml),
        lastMarkCovered: false,
        revision: REVISION,
      });
    });
    if (!result.ok) return `refused: ${result.reason}`;
    const tracked = body(store);
    // Rejecting the paste takes it out whole, other people's changes in it too
    resolveAll(store, 'rejectAllRevisions');
    expect(joinedRuns(body(store))).toContain('<w:t>HostText</w:t>');
    expect(body(store)).not.toContain('w:author="B"');
    return tracked;
  }

  // Word keeps another author's deletion and property records nested in the paste's insertion,
  // and absorbs that author's insertions (paste evidence 1b, `~$paste-out-e1b.docx`).
  test.each([
    [
      'an insertion (absorbed)',
      '<w:p><w:ins w:id="1" w:author="B"><w:r><w:t>x</w:t></w:r></w:ins></w:p>',
      [`<w:ins ${INS}><w:r><w:t>x</w:t></w:r></w:ins>`],
      false,
    ],
    [
      'a deletion (nested)',
      '<w:p><w:del w:id="1" w:author="B"><w:r><w:delText>x</w:delText></w:r></w:del></w:p>',
      [
        `<w:ins ${INS}><w:del w:author="B" w:id="1"><w:r><w:delText>x</w:delText></w:r></w:del></w:ins>`,
      ],
      true,
    ],
    [
      'a format record (kept)',
      '<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="1" w:author="B"><w:rPr/></w:rPrChange></w:rPr>' +
        '<w:t>x</w:t></w:r></w:p>',
      [`<w:ins ${INS}><w:r><w:rPr><w:b/><w:rPrChange w:author="B"`],
      true,
    ],
    [
      'a paragraph record (kept)',
      '<w:p><w:pPr><w:jc w:val="left"/><w:pPrChange w:id="1" w:author="B"><w:pPr/></w:pPrChange>' +
        '</w:pPr><w:r><w:t>x</w:t></w:r></w:p><w:p/>',
      [`<w:jc w:val="left"/><w:rPr><w:ins ${INS}/></w:rPr><w:pPrChange w:author="B"`],
      true,
    ],
    [
      'a deleted paragraph mark (nested)',
      '<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="B"/></w:rPr></w:pPr>' +
        '<w:r><w:t>x</w:t></w:r></w:p><w:p/>',
      [`<w:rPr><w:ins ${INS}/><w:del w:author="B"`],
      true,
    ],
    [
      'a numbering insertion (absorbed)',
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/>' +
        '<w:ins w:id="1" w:author="B"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p><w:p/>',
      ['<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'],
      false,
    ],
    [
      'a move (destination absorbed, source nested)',
      '<w:p><w:moveFromRangeStart w:id="5" w:name="m1" w:author="B"/>' +
        '<w:moveFrom w:id="6" w:author="B"><w:r><w:t>a</w:t></w:r></w:moveFrom>' +
        '<w:moveFromRangeEnd w:id="5"/><w:moveToRangeStart w:id="7" w:name="m1" w:author="B"/>' +
        '<w:moveTo w:id="8" w:author="B"><w:r><w:t>b</w:t></w:r></w:moveTo>' +
        '<w:moveToRangeEnd w:id="7"/></w:p>',
      [
        `<w:ins ${INS}><w:moveFromRangeStart`,
        '</w:moveFrom><w:moveFromRangeEnd',
        '<w:r><w:t>b</w:t></w:r></w:ins>',
      ],
      true,
    ],
    [
      'a cell insertion (absorbed)',
      '<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:tcPr>' +
        '<w:cellIns w:id="1" w:author="B"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>',
      [`<w:trPr><w:ins ${INS}/></w:trPr><w:tc><w:tcPr/>`],
      false,
    ],
  ])('a fragment holding %s lands as Word tracks it', (_name, xml, pieces, kept) => {
    const tracked = trackedBody(xml);
    for (const piece of pieces) expect(tracked).toContain(piece);
    expect(tracked.includes('w:author="B"')).toBe(kept);
    expect(insert(xml, false).ok).toBe(true);
  });

  test('a fragment holding a custom XML tracked range is refused when tracked', () => {
    const xml =
      '<w:p><w:customXmlInsRangeStart w:id="1" w:author="B"/><w:r><w:t>x</w:t></w:r>' +
      '<w:customXmlInsRangeEnd w:id="1"/></w:p>';
    expect(insert(xml)).toEqual({ ok: false, reason: 'fragment-holds-revisions' });
    expect(insert(xml, false).ok).toBe(true);
  });

  test('a fragment holding content no w:ins can carry is refused when tracked', () => {
    const subDoc = `<w:p><w:subDoc xmlns:r="${R}" r:id="rId9"/></w:p>`;
    expect(insert(subDoc)).toEqual({ ok: false, reason: 'fragment-untrackable-content' });
    const control =
      '<w:sdt><w:sdtPr><w:id w:val="4"/></w:sdtPr><w:sdtContent>' +
      '<w:p><w:r><w:t>inside</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p/>';
    expect(insert(control)).toEqual({ ok: false, reason: 'fragment-untrackable-content' });
    expect(insert(control, false).ok).toBe(true);
    const equation =
      '<w:p><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">' +
      '<m:r><m:t>x</m:t></m:r></m:oMath></w:p>';
    expect(insert(equation)).toEqual({ ok: false, reason: 'fragment-untrackable-content' });
  });

  test('a tracked fragment with an invalid attribution is refused', () => {
    const store = openStore(buildPackage(HOST));
    const result = store.transact({ kind: 'body' }, (ctx) => {
      ctx.apply({
        op: 'insertFragment',
        paragraphId: hostParagraphId(store),
        offset: 4,
        blocks: blocksOf('<w:p><w:r><w:t>x</w:t></w:r></w:p>'),
        revision: { author: '' },
      });
    });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('invalidArgs');
  });
});

describe('the gate a fragment paste asks', () => {
  test('it sees every op the paste commits and the merged package, and can refuse it', () => {
    const fragment = clipboard(
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
        '<w:r><w:t>Item</w:t></w:r></w:p>',
      { 'word/numbering.xml': NUMBERING }
    );
    for (const refusal of [null, 'collaboration-session-not-ready']) {
      const store = openStore(buildPackage(HOST));
      const before = body(store);
      const hostId = hostParagraphId(store);
      const seen: { ops: readonly TreeDocOp[]; numbering: boolean }[] = [];
      const result = store.applyFragmentPaste(
        { kind: 'body' },
        {
          paragraphId: hostId,
          offset: 4,
          fragmentBytes: fragment,
          lastMarkCovered: true,
          revision: REVISION,
          priorOps: [{ op: 'deleteText', paragraphId: hostId, start: 4, end: 8 }],
          gate: (ops, pkg) => {
            seen.push({ ops, numbering: pkg.parts.has('/word/numbering.xml') });
            return refusal;
          },
        }
      );
      expect(seen).toHaveLength(1);
      expect(seen[0]!.ops.map((op) => op.op)).toEqual(['deleteText', 'insertFragment']);
      const insertOp = seen[0]!.ops[1] as Extract<TreeDocOp, { op: 'insertFragment' }>;
      expect(insertOp.revision).toEqual(REVISION);
      expect(insertOp.lastMarkCovered).toBe(true);
      expect(seen[0]!.numbering).toBe(true);
      if (refusal === null) {
        expect(result.ok).toBe(true);
        expect(body(store)).toContain('Item');
      } else {
        expect(result).toEqual({ ok: false, reason: 'invalidArgs', detail: refusal });
        expect(body(store)).toBe(before);
        expect(store.currentPackage().parts.has('/word/numbering.xml')).toBe(false);
      }
    }
  });
});

describe('a fragment pasted at a control prompt', () => {
  // A prompt is state, not text: a paste replaces it as typing does, and as Word pastes over
  // placeholder text. Landed inside it, the pasted blocks became prompt text, still marked
  // `w:showingPlcHdr`.
  const PROMPTED =
    '<w:sdt><w:sdtPr><w:alias w:val="Clause"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent>' +
    '<w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>' +
    '<w:t>Click here to enter text.</w:t></w:r></w:p></w:sdtContent></w:sdt>';
  const PASTED = '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Pasted</w:t></w:r></w:p>';

  test('replaces the prompt with the pasted runs; undo shows the prompt again', () => {
    const { store, before } = paste(clipboard(PASTED), 3, false, null, PROMPTED);
    const xml = body(store);
    expect(xml).not.toContain('showingPlcHdr');
    expect(xml).not.toContain('Click here');
    expect(xml).toContain('<w:b/></w:rPr><w:t>Pasted</w:t>');
    // No empty run is left where the prompt was
    expect(xml).not.toContain('<w:t></w:t>');
    store.undo();
    expect(body(store)).toBe(before);
  });

  test('tracked: one insertion in place of the prompt, and a reject shows the prompt again', () => {
    const { store } = paste(clipboard(PASTED), 3, false, REVISION, PROMPTED);
    const xml = body(store);
    expect(xml).not.toContain('showingPlcHdr');
    expect(xml).toContain(`<w:ins ${INS}><w:r><w:rPr><w:b/></w:rPr><w:t>Pasted</w:t>`);
    // The control names its prompt, so it can show again once the paste goes
    expect(xml).toContain('<w:placeholder>');
    resolveAll(store, 'rejectAllRevisions');
    const rejected = body(store);
    expect(rejected).toContain('<w:showingPlcHdr/>');
    expect(rejected).toContain('Click here to enter text.');
    expect(rejected).not.toContain('Pasted');
  });

  test('at an inline prompt, a complex field elsewhere in the paragraph stays whole', () => {
    // Only the empty run the cleared prompt leaves goes; a field's runs after its begin have
    // no offset length of their own and are not empty
    const field =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> AUTHOR </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Ann</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const host =
      `<w:p>${field}<w:r><w:t xml:space="preserve"> Name: </w:t></w:r>` +
      '<w:sdt><w:sdtPr><w:id w:val="7"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent>' +
      '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>' +
      '<w:t>Click here to enter text.</w:t></w:r></w:sdtContent></w:sdt></w:p>';
    // Offset 8: the field is one unit, " Name: " six more, then the prompt
    const { store } = paste(clipboard(PASTED), 8, false, null, host);
    const xml = body(store);
    expect(xml).not.toContain('showingPlcHdr');
    expect(xml).toContain('<w:t>Pasted</w:t>');
    expect(xml).toContain(field);
  });
});

describe('the gate and the store', () => {
  test('a fragment the store refuses never reaches the gate, so the editor can fall back', () => {
    // Inside a plain text control the store lands no fragment; the editor then pastes plain
    // text, which a session judges on its own. Asked first, the gate refused what the store
    // refuses.
    const host =
      '<w:p><w:r><w:t xml:space="preserve">Alpha </w:t></w:r><w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent>' +
      '<w:r><w:t>Rich</w:t></w:r></w:sdtContent></w:sdt></w:p>';
    const store = openStore(buildPackage(host));
    let asked = 0;
    const result = store.applyFragmentPaste(
      { kind: 'body' },
      {
        paragraphId: hostParagraphId(store),
        offset: 8,
        fragmentBytes: clipboard('<w:p><w:r><w:t>Pasted</w:t></w:r></w:p>'),
        lastMarkCovered: false,
        gate: () => {
          asked += 1;
          return 'collaboration-session-not-ready';
        },
      }
    );
    expect(result.ok).toBe(false);
    expect(result.ok ? null : result.detail).not.toBe('collaboration-session-not-ready');
    expect(asked).toBe(0);
  });
});
