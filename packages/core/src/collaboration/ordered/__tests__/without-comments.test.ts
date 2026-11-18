// The Word file without comments (`comments: 'remove'`): no comment in any part, and everything
// else as the file with comments has it.

import { describe, expect, test } from 'bun:test';
import { strToU8, unzipSync, zipSync } from 'fflate';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { diffSemanticDigests, semanticDigest } from '../../../store/package/ooxml-digest.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes, writeOrderedPackage } from '../compose/package.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { apply, at, EDIT } from './change-apply-fixture.ts';
import { commentedFile, storyOf, withEntries } from './comments-fixture.ts';
import { expectSameFingerprint } from './oracle.ts';
import { snapshot } from './snapshot.ts';
import { docx, p, R, t } from './word-fixtures.ts';

const STORIES = ['word/document.xml', 'word/header1.xml', 'word/footer1.xml', 'word/footnotes.xml'];

/** The entries of the comments: the four parts, the comments' relationships and their picture. */
const COMMENT_ENTRIES = [
  'word/comments.xml',
  'word/_rels/comments.xml.rels',
  'word/media/remark.png',
  'word/commentsExtended.xml',
  'word/commentsIds.xml',
  'word/commentsExtensible.xml',
];

/** The entries a removal rewrites: the stories, and the two that name the comment parts. */
const REWRITTEN = new Set([...STORIES, '[Content_Types].xml', 'word/_rels/document.xml.rels']);

interface Opened {
  readonly doc: OrderedDocument;
  readonly base: OoxmlPackage;
}

const opened = (bytes: Uint8Array): Opened => {
  const file = openOrderedWordFile(bytes, { file: 1 });
  return { doc: file.document, base: file.base };
};

const text = (data: Uint8Array): string => new TextDecoder().decode(data);

/** PNG bytes of `size`, different for each `seed`. */
function pictureBytes(seed: number, size = 4096): Uint8Array {
  const bytes = new Uint8Array(size);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  bytes.fill(seed, 8);
  return bytes;
}

/** Each picture inserted as a suggestion at the start of the first paragraph. */
function addPictures(doc: OrderedDocument, pictures: readonly Uint8Array[]): void {
  pictures.forEach((bytes, k) => {
    const media = {
      sha256: orderedBlobHash(bytes),
      contentType: 'image/png' as const,
      size: bytes.length,
    };
    const object = { type: 'picture', media, extent: { cx: 914_400, cy: 457_200 } } as const;
    apply(doc, `ben-1:${k + 1}`, { ...EDIT, kind: 'insert-object', place: at(doc, 0, 0), object });
  });
}

function joined(chunks: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(chunks.reduce((sum, chunk) => sum + chunk.length, 0));
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }

  return out;
}

/** Equal bytes, without printing two whole files when they differ. */
function expectSameBytes(actual: Uint8Array, expected: Uint8Array): void {
  expect(actual.length).toBe(expected.length);
  expect(Buffer.compare(actual, expected)).toBe(0);
}

const CHART_PARTS = ['word/charts/chart1.xml', 'word/charts/style1.xml', 'word/charts/colors1.xml'];
const WORKBOOK = 'word/embeddings/Microsoft_Excel_Worksheet.xlsx';
const CHART_STYLE = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle';
const MS_RELATIONSHIPS = 'http://schemas.microsoft.com/office/2011/relationships';
const CHART_TYPES =
  '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>' +
  '<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' +
  '<Override PartName="/word/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>' +
  '<Override PartName="/word/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>';

const relationship = (id: string, type: string, target: string): string =>
  `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;

/**
 * The commented file with a chart that the comments part names (Word puts none there; a crafted
 * or converted file can): the chart names its workbook, style and colors. With
 * `documentNamesWorkbook`, the document names the workbook too.
 */
function chartInComment(documentNamesWorkbook: boolean): Uint8Array {
  const entries = unzipSync(commentedFile());
  const edit = (name: string, change: (xml: string) => string): void => {
    entries[name] = strToU8(change(text(entries[name]!)));
  };
  const chart = relationship('rIdChart', `${R}/chart`, 'charts/chart1.xml');
  edit('word/_rels/comments.xml.rels', (xml) =>
    xml.replace('</Relationships>', `${chart}</Relationships>`)
  );
  if (documentNamesWorkbook) {
    const book = relationship(
      'rIdBook',
      `${R}/package`,
      'embeddings/Microsoft_Excel_Worksheet.xlsx'
    );
    edit('word/_rels/document.xml.rels', (xml) =>
      xml.replace('</Relationships>', `${book}</Relationships>`)
    );
  }

  edit('[Content_Types].xml', (xml) => xml.replace('</Types>', `${CHART_TYPES}</Types>`));
  entries['word/charts/chart1.xml'] = strToU8(
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
      `xmlns:r="${R}"><c:externalData r:id="rIdBook"/></c:chartSpace>`
  );
  entries['word/charts/_rels/chart1.xml.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      relationship('rIdBook', `${R}/package`, '../embeddings/Microsoft_Excel_Worksheet.xlsx') +
      relationship('rIdStyle', `${MS_RELATIONSHIPS}/chartStyle`, 'style1.xml') +
      relationship('rIdColors', `${MS_RELATIONSHIPS}/chartColorStyle`, 'colors1.xml') +
      '</Relationships>'
  );
  entries['word/charts/style1.xml'] = strToU8(`<cs:chartStyle xmlns:cs="${CHART_STYLE}"/>`);
  entries['word/charts/colors1.xml'] = strToU8(`<cs:colorStyle xmlns:cs="${CHART_STYLE}"/>`);
  entries[WORKBOOK] = zipSync({ 'xl/workbook.xml': strToU8('<workbook/>') });
  return zipSync(entries);
}

/** The reason of the refusal `write` throws, or `written`. */
async function refusalOf(write: () => unknown): Promise<string> {
  try {
    await write();
  } catch (error) {
    if (error instanceof OrderedRefusal) {
      return error.reason;
    }

    throw error;
  }

  return 'written';
}

describe("comments: 'remove'", () => {
  const { doc, base } = opened(commentedFile());
  const keep = orderedPackageBytes(doc, base);
  const remove = orderedPackageBytes(doc, base, { comments: 'remove' });
  const kept = unzipSync(keep);
  const removed = unzipSync(remove);

  test('the fixture has a thread with replies and every comment part', () => {
    expect(Object.keys(kept)).toEqual(expect.arrayContaining(COMMENT_ENTRIES));
    expect(text(kept['word/commentsExtended.xml']!).match(/w15:paraIdParent=/g)).toHaveLength(2);
    for (const name of STORIES) {
      expect(text(kept[name]!)).toContain('<w:commentReference ');
    }
  });

  test('no comment element, attribute, part, relationship or content type is left', () => {
    for (const name of COMMENT_ENTRIES) {
      expect(removed[name]).toBeUndefined();
    }

    const named = Object.entries(removed).filter(
      ([name, data]) => !name.endsWith('.png') && /comment/i.test(`${name} ${text(data)}`)
    );
    expect(named.map(([name]) => name)).toEqual([]);
  });

  test('every other entry is the same bytes, the people part too', () => {
    const others = Object.keys(kept).filter(
      (name) => !REWRITTEN.has(name) && !COMMENT_ENTRIES.includes(name)
    );
    expect(others).toContain('word/people.xml');
    expect(Object.keys(removed).sort()).toEqual([...REWRITTEN, ...others].sort());
    for (const name of others) {
      expectSameBytes(removed[name]!, kept[name]!);
    }
  });

  test('each story is the kept story without its comment nodes; tracked changes stay', () => {
    for (const name of STORIES) {
      expectSameFingerprint(storyOf(remove, name, false), storyOf(keep, name, true), name);
    }

    const body = text(removed['word/document.xml']!);
    expect(body).toContain('<w:ins w:author="Cara" w:date="2026-01-01T00:00:00Z" w:id="101">');
    expect(body).toContain('<w:del w:author="Cara" w:date="2026-01-01T00:00:00Z" w:id="103">');

    // The insertion that held only a reference stays, empty
    expect(body).toContain('<w:ins w:author="Cara" w:date="2026-01-01T00:00:00Z" w:id="102"/>');
  });

  test("the semantic digest is the kept file's without its comment nodes", () => {
    const stripped = STORIES.map((name) => storyOf(keep, name, true));
    const written = STORIES.map((name) => storyOf(remove, name, false));
    expect(diffSemanticDigests(semanticDigest(stripped), semanticDigest(written))).toEqual([]);
  });

  test('the file opens again as an ordered document with nothing left to remove', () => {
    const again = opened(remove);
    expectSameBytes(
      orderedPackageBytes(again.doc, again.base, { comments: 'remove' }),
      orderedPackageBytes(again.doc, again.base)
    );
  });

  test('keep writes the bytes the writer gives with no option', () => {
    expectSameBytes(orderedPackageBytes(doc, base, { comments: 'keep' }), keep);
  });
});

describe('the document does not change', () => {
  test('keep, remove, keep: the same keep bytes both times', async () => {
    const file = opened(commentedFile());
    apply(file.doc, 'ana-1:1', {
      ...EDIT,
      kind: 'insert',
      place: at(file.doc, 0, 6),
      text: 'new ',
    });
    const before = snapshot(file.doc);
    const first = orderedPackageBytes(file.doc, file.base, { comments: 'keep' });
    const chunks: Uint8Array[] = [];
    await writeOrderedPackage(file.doc, file.base, {
      media: async () => new Uint8Array(0),
      write: async (chunk) => {
        chunks.push(chunk);
      },
      comments: 'remove',
    });
    expect(joined(chunks).length).toBeLessThan(first.length);
    expectSameBytes(orderedPackageBytes(file.doc, file.base, { comments: 'keep' }), first);
    expect(snapshot(file.doc)).toEqual(before);
  });

  test('a file with no comments: remove writes the same bytes as keep', () => {
    const ins = '<w:ins w:id="1" w:author="Cara" w:date="2026-01-01T00:00:00Z">';
    const file = opened(docx(p(t('Plain words.')) + p(`${ins}${t('Added.')}</w:ins>`)));
    const keep = orderedPackageBytes(file.doc, file.base);
    expectSameBytes(orderedPackageBytes(file.doc, file.base, { comments: 'remove' }), keep);
    expect(() => opened(keep)).not.toThrow();
  });

  test("the room's comments, replies and resolves are not written, nor their ranges", () => {
    const file = opened(commentedFile());
    apply(file.doc, 'ana-1:1', {
      kind: 'comment',
      part: '/word/document.xml',
      from: at(file.doc, 2, 0),
      to: at(file.doc, 2, 6),
      text: 'Why this date?',
    });
    const [added] = [...file.doc.comments.keys()].slice(-1);
    apply(file.doc, 'ana-1:2', { kind: 'comment-reply', comment: added!, text: 'Agreed.' });
    apply(file.doc, 'ana-1:3', { kind: 'comment-resolve', comment: added!, note: 'Done.' });
    const kept = unzipSync(orderedPackageBytes(file.doc, file.base));
    expect(text(kept['word/comments.xml']!)).toContain('Why this date?');
    expect(text(kept['word/commentsExtended.xml']!)).toContain('w15:done="1"');
    const removed = unzipSync(orderedPackageBytes(file.doc, file.base, { comments: 'remove' }));
    expect(Object.keys(removed).filter((name) => /comment/i.test(name))).toEqual([]);
    expect(text(removed['word/document.xml']!)).not.toContain('comment');
  });

  test("a file with no comments: remove writes no part for the room's comment", () => {
    const file = opened(docx(p(t('Plain words.'))));
    const before = unzipSync(orderedPackageBytes(file.doc, file.base));
    apply(file.doc, 'ana-1:1', {
      kind: 'comment',
      part: '/word/document.xml',
      from: at(file.doc, 0, 0),
      to: at(file.doc, 0, 5),
      text: 'Plain?',
    });
    const kept = unzipSync(orderedPackageBytes(file.doc, file.base));
    expect(Object.keys(kept).filter((name) => /comment/i.test(name))).toEqual([
      'word/comments.xml',
      'word/commentsExtended.xml',
    ]);
    const removed = unzipSync(orderedPackageBytes(file.doc, file.base, { comments: 'remove' }));
    expect(Object.keys(removed)).toEqual(Object.keys(before));
    expect(text(removed['word/document.xml']!)).not.toContain('comment');
    expect(text(removed['[Content_Types].xml']!)).toBe(text(before['[Content_Types].xml']!));
  });
});

describe('writeOrderedPackage with remove', () => {
  test('writes what orderedPackageBytes gives, reading each added picture once, one at a time', async () => {
    const file = opened(commentedFile());
    const pictures = [1, 2].map((seed) => pictureBytes(seed, 300_000));
    addPictures(file.doc, pictures);
    const store = new Map(pictures.map((bytes) => [orderedBlobHash(bytes), bytes]));
    const chunks: Uint8Array[] = [];
    const calls: string[] = [];
    let busy = 0;
    let most = 0;
    const slow = async <T>(work: () => T): Promise<T> => {
      busy += 1;
      most = Math.max(most, busy);
      await new Promise((done) => setTimeout(done, 1));
      busy -= 1;
      return work();
    };
    await writeOrderedPackage(file.doc, file.base, {
      media: (sha256) => {
        calls.push(sha256);
        return slow(() => store.get(sha256) ?? new Uint8Array(0));
      },
      write: (chunk) => slow(() => void chunks.push(chunk)),
      comments: 'remove',
    });
    expect(most).toBe(1);
    expect(calls.sort()).toEqual([...store.keys()].sort());
    expectSameBytes(
      joined(chunks),
      orderedPackageBytes(file.doc, file.base, {
        blobs: (sha256) => store.get(sha256),
        comments: 'remove',
      })
    );
  });

  test("refuses by name when an owner's relationships are kept as bytes", async () => {
    // With no content type for `.rels`, the package keeps every relationships part as bytes
    const bytes = withEntries(commentedFile(), (name, xml) =>
      name === '[Content_Types].xml' ? xml.replace(/<Default Extension="rels"[^>]*\/>/, '') : xml
    );
    const file = opened(bytes);
    expect(orderedPackageBytes(file.doc, file.base).length).toBeGreaterThan(0);
    expect(
      await refusalOf(() => orderedPackageBytes(file.doc, file.base, { comments: 'remove' }))
    ).toBe('engine-refused');
    expect(
      await refusalOf(() =>
        writeOrderedPackage(file.doc, file.base, {
          media: async () => new Uint8Array(0),
          write: async () => {},
          comments: 'remove',
        })
      )
    ).toBe('engine-refused');
  });
});

describe('the parts the comment parts name in turn', () => {
  test('a chart in a comment goes with its workbook, style and colors, and their content types', () => {
    const file = opened(chartInComment(false));
    const kept = unzipSync(orderedPackageBytes(file.doc, file.base));
    expect(Object.keys(kept)).toEqual(expect.arrayContaining([...CHART_PARTS, WORKBOOK]));
    const removed = unzipSync(orderedPackageBytes(file.doc, file.base, { comments: 'remove' }));
    expect(Object.keys(removed).filter((name) => /charts|embeddings/.test(name))).toEqual([]);
    expect(text(removed['[Content_Types].xml']!)).not.toContain('/word/charts/');
  });

  test('a part that another part still names stays', () => {
    const file = opened(chartInComment(true));
    const removed = unzipSync(orderedPackageBytes(file.doc, file.base, { comments: 'remove' }));
    expect(Object.keys(removed).filter((name) => /charts|embeddings/.test(name))).toEqual([
      WORKBOOK,
    ]);
    expect(text(removed['word/_rels/document.xml.rels']!)).toContain('rIdBook');
  });
});
