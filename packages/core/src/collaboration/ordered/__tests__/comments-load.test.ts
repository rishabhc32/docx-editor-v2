// The Word file's comments in the ordered model: each comment with its thread, its resolved
// state and its body as the file wrote it, the same as the editor reads them, through save,
// reopen and a checkpoint.

import { describe, expect, test } from 'bun:test';
import { strToU8, unzipSync, zipSync } from 'fflate';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { diffSemanticDigests, semanticDigest } from '../../../store/package/ooxml-digest.ts';
import type { OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import {
  commentPartNameOf,
  commentsExtendedPartNameOf,
} from '../../../store/store/comment-writes.ts';
import { collectReviewItems } from '../../../store/store/review-reads.ts';
import { openOrderedCheckpoint, openOrderedWordFile } from '../client/open.ts';
import { writeOrderedCheckpoint } from '../compose/checkpoint.ts';
import { writeOrderedPackage } from '../compose/package.ts';
import { loadOrderedPackage, orderedStoryParts } from '../load/load.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import type { OrderedDocument } from '../model/types.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { attributeOf } from '../ooxml-nodes.ts';
import { commentedFile, commentsOf, withEntries } from './comments-fixture.ts';
import { expectSameFingerprint } from './oracle.ts';
import { docx, p, packageOf, t, W, WML_CT } from './word-fixtures.ts';

const COMMENT_PARTS = [
  '/word/comments.xml',
  '/word/commentsExtended.xml',
  '/word/commentsIds.xml',
  '/word/commentsExtensible.xml',
];

const ALL = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10'];

/** The file opened as a room opens it. */
const opened = (bytes: Uint8Array, file = 1) => openOrderedWordFile(bytes, { file });

/** `bytes` without the entry `name`, its content type and its relationship. */
function withoutEntry(bytes: Uint8Array, name: string): Uint8Array {
  const entries = unzipSync(bytes);
  delete entries[name];
  const file = name.slice(name.lastIndexOf('/') + 1);
  const named = new RegExp(`<(?:Override|Relationship) [^>]*${file.replace('.', '\\.')}"/>`, 'g');
  for (const list of ['[Content_Types].xml', 'word/_rels/document.xml.rels']) {
    const text = new TextDecoder().decode(entries[list]);
    entries[list] = strToU8(text.replace(named, ''));
  }

  return zipSync(entries);
}

/** The words of a comment's body. */
function textOf(node: OoxmlNode): string {
  if (node.kind === 'textValue') {
    return node.value;
  }

  return node.children.map(textOf).join('');
}

/** Whether an element holds an element named `localName` anywhere below it. */
function holds(node: OoxmlNode, localName: string): boolean {
  return (
    node.kind !== 'textValue' &&
    node.children.some(
      (child) =>
        child.kind !== 'textValue' && (child.localName === localName || holds(child, localName))
    )
  );
}

/** Each comment's parent, by id: null for a comment that answers none. */
function parentsOf(doc: OrderedDocument): Record<string, string | null> {
  return Object.fromEntries([...doc.comments.values()].map((c) => [c.id, c.parent]));
}

/** The ids of the resolved comments. */
const resolved = (doc: OrderedDocument): string[] =>
  [...doc.comments.values()].filter((c) => c.done).map((c) => c.id);

/** Each comment's parent and resolved state as the editor's own review reads the package. */
function editorThreads(
  pkg: OoxmlPackage
): Record<string, { parent: string | null; done: boolean }> {
  const stories = [...orderedStoryParts(pkg).values()];
  const [storyPart, ...furnitureParts] = stories;
  if (storyPart === undefined) {
    throw new Error('no main document');
  }

  const main = pkg.mainDocumentPart;
  const items = collectReviewItems({
    storyPart,
    furnitureParts,
    commentsPart: pkg.parts.get(commentPartNameOf(pkg, main)),
    commentsExtendedPart: pkg.parts.get(commentsExtendedPartNameOf(pkg, main)),
  });
  const out: Record<string, { parent: string | null; done: boolean }> = {};
  for (const item of items) {
    if (item.kind === 'comment' && out[item.id] === undefined) {
      out[item.id] = { parent: item.parentId ?? null, done: item.resolved };
    }
  }

  return out;
}

/** The model's threads in the shape of {@link editorThreads}. */
function modelThreads(
  doc: OrderedDocument
): Record<string, { parent: string | null; done: boolean }> {
  return Object.fromEntries(
    [...doc.comments.values()].map((c) => [c.id, { parent: c.parent, done: c.done }])
  );
}

/** The whole file `writeOrderedPackage` writes, or a checkpoint. */
async function written(
  write: (options: {
    readonly media: (sha256: string) => Promise<Uint8Array>;
    readonly write: (chunk: Uint8Array) => Promise<void>;
  }) => Promise<void>
): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  await write({
    media: async (sha256) => {
      throw new Error(`no media is read: ${sha256}`);
    },
    write: async (chunk) => {
      chunks.push(chunk);
    },
  });
  return new Uint8Array(Buffer.concat(chunks));
}

const FILE = commentedFile();

/** The thread of comments 0, 1 and 2 resolved. */
const RESOLVED = withEntries(FILE, (name, text) =>
  name === 'word/commentsExtended.xml'
    ? text.replace(/(w15:paraId="1A00000[012]"[^/]*w15:done=)"0"/g, '$1"1"')
    : text
);

/** No `commentsExtended`: comments 1 and 2 cover the same words. */
const NO_EXTENDED = withoutEntry(FILE, 'word/commentsExtended.xml');

/** Comment 3 with two paragraphs, the second one bold. */
const RICH = withEntries(FILE, (name, text) =>
  name === 'word/comments.xml'
    ? text.replace(
        /(<w:comment w:id="3" [^>]*>)<w:p w14:paraId="1A000003"/,
        '$1<w:p w14:paraId="1A0000A3"><w:r><w:t>First line.</w:t></w:r></w:p>' +
          '<w:p w14:paraId="1A000003"><w:r><w:rPr><w:b/></w:rPr><w:t>Bold words.</w:t></w:r></w:p>' +
          '<w:p w14:paraId="1A0000B3"'
      )
    : text
);

/** Thread entries the editor ignores: an unknown parent, a parent of its own, a cycle, no key. */
const MALFORMED = withEntries(FILE, (name, text) => {
  if (name !== 'word/commentsExtended.xml') {
    return text;
  }

  const entry = (paraId: string, parent: string, done = '0') =>
    `<w15:commentEx w15:paraId="${paraId}" w15:paraIdParent="${parent}" w15:done="${done}"/>`;
  return text
    .replace(/<w15:commentEx w15:paraId="1A000001"[^>]*\/>/, entry('1A000001', 'DEADBEEF'))
    .replace(/<w15:commentEx w15:paraId="1A000002"[^>]*\/>/, entry('1A000002', '1A000002'))
    .replace(/<w15:commentEx w15:paraId="1A000003"[^>]*\/>/, entry('1A000003', '1A000004'))
    .replace(/<w15:commentEx w15:paraId="1A000004"[^>]*\/>/, entry('1A000004', '1A000003'))
    .replace(
      '</w15:commentsEx>',
      '<w15:commentEx w15:done="1"/><w15:commentEx w15:paraId="1A000005" w15:done="1"/>' +
        '</w15:commentsEx>'
    );
});

const NONE = docx(p(t('No remark here.')));

const FIXTURES: readonly (readonly [string, Uint8Array])[] = [
  ['every comment place', FILE],
  ['a resolved thread', RESOLVED],
  ['no commentsExtended', NO_EXTENDED],
  ['a rich comment body', RICH],
  ['malformed thread entries', MALFORMED],
  ['no comments', NONE],
];

describe("the Word file's comments in the ordered model", () => {
  test('a comment and its two replies, with every comment part', () => {
    const doc = opened(FILE).document;
    expect([...doc.comments.keys()]).toEqual(ALL);
    expect(parentsOf(doc)).toEqual({
      ...Object.fromEntries(ALL.map((id) => [id, null])),
      '1': '0',
      '2': '0',
    });
    expect(resolved(doc)).toEqual([]);

    const first = doc.comments.get('0');
    expect(first).toMatchObject({
      id: '0',
      author: 'Cara',
      initials: 'C',
      date: '2026-01-01T00:00:00Z',
      paraIds: ['1A000000'],
      paraId: '1A000000',
    });

    // The body as the file wrote it: its words and its picture
    expect(first && textOf(first.element)).toBe('Remark 0.');
    expect(first && holds(first.element, 'drawing')).toBe(true);
  });

  test('a resolved thread', () => {
    expect(resolved(opened(RESOLVED).document)).toEqual(['0', '1', '2']);
  });

  test('a file with no commentsExtended threads comments that cover the same words', () => {
    const doc = opened(NO_EXTENDED).document;
    expect(parentsOf(doc)).toEqual({
      ...Object.fromEntries(ALL.map((id) => [id, null])),
      '2': '1',
    });
  });

  test('comments in a header, a footer, a footnote and a text box', () => {
    const doc = opened(FILE).document;
    const storyOfRange = (id: string): string[] => {
      const stories: string[] = [];
      for (const [name, story] of doc.stories) {
        for (const item of story.tree.toArray()) {
          const node = item.node;
          const marker =
            item.kind === 'zero' &&
            node !== null &&
            node.kind !== 'textValue' &&
            node.localName.startsWith('commentRange') &&
            attributeOf(node, 'id') === id;
          if (marker) {
            stories.push(isTextBoxStoryName(name) ? 'text box' : name);
          }
        }
      }

      return stories;
    };
    expect(storyOfRange('7')).toEqual(['text box', 'text box']);
    expect(storyOfRange('8')).toEqual(['/word/footnotes.xml', '/word/footnotes.xml']);
    expect(storyOfRange('9')).toEqual(['/word/header1.xml', '/word/header1.xml']);
    expect(storyOfRange('10')).toEqual(['/word/footer1.xml', '/word/footer1.xml']);
    for (const id of ['7', '8', '9', '10']) {
      const comment = doc.comments.get(id);
      expect(comment && textOf(comment.element)).toBe(`Remark ${id}.`);
      expect(comment?.parent).toBeNull();
    }
  });

  test('a rich comment body is kept as the file wrote it', () => {
    const { base, document } = opened(RICH);
    const comment = document.comments.get('3');
    expect(comment?.paraIds).toEqual(['1A0000A3', '1A000003', '1A0000B3']);
    expect(comment?.paraId).toBe('1A0000B3');
    expect(comment && textOf(comment.element)).toBe('First line.Bold words.Remark 3.');
    expect(comment && holds(comment.element, 'b')).toBe(true);

    // The file's own element, not a copy
    const element = base.parts
      .get('/word/comments.xml')
      ?.root.children.find(
        (child) => child.kind !== 'textValue' && attributeOf(child, 'id') === '3'
      );
    expect(element).toBe(comment?.element);
  });

  test('malformed thread entries are ignored as the editor ignores them', () => {
    const { base, document } = opened(MALFORMED);
    expect(modelThreads(document)).toEqual(editorThreads(base));
    expect(parentsOf(document)).toMatchObject({ '1': null, '2': null });
    expect(resolved(document)).toEqual(['5']);
  });

  test('every fixture: the threads are the editor review threads', () => {
    for (const [, bytes] of FIXTURES) {
      const { base, document } = opened(bytes);
      expect(modelThreads(document)).toEqual(editorThreads(base));
    }
  });

  test('a file with no comments has none, and its saved file has no comments part', async () => {
    const { base, document } = opened(NONE);
    expect(document.comments.size).toBe(0);
    const saved = await written((options) => writeOrderedPackage(document, base, options));
    expect(Object.keys(unzipSync(saved)).filter((name) => /comment/i.test(name))).toEqual([]);
  });

  test("a comment's nodes count against the load's item limit", () => {
    const words = Array.from({ length: 40 }, (_, k) => `<w:r><w:t>w${k}</w:t></w:r>`).join('');
    const withComment = docx(p(t('Short.')), {
      parts: {
        'word/comments.xml':
          `<w:comments xmlns:w="${W}"><w:comment w:id="1" w:author="Cara">` +
          `<w:p>${words}</w:p></w:comment></w:comments>`,
      },
      contentTypes: `<Override PartName="/word/comments.xml" ContentType="${WML_CT}.comments+xml"/>`,
      relationships:
        '<Relationship Id="rIdComments" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" ' +
        'Target="comments.xml"/>',
    });
    const refusal = (bytes: Uint8Array): string | null => {
      try {
        loadOrderedPackage(packageOf(bytes), { maxItems: 60 });
        return null;
      } catch (error) {
        if (error instanceof OrderedRefusal) {
          return error.reason;
        }

        throw error;
      }
    };
    expect(refusal(docx(p(t('Short.'))))).toBeNull();
    expect(refusal(withComment)).toBe('too-many-items');
  });
});

describe('the comments through save, reopen and a checkpoint', () => {
  for (const [name, bytes] of FIXTURES) {
    test(`${name}: the saved comment parts are the file's, and reopen to the same comments`, async () => {
      const { base, document } = opened(bytes);
      const saved = await written((options) => writeOrderedPackage(document, base, options));
      const savedPackage = packageOf(saved);
      const savedParts = COMMENT_PARTS.flatMap((part) => savedPackage.parts.get(part) ?? []);
      const fileParts = COMMENT_PARTS.flatMap((part) => base.parts.get(part) ?? []);
      expect(savedParts.map((part) => part.name)).toEqual(fileParts.map((part) => part.name));
      savedParts.forEach((part, k) => {
        const filePart = fileParts[k];
        if (filePart === undefined) {
          throw new Error(`no file part for ${part.name}`);
        }

        expectSameFingerprint(part, filePart, part.name);
      });
      expect(diffSemanticDigests(semanticDigest(savedParts), semanticDigest(fileParts))).toEqual(
        []
      );
      expect(commentsOf(opened(saved).document)).toEqual(commentsOf(document));
    });
  }

  test('a checkpoint opens to the same comments, under its own file number and as file 1', async () => {
    for (const [, bytes] of FIXTURES) {
      const { base, document } = opened(bytes, 3);
      const checkpoint = await written((options) =>
        writeOrderedCheckpoint(document, base, options)
      );
      for (const file of [3, 1]) {
        const reopened = openOrderedCheckpoint(checkpoint, { file }).document;
        expect(commentsOf(reopened)).toEqual(commentsOf(document));
      }
    }
  });
});
