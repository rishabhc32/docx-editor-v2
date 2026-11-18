// Typing in SUGGESTING mode beside ANOTHER author's strike, through the keyboard lane.
//
// The store pins the markup (`store/__tests__/tracked-insert-struck-face.test.ts`). What is
// pinned here is the path a reviewer takes: the caret rests before, inside or after struck
// words, `type()` relocates an insert aimed inside them to their end, and the store places
// the words beside the `w:del`. The typed words take the face of the text the caret continues
// — the struck words count — and that is the face the toolbar shows at the caret.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

const FACE = '<w:rPr><w:b/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>';
const run = (text: string) => `<w:r>${FACE}<w:t xml:space="preserve">${text}</w:t></w:r>`;
/** Ana's strike, already in the file when Ben opens it. */
const struck = (text: string, face = FACE) =>
  '<w:del w:id="7" w:author="Ana" w:date="2026-10-18T00:00:00Z">' +
  `<w:r>${face}<w:delText>${text}</w:delText></w:r></w:del>`;

/** Ben's insertion, with the face it was given. */
const BENS_RUN = new RegExp(
  `<w:ins [^>]*w:author="Ben"[^>]*><w:r>${FACE}<w:t xml:space="preserve">X </w:t></w:r></w:ins>`
);

/** Ben, suggesting, puts the caret at `offset` and types "X " a key at a time. */
function benTypesAt(body: string, offset: number, check: (surface: PaginatedSurface) => void) {
  const container = document.createElement('div');
  document.body.append(container);
  const opened = mountPaginatedSurface(container, docx(body), { author: 'Ben' });
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const surface = opened.surface;
    surface.setEditingMode('suggest');
    const paragraphId = surface.session.paragraphIds()[0]!;
    surface.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
    // The toolbar reads the face the next character takes; the typed words must agree.
    expect(surface.formatting().bold).toBe(true);
    for (const character of 'X ') surface.type(character);
    expect(surface.state().lastRejection).toBeNull();
    check(surface);
  } finally {
    opened.surface.destroy();
    container.remove();
  }
}

const xmlOf = (surface: PaginatedSurface) => serializeOoxmlPart(surface.session.part());

describe("typing beside another author's strike, in suggesting mode", () => {
  test("before a strike at the paragraph's start", () => {
    benTypesAt(`<w:p>${struck('Alpha')}${run(' beta gamma')}</w:p>`, 0, (surface) => {
      const xml = xmlOf(surface);
      expect(xml).toMatch(BENS_RUN);
      // Still in FRONT of the struck words: only the face changed.
      expect(xml).toMatch(/<w:p[^>]*><w:ins [^>]*w:author="Ben"[^>]*>.*<\/w:ins><w:del /);
      expect(surface.state().selection.head.offset).toBe('X '.length);
    });
  });

  test('inside, and right after, a strike that ends the paragraph', () => {
    for (const offset of [13, 16]) {
      benTypesAt(`<w:p>${run('Alpha beta ')}${struck('gamma')}</w:p>`, offset, (surface) => {
        const xml = xmlOf(surface);
        expect(xml).toMatch(BENS_RUN);
        expect(xml).toMatch(/<\/w:del><w:ins [^>]*w:author="Ben"[^>]*>.*<\/w:ins><\/w:p>/);
        expect(surface.state().selection.head.offset).toBe('Alpha beta gammaX '.length);
      });
    }
  });

  test('inside a strike that a differently formatted run follows', () => {
    const body =
      `<w:p>${run('Alpha ')}${struck('beta')}` +
      '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> gamma</w:t></w:r></w:p>';
    benTypesAt(body, 8, (surface) => {
      const xml = xmlOf(surface);
      // The struck words are the text the caret continues, not the italic run after them.
      expect(xml).toMatch(BENS_RUN);
      expect(xml).toMatch(
        /<\/w:del><w:ins [^>]*w:author="Ben"[^>]*>.*<\/w:ins><w:r><w:rPr><w:i\/>/
      );
      expect(surface.state().selection.head.offset).toBe('Alpha betaX '.length);
    });
  });
});
