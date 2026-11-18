// A host selection outside the open story is refused by name.
//
// The file has a header part that no section shows. Its paragraph ids are real, so a host could
// hand one to `setSelection`; the caret then sat in a story nobody had open, and typing changed a
// part no page paints. Word cannot put the caret there. The surface refuses the selection and the
// caret stays where it was.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';
import { OUTSIDE_STORY_REFUSAL } from '../surface-scope.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const HIDDEN = '/word/header2.xml';

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/**
 * A shown default header (header1) and a header part no section references (header2). With
 * `note`, the second body paragraph also holds footnote 1.
 */
function file(note = false): Uint8Array {
  const header = (text: string) => strToU8(`<w:hdr xmlns:w="${W}">${p(text)}</w:hdr>`);
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${MAIN}.document.main+xml"/>` +
        `<Override PartName="/word/header1.xml" ContentType="${MAIN}.header+xml"/>` +
        `<Override PartName="/word/header2.xml" ContentType="${MAIN}.header+xml"/>` +
        (note
          ? `<Override PartName="/word/footnotes.xml" ContentType="${MAIN}.footnotes+xml"/>`
          : '') +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}">` +
        `<Relationship Id="rIdH1" Type="${R}/header" Target="header1.xml"/>` +
        `<Relationship Id="rIdH2" Type="${R}/header" Target="header2.xml"/>` +
        (note ? `<Relationship Id="rIdF" Type="${R}/footnotes" Target="footnotes.xml"/>` : '') +
        '</Relationships>'
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${p('Body one.')}` +
        (note
          ? '<w:p><w:r><w:t>Body two.</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
          : p('Body two.')) +
        '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/>' +
        '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>'
    ),
    'word/header1.xml': header('Shown header.'),
    'word/header2.xml': header('Hidden header.'),
    ...(note
      ? {
          'word/footnotes.xml': strToU8(
            `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1">${p('Note text.')}</w:footnote></w:footnotes>`
          ),
        }
      : {}),
  });
}

function withSurface(run: (surface: PaginatedSurface) => void, note = false): void {
  const container = document.createElement('div');
  document.body.append(container);
  const opened = mountPaginatedSurface(container, file(note), { scale: 1 });
  if (!opened.ok) throw new Error(opened.reason);
  try {
    run(opened.surface);
  } finally {
    opened.surface.destroy();
    container.remove();
  }
}

const at = (paragraphId: string, offset: number) => ({
  anchor: { paragraphId, offset },
  head: { paragraphId, offset },
});

function firstParagraphOf(surface: PaginatedSurface, partName = HIDDEN): string {
  const part = surface.session.currentPackage().parts.get(partName);
  let found: string | null = null;
  const walk = (node: NonNullable<typeof part>['root']): void => {
    if (found || node.kind === 'textValue') return;
    if (node.kind === 'paragraph') found = node.id;
    for (const child of node.children) walk(child);
  };
  if (part) walk(part.root);
  if (!found) throw new Error(`no paragraph in ${partName}`);
  return found;
}

describe('setSelection refuses a paragraph outside the open story', () => {
  test('a header part no section shows: refused by name, and typing stays in the body', () => {
    withSurface((surface) => {
      const body = surface.session.paragraphIds()[0]!;
      surface.setSelection(at(body, 4));
      surface.setSelection(at(firstParagraphOf(surface), 2));
      expect(surface.state().lastRejection).toBe(OUTSIDE_STORY_REFUSAL);
      expect(surface.state().selection.head).toEqual({ paragraphId: body, offset: 4 });
      surface.type('X');
      expect(surface.session.storyText({ kind: 'body' })).toBe('BodyX one.\nBody two.');
      const hidden = serializeOoxmlPart(surface.session.currentPackage().parts.get(HIDDEN)!);
      expect(hidden).toContain('Hidden header.');
      expect(hidden).not.toContain('X');
    });
  });

  test('a body paragraph while a header is open is refused too', () => {
    withSurface((surface) => {
      const body = surface.session.paragraphIds()[0]!;
      expect(surface.enterHeaderFooter?.({ rId: 'rIdH1', pageIndex: 0, sectionIndex: 0 })).toBe(
        true
      );
      const header = firstParagraphOf(surface, '/word/header1.xml');
      surface.setSelection(at(header, 1));
      surface.setSelection(at(body, 2));
      expect(surface.state().lastRejection).toBe(OUTSIDE_STORY_REFUSAL);
      expect(surface.state().selection.head).toEqual({ paragraphId: header, offset: 1 });
    });
  });

  test('a body paragraph while a footnote is open is refused; the note takes its own', () => {
    withSurface((surface) => {
      const body = surface.session.paragraphIds()[0]!;
      expect(surface.enterNote('footnote:1')).toBe(true);
      const note = firstParagraphOf(surface, '/word/footnotes.xml');
      surface.setSelection(at(note, 2));
      expect(surface.state().lastRejection).toBeNull();
      expect(surface.state().selection.head).toEqual({ paragraphId: note, offset: 2 });
      surface.setSelection(at(body, 2));
      expect(surface.state().lastRejection).toBe(OUTSIDE_STORY_REFUSAL);
      expect(surface.state().selection.head).toEqual({ paragraphId: note, offset: 2 });
    }, true);
  });

  test('a paragraph of the open story is still selected', () => {
    withSurface((surface) => {
      const second = surface.session.paragraphIds()[1]!;
      surface.setSelection(at(second, 3));
      expect(surface.state().lastRejection).toBeNull();
      expect(surface.state().selection.head).toEqual({ paragraphId: second, offset: 3 });
    });
  });
});
