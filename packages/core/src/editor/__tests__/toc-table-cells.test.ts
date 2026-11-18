// A table of contents inside a table cell is as read-only as one in the body.
//
// The editor found tables of contents in body paragraphs and body content controls only, so the
// rows of one inside a table cell took typing that the same rows in the body refuse. The heading
// paragraph that Word puts inside the control, before the field, stays editable, as it is in
// Word.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { mount } from './paginated-surface-fixtures.ts';

const READ_ONLY = 'the table of contents is generated and read-only';

const t = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (runs: string) => `<w:p>${runs}</w:p>`;
const fieldStart =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
const fieldEnd = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const control = (content: string) =>
  '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>' +
  `<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const table = (cell: string) =>
  '<w:tbl><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>' +
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>${cell}</w:tc></w:tr></w:tbl>`;

/** A Word table of contents: its heading, the field over two rows, and the field's end. */
const WORD_TOC = control(
  p(t('Contents')) + p(fieldStart + t('Terms\t1')) + p(t('Payment\t2')) + p(fieldEnd)
);
/** A pasted table of contents: the control and its rows, with no field. */
const PASTED_TOC = control(p(t('Terms\t1')) + p(t('Payment\t2')));
/** A TOC field with no control. */
const BARE_FIELD = p(fieldStart + t('Terms\t1')) + p(t('Payment\t2')) + p(fieldEnd);

function typeAt(body: string, paragraphIndex: number, offset: number) {
  const { surface, container } = mount(body + p(t('After.')));
  document.body.append(container);
  try {
    const paragraphId = surface.session.paragraphIds()[paragraphIndex]!;
    const before = surface.session.storyText({ kind: 'body' });
    surface.setSelection({
      anchor: { paragraphId, offset },
      head: { paragraphId, offset },
    });
    surface.type('X');
    return {
      rejection: surface.state().lastRejection,
      changed: surface.session.storyText({ kind: 'body' }) !== before,
      surface: surface as PaginatedSurface,
    };
  } finally {
    surface.destroy();
    container.remove();
  }
}

describe('a table of contents in a table cell is read-only', () => {
  test('a pasted table of contents control in a cell', () => {
    const typed = typeAt(table(PASTED_TOC), 1, 2);
    expect(typed.rejection).toBe(READ_ONLY);
    expect(typed.changed).toBe(false);
  });

  test("a Word table of contents' rows in a cell", () => {
    const typed = typeAt(table(WORD_TOC), 2, 2);
    expect(typed.rejection).toBe(READ_ONLY);
    expect(typed.changed).toBe(false);
  });

  test('a TOC field in a cell, with no control', () => {
    const typed = typeAt(table(BARE_FIELD), 1, 2);
    expect(typed.rejection).toBe(READ_ONLY);
    expect(typed.changed).toBe(false);
  });

  test('the same rows in the body, as before', () => {
    const typed = typeAt(PASTED_TOC, 0, 2);
    expect(typed.rejection).toBe(READ_ONLY);
    expect(typed.changed).toBe(false);
  });
});

describe('the heading inside a table of contents control stays editable, as in Word', () => {
  test('in the body', () => {
    const typed = typeAt(WORD_TOC, 0, 2);
    expect(typed.rejection).toBeNull();
    expect(typed.changed).toBe(true);
  });

  test('in a cell', () => {
    const typed = typeAt(table(WORD_TOC), 0, 2);
    expect(typed.rejection).toBeNull();
    expect(typed.changed).toBe(true);
  });

  test('a paragraph after the table stays editable', () => {
    const typed = typeAt(table(PASTED_TOC), 2, 2);
    expect(typed.rejection).toBeNull();
    expect(typed.changed).toBe(true);
  });
});
