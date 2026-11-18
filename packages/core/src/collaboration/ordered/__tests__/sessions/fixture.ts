// The document random sessions edit: many kinds of the item-kind matrix side by side (text,
// formats, a tab, a picture, note references, a bookmark, a comment, a link, controls, a field,
// tracked changes, a list, a heading, a block control, a text box, tables, a merged table and a
// nested one), with the matrix package's notes, comments and numbering, and a header.

import type { OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { loadOrderedPackage } from '../../load/load.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { bodyOf, docx, MATRIX_EXTRAS, MATRIX_ITEMS, packageOf, R, W } from '../word-fixtures.ts';

const KINDS = [
  'bold',
  'tab',
  'picture',
  'fnref',
  'enref',
  'bookmark',
  'comment',
  'link',
  'plain',
  'placeholder',
  'checkbox',
  'pagefield',
  'ins',
  'del',
  'fmt',
  'list',
  'heading',
  'sdtblock',
  'textbox',
  'table',
  'vmerge',
  'nested',
];

const HEADER = `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header text of the agreement.</w:t></w:r></w:p></w:hdr>`;

/** The session file's bytes. */
export const SESSION_FILE = docx(
  KINDS.map((key) => {
    const item = MATRIX_ITEMS.find((candidate) => candidate.key === key);
    if (!item) throw new Error(`no matrix item ${key}`);
    return bodyOf(item);
  }).join(''),
  {
    ...MATRIX_EXTRAS,
    parts: { ...MATRIX_EXTRAS.parts, 'word/header1.xml': HEADER },
    relationships: `${MATRIX_EXTRAS.relationships ?? ''}<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes: `${MATRIX_EXTRAS.contentTypes ?? ''}<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`,
  }
);

/** The session file's package. */
export const sessionPackage = (): OoxmlPackage => packageOf(SESSION_FILE);

/** A fresh copy of the session document. */
export const sessionDocument = (): OrderedDocument => loadOrderedPackage(sessionPackage());
