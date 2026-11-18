// A Word file with comments in every place Word puts them: the body (a thread with replies, a
// range across a table, references inside a tracked insertion and a tracked deletion), a text
// box, a header, a footer and a footnote, with the four comment parts and the people part.

import { strToU8, unzipSync, zipSync } from 'fflate';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPart,
  serializeOoxmlPart,
  type OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  docx,
  MC,
  p,
  packageOf,
  pictureRun,
  R,
  RED_PNG,
  t,
  V,
  W,
  WML_CT,
  WPS,
} from './word-fixtures.ts';

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
const W16CEX = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MS_REL = 'http://schemas.microsoft.com/office';
const REV = 'w:author="Cara" w:date="2026-01-01T00:00:00Z"';

export const start = (id: number) => `<w:commentRangeStart w:id="${id}"/>`;
export const end = (id: number) => `<w:commentRangeEnd w:id="${id}"/>`;
export const reference = (id: number) =>
  `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;

/** A comment's whole range around `xml`, with its reference after the end. */
const around = (id: number, xml: string) => start(id) + xml + end(id) + reference(id);

const paraId = (id: number) => `1A0000${id.toString(16).padStart(2, '0').toUpperCase()}`;

/** The thread: comment 0, answered by 1 and 2. */
const REPLIES: Readonly<Record<number, number>> = { 1: 0, 2: 0 };

/** Comment ids: 0-2 the thread, 3 across a table, 4-6 tracked changes, 7 the text box, 8-10. */
const IDS = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10];

const textBox = (words: string) =>
  `<w:r><mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:wps="${WPS}" Requires="wps"><w:drawing>` +
  '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/>' +
  `<wp:docPr id="8" name="Text Box 8"/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  `<wps:txbx><w:txbxContent>${words}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>` +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>' +
  `<mc:Fallback><w:pict><v:shape xmlns:v="${V}" style="width:72pt;height:36pt">` +
  `<v:textbox><w:txbxContent>${words}</w:txbxContent></v:textbox></v:shape></w:pict>` +
  '</mc:Fallback></mc:AlternateContent></w:r>';

const cell = (content: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${content}</w:tc>`;

/** The body: every comment place of the main document part. */
export const COMMENTED_BODY =
  p(start(0) + start(1) + start(2) + t('Clause one.') + end(0) + reference(0)) +
  p(t('Its second line.') + end(1) + reference(1) + end(2) + reference(2)) +
  p(t('Before the table, ') + start(3) + t('across it.')) +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>` +
  `<w:tr>${cell(p(t('In the cell.')))}</w:tr></w:tbl>` +
  p(t('After the table.') + end(3) + reference(3)) +
  p(
    t('Kept ') +
      start(4) +
      `<w:ins w:id="101" ${REV}>${t('inserted words')}${end(4)}${reference(4)}</w:ins>` +
      start(5) +
      t('plain') +
      end(5) +
      `<w:ins w:id="102" ${REV}>${reference(5)}</w:ins>`
  ) +
  p(
    t('Kept ') +
      start(6) +
      `<w:del w:id="103" ${REV}><w:r><w:delText>deleted words</w:delText></w:r>${end(6)}` +
      `${reference(6)}</w:del>`
  ) +
  p(t('A box: ') + textBox(p(around(7, t('Box text'))))) +
  p(t('A note.') + '<w:r><w:footnoteReference w:id="1"/></w:r>') +
  '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/>' +
  '<w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr>';

const separators =
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';

/** A comment's body, with a picture in comment 0. */
const commentBody = (id: number) =>
  `<w:comment w:id="${id}" ${REV} w:initials="C"><w:p w14:paraId="${paraId(id)}" w14:textId="77777777">` +
  '<w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:annotationRef/></w:r>' +
  `${t(`Remark ${id}.`)}${id === 0 ? pictureRun(30) : ''}</w:p></w:comment>`;

const COMMENT_PARTS = {
  'word/comments.xml':
    `<w:comments xmlns:w="${W}" xmlns:w14="${W14}" xmlns:r="${R}" ` +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `${IDS.map(commentBody).join('')}</w:comments>`,
  'word/_rels/comments.xml.rels':
    `<Relationships xmlns="${REL}">` +
    `<Relationship Id="rIdImg1" Type="${R}/image" Target="media/remark.png"/></Relationships>`,
  'word/media/remark.png': RED_PNG,
  'word/commentsExtended.xml':
    `<w15:commentsEx xmlns:w15="${W15}">` +
    IDS.map((id) => {
      const parent = REPLIES[id];
      const link = parent === undefined ? '' : ` w15:paraIdParent="${paraId(parent)}"`;
      return `<w15:commentEx w15:paraId="${paraId(id)}"${link} w15:done="0"/>`;
    }).join('') +
    '</w15:commentsEx>',
  'word/commentsIds.xml':
    `<w16cid:commentsIds xmlns:w16cid="${W16CID}">` +
    IDS.map(
      (id) =>
        `<w16cid:commentId w16cid:paraId="${paraId(id)}" w16cid:durableId="2B0000${id + 10}"/>`
    ).join('') +
    '</w16cid:commentsIds>',
  'word/commentsExtensible.xml':
    `<w16cex:commentsExtensible xmlns:w16cex="${W16CEX}">` +
    IDS.map(
      (id) =>
        `<w16cex:commentExtensible w16cex:durableId="2B0000${id + 10}" w16cex:dateUtc="2026-01-01T00:00:00Z"/>`
    ).join('') +
    '</w16cex:commentsExtensible>',
};

const STORY_PARTS = {
  'word/header1.xml': `<w:hdr xmlns:w="${W}" xmlns:r="${R}">${p(around(9, t('Header words')))}</w:hdr>`,
  'word/footer1.xml': `<w:ftr xmlns:w="${W}" xmlns:r="${R}">${p(around(10, t('Footer words')))}</w:ftr>`,
  'word/footnotes.xml':
    `<w:footnotes xmlns:w="${W}" xmlns:r="${R}">${separators}<w:footnote w:id="1">` +
    `<w:p><w:r><w:footnoteRef/></w:r>${around(8, t(' Fees are due net of tax.'))}</w:p>` +
    '</w:footnote></w:footnotes>',
  'word/people.xml':
    `<w15:people xmlns:w15="${W15}"><w15:person w15:author="Cara">` +
    '<w15:presenceInfo w15:providerId="None" w15:userId="Cara"/></w15:person></w15:people>',
};

const override = (part: string, type: string) =>
  `<Override PartName="/word/${part}" ContentType="${type}"/>`;

/** The Word file with every comment place, its four comment parts and its people part. */
export function commentedFile(body = COMMENTED_BODY): Uint8Array {
  return docx(body, {
    parts: { ...STORY_PARTS, ...COMMENT_PARTS },
    contentTypes:
      '<Default Extension="png" ContentType="image/png"/>' +
      override('header1.xml', `${WML_CT}.header+xml`) +
      override('footer1.xml', `${WML_CT}.footer+xml`) +
      override('footnotes.xml', `${WML_CT}.footnotes+xml`) +
      override('people.xml', `${WML_CT}.people+xml`) +
      override('comments.xml', `${WML_CT}.comments+xml`) +
      override('commentsExtended.xml', `${WML_CT}.commentsExtended+xml`) +
      override('commentsIds.xml', `${WML_CT}.commentsIds+xml`) +
      override('commentsExtensible.xml', `${WML_CT}.commentsExtensible+xml`),
    relationships:
      `<Relationship Id="rIdHeader" Type="${R}/header" Target="header1.xml"/>` +
      `<Relationship Id="rIdFooter" Type="${R}/footer" Target="footer1.xml"/>` +
      `<Relationship Id="rIdNotes" Type="${R}/footnotes" Target="footnotes.xml"/>` +
      `<Relationship Id="rIdPeople" Type="${MS_REL}/2011/relationships/people" Target="people.xml"/>` +
      `<Relationship Id="rIdComments" Type="${R}/comments" Target="comments.xml"/>` +
      `<Relationship Id="rIdEx" Type="${MS_REL}/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>` +
      `<Relationship Id="rIdIds" Type="${MS_REL}/2016/relationships/commentsIds" Target="commentsIds.xml"/>` +
      `<Relationship Id="rIdCex" Type="${MS_REL}/2018/08/relationships/commentsExtensible" Target="commentsExtensible.xml"/>`,
  });
}

/** `bytes` with each entry's text changed by `edit`, for a file the fixture builder cannot make. */
export function withEntries(
  bytes: Uint8Array,
  edit: (name: string, text: string) => string
): Uint8Array {
  const entries = unzipSync(bytes);
  const out: Record<string, Uint8Array> = {};
  for (const [name, data] of Object.entries(entries)) {
    out[name] = name.endsWith('.png') ? data : strToU8(edit(name, new TextDecoder().decode(data)));
  }

  return zipSync(out);
}

/** A story's XML with the comment nodes taken out by text: each range edge, each reference run. */
const withoutCommentText = (xml: string): string =>
  xml
    .replace(/<w:commentRange(?:Start|End) w:id="[^"]*"\/>/g, '')
    .replace(
      /<w:r><w:rPr>(?:(?!<\/w:rPr>).)*<\/w:rPr><w:commentReference w:id="[^"]*"\/><\/w:r>/g,
      ''
    );

/** The file's story part `name`, with the comment nodes taken out when `strip` is true. */
export function storyOf(bytes: Uint8Array, name: string, strip: boolean): OoxmlPart {
  const part = packageOf(bytes).parts.get(`/${name}`);
  if (!part) {
    throw new Error(`no part ${name}`);
  }

  if (!strip) {
    return part;
  }

  const read = readOoxmlPart(withoutCommentText(serializeOoxmlPart(part)), {
    name: part.name,
    contentType: part.contentType,
  });
  if (!read.ok) {
    throw new Error(`unreadable ${name}: ${read.reason}`);
  }

  return read.part;
}

/** A document's comments as plain data: each record, with its nodes by their fingerprints. */
export function commentsOf(doc: OrderedDocument): unknown[] {
  return [...doc.comments.values()].map(({ element, ...fields }) => ({
    ...fields,
    element: canonicalOoxmlFingerprint(element),
  }));
}
