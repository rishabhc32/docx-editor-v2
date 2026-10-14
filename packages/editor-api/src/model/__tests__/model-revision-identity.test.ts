/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `Revision.id` and `Revision.partName`: where a tracked change is written in the file.
//
// DocxEditor additions, recorded in `compat/manifest.json` `omissions`: Office.js has neither. Both
// answer what the file wrote, never an engine identity, so these tests check them against the XML
// and against the review items an editor lists for the same bytes.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { ReviewRevisionItem } from '@docx-editor.dev/core/store';
import { reviewModule } from '@docx-editor.dev/pro';
import { createBrowser } from '../../runtime/browser.ts';
import { createServer } from '../../runtime/server.ts';
import type { DocxEditorRuntime } from '../../runtime/runtime.ts';
import type { RevisionType } from '../review.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DATE = '2026-03-01T10:00:00Z';

const change = (tag: 'ins' | 'del', id: string | null, author: string, text: string): string => {
  const textTag = tag === 'ins' ? 't' : 'delText';
  return (
    `<w:${tag}${id === null ? '' : ` w:id="${id}"`} w:author="${author}" w:date="${DATE}">` +
    `<w:r><w:${textTag} xml:space="preserve">${text}</w:${textTag}></w:r></w:${tag}>`
  );
};
const para = (...content: string[]): string => `<w:p>${content.join('')}</w:p>`;

/**
 * A body and a header that both use `w:id="7"`, which is legal: revision ids are numbered per
 * part. The body also holds one decision over two marks, a typed-over word, and a change with
 * no `w:id`, which is read-only and so never listed.
 */
const REVIEWED = zipSync({
  '[Content_Types].xml': strToU8(
    `<Types xmlns="${CT}">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
      '</Types>'
  ),
  '_rels/.rels': strToU8(
    `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
  ),
  'word/_rels/document.xml.rels': strToU8(
    `<Relationships xmlns="${REL}"><Relationship Id="rId7" Type="${R}/header" Target="header1.xml"/></Relationships>`
  ),
  'word/document.xml': strToU8(
    `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>` +
      para(change('ins', '7', 'Ada', 'added')) +
      para(change('del', '21', 'Ada', 'one '), change('del', '22', 'Ada', 'two')) +
      para(change('del', '31', 'Grace', 'old'), change('ins', '32', 'Grace', 'new')) +
      para(change('ins', null, 'Ada', 'unnumbered')) +
      '<w:sectPr><w:headerReference w:type="default" r:id="rId7"/></w:sectPr>' +
      '</w:body></w:document>'
  ),
  'word/header1.xml': strToU8(
    `<w:hdr xmlns:w="${W}">${para(change('ins', '7', 'Grace', 'header added'))}</w:hdr>`
  ),
});

const BODY = '/word/document.xml';
const HEADER = '/word/header1.xml';

interface Identity {
  readonly id: string;
  readonly partName: string;
  readonly type: RevisionType;
  readonly author: string;
}

const EXPECTED: Identity[] = [
  { id: '7', partName: BODY, type: 'Insert', author: 'Ada' },
  // One decision over `w:id` 21 and 22: the id is its first mark's.
  { id: '21', partName: BODY, type: 'Delete', author: 'Ada' },
  { id: '31', partName: BODY, type: 'Delete', author: 'Grace' },
  { id: '32', partName: BODY, type: 'Insert', author: 'Grace' },
  { id: '7', partName: HEADER, type: 'Insert', author: 'Grace' },
];

/** Every listed revision of the body and the header, with its file identity. */
async function identities(runtime: DocxEditorRuntime): Promise<Identity[]> {
  return runtime.run(async (context) => {
    const body = context.document.revisions;
    const sections = context.document.sections;
    body.load('items');
    sections.load('items');
    await context.sync();
    const header = sections.items[0]!.getHeader('Primary');
    await context.sync();
    const inHeader = header.revisions;
    inHeader.load('items');
    await context.sync();
    const all = [...body.items, ...inHeader.items];
    for (const revision of all) revision.load(['id', 'partName', 'type', 'author']);
    await context.sync();
    return all.map(({ id, partName, type, author }) => ({ id, partName, type, author }));
  });
}

test('a body revision and a header revision read the w:id and the part name the file wrote', async () => {
  const runtime = await createServer(REVIEWED);
  try {
    expect(await identities(runtime)).toEqual(EXPECTED);
  } finally {
    runtime.dispose();
  }
});

test('id and partName load like the other revision properties', async () => {
  const runtime = await createServer(REVIEWED);
  try {
    await runtime.run(async (context) => {
      const revisions = context.document.revisions;
      revisions.load('items');
      await context.sync();
      const revision = revisions.items[0]!;
      expect(() => revision.id).toThrow(expect.objectContaining({ code: 'PropertyNotLoaded' }));
      revision.load('id');
      await context.sync();
      expect(revision.id).toBe('7');
      expect(() => revision.partName).toThrow(
        expect.objectContaining({ code: 'PropertyNotLoaded' })
      );
      revision.load();
      await context.sync();
      expect(revision.partName).toBe(BODY);
    });
  } finally {
    runtime.dispose();
  }
});

/** The revision items a reviewer's editor lists, without the read-only ones `items` omits. */
function revisionItems(
  editor: DocxEditorInstance,
  pairReplacements: boolean
): ReviewRevisionItem[] {
  return editor
    .getReviewItems({ placement: false, pairReplacements })
    .flatMap((entry) => (entry.kind === 'revision' && !entry.readOnly ? [entry.item] : []));
}

test('each revision names the address and part of its review item', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: REVIEWED,
    modules: [reviewModule()],
  });
  const runtime = createBrowser(editor);
  try {
    const read = await identities(runtime);
    expect(read).toEqual(EXPECTED);
    const items = revisionItems(editor, false);
    const grouped = items.find((item) => item.address.id === '21');
    expect(grouped?.addresses.map((address) => address.id)).toEqual(['21', '22']);
    // The editor orders its items by page position rather than story by story: compare as sets.
    const key = (id: string, partName: string, author: string) => `${partName} ${id} ${author}`;
    expect(read.map(({ id, partName, author }) => key(id, partName, author)).sort()).toEqual(
      items.map((item) => key(item.address.id, item.ranges[0]!.partName, item.author)).sort()
    );
  } finally {
    runtime.dispose();
    editor.destroy();
  }
});

test('a paired replacement names both of its revisions through its addresses', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: REVIEWED,
    modules: [reviewModule()],
  });
  const runtime = createBrowser(editor);
  try {
    const read = await identities(runtime);
    const pair = revisionItems(editor, true).find((item) => item.revisionKind === 'replace')!;
    expect(pair.address.id).toBe('31');
    const halves = read.filter(
      (revision) =>
        revision.partName === pair.ranges[0]!.partName &&
        pair.addresses.some(
          (address) => address.id === revision.id && address.author === revision.author
        )
    );
    expect(halves).toEqual([EXPECTED[2], EXPECTED[3]]);
  } finally {
    runtime.dispose();
    editor.destroy();
  }
});
