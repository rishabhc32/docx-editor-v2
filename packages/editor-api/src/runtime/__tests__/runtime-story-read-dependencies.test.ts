/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A header, footer or note body is a promise until its owner's read answers. Its members queue
// behind that read in the SAME sync, as Office.js allows, rather than refusing with
// `InvalidObjectPath` until an earlier `load('text')` and `sync()` resolved the story.

import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { DocxEditor } from '../../index.ts';
import {
  WITH_FURNITURE,
  docx,
  p,
  reopen,
  withNumbering,
} from '../../model/__tests__/support/documents.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const ins = (id: number, author: string, text: string): string =>
  `<w:ins w:id="${id}" w:author="${author}"><w:r><w:t>${text}</w:t></w:r></w:ins>`;
const del = (id: number, author: string, text: string): string =>
  `<w:del w:id="${id}" w:author="${author}"><w:r><w:delText>${text}</w:delText></w:r></w:del>`;

/**
 * The furniture fixture with a tracked change in every story, first-page and even-page headers,
 * a nested control and a bookmark in the primary header, and a list paragraph in its footer.
 */
const STORIES: Uint8Array = (() => {
  const parts = unzipSync(withNumbering(WITH_FURNITURE, ['1']));
  const edit = (name: string, from: string, to: string): void => {
    parts[name] = strToU8(strFromU8(parts[name]!).replace(from, to));
  };
  edit('word/document.xml', '<w:r><w:t>in the body</w:t></w:r>', ins(10, 'Bea', 'in the body'));
  edit(
    'word/document.xml',
    '<w:footerReference',
    '<w:headerReference w:type="first" r:id="rIdFirst"/>' +
      '<w:headerReference w:type="even" r:id="rIdEven"/><w:footerReference'
  );
  edit(
    'word/_rels/document.xml.rels',
    '</Relationships>',
    `<Relationship Id="rIdFirst" Type="${OFFICE}/header" Target="header2.xml"/>` +
      `<Relationship Id="rIdEven" Type="${OFFICE}/header" Target="header3.xml"/>` +
      `<Relationship Id="rIdStyles" Type="${OFFICE}/styles" Target="styles.xml"/></Relationships>`
  );
  edit(
    '[Content_Types].xml',
    '</Types>',
    `<Override PartName="/word/header2.xml" ContentType="${WML}.header+xml"/>` +
      `<Override PartName="/word/header3.xml" ContentType="${WML}.header+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${WML}.styles+xml"/></Types>`
  );
  edit(
    'word/footer1.xml',
    '<w:p><w:r><w:t>in the footer</w:t></w:r>',
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
      `<w:r><w:t xml:space="preserve">Footer </w:t></w:r>${del(14, 'Grace', 'gone')}`
  );
  edit('word/footnotes.xml', '<w:r><w:t>in the footnote</w:t></w:r>', ins(15, 'Noor', 'note'));
  parts['word/header1.xml'] = strToU8(
    `<w:hdr xmlns:w="${W}"><w:p><w:bookmarkStart w:id="1" w:name="letterhead"/>` +
      '<w:r><w:t xml:space="preserve">Primary </w:t></w:r><w:bookmarkEnd w:id="1"/>' +
      `${ins(11, 'Ada', 'inserted')}</w:p>` +
      '<w:sdt><w:sdtPr><w:tag w:val="outer"/></w:sdtPr><w:sdtContent>' +
      '<w:sdt><w:sdtPr><w:tag w:val="inner"/></w:sdtPr><w:sdtContent>' +
      '<w:p><w:r><w:t>deep</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt></w:hdr>'
  );
  parts['word/header2.xml'] = strToU8(
    `<w:hdr xmlns:w="${W}"><w:p>${del(12, 'Grace', 'removed')}</w:p></w:hdr>`
  );
  parts['word/header3.xml'] = strToU8(
    `<w:hdr xmlns:w="${W}"><w:p>${ins(13, 'Linus', 'added')}</w:p></w:hdr>`
  );
  parts['word/styles.xml'] = strToU8(
    `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/></w:style></w:styles>`
  );
  return zipSync(parts);
})();

test('header and footer revisions load in the same sync that names the story', async () => {
  const runtime = await DocxEditor.createServer(STORIES);
  try {
    const changes = await runtime.run(async (context) => {
      const section = context.document.sections.getFirst();
      const collections = [
        section.getHeader('Primary').revisions,
        section.getHeader('FirstPage').revisions,
        section.getHeader('EvenPages').revisions,
        section.getFooter('Primary').revisions,
        context.document.footnotes.getFirst().body.revisions,
        context.document.revisions,
      ];
      for (const collection of collections) collection.load('items/author,items/type');
      await context.sync();
      return collections.map((collection) =>
        collection.items.map((revision) => ({ author: revision.author, type: revision.type }))
      );
    });
    expect(changes).toEqual([
      [{ author: 'Ada', type: 'Insert' }],
      [{ author: 'Grace', type: 'Delete' }],
      [{ author: 'Linus', type: 'Insert' }],
      [{ author: 'Grace', type: 'Delete' }],
      [{ author: 'Noor', type: 'Insert' }],
      [{ author: 'Bea', type: 'Insert' }],
    ]);
  } finally {
    runtime.dispose();
  }
});

test('an absent header or footer answers empty revisions in the same sync and creates nothing', async () => {
  const runtime = await DocxEditor.createServer(docx(p('Body only') + '<w:sectPr/>'));
  try {
    const before = await runtime.save();
    const counts = await runtime.run(async (context) => {
      const section = context.document.sections.getFirst();
      const stories = (['Primary', 'FirstPage', 'EvenPages'] as const).flatMap((type) => [
        section.getHeader(type),
        section.getFooter(type),
      ]);
      for (const story of stories) story.revisions.load('items');
      await context.sync();
      return stories.map((story) => story.revisions.items.length);
    });
    expect(counts).toEqual([0, 0, 0, 0, 0, 0]);
    expect(await runtime.save()).toEqual(before);
  } finally {
    runtime.dispose();
  }
});

test('collections and searches on a header or footer resolve in the same sync', async () => {
  const runtime = await DocxEditor.createServer(STORIES);
  try {
    const found = await runtime.run(async (context) => {
      const section = context.document.sections.getFirst();
      const header = section.getHeader('Primary');
      const bookmarks = header.bookmarks.load('items/name');
      const outer = header.contentControls.getFirst().load('tag');
      const inner = outer.contentControls.load('items/tag');
      const comments = header.getComments().load('items');
      const hits = header.search('inserted').load('items/text');
      const lists = section.getFooter('Primary').lists.load('items');
      await context.sync();
      return {
        bookmarks: bookmarks.items.map((bookmark) => bookmark.name),
        controls: [outer.tag, ...inner.items.map((control) => control.tag)],
        comments: comments.items.length,
        hits: hits.items.map((hit) => hit.text),
        lists: lists.items.length,
      };
    });
    expect(found).toEqual({
      bookmarks: ['letterhead'],
      controls: ['outer', 'inner'],
      comments: 0,
      hits: ['inserted'],
      lists: 1,
    });
  } finally {
    runtime.dispose();
  }
});

test('story-wide writes on a header, footer or note need no earlier sync', async () => {
  const runtime = await DocxEditor.createServer(STORIES);
  try {
    // One story per write sync, so each edit is its own run.
    await runtime.run(async (context) => {
      context.document.sections.getFirst().getHeader('Primary').revisions.acceptAll();
      await context.sync();
    });
    await runtime.run(async (context) => {
      context.document.sections.getFirst().getFooter('Primary').insertParagraph('Appended', 'End');
      await context.sync();
    });
    await runtime.run(async (context) => {
      context.document.sections.getFirst().getHeader('EvenPages').style = 'Quote';
      await context.sync();
    });
    await runtime.run(async (context) => {
      context.document.footnotes.getFirst().body.clear();
      await context.sync();
    });
    const saved = await reopen(runtime);
    try {
      const state = await saved.run(async (context) => {
        const section = context.document.sections.getFirst();
        const header = section.getHeader('Primary').load('text');
        const revisions = header.revisions.load('items');
        const footer = section.getFooter('Primary').load('text');
        const even = section.getHeader('EvenPages').load('style');
        const note = context.document.footnotes.getFirst().body.load('text');
        await context.sync();
        return {
          header: header.text,
          revisions: revisions.items.length,
          footer: footer.text,
          even: even.style,
          note: note.text,
        };
      });
      expect(state).toEqual({
        header: 'Primary inserted\rdeep',
        revisions: 0,
        footer: 'Footer gone\rAppended',
        even: 'Quote',
        note: '',
      });
    } finally {
      saved.dispose();
    }
  } finally {
    runtime.dispose();
  }
});
