/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Tracked writes over several places, through a browser runtime in a collaboration room: a
// replacement across paragraphs, a table's values, table and cell properties, and a column
// deletion. An ordered room
// takes each as one batch of changes, because its session sets `trackedRangeReplacement`; the
// suggestion reaches every participant, and the organiser's decision gives what the editor
// alone gives. A session without that capability refuses them with `NotSupported`.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  createOrderedCollaboration,
  OrderedLocalRoom,
} from '@docx-editor.dev/core/collaboration/ordered';
import { reviewModule } from '../../../../pro/src/index.ts';
import { DocxEditor } from '../../browser.ts';
import type { RequestContext } from '../request-context.ts';
import { docx, p } from './support/docx.ts';

const cell = (text: string) => `<w:tc>${p(text)}</w:tc>`;
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
// A bookmark from "thirty" to the end of "Late": a range across the two paragraphs
const FILE = docx(
  `<w:p>${run('Payment is due in ')}<w:bookmarkStart w:id="1" w:name="Span"/>${run('thirty days.')}</w:p>` +
    `<w:p>${run('Late')}<w:bookmarkEnd w:id="1"/>${run(' fees apply.')}</w:p>` +
    `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell('Fee')}${cell('Rate')}</w:tr></w:tbl>` +
    p('Signed here.')
);

type Script = (context: RequestContext) => Promise<void>;

/** A tracked replacement from "thirty" in the first paragraph to "Late" in the second. */
const acrossParagraphs: Script = async (context) => {
  context.document.changeTrackingMode = 'TrackMineOnly';
  context.document.getBookmarkRange('Span').insertText('ten days. Higher', 'Replace');
  await context.sync();
};

/** A tracked write of the table's values. */
const tableValues: Script = async (context) => {
  context.document.changeTrackingMode = 'TrackMineOnly';
  context.document.body.tables.getFirst().values = [['Fee', 'Interest']];
  await context.sync();
};

/** A tracked fill of the first cell and a header row: Word's property records. */
const tableProperties: Script = async (context) => {
  context.document.changeTrackingMode = 'TrackMineOnly';
  const table = context.document.body.tables.getFirst();
  table.getCell(0, 0).shadingColor = '#FFCC00';
  table.headerRowCount = 1;
  await context.sync();
};

/** A tracked deletion of the second column: `w:cellDel` on its cell. */
const tableColumn: Script = async (context) => {
  context.document.changeTrackingMode = 'TrackMineOnly';
  context.document.body.tables.getFirst().deleteColumns(1, 1);
  await context.sync();
};

const closers: (() => void)[] = [];
afterEach(() => {
  for (const close of closers.splice(0)) close();
});

function mount(
  document: Uint8Array,
  author: string,
  mode: 'suggesting' | 'edit' = 'suggesting',
  modules: readonly object[] = []
) {
  const container = globalThis.document.createElement('div');
  globalThis.document.body.append(container);
  const editor = createDocxEditor({
    container,
    document,
    author,
    mode,
    modules: [reviewModule(), ...(modules as never[])],
  });
  const runtime = DocxEditor.createBrowser(editor, { author });
  closers.push(() => {
    runtime.dispose();
    editor.destroy();
    container.remove();
  });
  return { editor, runtime };
}

const settle = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await new Promise((resolve) => setTimeout(resolve, 0));
};

/** The body's XML as the editor saves it, without revision ids and dates. */
async function bodyOf(editor: DocxEditorInstance): Promise<string> {
  const saved = new Uint8Array(await editor.save());
  const { unzipSync, strFromU8 } = await import('fflate');
  const xml = strFromU8(unzipSync(saved)['word/document.xml']!);
  // An empty row or cell properties element a rejected proposal leaves behind says nothing
  return xml
    .replace(/^[\s\S]*<w:body>|<w:sectPr[\s\S]*$/g, '')
    .replace(/ w:(id|date|rsid\w*|paraId|textId)="[^"]*"| w14:\w+="[^"]*"/g, '')
    .replace(/<w:(trPr|tcPr)\/>/g, '');
}

async function orderedRoom() {
  const room = new OrderedLocalRoom({ room: 'room-1', wordFile: FILE, file: 1, code: 'test' });
  const join = async (name: string, organiser: boolean) => {
    const person = { id: name.toLowerCase(), name, org: '' };
    const login = {
      person,
      suggests: !organiser,
      decides: organiser,
      readOnly: false,
      comments: true,
    };
    const handle = await createOrderedCollaboration({
      transport: room.transport(login),
      room: 'room-1',
      code: 'test',
      identity: { actorId: name.toLowerCase(), name },
      wordFile: async () => FILE,
      upload: room.uploader(login),
      schedule: null,
    });
    const seat = mount(handle.document, name, organiser ? 'edit' : 'suggesting', [
      { id: 'collaboration', collaboration: { session: handle.session } },
    ]);
    return { ...seat, session: handle.session };
  };
  const ana = await join('Ana', false);
  const olive = await join('Olive', true);
  return { room, ana, olive };
}

for (const [name, script, marks] of [
  ['a replacement across paragraphs', acrossParagraphs, ['<w:del ', '<w:ins ']],
  ['a write of table values', tableValues, ['<w:del ', '<w:ins ']],
  ['a table and cell property change', tableProperties, ['<w:tcPrChange ', '<w:trPrChange ']],
  ['a column deletion', tableColumn, ['<w:cellDel ']],
] as const) {
  test(`an ordered room takes ${name} as a party's suggestion`, async () => {
    const { ana, olive } = await orderedRoom();
    expect(ana.session.trackedRangeReplacement).toBe(true);
    await ana.runtime.run(script);
    await settle();
    // The same suggestion in every copy, as the editor alone writes it
    const alone = mount(FILE, 'Ana');
    await alone.runtime.run(script);
    expect(await bodyOf(olive.editor)).toBe(await bodyOf(alone.editor));
    expect(await bodyOf(ana.editor)).toBe(await bodyOf(alone.editor));
    const pending = await bodyOf(olive.editor);
    for (const mark of marks) expect(pending).toContain(mark);
    for (const action of ['reject', 'accept'] as const) {
      const decided = await orderedRoom();
      await decided.ana.runtime.run(script);
      await settle();
      const command = { type: 'resolveAllReviewChanges', action, scope: 'document' } as const;
      expect(decided.olive.editor.exec(command).ok).toBe(true);
      await settle();
      const reference = mount(FILE, 'Ana');
      await reference.runtime.run(script);
      expect(reference.editor.exec(command).ok).toBe(true);
      const result = await bodyOf(reference.editor);
      expect(result).not.toMatch(/<w:(ins|del|cellDel|\w+Change) /);
      expect(await bodyOf(decided.ana.editor)).toBe(result);
      expect(await bodyOf(decided.olive.editor)).toBe(result);
    }
  });

  test(`a session without the capability refuses ${name}`, async () => {
    const { ana } = await orderedRoom();
    // A session that does not set `trackedRangeReplacement`, as a Yjs session does not
    Object.defineProperty(ana.session, 'trackedRangeReplacement', { value: undefined });
    const before = await bodyOf(ana.editor);
    await expect(ana.runtime.run(script)).rejects.toMatchObject({ code: 'NotSupported' });
    expect(await bodyOf(ana.editor)).toBe(before);
  });
}

test('an ordered room takes a tracked content control as the editor alone does', async () => {
  // The words are struck and a copy goes into the new control; only the control's id differs
  const control: Script = async (context) => {
    context.document.changeTrackingMode = 'TrackMineOnly';
    context.document.body.search('Signed', { matchCase: true }).getFirst().insertContentControl();
    await context.sync();
  };
  const withoutControlId = (xml: string) => xml.replace(/<w:id w:val="\d+"\/>/g, '<w:id/>');
  const alone = mount(FILE, 'Ana');
  await alone.runtime.run(control);
  const expected = withoutControlId(await bodyOf(alone.editor));
  expect(expected).toContain('<w:sdt>');
  const { ana, olive } = await orderedRoom();
  await ana.runtime.run(control);
  await settle();
  expect(withoutControlId(await bodyOf(ana.editor))).toBe(expected);
  expect(withoutControlId(await bodyOf(olive.editor))).toBe(expected);
});
