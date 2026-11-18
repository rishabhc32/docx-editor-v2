/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// An edit an ordered room refuses, through the editing API. The API used to drop
// such an edit with no name: a sync failed with a bare `GeneralException`. Now the sync names the
// refusal in `debugInfo.innerError` (`edit.collaboration-refused`), and the session's report names
// the room's rule: here a header write, which the room does not carry yet. (`exec` answers
// `ok: false` for a refused write: core's `editor-exec-write-refusal.test.ts`.) The table of
// contents' heading, which the room once refused, takes edits in a room as in the editor alone
// and in Word.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  createOrderedCollaboration,
  OrderedLocalRoom,
  type OrderedClientRefusalReport,
} from '@docx-editor.dev/core/collaboration/ordered';
import { reviewModule } from '../../../../pro/src/index.ts';
import { DocxEditor } from '../../browser.ts';
import { isDocxEditorError } from '../errors.ts';
import { docx, p } from './support/docx.ts';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
/** A Word table of contents: the heading, then the field over one row, in its control. */
const FILE = docx(
  '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>' +
    '<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>' +
    `<w:p>${run('Contents')}</w:p>` +
    '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run('Terms\t1')}</w:p>` +
    '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
    '</w:sdtContent></w:sdt>' +
    p('Terms') +
    p('Body text.')
);

const closers: (() => void)[] = [];
afterEach(() => {
  for (const close of closers.splice(0)) close();
});

const settle = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await new Promise((resolve) => setTimeout(resolve, 0));
};

function mount(document: Uint8Array, author: string, edit: boolean, modules: readonly object[]) {
  const container = globalThis.document.createElement('div');
  globalThis.document.body.append(container);
  const editor = createDocxEditor({
    container,
    document,
    author,
    mode: edit ? 'edit' : 'suggesting',
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

async function seat(room: OrderedLocalRoom, name: string, organiser: boolean) {
  const refusals: OrderedClientRefusalReport[] = [];
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
    onRefused: (report) => refusals.push(report),
  });
  const mounted = mount(handle.document, name, organiser, [
    { id: 'collaboration', collaboration: { session: handle.session } },
  ]);
  await settle();
  return { ...mounted, session: handle.session, refusals };
}

/** Type at the end of a body paragraph through the Office.js shape. */
async function typeAtEnd(
  runtime: ReturnType<typeof mount>['runtime'],
  tracked: boolean,
  k: number
) {
  await runtime.run(async (context) => {
    if (tracked) context.document.changeTrackingMode = 'TrackMineOnly';
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items/text');
    await context.sync();
    paragraphs.items[k]!.getRange('End').insertText('X', 'Replace');
    await context.sync();
  });
}

/** Write in the first section's primary header through the Office.js shape. */
async function writeHeader(runtime: ReturnType<typeof mount>['runtime']) {
  await runtime.run(async (context) => {
    context.document.sections.getFirst().getHeader('Primary').insertText('X', 'End');
    await context.sync();
  });
}

const bodyText = (editor: ReturnType<typeof mount>['editor']) =>
  editor.surface!.session.storyText({ kind: 'body' });

const failure = (write: Promise<void>) =>
  write.then(
    () => null,
    (error: unknown) => error
  );

/** A sync that the room refused: named in `debugInfo.innerError`, with the session's report. */
function expectNamedRefusal(
  thrown: unknown,
  refusals: readonly OrderedClientRefusalReport[],
  report: Partial<OrderedClientRefusalReport>
) {
  expect(isDocxEditorError(thrown)).toBe(true);
  if (!isDocxEditorError(thrown)) return;
  expect(thrown.code).toBe('GeneralException');
  expect(thrown.debugInfo.innerError?.code).toBe('edit.collaboration-refused');
  expect(refusals.at(-1)).toMatchObject(report);
}

describe("an ordered room's refusal through the editing API", () => {
  // The room carries no header or footer write of the editing API yet (the automation
  // API in rooms comes later), so it refuses one by name; the editor alone takes it. The
  // organiser and a party meet the same refusal: other room refusals are not reachable this way
  // (the editor refuses a tracked edit inside another person's insertion first, and a party's
  // decision is refused after the sync, through the session's report)
  for (const [name, organiser] of [
    ['Olive', true],
    ['Ana', false],
  ] as const) {
    test(`${name}: a header write through the API is refused by name`, async () => {
      const room = new OrderedLocalRoom({ room: 'room-1', wordFile: FILE, file: 1, code: 'test' });
      const person = await seat(room, name, organiser);
      const before = bodyText(person.editor);
      expectNamedRefusal(await failure(writeHeader(person.runtime)), person.refusals, {
        code: 'ordered-opaque-write',
        reason: 'opaque-write',
      });
      await settle();
      expect(bodyText(person.editor)).toBe(before);
      expect(person.session.unsentChanges()).toEqual([]);
    });
  }

  // The room refuses no more than the editor alone in a table of contents: its heading takes edits
  for (const [name, organiser] of [
    ['Olive', true],
    ['Ana', false],
  ] as const) {
    test(`${name}: the table of contents' heading is editable in a room`, async () => {
      const room = new OrderedLocalRoom({ room: 'room-1', wordFile: FILE, file: 1, code: 'test' });
      const person = await seat(room, name, organiser);
      await typeAtEnd(person.runtime, !organiser, 0);
      await settle();
      expect(bodyText(person.editor)?.split('\n')[0]).toBe('ContentsX');
      expect(person.refusals).toEqual([]);
    });
  }

  test('without a room the heading is editable, as in Word', async () => {
    const alone = mount(FILE, 'Olive', true, []);
    await typeAtEnd(alone.runtime, false, 0);
    expect(bodyText(alone.editor)?.split('\n')[0]).toBe('ContentsX');
  });
});
