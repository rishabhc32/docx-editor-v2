/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// An agent that edits a live editor through `DocxEditor.createBrowser`: its refused batches reach
// the editor's own diagnostics, so the telemetry module registered on that editor hears them.
import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import type { EditorDiagnostic } from '@docx-editor.dev/core/contracts/editor';
// Source rather than the package name, so the test does not need a built `pro`.
import { reviewModule } from '../../../../pro/src/index.ts';
import { DocxEditor, getReviewItemRevisions } from '../../browser.ts';
import type { DocxEditorError } from '../errors.ts';
import { docx, p } from './support/docx.ts';

test('a refused batch from an in-browser agent is an automation diagnostic', async () => {
  const events: EditorDiagnostic[] = [];
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docx(p('Confidential clause.')),
    modules: [{ id: 'observer', onDiagnostic: (event) => events.push(event) }],
  });
  const runtime = DocxEditor.createBrowser(editor, { author: 'Copilot' });
  try {
    await runtime.run(async (context) => {
      context.document.changeTrackingMode = 'TrackMineOnly';
      context.document.body.clear();
      await context.sync().catch(() => {});
    });
    const rejected = events.filter((event) => event.kind === 'edit.rejected');
    expect(rejected).toEqual([
      expect.objectContaining({
        origin: 'automation',
        tracked: true,
        code: 'edit.tracked-operation-unsupported',
      }),
    ]);
    expect(JSON.stringify(events)).not.toContain('Confidential');
  } finally {
    runtime.dispose();
    editor.destroy();
  }
});

test('a refusal inside an agent batch is reported once, as automation, never as input', async () => {
  const events: EditorDiagnostic[] = [];
  const container = document.createElement('div');
  document.body.append(container);
  const locked =
    '<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>' +
    p('Locked clause.') +
    '</w:sdtContent></w:sdt>';
  const editor = createDocxEditor({
    container,
    document: docx(locked + p('Open clause.')),
    modules: [{ id: 'observer', onDiagnostic: (event) => events.push(event) }],
  });
  const runtime = DocxEditor.createBrowser(editor);
  try {
    await runtime.run(async (context) => {
      const range = context.document.body.search('Locked clause').getFirst();
      range.insertText('Changed', 'Replace');
      await context.sync().catch(() => {});
    });
    const rejected = events.filter((event) => event.kind === 'edit.rejected');
    expect(rejected.map((event) => (event as { origin: string }).origin)).toEqual(['automation']);
  } finally {
    runtime.dispose();
    editor.destroy();
  }
});

test('a browser agent refused by a viewing editor gets the read-only code and hint', async () => {
  const events: EditorDiagnostic[] = [];
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docx(p('Viewing only.')),
    mode: 'view',
    modules: [{ id: 'observer', onDiagnostic: (event) => events.push(event) }],
  });
  const runtime = DocxEditor.createBrowser(editor);
  let failure: unknown;
  try {
    await runtime.run(async (context) => {
      context.document.body.insertParagraph('More', 'End');
      await context.sync().catch((error: unknown) => {
        failure = error;
      });
    });
    expect((failure as DocxEditorError).debugInfo.innerError).toEqual(
      expect.objectContaining({ code: 'edit.read-only' })
    );
    expect(events.filter((event) => event.kind === 'edit.rejected')).toEqual([
      expect.objectContaining({ origin: 'automation', code: 'edit.read-only' }),
    ]);
  } finally {
    runtime.dispose();
    editor.destroy();
  }
});

// Grace's formatting change inside Ada's insertion: two changes, one decision group.
const NESTED =
  '<w:p><w:ins w:id="1" w:author="Ada"><w:r><w:rPr><w:b/>' +
  '<w:rPrChange w:id="2" w:author="Grace"><w:rPr/></w:rPrChange></w:rPr>' +
  '<w:t>bold</w:t></w:r></w:ins></w:p>';

for (const decision of ['its revision', 'its review item key'] as const) {
  test(`a browser agent deciding one change of a group by ${decision} gets the review reason`, async () => {
    const events: EditorDiagnostic[] = [];
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({
      container,
      document: docx(NESTED),
      modules: [reviewModule(), { id: 'observer', onDiagnostic: (event) => events.push(event) }],
    });
    const runtime = DocxEditor.createBrowser(editor);
    const key = editor
      .getReviewItems({ placement: false })
      .find((item) => item.kind === 'revision' && item.author === 'Grace')!.key;
    let failure: unknown;
    try {
      await runtime.run(async (context) => {
        if (decision === 'its revision') {
          const revisions = context.document.revisions;
          revisions.load('items');
          await context.sync();
          revisions.items[1]!.accept();
        } else {
          getReviewItemRevisions(context, key).acceptAll();
        }
        await context.sync().catch((error: unknown) => {
          failure = error;
        });
      });
      expect(failure).toMatchObject({
        code: 'NotImplemented',
        debugInfo: { innerError: { code: 'review.incomplete-group' } },
      });
      expect(events.filter((event) => event.kind === 'edit.rejected')).toEqual([
        expect.objectContaining({ origin: 'automation', cause: 'incomplete-group' }),
      ]);
    } finally {
      runtime.dispose();
      editor.destroy();
    }
  });
}
