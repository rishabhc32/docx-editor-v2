/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// One tracked change accepted or rejected alone. When the engine cannot decide it alone, the
// sync refuses with the reason `RevisionCollection.resolve` reports for the same change, in
// `debugInfo.innerError`, and nothing in the batch is applied.
import { expect, test } from 'bun:test';
import type { EditorDiagnostic } from '@docx-editor.dev/core/contracts/editor';
import { explainDiagnostic } from '@docx-editor.dev/core/automation';
import { DocxEditor, isDocxEditorError } from '../../index.ts';
import { docx, p } from './support/docx.ts';

const at = (id: number, author: string) =>
  `w:id="${id}" w:author="${author}" w:date="2026-01-02T03:04:05Z"`;

/** Text Ada inserted, part of which Grace then deleted: one decision group in two changes. */
const NESTED =
  '<w:p><w:r><w:t xml:space="preserve">Start </w:t></w:r>' +
  `<w:ins ${at(1, 'Ada')}><w:r><w:t xml:space="preserve">kept </w:t></w:r>` +
  `<w:del ${at(2, 'Grace')}><w:r><w:delText>gone</w:delText></w:r></w:del></w:ins>` +
  '<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>';

const cell = (content: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
const table = (rows: string) =>
  '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr>' +
  `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows}</w:tbl>`;

/** An inserted row whose cell holds an existing nested table: rejecting it keeps the table. */
const ROW_WITH_NESTED_TABLE = table(
  `<w:tr><w:trPr><w:ins ${at(3, 'Ada')}/></w:trPr>` +
    cell(
      table(`<w:tr>${cell(p('Existing A'))}${cell(p('Existing B'))}</w:tr>`) +
        `<w:p><w:ins ${at(1, 'Ada')}><w:r><w:t>Outer</w:t></w:r></w:ins></w:p>`
    ) +
    cell(`<w:p><w:ins ${at(2, 'Ada')}><w:r><w:t>Neighbour</w:t></w:r></w:ins></w:p>`) +
    '</w:tr>'
);

/** An insertion holding a run whose content carries a marker where only run content can be. */
const MISPLACED_MARKER =
  `<w:p><w:ins ${at(1, 'Ada')}><w:r><w:t xml:space="preserve">outer </w:t></w:r>` +
  `<w:r><w:ins ${at(5, 'Bob')}><w:t>inner</w:t></w:ins></w:r></w:ins></w:p>`;

function observer() {
  const events: EditorDiagnostic[] = [];
  return {
    events,
    module: { id: 'observer', onDiagnostic: (event: EditorDiagnostic) => events.push(event) },
  };
}

for (const action of ['accept', 'reject'] as const) {
  test(`${action} on a nested deletion alone refuses with incomplete-group`, async () => {
    const { events, module } = observer();
    const runtime = await DocxEditor.createServer(docx(NESTED + p('Other paragraph')), {
      modules: [module],
    });
    try {
      const before = await runtime.save();
      await runtime.run(async (context) => {
        const revisions = context.document.revisions;
        const paragraphs = context.document.body.paragraphs;
        revisions.load('items');
        paragraphs.load('items');
        await context.sync();
        for (const revision of revisions.items) revision.load('author');
        await context.sync();
        const nested = revisions.items.find((revision) => revision.author === 'Grace')!;
        // An edit queued in the same batch is not applied either: the batch is one transaction.
        paragraphs.items[1]!.insertText('Added ', 'Start');
        nested[action]();
        const error = await context.sync().catch((caught: unknown) => caught);
        expect(isDocxEditorError(error)).toBe(true);
        if (!isDocxEditorError(error)) return;
        expect(error.code).toBe('NotImplemented');
        expect(error.debugInfo).toMatchObject({
          code: 'NotImplemented',
          errorLocation: `document.revisions.items[1].${action}`,
          innerError: {
            code: 'review.incomplete-group',
            message: explainDiagnostic('review.incomplete-group').hint,
          },
        });
      });
      expect(await runtime.save()).toEqual(before);
      expect(events.filter((event) => event.kind === 'edit.rejected')).toEqual([
        expect.objectContaining({ origin: 'automation', cause: 'incomplete-group' }),
      ]);
      // The reason is the one a batch decision reports for the same change.
      await runtime.run(async (context) => {
        const revisions = context.document.revisions;
        revisions.load('items');
        await context.sync();
        expect(revisions.items).toHaveLength(2);
        const batch = revisions.resolve(action, [revisions.items[1]!]);
        await context.sync();
        expect(batch.value.resolved).toEqual([]);
        expect(batch.value.skipped).toEqual([
          expect.objectContaining({
            reason: 'incomplete-group',
            revision: expect.objectContaining({ author: 'Grace' }),
          }),
        ]);
      });
    } finally {
      runtime.dispose();
    }
  });
}

const OTHER_REASONS = [
  [
    'a row that keeps a nested table',
    ROW_WITH_NESTED_TABLE,
    'reject',
    'retained-structure',
    'review.retained-structure',
  ],
  [
    'a misplaced marker',
    MISPLACED_MARKER,
    'accept',
    'unsupported-revision',
    'edit.revision-unsupported',
  ],
  [
    'a misplaced marker',
    MISPLACED_MARKER,
    'reject',
    'unsupported-revision',
    'edit.revision-unsupported',
  ],
] as const;

for (const [name, body, action, reason, innerError] of OTHER_REASONS) {
  test(`${action} on ${name} refuses with ${reason}`, async () => {
    const runtime = await DocxEditor.createServer(docx(body));
    try {
      const before = await runtime.save();
      await runtime.run(async (context) => {
        const revisions = context.document.revisions;
        revisions.load('items');
        await context.sync();
        expect(revisions.items).toHaveLength(1);
        revisions.items[0]![action]();
        await expect(context.sync()).rejects.toMatchObject({
          code: 'NotImplemented',
          debugInfo: { innerError: { code: innerError } },
        });
      });
      expect(await runtime.save()).toEqual(before);
      // The same reason a batch decision reports. A batch still applies independent parts.
      await runtime.run(async (context) => {
        const revisions = context.document.revisions;
        revisions.load('items');
        await context.sync();
        const batch = revisions.resolve(action, [revisions.items[0]!]);
        await context.sync();
        expect(batch.value.skipped).toEqual([expect.objectContaining({ reason })]);
      });
    } finally {
      runtime.dispose();
    }
  });
}

test('the whole nested group resolves when it is decided together', async () => {
  const runtime = await DocxEditor.createServer(docx(NESTED));
  try {
    await runtime.run(async (context) => {
      const revisions = context.document.revisions;
      revisions.load('items');
      await context.sync();
      const result = revisions.resolve('accept', revisions.items);
      await context.sync();
      expect(result.value).toMatchObject({ skipped: [], remaining: 0 });
      const body = context.document.body;
      body.load('text');
      await context.sync();
      expect(body.text).toBe('Start kept  end');
    });
  } finally {
    runtime.dispose();
  }
});
