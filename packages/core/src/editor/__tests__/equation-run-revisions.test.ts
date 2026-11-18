// Word's own tracked changes inside an equation: a `w:del` or a `w:ins` inside a math run
// ([MS-OI29500] 22.1.2.87(e)). Word 16 writes a whole equation deleted with tracking on as one
// `w:del` in each `m:r`, the `m:oMath` kept (Word 16 for Mac, 2026-10-29, sample
// `b-equation-word.docx`). The editor decides them as Word does: one card for the equation,
// Accept removes it, Reject keeps it unmarked; a change to some runs only is decided run by run.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { collectReviewItems, findNode, revisionItemsOf } from '@docx-editor.dev/core/store';
import type { EditorModule } from '../../contracts/modules.ts';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${MAIN}"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:m="${M}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

const STAMP = 'w:author="Grace Hopper" w:date="2026-10-29T23:18:00Z"';
const FONT = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
/** A math run whose content is in one mark, as Word writes it. */
const marked = (mark: 'del' | 'ins', id: number, text: string) =>
  `<m:r><w:${mark} w:id="${id}" ${STAMP}>${FONT}<m:t>${text}</m:t></w:${mark}></m:r>`;
const plain = (text: string) => `<m:r>${FONT}<m:t>${text}</m:t></m:r>`;
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** Word's struck equation, and the struck space after it: three marks, three ids. */
const WORD_STRUCK =
  `<w:p>${run('Alpha ')}<m:oMath>${marked('del', 0, 'x')}${marked('del', 1, '=1')}</m:oMath>` +
  `<w:del w:id="2" ${STAMP}><w:r><w:delText xml:space="preserve"> </w:delText></w:r></w:del>` +
  `${run('Omega.')}</w:p>`;

/** "x=1" made "x=2" with tracking on: one run struck, one inserted. */
const WORD_CHANGED =
  `<w:p>${run('Alpha ')}<m:oMath>${plain('x=')}${marked('del', 3, '1')}${marked('ins', 4, '2')}` +
  `</m:oMath>${run(' Omega.')}</w:p>`;

function reviewModule(): EditorModule {
  return {
    id: 'review',
    review: {
      displayModes: ['all-markup', 'proposed', 'original'],
      collectReviewItems,
      revisionItemsOfParagraph: (part, paragraphId) => {
        const paragraph = findNode(part, paragraphId);
        if (!paragraph || paragraph.kind !== 'paragraph') return [];
        return revisionItemsOf({ ...part, root: paragraph });
      },
    },
  };
}

async function withEditor(
  body: string,
  run: (editor: DocxEditorInstance) => Promise<void> | void
): Promise<void> {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docx(body),
    author: 'Reviewer',
    modules: [reviewModule()],
  });
  if (!editor.surface) throw new Error('surface failed to mount');
  try {
    await run(editor);
  } finally {
    editor.destroy();
    container.remove();
  }
}

async function documentXml(editor: DocxEditorInstance): Promise<string> {
  const files = unzipSync(new Uint8Array(await editor.save()));
  return strFromU8(files['word/document.xml']!);
}

const cards = (editor: DocxEditorInstance) =>
  editor.getReviewItems().flatMap((item) => (item.kind === 'revision' ? [item] : []));

describe("Word's tracked changes inside an equation", () => {
  test('a struck equation is one card: Accept removes it, Reject keeps it unmarked', async () => {
    for (const action of ['accept', 'reject'] as const) {
      await withEditor(WORD_STRUCK, async (editor) => {
        const equation = cards(editor).filter((card) => card.text === 'x=1');
        expect(equation).toHaveLength(1);
        expect(equation[0]!.revisionKind).toBe('delete');
        expect(equation[0]!.readOnly).toBe(false);
        const decided =
          action === 'accept'
            ? editor.acceptReviewItem(equation[0]!.key)
            : editor.rejectReviewItem(equation[0]!.key);
        expect(decided.ok).toBe(true);
        const xml = await documentXml(editor);
        if (action === 'accept') expect(xml).not.toContain('<m:oMath');
        else {
          expect(xml).toMatch(/<m:oMath><m:r><w:rPr>.*<m:t>x<\/m:t><\/m:r><m:r>/);
          expect(xml).toContain('<m:t>=1</m:t>');
        }
        // The struck space is a card of its own, still pending
        expect(xml).toContain('<w:delText xml:space="preserve"> </w:delText>');
        expect(cards(editor).map((card) => card.text)).toEqual([' ']);
      });
    }
  });

  test('Accept All and Reject All decide it', async () => {
    for (const action of ['accept', 'reject'] as const) {
      await withEditor(WORD_STRUCK, async (editor) => {
        const result = editor.exec({ type: 'resolveAllReviewChanges', action, scope: 'document' });
        expect(result.ok).toBe(true);
        const xml = await documentXml(editor);
        expect(xml).not.toMatch(/<w:del /);
        if (action === 'accept') expect(xml).not.toContain('<m:oMath');
        else expect(xml).toContain('<m:t>=1</m:t>');
        expect(cards(editor)).toHaveLength(0);
      });
    }
  });

  test('a change to some runs is decided run by run, and the equation stays', async () => {
    for (const action of ['accept', 'reject'] as const) {
      await withEditor(WORD_CHANGED, async (editor) => {
        expect(cards(editor).map((card) => [card.revisionKind, card.text])).toEqual([
          ['delete', '1'],
          ['insert', '2'],
        ]);
        for (const card of cards(editor)) {
          const decided =
            action === 'accept'
              ? editor.acceptReviewItem(card.key)
              : editor.rejectReviewItem(card.key);
          expect(decided.ok).toBe(true);
        }
        const xml = await documentXml(editor);
        expect(xml).not.toMatch(/<w:(del|ins) /);
        const text = [...xml.matchAll(/<m:t>([^<]*)<\/m:t>/g)].map((m) => m[1]).join('');
        expect(text).toBe(action === 'accept' ? 'x=2' : 'x=1');
      });
    }
  });
});
