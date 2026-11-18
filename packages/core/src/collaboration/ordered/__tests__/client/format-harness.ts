// Helpers for the format tests with real editors: a plain file, presses on a selection of its
// first paragraph, what the editor alone makes of the same presses, and the room's body checked
// after a save and reopen.

import { expect } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import type { EditorModule } from '../../../../contracts/modules.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { docx, MATRIX_EXTRAS, p, t } from '../word-fixtures.ts';
import { BODY, surface, type Editor, type Person } from './editor-harness.ts';
import { comparedBody } from './tables-harness.ts';

/** One paragraph of plain words. */
export const PLAIN_FILE = docx(p(t('Alpha Bold Omega.')), MATRIX_EXTRAS);

/** One person's presses in an editor. */
export type Step = (editor: Editor) => void;

/** Select `[start, end)` of the paragraph at `index` (the first by default). */
export function select(editor: Editor, start: number, end: number, index = 0): void {
  const s = editor.surface;
  if (!s) throw new Error('no surface');
  const id = s.session.paragraphIds()[index];
  if (id === undefined) throw new Error(`no paragraph ${index}`);
  s.setSelection({
    anchor: { paragraphId: id, offset: start },
    head: { paragraphId: id, offset: end },
  });
}

/** Run an editor command, which must succeed. */
export function exec(editor: Editor, command: Readonly<Record<string, unknown>>): void {
  expect(editor.exec(command as never).ok).toBe(true);
}

/**
 * What the editor alone makes of the steps: each person's steps in their own editor on the file
 * the one before saved, then, with `decide`, every revision decided in the last one.
 */
export async function alone(
  file: Uint8Array,
  module: EditorModule,
  steps: readonly { author: string; mode: 'suggesting' | 'edit'; step: Step }[],
  decide?: 'accept' | 'reject'
): Promise<string> {
  let bytes = file;
  let body = '';
  for (const [k, { author, mode, step }] of steps.entries()) {
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({
      container,
      document: bytes,
      author,
      mode,
      modules: [module],
    });
    try {
      step(editor);
      if (k === steps.length - 1 && decide !== undefined)
        exec(editor, { type: 'resolveAllReviewChanges', action: decide, scope: 'document' });
      const s = editor.surface;
      if (!s) throw new Error('no surface');
      body = comparedBody(s.session.part());
      bytes = new Uint8Array(await editor.save());
    } finally {
      editor.destroy();
      container.remove();
    }
  }
  return body;
}

/** The room's body as the matrix compares it, checked equal after a save and reopen. */
export async function roomBody(person: Person): Promise<string> {
  const body = comparedBody(surface(person).session.part());
  const read = readOoxmlPackage(new Uint8Array(await person.editor.save()), {});
  if (!read.ok) throw new Error(read.reason);
  const saved = read.package.parts.get(BODY);
  if (!saved) throw new Error('no body');
  expect(comparedBody(saved)).toBe(body);
  return body;
}

/** The authors of the revision cards an editor shows. */
export const revisionAuthors = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .flatMap((item) => (item.kind === 'revision' ? [item.author ?? ''] : []));
