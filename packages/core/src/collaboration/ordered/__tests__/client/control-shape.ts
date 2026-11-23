// A paragraph's controls as a reader of the file sees them, for room tests that compare with
// the editor alone: the text, each control, its prompt, and each person's insertions and
// deletions, without ids, dates, runs or faces.

import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { caretAt, editorXml, surface, type Person } from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

/**
 * The body's paragraphs as a reader sees them: `[#id` and `]` around a control, `?` for a
 * control that shows its prompt, `{+name:` or `{-name:` and `}` around a person's insertion or
 * deletion (one person's next to each other read as one), `(mark{+name:})` for a paragraph mark
 * a person inserted, and `¶` at each paragraph's end.
 */
export function shapeOf(xml: string): string {
  const body = /<w:body>(.*?)<w:sectPr/s.exec(xml)?.[1] ?? xml;
  const read = body
    .replace(/<w:sdt>/g, '[')
    .replace(/<\/w:sdt>/g, ']')
    .replace(/<w:id w:val="(\d+)"\/>/g, '#$1')
    .replace(/<w:showingPlcHdr\/>/g, '?')
    .replace(/<w:(ins|del) [^>]*w:author="([^"]*)"[^>]*>/g, (_, kind, name) =>
      kind === 'ins' ? `{+${name}:` : `{-${name}:`
    )
    .replace(/<\/w:(ins|del)>/g, '}')
    .replace(/<\/w:p>/g, '¶')
    .replace(/<w:(t|delText)(?: [^>]*)?>([^<]*)<\/w:\1>/g, '$2')
    .replace(/<w:pPr><w:rPr>(\{[^}]*\})<\/w:rPr><\/w:pPr>/g, '(mark$1)')
    .replace(/<[^>]+>/g, '');
  return read.replace(/\}\{([+-][^:]*):/g, (joint, mark: string, at: number, all: string) =>
    all.lastIndexOf(`{${mark}:`, at) > all.lastIndexOf('}', at - 1) ? '' : joint
  );
}

/** What the editor alone writes after `act` with the selection from `from` to `to`. */
export function aloneShape(
  file: Uint8Array,
  author: string,
  mode: 'edit' | 'suggesting',
  [paragraph, from, to]: readonly [number, number, number],
  act: (person: Person) => unknown
): string {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: file,
    author,
    mode,
    modules: [reviewTestModule()],
  });
  try {
    const person = { name: 'alone', editor, container } as unknown as Person;
    caretAt(person, paragraph, from, to);
    act(person);
    surface(person).flushPendingInput();
    return shapeOf(editorXml(person));
  } finally {
    editor.destroy();
    container.remove();
  }
}
