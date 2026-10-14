// The `headerFooterEditing` prop reaches the editor from `DocxEditor.Root` and the
// `DocxEditor` sugar. The editor reads it once, at creation.

import './dom-setup.ts';

(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, render } from '@testing-library/react';
import { zipSync, strToU8 } from 'fflate';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { DocxEditor } from '../src/components/DocxEditor.tsx';
import { DocxEditorRoot } from '../src/editor/DocxEditorRoot.tsx';
import { DocxEditorViewport } from '../src/editor/DocxEditorViewport.tsx';
import { DocxEditorContent } from '../src/editor/DocxEditorContent.tsx';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

const SOURCE = zipSync({
  '[Content_Types].xml': strToU8(
    `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  ),
  '_rels/.rels': strToU8(
    `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
  ),
  'word/document.xml': strToU8(
    `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>`
  ),
});

afterEach(cleanup);

/** Whether the editor would open (and create) the header: refused when furniture is read-only. */
function canEditHeader(editor: DocxEditorInstance): boolean {
  return editor.can({ type: 'editHeaderFooter', position: 'header' }).ok;
}

function captureWarnings(): { messages: string[]; restore: () => void } {
  const messages: string[] = [];
  const original = console.warn;
  console.warn = (...args: unknown[]) => {
    messages.push(String(args[0]));
  };
  return { messages, restore: () => (console.warn = original) };
}

const editingWarnings = (messages: readonly string[]) =>
  messages.filter((message) => message.includes('headerFooterEditing changed')).length;

describe('headerFooterEditing prop', () => {
  for (const host of ['root', 'sugar'] as const) {
    for (const headerFooterEditing of [false, undefined] as const) {
      test(`${host} with ${headerFooterEditing ?? 'the default'}`, async () => {
        let instance: DocxEditorInstance | null = null;
        const onReady = (editor: unknown) => {
          instance = editor as DocxEditorInstance;
        };
        const props = headerFooterEditing === false ? { headerFooterEditing } : {};
        await act(async () => {
          render(
            host === 'root' ? (
              <DocxEditorRoot document={SOURCE} onReady={onReady} {...props}>
                <DocxEditorViewport>
                  <DocxEditorContent />
                </DocxEditorViewport>
              </DocxEditorRoot>
            ) : (
              <DocxEditor document={SOURCE} onReady={onReady} {...props} />
            )
          );
        });
        expect(canEditHeader(instance!)).toBe(headerFooterEditing !== false);
      });
    }
  }

  // The warning prints once per module, so the silent cases run before the warning case.
  test('a changed prop is ignored and warns once, only for a real change', async () => {
    const warnings = captureWarnings();
    try {
      let instance: DocxEditorInstance | null = null;
      const tree = (headerFooterEditing: boolean | undefined) => (
        <DocxEditorRoot
          document={SOURCE}
          {...(headerFooterEditing !== undefined ? { headerFooterEditing } : {})}
          onReady={(editor) => {
            instance = editor as DocxEditorInstance;
          }}
        >
          <DocxEditorViewport>
            <DocxEditorContent />
          </DocxEditorViewport>
        </DocxEditorRoot>
      );
      let view: ReturnType<typeof render> | null = null;
      await act(async () => {
        view = render(tree(undefined));
      });
      const first = instance!;

      for (const headerFooterEditing of [undefined, true, true, undefined] as const) {
        await act(async () => {
          view!.rerender(tree(headerFooterEditing));
        });
      }
      expect(editingWarnings(warnings.messages)).toBe(0);

      await act(async () => {
        view!.rerender(tree(false));
      });
      await act(async () => {
        view!.rerender(tree(true));
      });
      await act(async () => {
        view!.rerender(tree(false));
      });
      expect(editingWarnings(warnings.messages)).toBe(1);
      expect(instance).toBe(first);
      expect(canEditHeader(first)).toBe(true);
    } finally {
      warnings.restore();
    }
  });
});
