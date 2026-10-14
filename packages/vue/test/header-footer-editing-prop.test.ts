// The `headerFooterEditing` prop reaches the editor from `DocxEditorRoot` and the
// `DocxEditor` sugar. The editor reads it once, at creation.

import './dom-setup.ts';

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, h, shallowRef, type Component } from 'vue';
import { zipSync, strToU8 } from 'fflate';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { DocxEditor } from '../src/components/DocxEditor';
import { DocxEditorRoot } from '../src/editor/DocxEditorRoot';
import { DocxEditorViewport } from '../src/editor/DocxEditorViewport';
import { DocxEditorContent } from '../src/editor/DocxEditorContent';
import { flush } from './helpers/mount';

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

afterEach(() => {
  document.body.innerHTML = '';
});

async function mount(component: Component, initial: boolean | undefined) {
  const container = document.createElement('div');
  document.body.append(container);
  const ready: DocxEditorInstance[] = [];
  const headerFooterEditing = shallowRef<boolean | undefined>(initial);
  const app = createApp({
    render: () =>
      h(
        component,
        {
          document: SOURCE,
          ...(headerFooterEditing.value !== undefined
            ? { headerFooterEditing: headerFooterEditing.value }
            : {}),
          onReady: (editor: unknown) => ready.push(editor as DocxEditorInstance),
        },
        component === DocxEditorRoot
          ? {
              default: () => h(DocxEditorViewport, null, { default: () => h(DocxEditorContent) }),
            }
          : undefined
      ),
  });
  app.mount(container);
  await flush();
  return {
    ready,
    editor: ready.at(-1)!,
    async set(value: boolean | undefined) {
      headerFooterEditing.value = value;
      await flush();
    },
    unmount: () => app.unmount(),
  };
}

/** Whether the editor would open (and create) the header: refused when furniture is read-only. */
function canEditHeader(editor: DocxEditorInstance): boolean {
  return editor.can({ type: 'editHeaderFooter', position: 'header' }).ok;
}

describe('headerFooterEditing prop', () => {
  for (const [host, component] of [
    ['root', DocxEditorRoot],
    ['sugar', DocxEditor],
  ] as const) {
    for (const headerFooterEditing of [false, undefined] as const) {
      test(`${host} with ${headerFooterEditing ?? 'the default'}`, async () => {
        const mounted = await mount(component, headerFooterEditing);
        try {
          expect(canEditHeader(mounted.editor)).toBe(headerFooterEditing !== false);
        } finally {
          mounted.unmount();
        }
      });
    }
  }

  // The warning prints once per module, so the silent cases run before the warning case.
  test('a changed prop is ignored and warns once, only for a real change', async () => {
    const messages: string[] = [];
    const original = console.warn;
    console.warn = (...args: unknown[]) => {
      messages.push(String(args[0]));
    };
    const count = () =>
      messages.filter((message) => message.includes('headerFooterEditing changed')).length;
    const mounted = await mount(DocxEditorRoot, undefined);
    try {
      const first = mounted.editor;
      for (const headerFooterEditing of [undefined, true, true, undefined] as const) {
        await mounted.set(headerFooterEditing);
      }
      expect(count()).toBe(0);

      await mounted.set(false);
      await mounted.set(true);
      await mounted.set(false);
      expect(count()).toBe(1);
      expect(mounted.ready.at(-1)).toBe(first);
      expect(canEditHeader(first)).toBe(true);
    } finally {
      console.warn = original;
      mounted.unmount();
    }
  });
});
