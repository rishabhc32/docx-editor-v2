// `headerFooterEditing: false` makes headers and footers read-only in the editor UI.
//
// Every lane that opens or creates a furniture story is refused: the double press on a painted
// or blank band, the surface's own `enterHeaderFooter` and `setActiveScope`, and the header and
// footer commands, which refuse BEFORE `editHeaderFooter` can create a part. The affordances
// that invite the double press are hidden like viewing's. The default keeps every lane open, and
// automation and review actions still write furniture.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { collectReviewItems, findNode, revisionItemsOf } from '@docx-editor.dev/core/store';
import type { EditorCommand } from '../../contracts/editor.ts';
import type { EditorModule } from '../../contracts/modules.ts';
import { createBrowserAutomationHost } from '../automation-host.ts';
import { createDocxEditor, type DocxEditorConfig } from '../docx-editor.ts';
import { HEADER_FOOTER_READ_ONLY_REASON } from '../docx-editor-hf.ts';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';
import { toolbarCommandState } from '../toolbar-commands.ts';
import { handleAt, handlesAt, roots } from '../../automation/__tests__/support/protocol.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const HEADER = { kind: 'headerFooter', rId: 'rId10' } as const;
const REFUSED = { ok: false, code: 'locked', reason: HEADER_FOOTER_READ_ONLY_REASON };

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

/** A document with an optional default header. */
function docx(header?: string): Uint8Array {
  const entries: Record<string, Uint8Array> = {};
  const hasHeader = header !== undefined;
  entries['[Content_Types].xml'] = strToU8(
    `<Types xmlns="${CT}">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      (hasHeader
        ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        : '') +
      '</Types>'
  );
  entries['_rels/.rels'] = strToU8(
    `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  if (hasHeader) {
    entries['word/_rels/document.xml.rels'] = strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId10" Type="${R}/header" Target="header1.xml"/></Relationships>`
    );
    entries['word/header1.xml'] = strToU8(`<w:hdr xmlns:w="${W}">${header}</w:hdr>`);
  }
  entries['word/document.xml'] = strToU8(
    `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>` +
      `${p('Body')}` +
      `<w:sectPr>${hasHeader ? '<w:headerReference w:type="default" r:id="rId10"/>' : ''}</w:sectPr>` +
      '</w:body></w:document>'
  );
  return zipSync(entries);
}

function mount(
  bytes: Uint8Array,
  headerFooterEditing?: boolean
): { surface: PaginatedSurface; container: HTMLElement; pages: HTMLElement } {
  const container = document.createElement('div');
  document.body.append(container);
  const result = mountPaginatedSurface(container, bytes, {
    scale: 1,
    ...(headerFooterEditing !== undefined ? { headerFooterEditing } : {}),
  });
  if (!result.ok) throw new Error(`${result.reason}: ${result.detail ?? ''}`);
  const pages = container.querySelector<HTMLElement>('.docx-pages')!;
  Object.defineProperty(pages, 'getBoundingClientRect', {
    configurable: true,
    value: () => ({ left: 0, top: 0, right: 2000, bottom: 4000, width: 2000, height: 4000 }),
  });
  pages.focus();
  return { surface: result.surface, container, pages };
}

/** Two presses at one point: the double click that opens or creates a band. */
function doublePress(pages: HTMLElement, clientX: number, clientY: number): void {
  for (let press = 0; press < 2; press += 1) {
    pages.dispatchEvent(
      new PointerEvent('pointerdown', {
        bubbles: true,
        cancelable: true,
        button: 0,
        pointerId: 1,
        pointerType: 'mouse',
        clientX,
        clientY,
      })
    );
  }
}

function packagePartNames(bytes: ArrayBuffer): string[] {
  return Object.keys(unzipSync(new Uint8Array(bytes)));
}

describe('headerFooterEditing on the surface', () => {
  for (const headerFooterEditing of [undefined, false] as const) {
    const readOnly = headerFooterEditing === false;
    const label = readOnly ? 'false' : 'the default';

    test(`with ${label}, a double press on the header band ${readOnly ? 'is refused' : 'opens it'}`, () => {
      const { surface, container, pages } = mount(docx(p('HEADER')), headerFooterEditing);
      const page = surface.layout().pages[0]!;
      const header = page.header!;
      doublePress(pages, header.box.x + 8, Math.max(page.box.y + 2, header.box.y - 4));

      expect(surface.activeScope()).toEqual(readOnly ? { kind: 'body' } : HEADER);
      expect(pages.classList.contains('docx-pages--hf-editing')).toBe(!readOnly);
      expect(container.querySelector('[data-docx-hf-active]') !== null).toBe(!readOnly);
      surface.destroy();
      container.remove();
    });

    test(`with ${label}, a double press on a blank band ${readOnly ? 'creates nothing' : 'creates a header'}`, () => {
      const { surface, container, pages } = mount(docx(), headerFooterEditing);
      const page = surface.layout().pages[0]!;
      doublePress(pages, page.contentBox.x + 8, page.box.y + 2);

      const headers = surface.session.headerFooterResolutionBySection()[0]?.headers.size ?? 0;
      expect(headers).toBe(readOnly ? 0 : 1);
      expect(surface.session.canUndo()).toBe(!readOnly);
      surface.destroy();
      container.remove();
    });

    test(`with ${label}, the surface ${readOnly ? 'carries' : 'omits'} the read-only class`, () => {
      const { surface, container } = mount(docx(p('HEADER')), headerFooterEditing);
      expect(container.classList.contains('docx-paginated-surface--hf-read-only')).toBe(readOnly);
      surface.destroy();
      container.remove();
    });
  }

  test('enterHeaderFooter and setActiveScope cannot open a header', () => {
    const { surface, container } = mount(docx(p('HEADER')), false);
    expect(surface.enterHeaderFooter({ rId: HEADER.rId })).toBe(false);
    expect(surface.setActiveScope(HEADER)).toBe(false);
    expect(surface.activeScope()).toEqual({ kind: 'body' });
    surface.destroy();
    container.remove();
  });
});

/** The engine's own review derivation, wired as a module (core may not import pro). */
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

describe('headerFooterEditing on the editor', () => {
  function mountEditor(bytes: Uint8Array, config: Partial<DocxEditorConfig> = {}) {
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({ ...config, container, document: bytes });
    return {
      editor,
      dispose() {
        editor.destroy();
        container.remove();
      },
    };
  }

  test('editHeaderFooter is refused before it creates a part', async () => {
    const { editor, dispose } = mountEditor(docx(), { headerFooterEditing: false });
    const command = { type: 'editHeaderFooter', position: 'header' } as const;

    expect(editor.can(command)).toEqual(REFUSED);
    expect(editor.exec(command)).toEqual(REFUSED);
    expect(editor.getActiveScope()).toEqual({ kind: 'body' });
    expect(packagePartNames(await editor.save()).filter((name) => name.includes('header'))).toEqual(
      []
    );
    dispose();
  });

  test('the default still creates the header and opens it', async () => {
    const { editor, dispose } = mountEditor(docx());
    expect(editor.exec({ type: 'editHeaderFooter', position: 'header' })).toEqual({
      ok: true,
      changed: true,
    });
    expect(editor.getActiveScope().kind).toBe('headerFooter');
    expect(packagePartNames(await editor.save())).toContain('word/header1.xml');
    dispose();
  });

  test('every other header and footer command is refused and changes nothing', async () => {
    const { editor, dispose } = mountEditor(docx(p('HEADER')), { headerFooterEditing: false });
    const before = new Uint8Array(await editor.save());
    const commands: EditorCommand[] = [
      { type: 'editHeaderFooter', position: 'header' },
      { type: 'removeHeaderFooter', position: 'header' },
      { type: 'setHeaderFooterOptions', titlePage: true },
      { type: 'linkHeaderFooterToPrevious', position: 'header' },
      { type: 'unlinkHeaderFooterFromPrevious', position: 'header' },
      { type: 'insertPageField', field: 'PAGE' },
    ];
    for (const command of commands) {
      expect([command.type, editor.exec(command)]).toEqual([command.type, REFUSED]);
    }
    expect(new Uint8Array(await editor.save())).toEqual(before);
    dispose();
  });

  test('setActiveScope and focus cannot open a header', () => {
    const { editor, dispose } = mountEditor(docx(p('HEADER')), { headerFooterEditing: false });
    editor.setActiveScope(HEADER);
    expect(editor.getActiveScope()).toEqual({ kind: 'body' });
    expect(editor.focus(HEADER)).toEqual({
      ok: false,
      code: 'invalidTarget',
      reason: 'that scope cannot be opened',
    });
    dispose();
  });

  test('the page-number control is disabled with the reason', () => {
    const { editor, dispose } = mountEditor(docx(p('HEADER')), { headerFooterEditing: false });
    expect(toolbarCommandState(editor, 'insert.pageNumber')).toMatchObject({
      enabled: false,
      disabledReason: HEADER_FOOTER_READ_ONLY_REASON,
    });
    dispose();
  });

  test('automation still writes the header', () => {
    const { editor, dispose } = mountEditor(docx(p('HEADER')), { headerFooterEditing: false });
    const host = createBrowserAutomationHost(editor);
    const { document } = roots(host);
    const section = handlesAt(
      host.execute({ operations: [{ op: 'getSections', document }] }),
      0
    )[0]!;
    const header = handleAt(
      host.execute({
        operations: [{ op: 'getFurniture', section, kind: 'header', variant: 'default' }],
      }),
      0
    );
    expect(
      host.execute({
        operations: [{ op: 'replaceSpan', span: { body: header }, text: 'Scripted' }],
      }).ok
    ).toBe(true);
    expect(editor.surface!.session.storyText(HEADER)).toBe('Scripted');
    host.dispose();
    dispose();
  });

  test('accepting every change still resolves a tracked change in the header', () => {
    const inserted =
      '<w:p><w:ins w:id="1" w:author="Ada"><w:r><w:t>Added</w:t></w:r></w:ins></w:p>';
    const { editor, dispose } = mountEditor(docx(inserted), {
      headerFooterEditing: false,
      modules: [reviewModule()],
    });
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' })).toMatchObject({
      ok: true,
      changed: true,
    });
    expect(editor.surface!.session.reviewItems()).toHaveLength(0);
    expect(editor.surface!.session.storyText(HEADER)).toBe('Added');
    dispose();
  });
});
