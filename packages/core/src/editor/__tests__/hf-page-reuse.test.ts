// Entering a header must not replace painted page sheets (#355 cause 3).

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { mountPaginatedSurface } from '../paginated-surface.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

function headeredPages(pageCount: number): Uint8Array {
  const pages: string[] = [];
  for (let index = 0; index < pageCount; index += 1) {
    pages.push(p(`body page ${index + 1}`));
    if (index < pageCount - 1) pages.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
  }
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId10" Type="${R}/header" Target="header1.xml"/></Relationships>`
    ),
    'word/header1.xml': strToU8(`<w:hdr xmlns:w="${W}">${p('HEADER')}</w:hdr>`),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${pages.join('')}` +
        `<w:sectPr><w:headerReference w:type="default" r:id="rId10"/></w:sectPr>` +
        '</w:body></w:document>'
    ),
  });
}

/** Pages with a footer on each, and a text control on page 1 for the body caret to sit in. */
function footeredPagesWithControl(pageCount: number): Uint8Array {
  const pages: string[] = [
    '<w:p><w:r><w:t xml:space="preserve">Party: </w:t></w:r>' +
      '<w:sdt><w:sdtPr><w:alias w:val="Party"/><w:text/></w:sdtPr>' +
      '<w:sdtContent><w:r><w:t>Acme Holdings</w:t></w:r></w:sdtContent></w:sdt></w:p>',
  ];
  for (let index = 1; index < pageCount; index += 1) {
    pages.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>', p(`body page ${index + 1}`));
  }
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId11" Type="${R}/footer" Target="footer1.xml"/></Relationships>`
    ),
    'word/footer1.xml': strToU8(`<w:ftr xmlns:w="${W}">${p('FOOTER')}</w:ftr>`),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${pages.join('')}` +
        `<w:sectPr><w:footerReference w:type="default" r:id="rId11"/></w:sectPr>` +
        '</w:body></w:document>'
    ),
  });
}

describe('header entry keeps painted pages', () => {
  test('enterHeaderFooter and crossing copies retint chrome without replacing sheets', () => {
    const container = document.createElement('div');
    document.body.append(container);
    const mounted = mountPaginatedSurface(container, headeredPages(6), { scale: 1 });
    if (!mounted.ok) throw new Error(`${mounted.reason}: ${mounted.detail ?? ''}`);
    const { surface } = mounted;
    const before = [...container.querySelectorAll<HTMLElement>('.docx-page')];
    expect(before.length).toBe(6);

    expect(surface.enterHeaderFooter({ rId: 'rId10', pageIndex: 0 })).toBe(true);
    const entered = [...container.querySelectorAll<HTMLElement>('.docx-page')];
    expect(entered).toHaveLength(before.length);
    for (let index = 0; index < before.length; index += 1) {
      expect(entered[index]).toBe(before[index]);
    }
    expect(before[0]!.querySelector('[data-docx-hf-active]')?.getAttribute('data-docx-hf')).toBe(
      'header'
    );
    expect(before[1]!.querySelector('[data-docx-hf-active]')).toBeNull();
    expect(container.querySelector('.docx-page-content')?.getAttribute('contenteditable')).toBe(
      'false'
    );

    expect(surface.enterHeaderFooter({ rId: 'rId10', pageIndex: 1 })).toBe(true);
    const moved = [...container.querySelectorAll<HTMLElement>('.docx-page')];
    for (let index = 0; index < before.length; index += 1) {
      expect(moved[index]).toBe(before[index]);
    }
    expect(before[0]!.querySelector('[data-docx-hf-active]')).toBeNull();
    expect(before[1]!.querySelector('[data-docx-hf-active]')).not.toBeNull();

    surface.exitHeaderFooter();
    const exited = [...container.querySelectorAll<HTMLElement>('.docx-page')];
    for (let index = 0; index < before.length; index += 1) {
      expect(exited[index]).toBe(before[index]);
    }
    expect(container.querySelector('[data-docx-hf-active]')).toBeNull();
    surface.destroy();
  });
});

// The body caret in a content control used to rebuild EVERY sheet on each entry and exit: the
// control at the caret was part of the paint-reuse key, and the caret crosses into the footer
// and back. A tool that probes several stories per edit repainted the whole document dozens of
// times in one task.
describe('footer entry with the body caret in a content control keeps painted pages', () => {
  for (const editingMode of ['edit', 'suggest'] as const) {
    test(`${editingMode} mode: entering and leaving retint the control without replacing sheets`, () => {
      const container = document.createElement('div');
      document.body.append(container);
      const mounted = mountPaginatedSurface(container, footeredPagesWithControl(6), {
        scale: 1,
        author: 'Olive',
        editingMode,
      });
      if (!mounted.ok) throw new Error(`${mounted.reason}: ${mounted.detail ?? ''}`);
      const { surface } = mounted;
      const sheets = () => [...container.querySelectorAll<HTMLElement>('.docx-page')];
      const chrome = () =>
        container.querySelector<HTMLElement>('.docx-content-control-chrome[data-alias="Party"]')!;
      const body = surface.session.paragraphIds()[0]!;
      surface.setSelection({
        anchor: { paragraphId: body, offset: 10 },
        head: { paragraphId: body, offset: 10 },
      });
      expect(surface.contentControls.atCaret()?.alias).toBe('Party');
      expect(chrome().hasAttribute('data-active')).toBe(true);

      /** Sheets the step replaced, and the control's chrome state after it. */
      const step = (run: () => void) => {
        const before = sheets();
        run();
        const after = sheets();
        expect(after).toHaveLength(before.length);
        return {
          replaced: after.filter((sheet, index) => sheet !== before[index]).length,
          active: chrome().hasAttribute('data-active'),
          visible: chrome().hasAttribute('data-boundary-visible'),
        };
      };
      const cycle = () => [
        step(() => expect(surface.enterHeaderFooter({ rId: 'rId11', pageIndex: 0 })).toBe(true)),
        step(() => surface.exitHeaderFooter()),
      ];

      const first = cycle();
      expect(first).toEqual([
        { replaced: 0, active: false, visible: false },
        { replaced: 0, active: true, visible: true },
      ]);

      // Edits in both stories between the cycles: a later cycle must cost no more.
      expect(surface.enterHeaderFooter({ rId: 'rId11', pageIndex: 0 })).toBe(true);
      surface.type('Page ');
      surface.exitHeaderFooter();
      surface.type('New ');
      expect(surface.contentControls.atCaret()?.alias).toBe('Party');
      const later = cycle();
      expect(later).toEqual(first);

      // The screen is the one a full paint of the same state shows.
      const fresh = document.createElement('div');
      document.body.append(fresh);
      const reopened = mountPaginatedSurface(fresh, surface.save(), {
        scale: 1,
        author: 'Olive',
        editingMode,
      });
      if (!reopened.ok) throw new Error(`${reopened.reason}: ${reopened.detail ?? ''}`);
      reopened.surface.setSelection(surface.state().selection);
      expect(sheets().map((sheet) => sheet.outerHTML)).toEqual(
        [...fresh.querySelectorAll<HTMLElement>('.docx-page')].map((sheet) => sheet.outerHTML)
      );
      reopened.surface.destroy();
      fresh.remove();
      surface.destroy();
      container.remove();
    });
  }
});
