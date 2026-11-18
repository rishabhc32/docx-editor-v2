// Typing in a header or footer repaints only its band on each sheet that shows it.
//
// A header shown on every page is a new story record on every page after each keystroke. Paint
// took that for a changed page and built every sheet again, body included: one keystroke in a
// 186-page contract's main header took seconds and gigabytes in the test DOM. The sheets, their
// body layer and its blocks stay; only the header and footer layer is painted again.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const PAGES = 12;

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** Pages with one header and one footer on each, three body paragraphs a page. */
function furnishedPages(): Uint8Array {
  const body: string[] = [];
  for (let page = 0; page < PAGES; page += 1) {
    for (let k = 0; k < 3; k += 1) body.push(p(`Body ${page + 1}.${k + 1}`));
    if (page < PAGES - 1) body.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
  }
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        `<Override PartName="/word/document.xml" ContentType="${MAIN}.document.main+xml"/>` +
        `<Override PartName="/word/header1.xml" ContentType="${MAIN}.header+xml"/>` +
        `<Override PartName="/word/footer1.xml" ContentType="${MAIN}.footer+xml"/>` +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}">` +
        `<Relationship Id="rId10" Type="${R}/header" Target="header1.xml"/>` +
        `<Relationship Id="rId11" Type="${R}/footer" Target="footer1.xml"/></Relationships>`
    ),
    'word/header1.xml': strToU8(`<w:hdr xmlns:w="${W}">${p('Main header')}</w:hdr>`),
    'word/footer1.xml': strToU8(`<w:ftr xmlns:w="${W}">${p('Main footer')}</w:ftr>`),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body.join('')}` +
        '<w:sectPr><w:headerReference w:type="default" r:id="rId10"/>' +
        '<w:footerReference w:type="default" r:id="rId11"/></w:sectPr>' +
        '</w:body></w:document>'
    ),
  });
}

const opened: { surface: PaginatedSurface; container: HTMLElement }[] = [];

afterEach(() => {
  for (const item of opened.splice(0)) {
    item.surface.destroy();
    item.container.remove();
  }
});

function mount(bytes: Uint8Array, editingMode: 'edit' | 'suggest'): PaginatedSurface {
  const container = document.createElement('div');
  document.body.append(container);
  const mounted = mountPaginatedSurface(container, bytes, {
    scale: 1,
    author: 'Olive',
    editingMode,
  });
  if (!mounted.ok) throw new Error(`${mounted.reason}: ${mounted.detail ?? ''}`);
  opened.push({ surface: mounted.surface, container });
  return mounted.surface;
}

const containerOf = (surface: PaginatedSurface) =>
  opened.find((item) => item.surface === surface)!.container;

/** The painted nodes a band repaint must keep: the sheets, their body layer and its blocks. */
function bodyNodes(container: HTMLElement): Element[] {
  return [
    ...container.querySelectorAll('.docx-page'),
    ...container.querySelectorAll('.docx-page > .docx-page-content'),
    ...container.querySelectorAll('.docx-page > .docx-page-content > *'),
  ];
}

/**
 * Each sheet as markup with its attributes in name order. A fresh mount paints the band and
 * then marks it active in place, which puts the attribute last; the screen is the same.
 */
function screen(container: HTMLElement): string[] {
  const markup = (node: Node): string => {
    if (node.nodeType !== node.ELEMENT_NODE) return node.textContent ?? '';
    const element = node as Element;
    const attributes = [...element.attributes]
      .map((attribute) => `${attribute.name}="${attribute.value}"`)
      .sort()
      .join(' ');
    const children = [...element.childNodes].map(markup).join('');
    return `<${element.localName} ${attributes}>${children}</${element.localName}>`;
  };
  return [...container.querySelectorAll('.docx-page')].map(markup);
}

/** Open a story on page 3 with the caret at the end of its text. */
function editStory(surface: PaginatedSurface, rId: string): void {
  expect(surface.enterHeaderFooter({ rId, pageIndex: 2 })).toBe(true);
  const [paragraphId] = surface.session.paragraphIdsIn({ kind: 'headerFooter', rId });
  surface.setSelection({
    anchor: { paragraphId: paragraphId!, offset: 11 },
    head: { paragraphId: paragraphId!, offset: 11 },
  });
}

describe('typing in a header or footer repaints only its band', () => {
  for (const editingMode of ['edit', 'suggest'] as const) {
    for (const [kind, rId] of [
      ['header', 'rId10'],
      ['footer', 'rId11'],
    ] as const) {
      test(`${editingMode} mode, ${kind}: no sheet or body block is replaced`, () => {
        const surface = mount(furnishedPages(), editingMode);
        const container = containerOf(surface);
        expect(container.querySelectorAll('.docx-page')).toHaveLength(PAGES);
        editStory(surface, rId);
        // The first suggestion adds its author to the paint's author map, which repaints in
        // full once; the keystrokes counted are the ones after it.
        surface.type('!');
        const bands = () => [
          ...container.querySelectorAll(`.docx-page > [data-docx-hf="${kind}"]`),
        ];
        for (const text of ['A', 'B', 'C']) {
          const body = bodyNodes(container);
          const before = bands();
          surface.type(text);
          const after = bodyNodes(container);
          expect(after).toHaveLength(body.length);
          expect(after.filter((node, index) => node !== body[index])).toHaveLength(0);
          // Every sheet shows the new text, in a band painted for it.
          expect(bands()).toHaveLength(PAGES);
          expect(bands().filter((band, index) => band === before[index])).toHaveLength(0);
          for (const band of bands()) expect(band.textContent).toContain(`Main ${kind}!`);
        }
        expect(bands()[0]!.textContent).toContain(`Main ${kind}!ABC`);

        // The screen is the one a fresh mount of the same state shows.
        const fresh = mount(surface.save(), editingMode);
        editStory(fresh, rId);
        fresh.setSelection(surface.state().selection);
        expect(screen(container)).toEqual(screen(containerOf(fresh)));
      });
    }
  }
});
