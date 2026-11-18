// A full repaint must not leave the old pages queued in a pages-layer observer.
//
// A subtree `MutationObserver` on the pages layer queues one record for each page a repaint
// replaces, and each record holds the removed page with its whole subtree. The records go out
// only at the end of the task, so a command that repainted every page several times kept every
// old generation of the page DOM alive until it returned.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const PAGES = 6;
const REPAINTS = 6;

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** Pages of plain paragraphs, one page break between pages. */
function pages(): Uint8Array {
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
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body.join('')}<w:sectPr/></w:body></w:document>`
    ),
  });
}

const view = document.defaultView as unknown as { MutationObserver: typeof MutationObserver };
const original = view.MutationObserver;
const observers: MutationObserver[] = [];
const opened: { surface: PaginatedSurface; container: HTMLElement }[] = [];

afterEach(() => {
  view.MutationObserver = original;
  observers.length = 0;
  for (const item of opened.splice(0)) {
    item.surface.destroy();
    item.container.remove();
  }
});

/** Mounts a surface while every observer the window makes is recorded. */
function mount(): { surface: PaginatedSurface; container: HTMLElement } {
  view.MutationObserver = class extends original {
    constructor(callback: MutationCallback) {
      super(callback);
      observers.push(this);
    }
  };
  const container = document.createElement('div');
  document.body.append(container);
  const mounted = mountPaginatedSurface(container, pages(), { scale: 1, author: 'Olive' });
  if (!mounted.ok) throw new Error(`${mounted.reason}: ${mounted.detail ?? ''}`);
  opened.push({ surface: mounted.surface, container });
  return { surface: mounted.surface, container };
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

/** Nodes (each removed node and its subtree) that the observers' queued records still hold. */
function queuedRemovedNodes(): number {
  let count = 0;
  const visit = (node: Node) => {
    count += 1;
    for (const child of node.childNodes) visit(child);
  };
  for (const observer of observers)
    for (const record of observer.takeRecords())
      for (const node of record.removedNodes) visit(node);
  return count;
}

test('repainting every page several times in one task queues none of the old pages', async () => {
  const { surface, container } = mount();
  expect(container.querySelectorAll('.docx-page')).toHaveLength(PAGES);
  await settle();
  const first = [...container.querySelectorAll('.docx-page')];
  let pageNodes = 0;
  for (const page of first) pageNodes += 1 + page.querySelectorAll('*').length;
  expect(pageNodes).toBeGreaterThan(PAGES * 4);

  // One task: each call repaints every page (paragraph marks are part of each page's paint key).
  for (let k = 0; k < REPAINTS; k += 1) surface.setShowParagraphMarks(k % 2 === 0);
  const now = [...container.querySelectorAll('.docx-page')];
  expect(now).toHaveLength(PAGES);
  expect(now.some((page) => first.includes(page))).toBe(false);
  expect(queuedRemovedNodes()).toBe(0);
});
