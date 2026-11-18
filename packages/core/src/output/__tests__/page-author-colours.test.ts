// Paint reuse when the author map changes. A person's first suggestion adds them to the map of
// revision authors, and the take-back of their last one removes them. A page that paints no
// mark of an author whose colour moved keeps its sheet; a page that does is painted again.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { mountPaginatedSurface, type PaginatedSurface } from '../../editor/paginated-surface.ts';
import type { SemanticLayout } from '../../layout/semantic-records.ts';
import { docx } from '../../editor/__tests__/paginated-surface-fixtures.ts';
import { authorSlotsOf } from '../revision-presentation.ts';
import {
  paintSemanticLayout,
  paintSemanticLayoutWithAuthorSlots,
  type PaintOptions,
} from '../semantic-paint.ts';

const LONG =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
  'ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation.';
const plain = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const zed = (text: string) =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r>` +
  '<w:ins w:id="7" w:author="Zed" w:date="2026-10-01T00:00:00Z">' +
  '<w:r><w:t xml:space="preserve"> added by Zed</w:t></w:r></w:ins></w:p>';
/** Zed's suggestion sits late in the document; paragraph 2 is short, on page 1. */
const BODY = Array.from({ length: 160 }, (_, k) =>
  k === 2 ? plain('Short.') : k === 120 ? zed(`${k} ${LONG}`) : plain(`${k} ${LONG}`)
).join('');
const OPTIONS: PaintOptions = { scale: 1 };

const surfaces: { surface: PaginatedSurface; container: HTMLElement }[] = [];
afterEach(() => {
  for (const { surface, container } of surfaces.splice(0)) {
    surface.destroy();
    container.remove();
  }
});

/** Layouts before and after Ana's first suggestion: one letter at the end of paragraph 2. */
function firstSuggestion(): { before: SemanticLayout; after: SemanticLayout } {
  const container = document.createElement('div');
  document.body.append(container);
  const mounted = mountPaginatedSurface(container, docx(BODY), { scale: 1, author: 'Ana' });
  if (!mounted.ok) throw new Error(mounted.reason);
  const surface = mounted.surface;
  surfaces.push({ surface, container });
  surface.setEditingMode('suggest');
  const before = surface.layout();
  const paragraphId = surface.session.paragraphIds()[2]!;
  surface.setSelection({ anchor: { paragraphId, offset: 6 }, head: { paragraphId, offset: 6 } });
  surface.type('x');
  return { before, after: surface.layout() };
}

/** The pages whose sheet a paint kept, by index. */
function keptSheets(container: HTMLElement, sheets: readonly Element[]): number[] {
  return Array.from(container.children).flatMap((sheet, index) =>
    sheet === sheets[index] ? [index] : []
  );
}

/** The page indexes whose records name `author`. */
function pagesOf(layout: SemanticLayout, author: string): number[] {
  return layout.pages.flatMap((page, index) =>
    authorSlotsOf({ ...layout, pages: [page] }).has(author) ? [index] : []
  );
}

function cold(layout: SemanticLayout, slots?: ReadonlyMap<string, number>): string {
  const container = document.createElement('div');
  if (slots) paintSemanticLayoutWithAuthorSlots(container, layout, OPTIONS, slots);
  else paintSemanticLayout(container, layout, OPTIONS);
  return container.innerHTML;
}

describe('paint reuse when the author map changes', () => {
  test("a new author's first suggestion keeps every sheet that shows none of it", () => {
    const { before, after } = firstSuggestion();
    expect(after.pages.length).toBeGreaterThan(8);
    // The precondition: the map gains Ana, on page 1 only, and Zed keeps his colour
    expect(pagesOf(after, 'Ana')).toEqual([0]);
    const slotsBefore = new Map([['Zed', 0]]);
    const slotsAfter = new Map([
      ['Zed', 0],
      ['Ana', 1],
    ]);
    const container = document.createElement('div');
    paintSemanticLayoutWithAuthorSlots(container, before, OPTIONS, slotsBefore);
    const sheets = Array.from(container.children);
    paintSemanticLayoutWithAuthorSlots(container, after, OPTIONS, slotsAfter);
    expect(container.innerHTML).toBe(cold(after, slotsAfter));
    // Page 1 keeps its sheet too: only its changed paragraph is painted again
    expect(keptSheets(container, sheets)).toEqual(after.pages.map((_, index) => index));
    // The take-back: Ana leaves the map, and again no sheet is built anew
    const shown = Array.from(container.children);
    paintSemanticLayoutWithAuthorSlots(container, before, OPTIONS, slotsBefore);
    expect(container.innerHTML).toBe(cold(before, slotsBefore));
    expect(keptSheets(container, shown)).toEqual(before.pages.map((_, index) => index));
  });

  test('an author whose colour moves has every page with their marks painted again', () => {
    const { before, after } = firstSuggestion();
    // A standalone paint numbers authors by first appearance: Ana on page 1 takes Zed's slot
    const zedPages = pagesOf(after, 'Zed');
    expect(zedPages.length).toBeGreaterThan(0);
    expect(zedPages).not.toContain(0);
    const container = document.createElement('div');
    paintSemanticLayout(container, before, OPTIONS);
    const sheets = Array.from(container.children);
    paintSemanticLayout(container, after, OPTIONS);
    expect(container.innerHTML).toBe(cold(after));
    const kept = keptSheets(container, sheets);
    for (const index of zedPages) expect(kept).not.toContain(index);
    expect(kept.length).toBe(after.pages.length - zedPages.length);
  });
});
