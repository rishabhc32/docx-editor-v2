// A toolbar format that restates what the text already has leaves the document as it was.
//
// Word makes no change for such a press. The editor used to commit it: the run was split at
// the selection edges and its `w:rPr` re-emitted in schema order, so the saved package moved
// and the press became an undo step that undid nothing.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { mount } from './paginated-surface-fixtures.ts';

// Both containers are authored out of schema order (`w:b` precedes `w:sz`, `w:pStyle`
// precedes `w:jc`), so any rewrite shows in the XML.
const BODY =
  '<w:p><w:pPr><w:jc w:val="center"/><w:pStyle w:val="Normal"/>' +
  '<w:rPr><w:sz w:val="28"/><w:b/></w:rPr></w:pPr>' +
  '<w:r><w:rPr><w:sz w:val="28"/><w:b/></w:rPr><w:t>Hello world</w:t></w:r></w:p>';

function withSurface(run: (surface: PaginatedSurface, paragraphId: string) => void): void {
  const { surface, container } = mount(BODY);
  document.body.append(container);
  try {
    run(surface, surface.session.paragraphIds()[0]!);
  } finally {
    surface.destroy();
    container.remove();
  }
}

function select(surface: PaginatedSurface, paragraphId: string, start: number, end: number) {
  surface.setSelection({
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  });
}

function expectNoChange(surface: PaginatedSurface, press: () => void, suggesting = false): void {
  if (suggesting) surface.setEditingMode('suggest');
  const xml = serializeOoxmlPart(surface.session.part());
  const revision = surface.session.packageRevision();
  press();
  expect(surface.session.packageRevision()).toBe(revision);
  expect(serializeOoxmlPart(surface.session.part())).toBe(xml);
}

describe('a format the text already has makes no change in the editor', () => {
  test('font size over part of a run', () => {
    withSurface((surface, paragraphId) => {
      select(surface, paragraphId, 2, 7);
      expectNoChange(surface, () => surface.setRunProperty('sz', { val: '28' }));
    });
  });

  test('font size over the whole paragraph, mark included', () => {
    withSurface((surface, paragraphId) => {
      select(surface, paragraphId, 0, 11);
      expectNoChange(surface, () => surface.setRunProperty('sz', { val: '28' }));
    });
  });

  test('font size in suggesting mode writes no suggestion', () => {
    withSurface((surface, paragraphId) => {
      select(surface, paragraphId, 2, 7);
      expectNoChange(surface, () => surface.setRunProperty('sz', { val: '28' }), true);
    });
  });

  test('paragraph alignment the paragraph already has', () => {
    withSurface((surface, paragraphId) => {
      select(surface, paragraphId, 2, 2);
      expectNoChange(surface, () => surface.setParagraphProperty('jc', { val: 'center' }));
    });
  });

  test('a new size still changes the document', () => {
    withSurface((surface, paragraphId) => {
      select(surface, paragraphId, 2, 7);
      const revision = surface.session.packageRevision();
      surface.setRunProperty('sz', { val: '32' });
      expect(surface.session.packageRevision()).not.toBe(revision);
    });
  });
});
