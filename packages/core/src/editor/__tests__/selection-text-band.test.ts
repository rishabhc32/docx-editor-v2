// A selection the engine draws stands in for the browser's own, so it covers the text and not
// the line spacing: a retained pin while a panel holds focus, or a paragraph mark beside a
// native highlight. Drawn a full line high, the highlight grew when a panel took focus.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { expect, test } from 'bun:test';
import { documentOrder } from '../../layout/document-order.ts';
import { paragraphRangeRects } from '../../layout/paragraph-range-rects.ts';
import { selectionRects } from '../../layout/selection-rects.ts';
import { docx, mount, openLayout, paragraph } from './paginated-surface-fixtures.ts';

const P0 = '/word/document.xml#0.0.0';
const P1 = '/word/document.xml#0.0.1';
const spaced = (spacing: string, text: string) =>
  `<w:p><w:pPr><w:spacing ${spacing}/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

test('a selection on a spaced line covers the text, as a text highlight does', () => {
  for (const spacing of ['w:line="480" w:lineRule="auto"', 'w:line="600" w:lineRule="exact"']) {
    const { layout } = openLayout(docx(spaced(spacing, 'abcd') + paragraph('efgh')));
    const fragment = layout.pages[0]!.fragments[0]!;
    if (fragment.kind !== 'paragraph') throw new Error('expected a paragraph');
    const line = fragment.lines[0]!;
    const [highlight] = paragraphRangeRects(
      layout,
      new Map([[P0, [{ key: 0, start: 1, end: 4 }]]])
    );
    // From inside the text through the paragraph mark: one band on the line.
    const rects = selectionRects(
      layout,
      { anchor: { paragraphId: P0, offset: 1 }, head: { paragraphId: P1, offset: 0 } },
      documentOrder(layout)
    );
    expect(rects).toHaveLength(1);
    expect(rects[0]!.y).toBe(line.box.y + line.leading);
    expect(rects[0]!.height).toBe(line.box.height - line.leading - (line.trailingSpacing ?? 0));
    expect(rects[0]!.height).toBeLessThan(line.box.height);
    expect([rects[0]!.y, rects[0]!.height]).toEqual([highlight!.y, highlight!.height]);
  }
});

test('the pinned range is as tall as the text the browser selects', () => {
  const { surface, container } = mount(spaced('w:line="480" w:lineRule="auto"', 'Visit example'));
  surface.setSelection({
    anchor: { paragraphId: P0, offset: 6 },
    head: { paragraphId: P0, offset: 13 },
  });
  surface.retainSelection();
  const pinned = container.querySelector<HTMLElement>('.docx-retained-selection-rect')!;
  // The run's own line box is the band the browser selects (see `paintSpan`).
  const run = container.querySelector<HTMLElement>('.layout-run-text')!;
  expect(parseFloat(pinned.style.height)).toBeCloseTo(parseFloat(run.style.lineHeight), 5);
  expect(parseFloat(pinned.style.height)).toBeLessThan(parseFloat(run.parentElement!.style.height));
  surface.destroy();
});
