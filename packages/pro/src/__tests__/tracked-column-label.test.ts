/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/

// A proposed column reads as a column.
//
// Suggesting mode proposes deleting a column as `w:cellDel` on every cell of it, and inserting
// one as `w:cellIns`, all under one revision. The card read "Deleted table cell" for what a
// reviewer sees as a deleted column. A decision that covers whole columns of its table now says
// so; one that leaves part of a column out is still a cell change.

import { describe, expect, test } from 'bun:test';
import { createT, en } from '@docx-editor.dev/i18n';
import {
  readOoxmlPart,
  revisionItemsOf,
  type ReviewRevisionItem,
} from '@docx-editor.dev/core/store';
import { revisionItemLabel } from '../react/review-labels.ts';
import { revisionItemLabel as vueRevisionItemLabel } from '../vue/review-labels.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STAMP = 'w:id="7" w:author="Bea" w:date="2026-10-01T00:00:00Z"';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
/** One cell; `mark` is `del` or `ins` for a proposed cell, `span` its grid span. */
const cell = (text: string, mark?: 'del' | 'ins', span = 1) =>
  `<w:tc><w:tcPr>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}` +
  `${mark === 'del' ? `<w:cellDel ${STAMP}/>` : mark === 'ins' ? `<w:cellIns ${STAMP}/>` : ''}` +
  `</w:tcPr>${p(text)}</w:tc>`;
const table = (columns: number, ...rows: string[]) =>
  `<w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(columns)}</w:tblGrid>` +
  `${rows.map((cells) => `<w:tr>${cells}</w:tr>`).join('')}</w:tbl>`;

function structuralItem(body: string): ReviewRevisionItem {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'application/xml',
  });
  if (!result.ok) throw new Error(result.reason);
  const items = revisionItemsOf(result.part).filter(
    (item) => item.kind === 'revision' && item.revisionKind === 'structural'
  );
  expect(items).toHaveLength(1);
  return items[0]!;
}

/** The card label, the same in both adapters. */
function label(body: string): string {
  const item = structuralItem(body);
  const t = createT(en);
  const react = revisionItemLabel(item, t);
  expect(vueRevisionItemLabel(item, t)).toBe(react);
  return react;
}

describe('a proposed column card', () => {
  test('a deleted column and an inserted column read as columns', () => {
    expect(
      label(
        table(
          3,
          cell('a1') + cell('a2', 'del') + cell('a3'),
          cell('b1') + cell('b2', 'del') + cell('b3')
        )
      )
    ).toBe('Deleted column');
    expect(label(table(2, cell('a1') + cell('new', 'ins'), cell('b1') + cell('new', 'ins')))).toBe(
      'Inserted column'
    );
  });

  test('several columns in one decision read in the plural', () => {
    expect(
      label(
        table(
          3,
          cell('a1', 'del') + cell('a2', 'del') + cell('a3'),
          cell('b1', 'del') + cell('b2', 'del') + cell('b3')
        )
      )
    ).toBe('Deleted columns');
  });

  test('only cells that make up whole columns read as columns', () => {
    // A column drawn across two grid columns in every row is one column.
    expect(
      label(table(3, cell('a1', 'del', 2) + cell('a3'), cell('b1', 'del', 2) + cell('b3')))
    ).toBe('Deleted column');
    // One row only.
    expect(label(table(2, cell('a1') + cell('a2', 'del'), cell('b1') + cell('b2')))).toBe(
      'Deleted table cell'
    );
    // A cell beside the column shares its grid columns in the second row.
    expect(label(table(2, cell('a1', 'del') + cell('a2'), cell('b', 'del', 2)))).toBe(
      'Deleted table cell'
    );
  });
});
