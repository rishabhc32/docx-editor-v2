// A suggested continuous section break on its author's own new paragraph mark (design D14). Word
// and the editor write the type it gives the next section untracked: no change record holds it as
// a change, so a reject of the break leaves it, and another person's pending record on that
// section holds the section before its own change with the type. The room, a save and reopen, and
// one and two folds agree.

import { describe, expect, test } from 'bun:test';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { openOrderedWordFile } from '../client/open.ts';
import type { OrderedDocument } from '../model/types.ts';
import { apply, at, EDIT, markOf, partsOf } from './change-apply-fixture.ts';
import { checkFold, folded, mediaStore } from './fold-check.ts';
import { checkSaveAndLoad } from './sessions/session.ts';
import { docx, p, t } from './word-fixtures.ts';

const FILE = docx(
  p(t('One paragraph.')) +
    p(t('Two.')) +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708"' +
    ' w:footer="708" w:gutter="0"/><w:pgNumType w:start="5"/><w:cols w:num="2" w:space="720"/>' +
    '<w:docGrid w:linePitch="360"/></w:sectPr>'
);

/** Each section of the main part as compose draws it, without ids. */
const sectionsOf = (doc: OrderedDocument): string[] =>
  (partsOf(doc)[0] ?? '')
    .match(/<w:sectPr>(?:(?!<\/w:sectPr>)[\s\S])*<\/w:sectPr>(?:<\/w:sectPrChange><\/w:sectPr>)?/g)
    ?.map((xml) => xml.replace(/ w:id="[^"]*"/g, '')) ?? [];

/** Ana's landscape page for every section. */
const landscape = (doc: OrderedDocument) =>
  apply(
    doc,
    'ana:1',
    { kind: 'section-property', ...EDIT, sections: [], page: { orientation: 'landscape' } },
    'Ana'
  );

/** Olive's new paragraph mark in the first paragraph, and a continuous break on it. */
function oliveBreaks(doc: OrderedDocument): void {
  apply(doc, 'olive:1', { kind: 'split', ...EDIT, place: at(doc, 0, 4) }, 'Olive');
  apply(
    doc,
    'olive:2',
    { kind: 'section-break', ...EDIT, mark: markOf(doc, 0), type: 'continuous' },
    'Olive'
  );
}

describe('a suggested continuous section break', () => {
  for (const [label, before, after] of [
    ['alone', false, false],
    ["after Ana's pending page", true, false],
    ["before Ana's pending page", false, true],
  ] as const) {
    test(`its type is under every change record (${label})`, async () => {
      const opened = openOrderedWordFile(FILE, { file: 1 });
      if (before) {
        landscape(opened.document);
      }

      oliveBreaks(opened.document);
      if (after) {
        landscape(opened.document);
      }

      const { document: doc, base } = opened;
      const blobs = mediaStore(base);
      const [first, next] = sectionsOf(doc);
      expect(next).toContain('<w:type w:val="continuous"/>');

      // No change record holds the old type, and the new section's holds the whole section it copied
      for (const record of [first, next].map((xml) => /<w:sectPrChange\b.*/.exec(xml ?? '')?.[0])) {
        expect(record ?? '').not.toContain('nextPage');
        expect(record ?? '<w:pgNumType').toContain('<w:pgNumType');
      }

      checkSaveAndLoad(doc, label, base);
      await checkFold(doc, base, blobs, label);

      // Olive's card rejected: her section goes, and the next section stays continuous
      for (const copy of [doc, (await folded(doc, base, 3, blobs)).document]) {
        const card = orderedCardSeen(copy, 'S:olive:1.1') ?? [];
        apply(
          copy,
          'olive:3',
          { kind: 'decide', cards: [{ suggestions: card }], status: 'rejected' },
          'Olive'
        );
        const sections = sectionsOf(copy);
        expect(sections).toHaveLength(1);
        expect(sections[0]).toContain('<w:type w:val="continuous"/>');
      }
    });
  }
});
