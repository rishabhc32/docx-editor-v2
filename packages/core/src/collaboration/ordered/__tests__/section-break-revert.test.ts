// A direct section break over a section that holds another person's pending section record, then
// its author's revert (design D14). The break copies the section with that record on it, as the
// engine's clone keeps the section's change record. The revert takes the break away whole, as
// the editor's undo does: one section, with every setting it had, live, after a save and reopen,
// and after one and two folds. A later page setup rejected gives the section back whole too.

import { describe, expect, test } from 'bun:test';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { openOrderedWordFile } from '../client/open.ts';
import type { OrderedDocument } from '../model/types.ts';
import { apply, DIRECT, EDIT, markOf, partsOf } from './change-apply-fixture.ts';
import { checkFold, folded, mediaStore } from './fold-check.ts';
import { checkSaveAndLoad } from './sessions/session.ts';
import { docx, p, t } from './word-fixtures.ts';

const FILE = docx(
  p(t('One.')) +
    p(t('Two.')) +
    p(t('Three.')) +
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

/** The file, with Cy's landscape page pending: in the room, or read from a fold. */
async function withCy(read: boolean) {
  const opened = openOrderedWordFile(FILE, { file: 1 });
  apply(
    opened.document,
    'cy:1',
    { kind: 'section-property', ...EDIT, sections: [], page: { orientation: 'landscape' } },
    'Cy'
  );
  return read ? folded(opened.document, opened.base, 2, mediaStore(opened.base)) : opened;
}

describe("a direct section break over another person's pending section record, reverted", () => {
  for (const [label, read] of [
    ['pending in the room', false],
    ['read from a fold', true],
  ] as const) {
    test(`the revert takes the break away (Cy's record ${label})`, async () => {
      const { document: doc, base } = await withCy(read);
      const before = sectionsOf(doc);
      expect(before).toHaveLength(1);
      apply(
        doc,
        'olive:1',
        { kind: 'section-break', ...DIRECT, mark: markOf(doc, 1), type: 'nextPage' },
        'Olive'
      );
      expect(sectionsOf(doc)).toHaveLength(2);
      apply(
        doc,
        'olive:2',
        { kind: 'revert', suggestions: ['S:olive:1.1'], status: 'rejected' },
        'Olive'
      );
      expect(sectionsOf(doc)).toEqual(before);
      checkSaveAndLoad(doc, label, base);
      await checkFold(doc, base, mediaStore(base), label);

      // Made again, the break copies Cy's record again: both sections show it, as at first
      apply(
        doc,
        'olive:3',
        { kind: 'revert', suggestions: ['S:olive:1.1'], status: 'accepted' },
        'Olive'
      );
      expect(sectionsOf(doc)).toEqual([...before, ...before]);
      checkSaveAndLoad(doc, `${label}, made again`, base);
      await checkFold(doc, base, mediaStore(base), `${label}, made again`);
    });

    test(`a page setup made after it, rejected, gives the section back whole (Cy's record ${label})`, async () => {
      const { document: doc, base } = await withCy(read);
      apply(
        doc,
        'olive:1',
        { kind: 'section-break', ...DIRECT, mark: markOf(doc, 1), type: 'nextPage' },
        'Olive'
      );
      apply(
        doc,
        'olive:2',
        { kind: 'revert', suggestions: ['S:olive:1.1'], status: 'rejected' },
        'Olive'
      );
      apply(
        doc,
        'olive:3',
        {
          kind: 'section-property',
          ...EDIT,
          sections: [],
          page: { pageWidthTwips: 11906, pageHeightTwips: 16838 },
        },
        'Olive'
      );
      await checkFold(doc, base, mediaStore(base), label);
      apply(
        doc,
        'olive:4',
        {
          kind: 'decide',
          cards: [{ suggestions: orderedCardSeen(doc, 'S:olive:3.1') ?? [] }],
          status: 'rejected',
        },
        'Olive'
      );

      // Olive's page setup took Cy's record over (one record for each section): the reject gives
      // the section as the file had it
      expect(sectionsOf(doc)).toEqual(sectionsOf(openOrderedWordFile(FILE, { file: 1 }).document));
    });
  }
});
