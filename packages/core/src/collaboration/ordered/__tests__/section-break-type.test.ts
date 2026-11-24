// A suggested continuous section break on its author's own new paragraph mark (design D14). Word
// and the editor write the type it gives the next section untracked: no change record holds it as
// a change, so a reject of the break leaves it, and another person's pending record on that
// section holds the section before its own change with the type, whether the room made that
// record or a fold before the break wrote it into the file. That person's card holds their own
// change alone, and its reject leaves the break, as in the editor alone. The room, a save and
// reopen, and one and two folds agree.

import { describe, expect, test } from 'bun:test';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { openOrderedWordFile, type OrderedWordFile } from '../client/open.ts';
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

/** One step of a case: Olive's new paragraph mark, Ana's page, a fold, or Olive's break. */
type Step = 'split' | 'page' | 'fold' | 'break';

/** The room after the steps, from the file. */
async function roomAfter(steps: readonly Step[]): Promise<OrderedWordFile> {
  let opened = openOrderedWordFile(FILE, { file: 1 });
  for (const step of steps) {
    const doc = opened.document;
    if (step === 'fold') {
      opened = await folded(doc, opened.base, doc.file + 1, mediaStore(opened.base));
    } else if (step === 'page') {
      landscape(doc);
    } else if (step === 'split') {
      apply(doc, 'olive:1', { kind: 'split', ...EDIT, place: at(doc, 0, 4) }, 'Olive');
    } else {
      apply(
        doc,
        'olive:2',
        { kind: 'section-break', ...EDIT, mark: markOf(doc, 0), type: 'continuous' },
        'Olive'
      );
    }
  }

  return opened;
}

/** The change records of a document's sections. */
const recordsOf = (doc: OrderedDocument): string[] =>
  sectionsOf(doc).map((xml) => /<w:sectPrChange\b.*/.exec(xml)?.[0] ?? '');

/** Each case: its name and its steps. */
const CASES: readonly (readonly [string, readonly Step[]])[] = [
  ['alone', ['split', 'break']],
  ["after Ana's pending page", ['page', 'split', 'break']],
  ["before Ana's pending page", ['split', 'break', 'page']],
  ["after Olive's mark and Ana's pending page", ['split', 'page', 'break']],
  ["after Olive's mark, Ana's pending page and a fold", ['split', 'page', 'fold', 'break']],
];

describe('a suggested continuous section break', () => {
  for (const [label, steps] of CASES) {
    test(`its type is under every change record (${label})`, async () => {
      const { document: doc, base } = await roomAfter(steps);
      const blobs = mediaStore(base);
      const [, next] = sectionsOf(doc);
      expect(next).toContain('<w:type w:val="continuous"/>');

      // No change record holds the old type, and the new section's holds the whole section it
      // copied: live, and after one and two folds
      const once = await folded(doc, base, doc.file + 1, blobs);
      const twice = await folded(once.document, once.base, doc.file + 2, blobs);
      for (const copy of [doc, once.document, twice.document]) {
        for (const record of recordsOf(copy)) {
          expect(record).not.toContain('nextPage');
          expect(record || '<w:pgNumType').toContain('<w:pgNumType');
        }
      }

      checkSaveAndLoad(doc, label, base);
      await checkFold(doc, base, blobs, label);

      // Ana's card holds her page alone, as in the editor alone: its reject leaves Olive's break
      if (steps.includes('page')) {
        const live = (await roomAfter(steps)).document;
        for (const copy of [live, (await folded(doc, base, doc.file + 4, blobs)).document]) {
          const card = orderedCardSeen(copy, 'S:ana:1.1') ?? [];
          expect(card.map((seen) => seen.suggestion)).toEqual(['S:ana:1.1']);
          apply(
            copy,
            'olive:3',
            { kind: 'decide', cards: [{ suggestions: card }], status: 'rejected' },
            'Olive'
          );
          const sections = sectionsOf(copy);
          expect(orderedCardSeen(copy, 'S:olive:1.1')).not.toBeNull();
          expect(sections).toHaveLength(2);
          expect(sections.join()).not.toContain('landscape');
          expect(sections[1]).toContain('<w:type w:val="continuous"/>');
        }
      }

      // Olive's card rejected: her section goes, and the next section stays continuous
      for (const copy of [doc, (await folded(doc, base, doc.file + 3, blobs)).document]) {
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
