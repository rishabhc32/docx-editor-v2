// A suggestion on a control's prompt, decided while typing over the prompt hides it, then the
// typing decided, live and after one and two folds (design D14). While content fills the control,
// the glossary keeps the prompt with its change records, and a load reads each record as its
// suggestion again. A record whose suggestion was decided follows its decision there as it does
// when the prompt shows: accepted, the format stays with no record; rejected, the face before
// it. A fold once kept a decided record in the glossary, so the folded copy read a pending
// suggestion that labels nothing and shows no card, and a reject of the typing showed it.

import { describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { openOrderedWordFile } from '../client/open.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import type { OrderedDocument } from '../model/types.ts';
import { apply, at, EDIT, markOf } from './change-apply-fixture.ts';
import { folded, mediaStore } from './fold-check.ts';
import { semanticBody } from './matrix/compare.ts';
import { docx, p, t } from './word-fixtures.ts';

const PROMPT = 'Click or tap here to enter text.';

/** "Name: <prompt>" with an inline plain text control showing its prompt. */
const INLINE = docx(
  p(
    t('Name: ') +
      '<w:sdt><w:sdtPr><w:showingPlcHdr/><w:text/></w:sdtPr>' +
      `<w:sdtContent>${t(PROMPT)}</w:sdtContent></w:sdt>`
  )
);

/** A block control showing its prompt, then a paragraph after it. */
const BLOCK = docx(
  '<w:sdt><w:sdtPr><w:showingPlcHdr/></w:sdtPr>' +
    `<w:sdtContent>${p(t(PROMPT))}</w:sdtContent></w:sdt>${p(t('After.'))}`
);

/** Ana's suggestion on the prompt: "to ente" bigger, or the prompt's paragraph centred. */
const SUGGESTIONS = [
  [
    'a format on part of an inline prompt',
    INLINE,
    (doc: OrderedDocument) =>
      apply(doc, 'ana:1', {
        kind: 'format',
        ...EDIT,
        from: at(doc, 0, 24),
        to: at(doc, 0, 31),
        properties: [{ localName: 'sz', attributes: { val: '28' } }],
      }),
    (doc: OrderedDocument) => at(doc, 0, 9),
  ],
  [
    "a paragraph property on a block prompt's paragraph",
    BLOCK,
    (doc: OrderedDocument) =>
      apply(doc, 'ana:1', {
        kind: 'paragraph-property',
        ...EDIT,
        marks: [markOf(doc, 0)],
        properties: [{ localName: 'jc', attributes: { val: 'center' } }],
      }),
    (doc: OrderedDocument) => at(doc, 0, 3),
  ],
] as const;

const DECISIONS = ['accepted', 'rejected'] as const;

let decisions = 0;

/** The organiser decides the card that holds `id`. */
function decide(doc: OrderedDocument, id: string, status: (typeof DECISIONS)[number]): void {
  decisions += 1;
  const suggestions = orderedCardSeen(doc, id);
  expect(suggestions).not.toBeNull();
  apply(
    doc,
    `olive:${decisions}`,
    { kind: 'decide', cards: [{ suggestions: suggestions! }], status },
    'Olive'
  );
}

/**
 * The main part and the glossary of a copy's Word file, each by what it says (`semanticBody`): a
 * decision on a folded copy may keep an empty `w:rPr` where the room keeps none.
 */
function fileOf(doc: OrderedDocument, base: OoxmlPackage) {
  const parts = unzipSync(orderedPackageBytes(doc, base, { blobs: mediaStore(base) }));
  const read = (name: string) =>
    semanticBody(
      strFromU8(parts[name] ?? new Uint8Array()).replace(
        / w:(id|rsid[A-Za-z]*)="[^"]*"| w14:(paraId|textId)="[^"]*"/g,
        ''
      )
    );
  return { body: read('word/document.xml'), glossary: read('word/glossary/document.xml') };
}

/** The pending suggestions of a copy: who made each. */
const pendingOf = (doc: OrderedDocument) =>
  [...doc.suggestions.values()]
    .filter((s) => s.status === 'pending')
    .map((s) => `${s.id} ${s.attribution.author}`)
    .sort();

describe('a suggestion on a prompt decided while typing hides it, after a fold', () => {
  for (const [shape, file, suggest, typeAt] of SUGGESTIONS) {
    for (const first of DECISIONS) {
      for (const typing of DECISIONS) {
        test(`${shape}, ${first}, the typing ${typing}`, async () => {
          const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
          suggest(doc);
          const [own] = [...doc.suggestions.keys()];
          apply(doc, 'ana:2', { kind: 'insert', ...EDIT, place: typeAt(doc), text: 'unless ' });
          const typed = [...doc.suggestions.keys()].find((id) => id !== own)!;
          decide(doc, own!, first);
          const blobs = mediaStore(base);
          const once = await folded(doc, base, 2, blobs);
          const twice = await folded(once.document, once.base, 3, blobs);
          const copies = [
            ['live', doc, base],
            ['folded once', once.document, once.base],
            ['folded twice', twice.document, twice.base],
          ] as const;
          const pending = pendingOf(doc);
          const files = copies.map(([copy, d, b]) => {
            expect({ copy, pending: pendingOf(d) }).toEqual({ copy, pending });
            decide(d, typed, typing);
            return { copy, file: fileOf(d, b) };
          });
          for (const { copy, file: decided } of files.slice(1)) {
            expect({ copy, file: decided }).toEqual({ copy, file: files[0]!.file });
          }
        });
      }
    }
  }
});
