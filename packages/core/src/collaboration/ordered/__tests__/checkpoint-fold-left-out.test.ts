// A fold keeps what shows (design D14). The room keeps the records of what shows nowhere now, for
// the undo of the change that hid them: a suggestion a direct or an accepted deletion took whole,
// a table whose rows its author took back, a note whose reference went, a decided format record,
// and a text box whose object went. The file holds none of them, so a fold drops them, and the
// room shows none of them either.

import { describe, test } from 'bun:test';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { openOrderedWordFile } from '../client/open.ts';
import type { OrderedDocument } from '../model/types.ts';
import { apply, APPLY_FILE, at, DIRECT, EDIT, markOf } from './change-apply-fixture.ts';
import { checkFold, mediaStore } from './fold-check.ts';
import { docx, MATRIX_EXTRAS, p, t } from './word-fixtures.ts';

const TEXT = docx(p(t('Cloud and serverless.')) + p(t('Signed here.')));
const NOTE = '<w:r><w:footnoteReference w:id="1"/></w:r>';
const TEXT_BOX = { type: 'textbox', extent: { cx: 1_828_800 } } as const;

/** The organiser's decision on one suggestion's card. */
function decide(doc: OrderedDocument, id: string, status: 'accepted' | 'rejected'): void {
  const cards = [{ suggestions: orderedCardSeen(doc, id) ?? [] }];
  apply(doc, `olive:${doc.suggestions.size}`, { kind: 'decide', cards, status }, 'Olive');
}

async function foldOf(file: Uint8Array, edit: (doc: OrderedDocument) => void, label: string) {
  const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
  edit(doc);
  await checkFold(doc, base, mediaStore(base), label);
}

describe('a fold of what shows nowhere', () => {
  test("a direct deletion over another person's pending words", async () => {
    await foldOf(
      TEXT,
      (doc) => {
        apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 6), text: 'net ' });
        apply(
          doc,
          'olga:1',
          { kind: 'delete', ...DIRECT, from: at(doc, 0, 0), to: at(doc, 0, 14) },
          'Olga'
        );
      },
      'a direct deletion over pending words'
    );
  });

  test("an accepted deletion over another person's pending format", async () => {
    await foldOf(
      TEXT,
      (doc) => {
        apply(doc, 'ana:1', {
          kind: 'format',
          ...EDIT,
          from: at(doc, 0, 0),
          to: at(doc, 0, 5),
          properties: [{ localName: 'b' }],
        });
        apply(
          doc,
          'bob:1',
          { kind: 'delete', ...EDIT, from: at(doc, 0, 0), to: at(doc, 0, 6) },
          'Bob'
        );
        decide(doc, 'S:bob:1.1', 'accepted');
      },
      'an accepted deletion over a pending format'
    );
  });

  test('a suggested table its author took back', async () => {
    await foldOf(
      APPLY_FILE,
      (doc) => {
        apply(doc, 'ana:1', {
          kind: 'insert-table',
          ...EDIT,
          place: at(doc, 9, 0),
          rows: 2,
          columns: 2,
          columnWidth: 3000,
        });
        apply(doc, 'ana:2', { kind: 'retract', suggestions: ['S:ana:1.1'] });
      },
      'a table taken back'
    );
  });

  test('a note whose reference a direct deletion took', async () => {
    const file = docx(p(t('Fees exclude tax.') + NOTE) + p(t('Signed here.')), MATRIX_EXTRAS);
    await foldOf(
      file,
      (doc) => {
        apply(
          doc,
          'olga:1',
          { kind: 'delete', ...DIRECT, from: at(doc, 0, 5), to: at(doc, 0, 18) },
          'Olga'
        );
      },
      'a note with no reference'
    );
  });

  test("a format that took over the same person's format, rejected", async () => {
    await foldOf(
      TEXT,
      (doc) => {
        for (const [n, localName] of ['b', 'i'].entries()) {
          apply(
            doc,
            `bob:${n + 1}`,
            {
              kind: 'format',
              ...EDIT,
              from: at(doc, 0, 0),
              to: at(doc, 0, 5),
              properties: [{ localName }],
            },
            'Bob'
          );
        }

        decide(doc, 'S:bob:2.1', 'rejected');
      },
      'a rejected takeover'
    );
  });

  test('a suggested text box, rejected', async () => {
    await foldOf(
      TEXT,
      (doc) => {
        apply(doc, 'ana:1', {
          kind: 'insert-object',
          ...EDIT,
          place: at(doc, 1, 0),
          object: TEXT_BOX,
        });
        decide(doc, 'S:ana:1.1', 'rejected');
      },
      'a rejected text box'
    );
  });

  test("a direct deletion of a suggested text box's paragraph", async () => {
    await foldOf(
      TEXT,
      (doc) => {
        apply(doc, 'ana:1', {
          kind: 'insert-object',
          ...EDIT,
          place: at(doc, 1, 0),
          object: TEXT_BOX,
        });
        apply(doc, 'olga:1', { kind: 'delete-paragraph', ...DIRECT, mark: markOf(doc, 1) }, 'Olga');
      },
      "a text box's paragraph deleted"
    );
  });
});
