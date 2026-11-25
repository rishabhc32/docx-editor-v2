// A fold changes nothing the room uses (design D14): random sessions replayed from the room's log,
// folded at the end of rounds, and their folded copies, with suggestions read from the file taken
// back and then edited at random, folded again, give back every shown unit and its labels, every
// pending suggestion whole, its card, who typed its units, the takeovers, comments, part
// revisions, what a page shows, the Word file, and the Word file after an accept and after a
// reject of each suggestion.

import { describe, expect, test } from 'bun:test';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { orderedChangeId, orderedIdMinter } from '../change/ids.ts';
import { validateOrderedChange } from '../change/validate.ts';
import { openOrderedWordFile } from '../client/open.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { orderedCards } from '../decide.ts';
import { setStatus } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import { orderedUnitTyper } from '../model/typers.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import { isWml } from '../ooxml-nodes.ts';
import { orderedStampDate } from '../room/room.ts';
import {
  addressOf,
  apply,
  APPLY_FILE,
  at,
  BODY,
  DATE,
  DIRECT,
  EDIT,
  FOOTNOTES,
  markOf,
  opens,
  partsOf,
  textOf,
} from './change-apply-fixture.ts';
import { checkFold, folded, mediaStore } from './fold-check.ts';
import { anas, FOLD_SHAPES } from './fold-shapes.ts';
import { seededRandom } from './random.ts';
import { SESSION_FILE } from './sessions/fixture.ts';
import { randomChange, type ChangeBody } from './sessions/random-change.ts';
import { runSession } from './sessions/session.ts';
import { testPerson } from './test-persons.ts';
import { bodyOf, docx, MATRIX_ITEMS, p, t } from './word-fixtures.ts';

const PEOPLE = ['Olga', 'Pia', 'Quin'];

describe('a fold of a random session', () => {
  for (const seed of [7, 23]) {
    test(`seed ${seed}: the folded copy gives back what the room uses`, async () => {
      const { log, rounds } = runSession(seed, PEOPLE, 240, 3);
      const { document: doc, base } = openOrderedWordFile(SESSION_FILE, { file: 0 });
      const blobs = mediaStore(base);
      let applied = 0;
      for (const end of rounds) {
        for (; applied < end; applied += 1) {
          const change = log[applied]!;
          withUndoRecord(doc, () =>
            applyOrderedChange(doc, change, {
              person: change.by,
              date: change.at,
              ids: orderedIdMinter(change.id),
            })
          );
        }

        await checkFold(doc, base, blobs, `seed ${seed} after ${end} changes`);
      }

      // The folded copy edited as a room goes on from its checkpoint, and folded again: what the
      // copy read back from the file takes new edits and decisions. Each person's edits share one
      // date, so the decisions of each maker stay few. The copy refuses some changes by name
      const copy = await folded(doc, base, doc.file + 1, blobs);
      let n = 0;
      const send = (who: number, body: ChangeBody) => {
        const author = PEOPLE[who]!;
        const namespace = author.toLowerCase();
        n += 1;
        const checked = validateOrderedChange(
          { v: 1, id: orderedChangeId(namespace, n), ...body },
          { namespace, parts: copy.document.stories }
        );
        if (!checked.ok) {
          return;
        }

        try {
          withUndoRecord(copy.document, () =>
            applyOrderedChange(copy.document, checked.value, {
              person: testPerson(author),
              date: orderedStampDate(Date.UTC(2026, 9, 29) + 1000 * who),
              ids: orderedIdMinter(checked.value.id),
            })
          );
        } catch (error) {
          if (!(error instanceof OrderedRefusal)) {
            throw error;
          }
        }
      };

      // First each person takes back every other suggestion of theirs that the copy read from
      // the file: its paragraph or block may hold no other change, and must not show it again
      const loaded = [...copy.document.suggestions.values()].filter((s) => s.status === 'pending');
      for (const [k, suggestion] of loaded.entries()) {
        const who = PEOPLE.indexOf(suggestion.attribution.author);
        if (who >= 0 && k % 2 === 1) {
          send(who, { kind: 'retract', suggestions: [suggestion.id] });
        }
      }

      await checkFold(copy.document, copy.base, blobs, `seed ${seed} after take-backs`);

      // Then at random
      const rand = seededRandom(seed);
      for (let k = 1; k <= 80; k += 1) {
        const who = k % PEOPLE.length;
        const author = PEOPLE[who]!;
        const body = randomChange(copy.document, { author, organiser: author === 'Olga' }, rand);
        if (body !== null) {
          send(who, body);
        }
      }

      await checkFold(copy.document, copy.base, blobs, `seed ${seed} on the folded copy`);
    }, 60_000);
  }
});

describe('a fold of a formatted field', () => {
  const pageField = MATRIX_ITEMS.find((item) => item.key === 'pagefield');

  // The room reads a field's face from the run a format reaches, as the engine does: a fold
  // writes a direct format there, not on the run that begins the field
  for (const [label, mode] of [
    ['a direct format', DIRECT],
    ['a suggested format', EDIT],
  ] as const) {
    test(`${label} on a field shows the same face after the fold`, async () => {
      const { document: doc, base } = openOrderedWordFile(docx(bodyOf(pageField!)), { file: 1 });

      // The field's paragraph, between the item's neighbours
      apply(doc, 'olga:1', {
        kind: 'format',
        ...mode,
        from: at(doc, 1, 0),
        to: at(doc, 1, textOf(doc, 1).length),
        properties: [{ localName: 'b' }],
      });
      await checkFold(doc, base, mediaStore(base), label);
    });
  }
});

describe('a fold of a formatted simple field', () => {
  // A simple field's run has its own face: a format over the field shows on top of it, so the
  // room's face and the face read back after a fold agree, and words typed next to the field
  // take it on every copy
  const SIZED = '<w:rPr><w:sz w:val="28"/></w:rPr>';
  const FIELD = `<w:fldSimple w:instr=" NUMPAGES "><w:r>${SIZED}<w:t>9</w:t></w:r></w:fldSimple>`;

  test('a suggested format over a field with a sized run', async () => {
    const { document: doc, base } = openOrderedWordFile(docx(p(t('Page of ') + FIELD)), {
      file: 1,
    });
    apply(doc, 'olga:1', {
      kind: 'format',
      ...EDIT,
      from: at(doc, 0, 5),
      to: at(doc, 0, 9),
      properties: [{ localName: 'color', attributes: { val: 'C00000' } }],
    });
    await checkFold(doc, base, mediaStore(base), 'a formatted simple field', (copy) => {
      apply(copy, 'pia:1', { kind: 'insert', ...EDIT, place: at(copy, 0, 9), text: '!' }, 'Pia');
    });
  });
});

describe('a fold of words typed after a field', () => {
  // An empty run inside a field rides with it and draws inside it: words typed right after the
  // field go past it, where a fold reads it back
  const fldChar = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
  const FIELD =
    fldChar('begin') +
    '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
    '<w:r><w:rPr><w:b/></w:rPr></w:r>' +
    fldChar('separate') +
    '<w:r><w:t>3</w:t></w:r>' +
    fldChar('end');

  test('a suggested word after a field with an empty run in it', async () => {
    const { document: doc, base } = openOrderedWordFile(docx(p(t('Page ') + FIELD + t('.'))), {
      file: 1,
    });
    apply(doc, 'olga:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 6), text: ' only' });
    await checkFold(doc, base, mediaStore(base), 'a word after a field');
  });
});

describe('a fold of a hidden prompt', () => {
  // A suggested format on a prompt that direct typing then filled: the glossary keeps the prompt
  // with its record, and an emptying after the fold gives it back with that record. The prompt
  // starts at `start` in the item's paragraph
  for (const [key, start] of [
    ['placeholder', 6],
    ['blockprompt', 0],
  ] as const) {
    test(`${key}: the prompt comes back with its suggested format when its control empties`, async () => {
      const item = MATRIX_ITEMS.find((candidate) => candidate.key === key);
      const { document: doc, base } = openOrderedWordFile(docx(bodyOf(item!)), { file: 1 });
      apply(doc, 'ana:1', {
        kind: 'format',
        ...EDIT,
        from: at(doc, 1, start + 1),
        to: at(doc, 1, start + 5),
        properties: [{ localName: 'i' }],
      });
      apply(
        doc,
        'olive:2',
        { kind: 'insert', ...DIRECT, place: at(doc, 1, start), text: 'Q' },
        'Olive'
      );

      // Folded once, and twice, before the control empties
      await checkFold(doc, base, mediaStore(base), key, (copy) => {
        apply(
          copy,
          'olive:3',
          { kind: 'delete', ...DIRECT, from: at(copy, 1, start), to: at(copy, 1, start + 1) },
          'Olive'
        );
      });
    });
  }
});

describe('a fold of a filled prompt', () => {
  const item = MATRIX_ITEMS.find((candidate) => candidate.key === 'placeholder');

  // A suggested format on a prompt, then typing that fills its control: Word drops the prompt's
  // runs, and a fold keeps them with the record in the glossary only. The room shows them
  // nowhere either, so a fold changes no unit and no card. The format is the room's own, or
  // one read from a checkpoint
  for (const [label, mode, who] of [
    ['direct typing', DIRECT, 'Olive'],
    ['a suggested insertion', EDIT, 'Bob'],
  ] as const) {
    for (const loaded of [false, true]) {
      test(`${label} fills the prompt${loaded ? ' of a folded copy' : ''}`, async () => {
        let { document: doc, base } = openOrderedWordFile(docx(bodyOf(item!)), { file: 1 });
        apply(doc, 'ana:1', {
          kind: 'format',
          ...EDIT,
          from: at(doc, 1, 7),
          to: at(doc, 1, 11),
          properties: [{ localName: 'i' }],
        });
        if (loaded) {
          ({ document: doc, base } = await folded(doc, base, 2, mediaStore(base)));
        }

        apply(
          doc,
          `${who.toLowerCase()}:2`,
          { kind: 'insert', ...mode, place: at(doc, 1, 6), text: 'Q' },
          who
        );
        await checkFold(doc, base, mediaStore(base), label);
      });
    }
  }
});

describe('a prompt record the glossary keeps', () => {
  // Quin's format on a control's prompt, then his insertion fills the control: the file keeps
  // the prompt with its record in the glossary. A decision that empties the control shows the
  // prompt under the record's own decision, live, after a save and reopen, and after one and two
  // folds: accepted, the format with no record; rejected, the old face; pending, the record
  const TYPED = '2026-10-28T09:31:00Z';
  const item = MATRIX_ITEMS.find((candidate) => candidate.key === 'placeholder');
  const prompt = (doc: OrderedDocument): string => {
    const xml = partsOf(doc)[0] ?? '';
    return xml.slice(xml.indexOf('<w:sdtContent>'), xml.indexOf('</w:sdtContent>'));
  };
  const made = (doc: OrderedDocument, date: string): string =>
    [...doc.suggestions.values()].find((s) => s.attribution.date === date)?.id ?? 'none';

  async function copies(): Promise<[string, OrderedDocument][]> {
    const { document: doc, base } = openOrderedWordFile(docx(bodyOf(item!)), { file: 1 });
    apply(
      doc,
      'quin:1',
      {
        kind: 'format',
        ...EDIT,
        from: at(doc, 1, 7),
        to: at(doc, 1, 11),
        properties: [{ localName: 'i' }],
      },
      'Quin'
    );
    apply(
      doc,
      'quin:2',
      { kind: 'insert', ...EDIT, place: at(doc, 1, 6), text: 'Q' },
      'Quin',
      TYPED
    );
    const saved = openOrderedWordFile(orderedPackageBytes(doc, base), { file: 2 }).document;
    const once = await folded(doc, base, 3, mediaStore(base));
    const twice = await folded(once.document, once.base, 4, mediaStore(base));
    return [
      ['live', doc],
      ['saved', saved],
      ['folded once', once.document],
      ['folded twice', twice.document],
    ];
  }

  test('a decision that empties the control', async () => {
    const decided = (doc: OrderedDocument, format: 'accepted' | 'rejected' | null) => {
      const { record } = withUndoRecord(doc, () => {
        const formatted = made(doc, DATE);
        if (format !== null && doc.suggestions.has(formatted)) {
          setStatus(doc, formatted, format);
        }

        setStatus(doc, made(doc, TYPED), 'rejected');
      });
      const shown = prompt(doc);
      takeBack(doc, record);
      return shown;
    };
    for (const [label, doc] of await copies()) {
      const shown = [decided(doc, 'rejected'), decided(doc, 'accepted'), decided(doc, null)];
      expect([
        label,
        shown.map((xml) => [xml.includes('<w:i/>'), xml.includes('rPrChange')]),
      ]).toEqual([
        label,
        [
          [false, false],
          [true, false],
          [true, true],
        ],
      ]);
    }
  });

  test("the organiser's decisions, card by card", async () => {
    for (const [label, doc] of await copies()) {
      const decide = (id: string, suggestion: string) =>
        apply(
          doc,
          id,
          {
            kind: 'decide',
            cards: [
              { suggestions: [{ suggestion, version: doc.suggestions.get(suggestion)!.version }] },
            ],
            status: 'rejected',
          },
          'Olga'
        );
      decide('olga:1', made(doc, TYPED));
      const format = made(doc, DATE);
      expect([
        label,
        orderedCards(doc).map((card) => card.suggestions),
        prompt(doc).includes('rPrChange'),
      ]).toEqual([label, [[format]], true]);
      decide('olga:2', format);
      expect([label, prompt(doc).includes('<w:i/>'), prompt(doc).includes('rPrChange')]).toEqual([
        label,
        false,
        false,
      ]);
    }
  });
});

describe('a fold of a paragraph press that changes nothing', () => {
  // Word and the engine record no change for a press that gives a paragraph a value it has
  test('a suggested alignment the paragraph has makes no suggestion', async () => {
    const file = docx(p(t('Centred words.'), '<w:jc w:val="center"/>'));
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    apply(
      doc,
      'bob:1',
      {
        kind: 'paragraph-property',
        ...EDIT,
        marks: [markOf(doc, 0)],
        properties: [{ localName: 'jc', attributes: { val: 'center' } }],
      },
      'Bob'
    );
    expect(orderedCards(doc)).toEqual([]);
    await checkFold(doc, base, mediaStore(base), 'a press that changes nothing');
  });
});

describe('a fold of a section press', () => {
  // A section that states default values, as a file not made by Word may
  const file = docx(
    p(t('Words.')) +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720"' +
      ' w:footer="720" w:gutter="0"/><w:titlePg w:val="0"/></w:sectPr>'
  );
  const bodySection = (doc: OrderedDocument) =>
    addressOf([...doc.stories.get(BODY)!.tree.items()].find((item) => item.block === true)!);

  // Word and the engine record no change for a press that gives a section a value it has
  for (const [label, page, furniture] of [
    ['its orientation', { orientation: 'portrait' }, undefined],
    ['its page size', { pageWidthTwips: 12240, pageHeightTwips: 15840 }, undefined],
    ['its margins', { marginTopTwips: 1440, marginLeftTwips: 1440 }, undefined],
    ['its header distance', undefined, { headerDistanceTwips: 720 }],
    ['no different first page', undefined, { titlePage: false }],
  ] as const) {
    test(`a suggested press of ${label} makes no suggestion`, async () => {
      const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
      apply(
        doc,
        'bob:1',
        {
          kind: 'section-property',
          ...EDIT,
          sections: furniture === undefined ? [] : [bodySection(doc)],
          ...(page === undefined ? {} : { page }),
          ...(furniture === undefined ? {} : { furniture }),
        },
        'Bob'
      );
      expect(orderedCards(doc)).toEqual([]);
      await checkFold(doc, base, mediaStore(base), label);
    });
  }

  // The record keeps the values as the file stated them, and Bob's later press on a folded copy
  // takes it over with them. A reject gives the section back as Word restores it
  for (const [label, page] of [
    ['a landscape page', { orientation: 'landscape' }],
    ['a margin', { marginTopTwips: 2000 }],
    ['a margin with the orientation it has', { orientation: 'portrait', marginTopTwips: 2000 }],
  ] as const) {
    test(`${label} on a section that states default values`, async () => {
      const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
      apply(doc, 'ana:1', { kind: 'section-property', ...EDIT, sections: [], page });
      await checkFold(doc, base, mediaStore(base), label, (copy) => {
        apply(
          copy,
          'bob:1',
          { kind: 'section-property', ...EDIT, sections: [], page: { marginLeftTwips: 1800 } },
          'Bob'
        );
      });
      const record = /<w:sectPrChange [^>]*><w:sectPr>(.*?)<\/w:sectPr>/.exec(
        partsOf(doc)[0] ?? ''
      );
      expect(record?.[1]).toContain('<w:pgSz w:h="15840" w:orient="portrait" w:w="12240"/>');
      expect(record?.[1]).toContain('<w:titlePg w:val="0"/>');
    });
  }
});

describe('a fold of a takeover in a text box', () => {
  // Pia's press takes over Quin's record of the text box's paragraph (policy 2): once folded, the
  // text box reads as the file wrote it, and the next fold keeps the takeover too
  test("a press over another person's alignment", async () => {
    const item = MATRIX_ITEMS.find((candidate) => candidate.key === 'textbox');
    const { document: doc, base } = openOrderedWordFile(docx(bodyOf(item!)), { file: 1 });
    const part = [...doc.stories.keys()].find(isTextBoxStoryName)!;
    for (const [id, who, property] of [
      ['quin:1', 'Quin', { localName: 'jc', attributes: { val: 'center' } }],
      ['pia:1', 'Pia', { localName: 'spacing', attributes: { after: '240' } }],
    ] as const) {
      apply(
        doc,
        id,
        {
          kind: 'paragraph-property',
          part,
          mode: 'suggest',
          marks: [markOf(doc, 0, part)],
          properties: [property],
        },
        who
      );
    }

    await checkFold(doc, base, mediaStore(base), 'a takeover in a text box');
  });
});

describe('a fold of a table part inside a new row or table', () => {
  const newRow = (doc: OrderedDocument) =>
    apply(doc, 'ana:1', {
      kind: 'insert-row',
      ...EDIT,
      row: addressOf(opens(doc, 'tr')[0]!),
      where: 'below',
    });

  // Bob's new column, row or table inside Ana's new row or table sits in her insertion, as a
  // reopened Word file reads it: it goes when her insertion is rejected
  for (const [label, insert] of [
    [
      'a new column across a new row',
      (doc: OrderedDocument) => {
        newRow(doc);

        // The new row's first cell, after the first row's two
        apply(
          doc,
          'bob:1',
          { kind: 'insert-column', ...EDIT, cell: addressOf(opens(doc, 'tc')[2]!), where: 'right' },
          'Bob'
        );
      },
    ],
    [
      'a new row in a new table',
      (doc: OrderedDocument) => {
        apply(doc, 'ana:1', {
          kind: 'insert-table',
          ...EDIT,
          place: at(doc, 9, 0),
          rows: 1,
          columns: 2,
          columnWidth: 3000,
        });
        const row = addressOf(opens(doc, 'tr').at(-1)!);
        apply(doc, 'bob:1', { kind: 'insert-row', ...EDIT, row, where: 'below' }, 'Bob');
      },
    ],
    [
      'a new table in a new row',
      (doc: OrderedDocument) => {
        newRow(doc);
        apply(
          doc,
          'bob:1',
          {
            kind: 'insert-table',
            ...EDIT,
            place: at(doc, 6, 0),
            rows: 1,
            columns: 1,
            columnWidth: 2000,
          },
          'Bob'
        );
      },
    ],
  ] as const) {
    test(label, async () => {
      const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
      insert(doc);
      await checkFold(doc, base, mediaStore(base), label);
    });
  }
});

describe('a fold of a link inside an insertion', () => {
  // Ana's link on the end of her own insertion, and Bob's words typed before the link: the room
  // labels the link's edges with her insertion, a Word file holds no insertion on a link
  test("another person's words before the link stay on the insertion's card", async () => {
    const { document: doc, base } = openOrderedWordFile(docx(p(t('Cloud and serverless.'))), {
      file: 1,
    });
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 6), text: 'net ' });
    apply(doc, 'ana:2', {
      kind: 'insert-link',
      ...EDIT,
      from: at(doc, 0, 9),
      to: at(doc, 0, 10),
      href: 'https://a.example/',
    });
    apply(doc, 'bob:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 9), text: 'émigré ' }, 'Bob');
    await checkFold(doc, base, mediaStore(base), 'a link inside an insertion');
  });
});

describe('a fold of a move with words inside it', () => {
  // Bob's words inside Ana's moved words split the move into two ranges, which a fold keeps as
  // the file's own markers: once his words are rejected, the move is one range again
  test('a reject of the words inside gives one move range', async () => {
    const file = docx(p(t('Rich text here.')) + p(t('Other words.')));
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    apply(doc, 'ana:1', {
      kind: 'move',
      ...EDIT,
      from: at(doc, 0, 0),
      to: at(doc, 0, 2),
      place: at(doc, 1, 6),
    });
    apply(doc, 'bob:1', { kind: 'insert', ...EDIT, place: at(doc, 1, 7), text: 'X' }, 'Bob');
    await checkFold(doc, base, mediaStore(base), 'a move with words inside it');
  });

  // A folded move keeps the file's range markers; once the moved words went by a direct
  // deletion, the range holds nothing and draws nowhere, as the room draws no range around no
  // words, so a decision of the move gives one file before and after a fold
  test('a range whose words went', async () => {
    const file = docx(p(t('Rich text here.')) + p(t('Other words.')));
    const first = openOrderedWordFile(file, { file: 1 });
    apply(first.document, 'ana:1', {
      kind: 'move',
      ...EDIT,
      from: at(first.document, 0, 0),
      to: at(first.document, 0, 5),
      place: at(first.document, 1, 6),
    });
    const blobs = mediaStore(first.base);
    const { document: doc, base } = await folded(first.document, first.base, 2, blobs);
    apply(
      doc,
      'olive:1',
      { kind: 'delete', ...DIRECT, from: at(doc, 0, 0), to: at(doc, 0, 5) },
      'Olive'
    );
    await checkFold(doc, base, blobs, 'a range whose words went');
  });

  // A rejected move's words are in their run again, which a fold reads with no move: the room
  // draws them in that run too, also in a cell a pending column deletion strikes
  test('a rejected move out of a struck cell', async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    apply(doc, 'ana:1', {
      kind: 'move',
      ...EDIT,
      from: at(doc, 5, 1),
      to: at(doc, 5, 2),
      place: at(doc, 0, 0),
    });
    apply(doc, 'ana:2', { kind: 'delete-column', ...EDIT, cell: addressOf(opens(doc, 'tc')[1]!) });
    const suggestions = orderedCardSeen(doc, 'S:ana:1.1')!;
    const decide = { kind: 'decide', cards: [{ suggestions }], status: 'rejected' } as const;
    apply(doc, 'olive:1', decide, 'Olive');
    await checkFold(doc, base, mediaStore(base), 'a rejected move out of a struck cell');
  });

  // Olive's direct move is no suggestion: her moved words, which Ana then strikes, show as Ana's
  // plain deletion, though the room keeps the move's name on them; a fold writes them plain
  test("a direct move's words another person struck", async () => {
    const file = docx(p(t('Rich text here.')) + p(t('Other words.')));
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    const move = { kind: 'move', ...DIRECT, from: at(doc, 0, 0), to: at(doc, 0, 5) } as const;
    apply(doc, 'olive:1', { ...move, place: at(doc, 1, 6) }, 'Olive');
    apply(doc, 'ana:1', { kind: 'delete', ...EDIT, from: at(doc, 1, 6), to: at(doc, 1, 11) });
    await checkFold(doc, base, mediaStore(base), "a direct move's words struck");
  });

  // Olive's bold inside Ben's takes over his records; once his card is rejected, a fold reads the
  // run's one change record as one record of hers, the same face
  test("a format that took over another's, once that other is rejected", async () => {
    const file = docx(p(t('Rich text here.')) + p(t('Other words.')));
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    const bold = { kind: 'format', ...EDIT, properties: [{ localName: 'b' }] } as const;
    apply(doc, 'ben:1', { ...bold, from: at(doc, 0, 0), to: at(doc, 0, 10) }, 'Ben');
    apply(doc, 'olive:1', { ...bold, from: at(doc, 0, 2), to: at(doc, 0, 6) }, 'Olive');
    const seen = [{ suggestion: 'S:ben:1.1', version: doc.suggestions.get('S:ben:1.1')!.version }];
    const reject = { kind: 'decide', cards: [{ suggestions: seen }], status: 'rejected' } as const;
    apply(doc, 'olive:2', reject, 'Olive');
    await checkFold(doc, base, mediaStore(base), "a format that took over another's");
  });
});

describe('who typed inside a new row, cell, note, control or words', () => {
  // Olive's units in a part that `olives` picks: who holds each, as a reader of the labels finds
  // it (its own label, else the new row or cell around it), and who typed it
  const reads = (
    doc: OrderedDocument,
    part: string,
    olives: (item: OrderedItem, unit: string) => boolean
  ) =>
    [...(doc.stories.get(part)?.tree.items() ?? [])].flatMap((item) =>
      [...item.text].flatMap((unit, k) =>
        olives(item, unit)
          ? [
              {
                holder: item.by ?? item.within.by.at(-1) ?? null,
                typer: orderedUnitTyper(doc, item, item.offset + k),
              },
            ]
          : []
      )
    );
  const typed = (_: OrderedItem, unit: string) => unit === 'X' || unit === 'Y';
  const olive = { namespace: 'olive', person: testPerson('Olive') };
  const anas = { holder: 'S:ana:1.1', typer: olive };

  // Each copy of the room: live, and after one and two folds
  async function copiesOf(doc: OrderedDocument, base: OoxmlPackage) {
    const blobs = mediaStore(base);
    const once = await folded(doc, base, 2, blobs);
    const twice = await folded(once.document, once.base, 3, blobs);
    return [doc, once.document, twice.document];
  }

  // Ana's new control round "fee", and the k-th unit of the words it took
  const newControl = (doc: OrderedDocument) => {
    apply(doc, 'ana:1', {
      kind: 'wrap-control',
      ...EDIT,
      from: at(doc, 0, 4),
      to: at(doc, 0, 7),
      control: { type: 'richText' },
    });
    return opens(doc, 'sdt').find((open) => open.by === 'S:ana:1.1')!;
  };
  const inNewControl = (doc: OrderedDocument, k: number) => {
    const items = [...(doc.stories.get(BODY)?.tree.items() ?? [])];
    const open = items.findIndex((item) => item.token === 'sdt' && item.by === 'S:ana:1.1');
    const words = items.slice(open).find((item) => item.kind === 'text')!;
    return { space: words.space, offset: words.offset + k };
  };

  // Ana's new row, cell, note, control or words, and Olive's direct typing in it: the container
  // or the insertion covers the words, as Word writes them with no mark of their own, so a reader
  // of the labels finds them in Ana's insertion and the read names Olive who typed them, in the
  // room and after each fold. In a control and in words Ana typed in two parts, Olive types
  // where the parts meet: one insertion in the file on both sides
  const containers = [
    [
      'a new row',
      BODY,
      (doc: OrderedDocument) => {
        apply(doc, 'ana:1', {
          kind: 'insert-row',
          ...EDIT,
          row: addressOf(opens(doc, 'tr')[0]!),
          where: 'below',
        });

        // The new row's first cell
        return at(doc, 6, 0);
      },
    ],
    [
      'a new cell',
      BODY,
      (doc: OrderedDocument) => {
        apply(doc, 'ana:1', {
          kind: 'insert-column',
          ...EDIT,
          cell: addressOf(opens(doc, 'tc')[0]!),
          where: 'right',
        });

        // The new column's first cell, after A1
        return at(doc, 5, 0);
      },
    ],
    [
      'a new note',
      FOOTNOTES,
      (doc: OrderedDocument) => {
        apply(doc, 'ana:1', {
          kind: 'insert-note',
          ...EDIT,
          place: at(doc, 0, 3),
          note: 'footnote',
          text: '',
        });

        // After the new note's reference mark
        return at(doc, 3, 1, FOOTNOTES);
      },
    ],
    [
      'a new control',
      BODY,
      (doc: OrderedDocument) => {
        newControl(doc);
        apply(doc, 'ana:2', { kind: 'insert', ...EDIT, place: inNewControl(doc, 1), text: 'new ' });

        // Between Ana's typing and the rest of the words the control took
        return inNewControl(doc, 1);
      },
    ],
    [
      'new words typed in two parts',
      BODY,
      (doc: OrderedDocument) => {
        apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 4), text: 'new ' });
        apply(doc, 'ana:2', { kind: 'insert', ...EDIT, place: at(doc, 0, 6), text: 'ab' });

        // After "ab", before the "w" of "new "
        return at(doc, 0, 8);
      },
    ],
  ] as const;
  for (const [label, part, insert] of containers) {
    test(`${label}: the read names the organiser who typed in it`, async () => {
      const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
      const place = insert(doc);
      apply(doc, 'olive:2', { kind: 'insert', part, mode: 'edit', place, text: 'XY' }, 'Olive');
      const copies = await copiesOf(doc, base);
      expect(copies.map((copy) => reads(copy, part, typed))).toEqual(
        copies.map(() => [anas, anas])
      );
      await checkFold(doc, base, mediaStore(base), label);
    });
  }

  // Olive's typing at the start of Ana's new control, in it as the keyboard types on in a control
  // (`inside`): one `w:ins` holds the whole control, so the words are Ana's insertion's
  test("the start of a new control: the organiser's typing is in it", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const inside = addressOf(newControl(doc));
    const place = inNewControl(doc, 0);
    apply(doc, 'olive:2', { kind: 'insert', ...DIRECT, place, inside, text: 'XY' }, 'Olive');
    const copies = await copiesOf(doc, base);
    expect(copies.map((copy) => reads(copy, BODY, typed))).toEqual(copies.map(() => [anas, anas]));
    await checkFold(doc, base, mediaStore(base), 'the start of a new control');
  });

  // Olive's direct footnote in Ana's new row, cell, control or words: its reference goes where a
  // direct object goes, into the container or the insertion, as Word writes it there; the note
  // itself stays Olive's
  for (const [label, part, insert] of containers.filter(([, part]) => part === BODY)) {
    test(`${label}: a direct note's reference is the container's`, async () => {
      const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
      const place = insert(doc);
      apply(
        doc,
        'olive:2',
        { kind: 'insert-note', part, mode: 'edit', place, note: 'footnote', text: 'XY' },
        'Olive'
      );
      const reference = (item: OrderedItem) =>
        item.node !== null && isWml(item.node, 'footnoteReference');
      const copies = await copiesOf(doc, base);
      expect(copies.map((copy) => reads(copy, part, reference))).toEqual(copies.map(() => [anas]));
      await checkFold(doc, base, mediaStore(base), label);
    });
  }

  // Olive's direct paste in Ana's new row, cell, note or control: its words are the container's,
  // as her typing is, and the read names Olive who pasted them. In Ana's words a paste sits
  // beside them, as Word pastes (`client/editor-paste-inside.test.ts`)
  for (const [label, part, insert] of containers.filter(([name]) => !name.includes('words'))) {
    test(`${label}: a direct paste's words are the container's`, async () => {
      const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
      const place = insert(doc);
      const blocks = [{ e: 'w:p', c: [{ e: 'w:r', c: [{ e: 'w:t', c: ['XY'] }] }] }];
      apply(
        doc,
        'olive:2',
        { kind: 'insert-fragment', part, mode: 'edit', place, blocks },
        'Olive'
      );
      const copies = await copiesOf(doc, base);
      expect(copies.map((copy) => reads(copy, part, typed))).toEqual(
        copies.map(() => [anas, anas])
      );
      await checkFold(doc, base, mediaStore(base), label);
    });
  }

  // Ana's value in a plain text control: the engine's template wrote its words, and the read
  // names Ana who set them, in the room and after each fold
  test("a control's value: the read names who set it", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const control = addressOf(opens(doc, 'sdt')[0]!);
    apply(doc, 'ana:1', {
      kind: 'control-value',
      ...EDIT,
      control,
      value: { kind: 'text', text: 'XY' },
    });
    const ana = { namespace: 'ana', person: testPerson('Ana') };
    const value = { holder: 'S:ana:1.1', typer: ana };
    const copies = await copiesOf(doc, base);
    expect(copies.map((copy) => reads(copy, BODY, typed))).toEqual(
      copies.map(() => [value, value])
    );
    await checkFold(doc, base, mediaStore(base), 'a control value');
  });
});

const W16DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';

describe('a fold of a long list', () => {
  test('one suggested deletion over 300 paragraphs', async () => {
    const paragraphs = Array.from({ length: 300 }, (_, k) => p(t(`Paragraph ${k}.`))).join('');
    const { document: doc, base } = openOrderedWordFile(docx(paragraphs), { file: 1 });
    apply(doc, 'olga:1', {
      kind: 'delete',
      ...EDIT,
      from: at(doc, 0, 0),
      to: at(doc, 299, textOf(doc, 299).length),
    });
    await checkFold(doc, base, mediaStore(base), 'a deletion over 300 paragraphs');
  });

  test("a Word file's 20,000 own tracked changes", async () => {
    const mark = (id: number) =>
      `<w:ins w:id="${id}" w:author="Reviewer" w:date="${DATE}" xmlns:w16du="${W16DU}" ` +
      `w16du:dateUtc="${DATE}">${t('x')}</w:ins>`;
    const body = Array.from({ length: 1000 }, (_, k) =>
      p(Array.from({ length: 20 }, (_, j) => mark(k * 20 + j)).join(''))
    ).join('');
    const { document: doc, base } = openOrderedWordFile(docx(body), { file: 1 });
    const again = await folded(doc, base, 2, mediaStore(base));
    expect(again.document.suggestions.size).toBe(20_000);
  });
});

describe("a fold of a Word file's own values", () => {
  const inserted = (attributes: string) =>
    docx(p(`<w:ins ${attributes}>${t('x')}</w:ins>${t('y')}`));
  const many = Array.from({ length: 10 }, (_, k) => `a${k}="1"`).join(' ');

  // Every value a load keeps from a revision mark comes back
  for (const [label, attributes] of [
    ['an empty date', 'w:id="1" w:author="A" w:date=""'],
    ['a long author', `w:id="1" w:author="${'a'.repeat(300)}"`],
    ['a tab in the author', 'w:id="1" w:author="A&#9;B"'],
    ['an attribute with no namespace', 'w:id="1" w:author="A" foo="1"'],
    ['a long attribute value', `w:id="1" w:author="A" w:rsidR="${'0'.repeat(300)}"`],
    ['more attributes than a load keeps', `w:id="1" w:author="A" ${many}`],
    ['a long w:id', `w:id="${'9'.repeat(70)}" w:author="A"`],
  ] as const) {
    test(`an insertion with ${label}`, async () => {
      const { document: doc, base } = openOrderedWordFile(inserted(attributes), { file: 1 });
      await checkFold(doc, base, mediaStore(base), label);
    });
  }

  test('one w:id of two authors: two suggestions, each with its own units', async () => {
    const mark = (author: string, text: string) =>
      `<w:ins w:id="1" w:author="${author}" w:date="${DATE}">${t(text)}</w:ins>`;
    const file = docx(p(mark('A', 'ab') + mark('B', 'cd')));
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    const again = await folded(doc, base, 2, mediaStore(base));
    const unitsOf = (copy: OrderedDocument) =>
      [...copy.suggestions.values()].map(({ id, attribution }) => [
        attribution.author,
        [...(copy.labelled.get(id) ?? [])].map((item) => item.text).join(''),
      ]);
    expect(unitsOf(again.document)).toEqual([
      ['A', 'ab'],
      ['B', 'cd'],
    ]);
    expect(unitsOf(again.document)).toEqual(unitsOf(doc));
  });
});

describe('a take-back of a suggestion a fold loaded', () => {
  // Ana's suggestion on the Word file, folded, then taken back on the folded copy: the file the
  // copy writes holds what the copy shows, so the next folds bring none of it back. A take-back
  // removes the labels of what the file marked, whose paragraph and blocks then draw again
  for (const [label, file, suggest] of FOLD_SHAPES) {
    test(label, async () => {
      const first = openOrderedWordFile(file, { file: 1 });
      suggest(first.document);
      const blobs = mediaStore(first.base);
      const { document: doc, base } = await folded(first.document, first.base, 2, blobs);
      expect(anas(doc).length).toBeGreaterThan(0);
      apply(doc, 'ana:3', { kind: 'retract', suggestions: anas(doc) });
      await checkFold(doc, base, blobs, label);
    });
  }
});
