// A decision after a fold gives the Word file the same decision gives in the room (design D14).
// A pending suggestion splits the run it sits in, and a fold writes each piece as a run of its
// own; a decision that takes the suggestion away draws the room's run whole again. A new list's
// definition or style goes with its suggestion. A table record keeps what it took from another.

import { describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { openOrderedWordFile } from '../client/open.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  addressOf,
  apply,
  APPLY_FILE,
  at,
  DIRECT,
  EDIT,
  markOf,
  opens,
} from './change-apply-fixture.ts';
import { checkFold, checkpointBytes, mediaStore } from './fold-check.ts';
import {
  alphaOmega,
  bodyOf,
  docx,
  MATRIX_EXTRAS,
  MATRIX_ITEMS,
  p,
  type PackageExtras,
} from './word-fixtures.ts';

/** A run as Word writes it, with its revision session and its own properties. */
const run = (text: string, rPr = '') =>
  `<w:r w:rsidR="00A96A35"><w:rPr><w:sz w:val="24"/>${rPr}</w:rPr><w:t>${text}</w:t></w:r>`;

/** The file's own format record on a run, by another author. */
const FILE_FORMAT =
  '<w:rPrChange w:id="7" w:author="Author" w:date="2026-07-07T18:18:00Z">' +
  '<w:rPr><w:color w:val="000000"/></w:rPr></w:rPrChange>';

async function foldOf(
  body: string,
  edit: (doc: OrderedDocument) => void,
  label: string,
  extras?: PackageExtras
) {
  const { document: doc, base } = openOrderedWordFile(docx(body, extras), { file: 1 });
  edit(doc);
  await checkFold(doc, base, mediaStore(base), label);
}

describe('a decision after a fold, on a suggestion inside a run', () => {
  test('a suggested word', async () => {
    await foldOf(
      p(run('Cloud and serverless.')),
      (doc) => {
        apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 6), text: 'net ' });
      },
      'a word inside a run'
    );
  });

  test('a suggested deletion', async () => {
    await foldOf(
      p(run('Cloud and serverless.')),
      (doc) => {
        apply(doc, 'ana:1', { kind: 'delete', ...EDIT, from: at(doc, 0, 6), to: at(doc, 0, 10) });
      },
      'a deletion inside a run'
    );
  });

  test("a suggested word in a run with the file's format record", async () => {
    await foldOf(
      p(run('Lorem ipsum dolor.', FILE_FORMAT)),
      (doc) => {
        apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 7), text: 'X ' });
      },
      'a word inside a formatted run'
    );
  });

  // The room's run goes on in the new paragraph: its piece there is drawn with new ids
  test('a suggested paragraph break, beside a deletion in a combo box', async () => {
    const combo = MATRIX_ITEMS.find((item) => item.key === 'combo');
    await foldOf(
      p(run('Cloud and serverless.')) + bodyOf(combo!),
      (doc) => {
        apply(doc, 'ana:1', { kind: 'delete', ...EDIT, from: at(doc, 2, 6), to: at(doc, 2, 7) });
        apply(doc, 'bob:1', { kind: 'split', ...EDIT, place: at(doc, 0, 6) }, 'Bob');
      },
      'a paragraph break inside a run'
    );
  });
});

describe("a checkpoint's room run numbers", () => {
  /** The room run numbers (`oc:run`) the main document of a checkpoint holds, in order. */
  async function numbersOf(edit: (doc: OrderedDocument) => void) {
    const body = p(run('Cloud and serverless.')) + p(run('Untouched.'));
    const { document: doc, base } = openOrderedWordFile(docx(body), { file: 1 });
    edit(doc);
    const bytes = await checkpointBytes(doc, base, mediaStore(base));
    const xml = strFromU8(unzipSync(bytes)['word/document.xml']!);
    return {
      numbers: [...xml.matchAll(/ oc:run="(\d+)"/g)].map((m) => m[1]),
      declared: xml.includes('xmlns:oc='),
    };
  }

  test('only on the pieces of a run a suggestion splits: each run and its text', async () => {
    const { numbers, declared } = await numbersOf((doc) => {
      apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 6), text: 'net ' });
    });
    expect({ numbers, declared }).toEqual({
      numbers: [numbers[0], numbers[1], numbers[0], numbers[1]],
      declared: true,
    });
    expect(numbers[0]).not.toBe(numbers[1]);
  });

  test('none, and no namespace, where no suggestion splits a run', async () => {
    const drawn = await numbersOf((doc) => {
      const text = ' Yes.';
      apply(doc, 'olive:1', { kind: 'insert', ...DIRECT, place: at(doc, 0, 21), text }, 'Olive');
    });
    expect(drawn).toEqual({ numbers: [], declared: false });
  });
});

describe('a decision after a fold, on a new list', () => {
  /** A new list on paragraph `k`: Ana's suggestion, or Olive's direct edit. */
  const listOn = (doc: OrderedDocument, k: number, id = 'ana:1') =>
    apply(doc, id, {
      kind: 'new-list',
      ...(id.startsWith('olive') ? DIRECT : EDIT),
      list: {
        abstract: {
          e: 'w:abstractNum',
          c: [
            {
              e: 'w:lvl',
              a: { 'w:ilvl': '0' },
              c: [
                { e: 'w:numFmt', a: { 'w:val': 'lowerLetter' } },
                { e: 'w:lvlText', a: { 'w:val': '(%1)' } },
              ],
            },
          ],
        },
      },
      paragraphs: [{ mark: markOf(doc, k), level: 0 }],
    });
  const newList = (doc: OrderedDocument) => listOn(doc, 0);
  const ITEMS = p(run('First item.')) + p(run('Second item.'));

  test('in a file with no lists: the numbering part goes with it', async () => {
    await foldOf(ITEMS, newList, 'a new list, no numbering part');
  });

  test("in a file with lists of its own: the file's stay", async () => {
    await foldOf(ITEMS, newList, 'a new list, a numbering part', MATRIX_EXTRAS);
  });

  test('beside a direct new list: the definitions keep their order', async () => {
    await foldOf(
      ITEMS,
      (doc) => {
        listOn(doc, 1, 'olive:1');
        listOn(doc, 0);
      },
      'a new list beside a direct one'
    );
  });
});

describe('a decision after a fold, on a pasted style', () => {
  const STYLE = {
    e: 'w:style',
    a: { 'w:type': 'paragraph', 'w:styleId': 'PastedClause' },
    c: [
      { e: 'w:name', a: { 'w:val': 'Pasted Clause' } },
      { e: 'w:rPr', c: [{ e: 'w:i' }] },
    ],
  } as const;
  const BLOCK = {
    e: 'w:p',
    c: [
      { e: 'w:pPr', c: [{ e: 'w:pStyle', a: { 'w:val': 'PastedClause' } }] },
      { e: 'w:r', c: [{ e: 'w:t', c: ['Pasted.'] }] },
    ],
  } as const;

  test('whole paragraphs with their style: a reject takes the style away', async () => {
    await foldOf(
      p(run('Cloud and serverless.')),
      (doc) => {
        const place = at(doc, 0, 6);
        const paste = { place, blocks: [BLOCK], styles: [STYLE], whole: true };
        apply(doc, 'ana:1', { kind: 'insert-fragment', ...EDIT, ...paste });
      },
      'a pasted style'
    );
  });
});

describe('a decision after a fold, beside a pending control value', () => {
  // A check box's value record names elements of another namespace than the main one: a fold
  // keeps their prefix, so the part a decision edits still reads
  const xmlOf = (key: string) => MATRIX_ITEMS.find((item) => item.key === key)?.xml ?? '';

  test('an accepted deletion empties a drop-down list beside a ticked check box', async () => {
    await foldOf(
      p(alphaOmega(xmlOf('dropdown'))) + p(alphaOmega(xmlOf('checkbox'))),
      (doc) => {
        apply(doc, 'ana:1', {
          kind: 'control-value',
          ...EDIT,
          control: opens(doc, 'sdt')[1]!,
          value: { kind: 'checkbox', checked: true },
        });
        apply(doc, 'bob:1', { kind: 'delete', ...EDIT, from: at(doc, 0, 6), to: at(doc, 0, 9) });
      },
      'a drop-down list beside a check box'
    );
  });
});

describe("a fold of a table's records", () => {
  // Two dates: each is decided on its own
  test('two row heights of one person on a row with no properties', async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    for (const [id, heightTwips, date] of [
      ['ana:1', 300, '2026-10-28T09:01:00Z'],
      ['ana:2', 480, '2026-10-28T09:02:00Z'],
    ] as const) {
      const on = [addressOf(opens(doc, 'tr')[0]!)];
      const op = { op: 'setTableRowHeight', heightTwips } as const;
      apply(doc, id, { kind: 'table-property', ...EDIT, on, op }, 'Ana', date);
    }

    await checkFold(doc, base, mediaStore(base), 'two row heights');
  });
});
