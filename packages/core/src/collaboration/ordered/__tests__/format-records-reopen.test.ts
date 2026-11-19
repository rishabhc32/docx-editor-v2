// The room keeps only the format records the Word file holds, for every item kind of the matrix
// and for pasted ones. A format over a paragraph and its mark is saved and loaded again: each
// character and the mark show the same pending format suggestions, by author, and the same run
// properties as before the save. An equation has no run, so a format leaves it as it is, as Word
// does.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPackage, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { currentFace } from '../client/properties.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import { orderedIdMinter } from '../change/ids.ts';
import { applyOrderedEdit, type OrderedEdit } from '../edits/apply.ts';
import { orderedPlaceAt } from '../edits/places.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { statusOf } from '../model/document.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import { APPLY_FILE, apply, at, BODY, DIRECT, load } from './change-apply-fixture.ts';
import { bodyOf, extrasOf, KINDS, paragraphOf, type MatrixKind } from './matrix/catalogue.ts';
import { reopened, saved } from './matrix/word.ts';
import { testAttribution } from './test-persons.ts';
import { docx, packageOf } from './word-fixtures.ts';

const DATE = '2026-11-20T09:00:00Z';
const ITALIC = { properties: [{ localName: 'i' }] };

/** The kinds whose format is refused by name, as the editor refuses it: the refusal codes. */
const REFUSED: Readonly<Record<string, readonly string[]>> = {
  // The words of a control bound to custom XML, and a table of contents' generated result
  bound: ['not-covered'],
  toc: ['not-covered'],
};

/** A node's properties as a string to compare, ids aside. */
function propertiesKey(node: OoxmlNode): string {
  if (node.kind === 'textValue') {
    return node.value;
  }

  const attributes = node.attributes
    .map((a) => `${a.localName}=${a.value}`)
    .sort()
    .join(',');
  return `${node.localName}(${attributes})[${node.children.map(propertiesKey).join(',')}]`;
}

/** The story a kind's edits go to in the model. */
function storyOf(doc: OrderedDocument, kind: MatrixKind): string {
  if (kind.story?.scope !== 'textbox') {
    return kind.story?.part ?? BODY;
  }

  const found = [...doc.stories.keys()].find(isTextBoxStoryName);
  if (found === undefined) {
    throw new Error('the model has no text box story');
  }

  return found;
}

function viewOf(doc: OrderedDocument, story: string, index: number) {
  const view = orderedViewParagraph(doc, story, index);
  if (!view) {
    throw new Error(`no paragraph ${index} in ${story}`);
  }

  return view;
}

/**
 * The records of each character of a paragraph and of its mark, ids aside: the authors of the
 * suggestions that hold a pending format of it (one change record each in the file), and the run
 * properties it shows now. A whole field's properties are its runs' own, so its face is left out
 * (its records are not).
 */
function recordsOf(doc: OrderedDocument, name: string, index: number): string[] {
  const story = doc.stories.get(name);
  if (!story) {
    throw new Error(`no story ${name}`);
  }

  const records = (item: OrderedItem) => {
    const pending = item.fmts
      .map((record) => record.suggestion)
      .filter((id) => statusOf(doc, id) === 'pending');
    const authors = [...new Set(pending)].map(
      (id) => doc.suggestions.get(id)?.attribution.author ?? '?'
    );
    const face = item.pieces === undefined ? currentFace(doc, story, item) : null;
    return `[${authors.join(',')}] ${face && face.children.length > 0 ? propertiesKey(face) : ''}`;
  };
  const view = viewOf(doc, name, index);
  return [...view.chars.map((c) => `${c.ch} ${records(c.item)}`), `mark ${records(view.mark)}`];
}

/**
 * Italic over a paragraph and on its mark, by `author` (the organiser's edit when `direct`).
 * Returns the codes of the edits refused.
 */
function formatParagraph(
  doc: OrderedDocument,
  story: string,
  index: number,
  author: string,
  direct = false
): string[] {
  const view = viewOf(doc, story, index);
  const end = orderedPlaceAt(view, view.chars.length);
  const edits: OrderedEdit[] = [
    { kind: 'format', from: orderedPlaceAt(view, 0), to: end, ...ITALIC },
    { kind: 'formatMark', at: end, ...ITALIC },
  ];
  const refused: string[] = [];
  for (const [k, edit] of edits.entries()) {
    const id = `${author.toLowerCase()}-1:${k + 1}`;
    const ctx = {
      doc,
      story,
      attribution: testAttribution(author, DATE),
      ids: orderedIdMinter(id),
      ...(direct ? { direct: true } : {}),
    };
    try {
      applyOrderedEdit(ctx, edit);
    } catch (error) {
      if (!(error instanceof OrderedRefusal)) {
        throw error;
      }

      refused.push(error.reason);
    }
  }

  return refused;
}

/** A paragraph's records after a save and reopen equal its live ones. */
function expectSameAfterReopen(
  pkg: OoxmlPackage,
  doc: OrderedDocument,
  story: string,
  index: number
): void {
  const again = reopened(saved(pkg, doc));
  expect(recordsOf(again, story, index)).toEqual(recordsOf(doc, story, index));
}

/** A kind's file, and the model loaded from it. */
function opened(kind: MatrixKind): { pkg: OoxmlPackage; doc: OrderedDocument } {
  const read = readOoxmlPackage(docx(bodyOf(kind), extrasOf(kind)), {});
  if (!read.ok) {
    throw new Error(read.reason);
  }

  return { pkg: read.package, doc: loadOrderedPackage(read.package) };
}

describe('format records survive a save and reopen for every item kind', () => {
  for (const kind of KINDS) {
    for (const direct of [false, true]) {
      test(`${kind.key}: ${direct ? "the organiser's direct format" : "a party's format"}`, () => {
        const { pkg, doc } = opened(kind);
        const story = storyOf(doc, kind);
        const author = direct ? 'Olive' : 'Ana';
        expect(formatParagraph(doc, story, paragraphOf(kind), author, direct)).toEqual([
          ...(REFUSED[kind.key] ?? []),
        ]);
        expectSameAfterReopen(pkg, doc, story, paragraphOf(kind));
      });
    }
  }

  test("the organiser's new link over words and an equation: the equation takes no style", () => {
    const math = KINDS.find((kind) => kind.key === 'math');
    if (!math) {
      throw new Error('no math kind');
    }

    const { pkg, doc } = opened(math);
    const index = paragraphOf(math);
    const end = viewOf(doc, BODY, index).chars.length;
    const link = {
      ...DIRECT,
      kind: 'insert-link',
      from: at(doc, index, 0),
      to: at(doc, index, end),
      href: 'https://example.com',
      style: 'Hyperlink',
    } as const;
    apply(doc, 'olive-1:1', link, 'Olive');
    expect(recordsOf(doc, BODY, index).filter((entry) => entry.startsWith('￼'))).toEqual(['￼ [] ']);
    expectSameAfterReopen(pkg, doc, BODY, index);
  });

  // A suggested paste of an equation is refused by the engine: only the organiser pastes one
  test("a party's format over the organiser's pasted equation and its words", () => {
    const doc = load();
    const run = (text: string) => ({ e: 'w:r', c: [{ e: 'w:t', c: [text] }] });
    const equation = { e: 'm:oMath', c: [{ e: 'm:r', c: [{ e: 'm:t', c: ['x=1'] }] }] };
    const blocks = [{ e: 'w:p', c: [run('Let '), equation, run(' hold')] }];
    apply(
      doc,
      'olive-1:1',
      { ...DIRECT, kind: 'insert-fragment', place: at(doc, 0, 4), blocks },
      'Olive'
    );
    expect(formatParagraph(doc, BODY, 0, 'Cara')).toEqual([]);
    const live = recordsOf(doc, BODY, 0);
    expect(live.filter((entry) => entry.startsWith('￼'))).toEqual(['￼ [] ']);
    expect(live[0]).toBe('T [Cara] rPr()[i()[]]');
    expectSameAfterReopen(packageOf(APPLY_FILE), doc, BODY, 0);
  });
});
