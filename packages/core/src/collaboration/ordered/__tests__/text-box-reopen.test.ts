// A file with text boxes saves and reopens to the same document, its paragraph ids included
// (group 9.9). A room's addresses and a host's anchors read a paragraph by its `w14:paraId`, so
// a save must not leave a paragraph without one, or with one another paragraph holds, for the
// reopen to mint anew. Text boxes in the body, in a table cell and in a header; with no edit,
// with typing, Enter and a format inside the box, and with a party's new text box.

import { describe, expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint } from '../../../store/package/ooxml-serialize.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedIdMinter } from '../change/ids.ts';
import type { OrderedChange } from '../change/types.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import { composeOrderedPackage, orderedPackageBytes } from '../compose/package.ts';
import { orderedPlaceAt } from '../edits/places.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import type { OrderedDocument } from '../model/types.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { BODY, DATE } from './change-apply-fixture.ts';
import { bodyOf, docx, MATRIX_EXTRAS, MATRIX_ITEMS, p, R, t, W, WML_CT } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

const HEADER = '/word/header1.xml';

const xmlOf = (key: string): string => {
  const found = MATRIX_ITEMS.find((candidate) => candidate.key === key);
  if (!found?.xml) throw new Error(`no ${key} item`);
  return found.xml;
};

/** A text box in the body, one in a table cell, and one in the header. */
const FILE = docx(
  bodyOf(MATRIX_ITEMS.find((candidate) => candidate.key === 'textbox')!) +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/>' +
    `</w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t('Cell ') + xmlOf('textbox'))}</w:tc></w:tr></w:tbl>` +
    p(t('After.')) +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr>',
  {
    ...MATRIX_EXTRAS,
    parts: {
      ...MATRIX_EXTRAS.parts,
      'word/header1.xml': `<w:hdr xmlns:w="${W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${p(t('Head ') + xmlOf('textbox'))}</w:hdr>`,
    },
    contentTypes: `${MATRIX_EXTRAS.contentTypes ?? ''}<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>`,
    relationships: `${MATRIX_EXTRAS.relationships ?? ''}<Relationship Id="rIdHeader" Type="${R}/header" Target="header1.xml"/>`,
  }
);

type Body = OrderedChange extends infer C
  ? C extends OrderedChange
    ? Omit<C, 'v' | 'id'>
    : never
  : never;

const open = (bytes = FILE) => openOrderedWordFile(bytes, { file: 1 });

function apply(doc: OrderedDocument, id: string, body: Body, author = 'Ana'): void {
  const change = { v: 1, id, ...body } as OrderedChange;
  withUndoRecord(doc, () =>
    applyOrderedChange(doc, change, {
      person: testPerson(author),
      date: DATE,
      ids: orderedIdMinter(id),
    })
  );
}

/** The text box stories of a document, in the order their objects take in the body. */
const boxes = (doc: OrderedDocument): string[] =>
  [...doc.stories.keys()].filter(isTextBoxStoryName);

const at = (doc: OrderedDocument, story: string, index: number, n: number) => {
  const view = orderedViewParagraph(doc, story, index);
  if (!view) throw new Error(`no paragraph ${index} in ${story}`);
  return orderedPlaceAt(view, n);
};

/** Every `w14:paraId` of a part, in document order (a text box's two copies included). */
const paraIds = (pkg: OoxmlPackage, name: string): string[] =>
  [
    ...serializeOoxmlPart(pkg.parts.get(name)!).matchAll(/<w:p [^>]*w14:paraId="([0-9A-F]{8})"/g),
  ].map((m) => m[1] ?? '');

/** Every paragraph of a part, and how many of them have an id. */
const paragraphs = (pkg: OoxmlPackage, name: string) => {
  const xml = serializeOoxmlPart(pkg.parts.get(name)!);
  return { all: (xml.match(/<w:p[ >]/g) ?? []).length, ids: paraIds(pkg, name).length };
};

/** The document saved and opened again, as the room's next file opens. */
function reopened(doc: OrderedDocument, base: OoxmlPackage) {
  const before = composeOrderedPackage(doc, base);
  const again = open(orderedPackageBytes(doc, base));
  const after = composeOrderedPackage(again.document, again.base);
  return { before, after, again: again.document };
}

function expectStable(doc: OrderedDocument, base: OoxmlPackage): void {
  const { before, after, again } = reopened(doc, base);
  for (const name of [BODY, HEADER]) {
    expect({ name, ids: paraIds(after, name) }).toEqual({ name, ids: paraIds(before, name) });
    const counts = paragraphs(before, name);
    expect({ name, missing: counts.all - counts.ids }).toEqual({ name, missing: 0 });
    expect(canonicalOoxmlFingerprint(after.parts.get(name)!)).toBe(
      canonicalOoxmlFingerprint(before.parts.get(name)!)
    );
  }
  expect(boxes(again)).toHaveLength(boxes(doc).length);
}

describe('text boxes save and reopen with the same paragraph ids (group 9.9)', () => {
  test('with no edit: in the body, in a table cell and in the header', () => {
    const { document: doc, base } = open();
    // The header's text box is a story too, as the body's and the cell's are
    expect(boxes(doc)).toHaveLength(3);
    expectStable(doc, base);
  });

  test('with typing, Enter and a format inside each text box, by a party and the organiser', () => {
    const { document: doc, base } = open();
    const [body, cell] = boxes(doc);
    if (!body || !cell) throw new Error('two text boxes');
    apply(doc, 'ana-1:1', {
      kind: 'insert',
      part: body,
      mode: 'suggest',
      place: at(doc, body, 0, 3),
      text: 'xy',
    });
    apply(doc, 'ana-1:2', {
      kind: 'split',
      part: body,
      mode: 'suggest',
      place: at(doc, body, 0, 5),
    });
    apply(
      doc,
      'olive-1:1',
      { kind: 'insert', part: cell, mode: 'edit', place: at(doc, cell, 0, 0), text: 'Z' },
      'Olive'
    );
    apply(
      doc,
      'olive-1:2',
      { kind: 'split', part: cell, mode: 'edit', place: at(doc, cell, 0, 2) },
      'Olive'
    );
    apply(
      doc,
      'olive-1:3',
      {
        kind: 'format',
        part: cell,
        mode: 'edit',
        from: at(doc, cell, 1, 0),
        to: at(doc, cell, 1, 3),
        properties: [{ localName: 'b' }],
      },
      'Olive'
    );
    expectStable(doc, base);
  });

  test("a party's new text box, typed in and split, and the organiser's header edit", () => {
    const { document: doc, base } = open();
    apply(doc, 'ana-1:1', {
      kind: 'insert-object',
      part: BODY,
      mode: 'suggest',
      place: at(doc, BODY, 2, 0),
      object: { type: 'textbox', extent: { cx: 914_400 } },
    } as Body);
    const added = boxes(doc).find((name) => !name.includes('#0.'));
    if (!added) throw new Error('no new text box story');
    apply(doc, 'ana-1:2', {
      kind: 'insert',
      part: added,
      mode: 'suggest',
      place: at(doc, added, 0, 0),
      text: 'New box',
    });
    apply(doc, 'ana-1:3', {
      kind: 'split',
      part: added,
      mode: 'suggest',
      place: at(doc, added, 0, 3),
    });
    apply(
      doc,
      'olive-1:1',
      { kind: 'insert', part: HEADER, mode: 'edit', place: at(doc, HEADER, 0, 0), text: 'Draft ' },
      'Olive'
    );
    expectStable(doc, base);
  });
});
