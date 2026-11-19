// Text boxes in the ordered model (group 9.9): each text box of the main document is a story of
// its own, so typing, deletion, formatting, Enter and a join inside it are ordinary changes,
// suggested or direct, decided, taken back and saved as Word writes them (`w:ins` and `w:del`
// inside `w:txbxContent`); the object around it keeps its node.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPackage, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { orderedCardOf } from '../card-of.ts';
import { orderedIdMinter } from '../change/ids.ts';
import { applyOrderedChange } from '../change/apply.ts';
import type { OrderedChange } from '../change/types.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { composeOrderedPart, orderedViewParagraph } from '../compose/compose.ts';
import { composeOrderedPackage, orderedPackageBytes } from '../compose/package.ts';
import { decideOrdered, orderedCards } from '../decide.ts';
import { orderedPlaceAt } from '../edits/places.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { isTextBoxStoryName, textBoxHostOf } from '../model/text-boxes.ts';
import type { OrderedDocument } from '../model/types.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { suggestionsOnlyPolicy } from '../room/suggestions-only.ts';
import { orderedChangeTargets } from '../change/targets.ts';
import { BODY, DATE } from './change-apply-fixture.ts';
import { snapshot } from './snapshot.ts';
import { bodyOf, docx, MATRIX_EXTRAS, MATRIX_ITEMS } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

const item = (key: string) => {
  const found = MATRIX_ITEMS.find((candidate) => candidate.key === key);
  if (!found) throw new Error(`no ${key} item`);
  return found;
};

const FILE = docx(bodyOf(item('textbox')), MATRIX_EXTRAS);
const VML_FILE = docx(bodyOf(item('vml')), MATRIX_EXTRAS);

type Body = OrderedChange extends infer C
  ? C extends OrderedChange
    ? Omit<C, 'v' | 'id'>
    : never
  : never;

function open(bytes = FILE): { doc: OrderedDocument; base: OoxmlPackage } {
  const opened = openOrderedWordFile(bytes, { file: 1 });
  return { doc: opened.document, base: opened.base };
}

function apply(doc: OrderedDocument, id: string, body: Body, author = 'Ana'): string[] {
  const change = { v: 1, id, ...body } as OrderedChange;
  return [
    ...withUndoRecord(doc, () =>
      applyOrderedChange(doc, change, {
        person: testPerson(author),
        date: DATE,
        ids: orderedIdMinter(id),
      })
    ).value.suggestions,
  ];
}

/** The one text box story of a document. */
function boxOf(doc: OrderedDocument): string {
  const names = [...doc.stories.keys()].filter(isTextBoxStoryName);
  expect(names).toHaveLength(1);
  return names[0] ?? '';
}

const text = (doc: OrderedDocument, story: string, index = 0) =>
  orderedViewParagraph(doc, story, index)
    ?.chars.map((c) => c.ch)
    .join('') ?? '<none>';

const at = (doc: OrderedDocument, story: string, index: number, n: number) => {
  const view = orderedViewParagraph(doc, story, index);
  if (!view) throw new Error(`no paragraph ${index}`);
  return orderedPlaceAt(view, n);
};

const body = (doc: OrderedDocument) => serializeOoxmlPart(composeOrderedPart(doc, BODY).part);

/** The DrawingML copy of the text box in the body, as XML. */
const boxXml = (xml: string) => /<wps:txbx>([\s\S]*?)<\/wps:txbx>/.exec(xml)?.[1] ?? '<none>';

describe('text boxes (group 9.9)', () => {
  test('the text box is its own story, and an untouched one composes to the file', () => {
    const { doc, base } = open();
    const box = boxOf(doc);
    expect(text(doc, box)).toBe('Box text');
    expect(textBoxHostOf(doc, box)?.host).toBe(BODY);
    // The body's paragraph keeps the object as one unit
    expect(text(doc, BODY, 1)).toBe('Alpha ￼ Omega.');
    const xml = body(doc);
    expect(boxXml(xml)).toContain('Box text');
    expect(xml).toBe(serializeOoxmlPart(base.parts.get(BODY)!));
    // No part is written for the text box's story
    expect([...composeOrderedPackage(doc, base).parts.keys()].some(isTextBoxStoryName)).toBe(false);
  });

  test('a suggestion inside it is `w:ins` in the text box, on two copies alike, and a reject takes it back', () => {
    const run = () => {
      const { doc } = open();
      const box = boxOf(doc);
      const [sid = ''] = apply(doc, 'ana-1:1', {
        kind: 'insert',
        part: box,
        mode: 'suggest',
        place: at(doc, box, 0, 3),
        text: ' new',
      });
      return { doc, box, sid };
    };
    const one = run();
    const two = run();
    expect(text(one.doc, one.box)).toBe('Box new text');
    expect(snapshot(one.doc, false)).toEqual(snapshot(two.doc, false));
    const xml = body(one.doc);
    expect(boxXml(xml)).toMatch(
      /<w:ins [^>]*w:author="Ana"[^>]*><w:r><w:t xml:space="preserve"> new<\/w:t><\/w:r><\/w:ins>/
    );
    expect(xml).toBe(body(two.doc));
    const found = one.doc.suggestions.get(one.sid);
    apply(
      one.doc,
      'olive-1:1',
      {
        kind: 'decide',
        status: 'rejected',
        cards: [{ suggestions: [{ suggestion: one.sid, version: found?.version ?? 0 }] }],
      } as Body,
      'Olive'
    );
    expect(text(one.doc, one.box)).toBe('Box text');
    expect(boxXml(body(one.doc))).not.toContain('<w:ins');
  });

  test('a deletion, Enter and a format inside it, then a save that opens again', () => {
    const { doc, base } = open();
    const box = boxOf(doc);
    apply(doc, 'ana-1:1', {
      kind: 'delete',
      part: box,
      mode: 'suggest',
      from: at(doc, box, 0, 0),
      to: at(doc, box, 0, 4),
    });
    apply(doc, 'ana-1:2', { kind: 'split', part: box, mode: 'suggest', place: at(doc, box, 0, 4) });
    apply(doc, 'ana-1:3', {
      kind: 'format',
      part: box,
      mode: 'suggest',
      from: at(doc, box, 1, 0),
      to: at(doc, box, 1, 4),
      properties: [{ localName: 'b' }],
    });
    const xml = boxXml(body(doc));
    expect(xml).toContain('<w:delText xml:space="preserve">Box </w:delText>');
    expect(xml).toContain('<w:rPrChange');
    expect(xml.match(/<w:p[ >]/g)).toHaveLength(2);
    // The saved file holds the text box's suggestions, in both of its copies, and opens again
    const bytes = orderedPackageBytes(doc, base);
    const read = readOoxmlPackage(bytes, {});
    if (!read.ok) throw new Error(read.reason);
    const saved = serializeOoxmlPart(read.package.parts.get(BODY)!);
    expect(saved.match(/>Box <\/w:delText>/g)).toHaveLength(2);
    const again = open(bytes);
    const reopened = boxOf(again.doc);
    expect(text(again.doc, reopened, 0)).toBe('Box ');
    expect(text(again.doc, reopened, 1)).toBe('text');
    expect(again.doc.suggestions.size).toBeGreaterThanOrEqual(3);
  });

  test("a direct edit inside it, and its undo, as the organiser's", () => {
    const { doc } = open();
    const box = boxOf(doc);
    const before = body(doc);
    const [sid = ''] = apply(
      doc,
      'olive-1:1',
      { kind: 'insert', part: box, mode: 'edit', place: at(doc, box, 0, 0), text: 'My ' },
      'Olive'
    );
    expect(text(doc, box)).toBe('My Box text');
    expect(boxXml(body(doc))).not.toContain('<w:ins');
    apply(
      doc,
      'olive-1:2',
      { kind: 'revert', suggestions: [sid], status: 'rejected' } as Body,
      'Olive'
    );
    expect(body(doc)).toBe(before);
  });

  test('a legacy VML text box is a story too', () => {
    const { doc } = open(VML_FILE);
    const box = boxOf(doc);
    expect(text(doc, box)).toBe('Old box');
    apply(doc, 'ana-1:1', {
      kind: 'insert',
      part: box,
      mode: 'suggest',
      place: at(doc, box, 0, 0),
      text: 'An ',
    });
    expect(body(doc)).toMatch(/<v:textbox><w:txbxContent><w:p[^>]*><w:ins /);
  });

  test("another person's new text box: typing inside it is inside their suggestion", () => {
    const { doc, base } = open();
    const box = boxOf(doc);
    // Ana's copy of the text box is her insertion, and its text box a story of its own
    const view = orderedViewParagraph(doc, BODY, 1);
    const object = view?.chars.find((c) => c.item.kind === 'atom');
    if (!view || !object) throw new Error('fixture');
    apply(doc, 'ana-1:1', {
      kind: 'move',
      part: BODY,
      mode: 'suggest',
      from: { space: object.address.space, offset: object.address.offset },
      to: at(doc, BODY, 1, 7),
      place: at(doc, BODY, 2, 0),
    });
    const boxes = [...doc.stories.keys()].filter(isTextBoxStoryName);
    expect(boxes).toHaveLength(2);
    const copy = boxes.find((name) => name !== box) ?? '';
    expect(text(doc, copy)).toBe('Box text');
    const change = {
      v: 1,
      id: 'ben-1:1',
      kind: 'insert',
      part: copy,
      mode: 'suggest',
      place: at(doc, copy, 0, 4),
      text: 'x',
    } as OrderedChange;
    const targets = orderedChangeTargets(doc, change);
    if (!targets.ok) throw new Error('targets');
    const verdict = suggestionsOnlyPolicy().check({
      actor: {
        person: testPerson('Ben'),
        namespace: 'ben-1',
        suggests: true,
        decides: false,
        readOnly: false,
        comments: true,
      },
      change,
      document: doc,
      touched: targets.touched,
    });
    expect(verdict).toMatchObject({ ok: false, code: 'policy-nested' });
    // Ana's own typing there joins her insertion; a reject of it takes the copy and its text box
    const [sid = ''] = apply(doc, 'ana-1:2', {
      kind: 'insert',
      part: copy,
      mode: 'suggest',
      place: at(doc, copy, 0, 4),
      text: '!',
    });
    expect(text(doc, copy)).toBe('Box !text');
    const found = doc.suggestions.get(sid);
    apply(
      doc,
      'olive-1:1',
      {
        kind: 'decide',
        status: 'rejected',
        cards: [{ suggestions: [{ suggestion: sid, version: found?.version ?? 0 }] }],
      } as Body,
      'Olive'
    );
    expect(body(doc)).toBe(serializeOoxmlPart(base.parts.get(BODY)!));
  });

  test('a text box story the change does not name is refused by name', () => {
    const { doc } = open();
    expect(() =>
      apply(doc, 'ana-1:1', {
        kind: 'insert',
        part: `${BODY}!txbx:nothing`,
        mode: 'suggest',
        place: at(doc, BODY, 0, 0),
        text: 'x',
      })
    ).toThrow(OrderedRefusal);
  });

  test('a note in a text box is refused by name, as Word ignores a note reference there', () => {
    // [MS-OI29500] 17.11.14(a), 17.11.7(b); the editor refuses it too (the engine's insertNote)
    const { doc } = open();
    const box = boxOf(doc);
    const refusal = (place: ReturnType<typeof at>, part: string): string => {
      try {
        apply(doc, 'olive-1:1', {
          kind: 'insert-note',
          part,
          mode: 'edit',
          place,
          note: 'footnote',
          text: '',
        });
      } catch (error) {
        if (error instanceof OrderedRefusal) return error.reason;
        throw error;
      }
      return 'applied';
    };
    expect(refusal(at(doc, box, 0, 1), box)).toBe('note-outside-body');
    // The paragraph that holds the text box takes one
    expect(refusal(at(doc, BODY, 0, 1), BODY)).toBe('applied');
  });

  test('a change in a text box that no longer shows is refused by name; a struck one takes it', () => {
    const { doc } = open();
    const box = boxOf(doc);
    const reason = (id: string, change: Body): string => {
      try {
        apply(doc, id, change);
      } catch (error) {
        if (error instanceof OrderedRefusal) return error.reason;
        throw error;
      }
      return 'applied';
    };
    const typing = (): Body => ({
      kind: 'insert',
      part: box,
      mode: 'suggest',
      place: at(doc, box, 0, 4),
      text: 'x',
    });
    const centred = (): Body => ({
      kind: 'paragraph-property',
      part: box,
      mode: 'suggest',
      marks: [{ ...orderedViewParagraph(doc, box, 0)!.mark }],
      properties: [{ localName: 'jc', attributes: { val: 'center' } }],
    });
    const view = orderedViewParagraph(doc, BODY, 1);
    const k = view?.chars.findIndex((c) => c.item.kind === 'atom') ?? -1;
    if (k < 0) throw new Error('fixture');
    const [deletion = ''] = apply(doc, 'ana-1:1', {
      kind: 'delete',
      part: BODY,
      mode: 'suggest',
      from: at(doc, BODY, 1, k),
      to: at(doc, BODY, 1, k + 1),
    });
    // Struck, the text box still shows, as Word shows it: a change in it is a suggestion
    expect(reason('ana-1:2', typing())).toBe('applied');
    expect(text(doc, box)).toBe('Box xtext');
    const late = { typing: typing(), centred: centred() };
    apply(
      doc,
      'olive-1:1',
      {
        kind: 'decide',
        status: 'accepted',
        cards: [
          {
            suggestions: [
              { suggestion: deletion, version: doc.suggestions.get(deletion)!.version },
            ],
          },
        ],
      } as Body,
      'Olive'
    );
    // Gone with its object: nothing a change made there would show. The changes name places
    // from before the decision, as a copy that has not seen it yet sends them
    const before = JSON.stringify(snapshot(doc, false));
    expect(reason('ana-1:3', late.typing)).toBe('hidden-target');
    expect(reason('ana-1:4', late.centred)).toBe('hidden-target');
    expect(JSON.stringify(snapshot(doc, false))).toBe(before);
  });

  test('an edit that names its item in a text box that no longer shows is refused by name', () => {
    const { doc } = open();
    const box = boxOf(doc);
    const reason = (id: string, change: Body): string => {
      try {
        apply(doc, id, change, 'Olive');
      } catch (error) {
        if (error instanceof OrderedRefusal) return error.reason;
        throw error;
      }
      return 'applied';
    };
    // A link around "Box" and a second paragraph, so that each edit would apply in a shown box
    const from = at(doc, box, 0, 0);
    const to = at(doc, box, 0, 3);
    const href = 'https://example.com';
    apply(
      doc,
      'olive-1:1',
      { kind: 'insert-link', part: box, mode: 'edit', from, to, href },
      'Olive'
    );
    apply(doc, 'olive-1:2', { kind: 'split', part: box, mode: 'edit', place: to }, 'Olive');
    const items = [...doc.stories.get(box)!.tree.items()];
    const link = items.find((i) => i.kind === 'open' && i.token === 'hyperlink')!;
    const mark = orderedViewParagraph(doc, box, 0)!.mark;
    const unlink: Body = {
      kind: 'remove-link',
      part: box,
      mode: 'edit',
      link: { space: link.space, offset: link.offset },
    };
    const removal: Body = {
      kind: 'delete-paragraph',
      part: box,
      mode: 'edit',
      mark: { space: mark.space, offset: mark.offset },
    };
    // The text box object goes, with the deletion accepted
    const view = orderedViewParagraph(doc, BODY, 1);
    const k = view?.chars.findIndex((c) => c.item.kind === 'atom') ?? -1;
    const [deletion = ''] = apply(doc, 'ana-1:1', {
      kind: 'delete',
      part: BODY,
      mode: 'suggest',
      from: at(doc, BODY, 1, k),
      to: at(doc, BODY, 1, k + 1),
    });
    const version = doc.suggestions.get(deletion)!.version;
    const accept = {
      kind: 'decide',
      status: 'accepted',
      cards: [{ suggestions: [{ suggestion: deletion, version }] }],
    } as Body;
    apply(doc, 'olive-1:3', accept, 'Olive');
    const before = JSON.stringify(snapshot(doc, false));
    expect(reason('olive-1:4', unlink)).toBe('hidden-target');
    expect(reason('olive-1:5', removal)).toBe('hidden-target');
    expect(JSON.stringify(snapshot(doc, false))).toBe(before);
  });

  test('a suggestion in a text box that no longer shows is on no card', () => {
    // The object goes by an accepted deletion, or by the organiser's direct one
    for (const mode of ['suggest', 'edit'] as const) {
      const { doc } = open();
      const box = boxOf(doc);
      const [enter = ''] = apply(
        doc,
        'ben-1:1',
        { kind: 'split', part: box, mode: 'suggest', place: at(doc, box, 0, 4) },
        'Ben'
      );
      const view = orderedViewParagraph(doc, BODY, 1);
      const k = view?.chars.findIndex((c) => c.item.kind === 'atom') ?? -1;
      const removal: Body = {
        kind: 'delete',
        part: BODY,
        mode,
        from: at(doc, BODY, 1, k),
        to: at(doc, BODY, 1, k + 1),
      };
      const [deletion = ''] = apply(doc, 'olive-1:1', removal, 'Olive');
      if (mode === 'suggest') {
        const version = doc.suggestions.get(deletion)!.version;
        const accept = {
          kind: 'decide',
          status: 'accepted',
          cards: [{ suggestions: [{ suggestion: deletion, version }] }],
        } as Body;
        apply(doc, 'olive-1:2', accept, 'Olive');
      }

      // Compose writes no such text box, so its Enter has no card, as a reopen of the file has
      // none: a decision on it is refused by name
      expect({ mode, card: orderedCardOf(doc, enter) }).toEqual({ mode, card: null });
      expect(orderedCards(doc).flatMap((card) => card.suggestions)).not.toContain(enter);
      expect(() => decideOrdered(doc, enter, 'reject')).toThrow('hidden-target');
    }
  });
});
