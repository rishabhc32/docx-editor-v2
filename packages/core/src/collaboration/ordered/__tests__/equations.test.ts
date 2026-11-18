// Equations in the ordered model (group 9.9): a new linear form strikes the old equation object
// and inserts the new one at its place, in one suggestion, so a decision shows one or the other
// whole; a direct edit keeps the ids the editor names and its undo brings the old one back; an
// equation of a display that holds others goes alone; two copies make the same document. A file's
// equation struck as Word writes it (a `w:del` in each math run) is one suggestion.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { orderedIdMinter } from '../change/ids.ts';
import { applyOrderedChange } from '../change/apply.ts';
import type { OrderedChange } from '../change/types.ts';
import { validateOrderedChange } from '../change/validate.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { composeOrderedPart, orderedViewParagraph } from '../compose/compose.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { BODY, DATE } from './change-apply-fixture.ts';
import { snapshot } from './snapshot.ts';
import { docx, M, p, t } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

const INLINE = docx(
  p(t('Let ') + `<m:oMath xmlns:m="${M}"><m:r><m:t>x=1</m:t></m:r></m:oMath>` + t(' hold.'))
);
const DISPLAY = docx(
  p(
    t('Then ') +
      `<m:oMathPara xmlns:m="${M}"><m:oMath><m:r><m:t>a=1</m:t></m:r></m:oMath>` +
      '<m:oMath><m:r><m:t>b=2</m:t></m:r></m:oMath></m:oMathPara>'
  )
);

type Body = OrderedChange extends infer C
  ? C extends OrderedChange
    ? Omit<C, 'v' | 'id'>
    : never
  : never;

const open = (bytes = INLINE) => {
  const opened = openOrderedWordFile(bytes, { file: 1 });
  return { doc: opened.document, base: opened.base };
};

function apply(doc: OrderedDocument, id: string, body: Body, author = 'Ana'): string[] {
  const checked = validateOrderedChange({ v: 1, id, ...body }, { parts: doc.stories });
  if (!checked.ok) throw new Error(`${checked.error.code} ${checked.error.path}`);
  return [
    ...withUndoRecord(doc, () =>
      applyOrderedChange(doc, checked.value, {
        person: testPerson(author),
        date: DATE,
        ids: orderedIdMinter(id),
      })
    ).value.suggestions,
  ];
}

/** The first equation object of the body that shows. */
function objectOf(doc: OrderedDocument): OrderedItem {
  const view = orderedViewParagraph(doc, BODY, 0);
  const found = view?.chars.find((c) => c.item.bare && !c.deleted)?.item;
  if (!found) throw new Error('no equation');
  return found;
}

const body = (doc: OrderedDocument) => serializeOoxmlPart(composeOrderedPart(doc, BODY).part);
const text = (doc: OrderedDocument) =>
  orderedViewParagraph(doc, BODY, 0)
    ?.chars.map((c) => c.ch)
    .join('');

function decide(doc: OrderedDocument, ids: readonly string[], status: 'accepted' | 'rejected') {
  apply(
    doc,
    `olive-1:${doc.suggestions.size}`,
    {
      kind: 'decide',
      status,
      cards: [
        {
          suggestions: ids.map((suggestion) => ({
            suggestion,
            version: doc.suggestions.get(suggestion)?.version ?? 0,
          })),
        },
      ],
    } as Body,
    'Olive'
  );
}

describe('equations (group 9.9)', () => {
  test('a suggested new form: the old equation struck, the new one inserted, one suggestion', () => {
    const run = () => {
      const { doc } = open();
      const object = objectOf(doc);
      const ids = apply(doc, 'ana-1:1', {
        kind: 'equation',
        part: BODY,
        mode: 'suggest',
        object: { space: object.space, offset: object.offset },
        equation: 0,
        linear: 'x=2',
      });
      return { doc, ids };
    };
    const one = run();
    const two = run();
    expect(one.ids).toHaveLength(1);
    expect(snapshot(one.doc, false)).toEqual(snapshot(two.doc, false));
    const xml = body(one.doc);
    expect(xml).toMatch(
      /<w:del [^>]*w:author="Ana"[^>]*><m:oMath[^>]*>[\s\S]*?x=1[\s\S]*?<\/m:oMath><\/w:del>/
    );
    expect(xml).toMatch(
      /<w:ins [^>]*w:author="Ana"[^>]*><m:oMath[^>]*>[\s\S]*?x=2[\s\S]*?<\/m:oMath><\/w:ins>/
    );
    expect(xml.indexOf('x=1')).toBeLessThan(xml.indexOf('x=2'));
    expect(text(one.doc)).toBe('Let ￼￼ hold.');
    // Rejected: the old equation alone, as the file had it
    const { doc: rejected, base } = open();
    const before = serializeOoxmlPart(base.parts.get(BODY)!);
    const object = objectOf(rejected);
    const ids = apply(rejected, 'ana-1:1', {
      kind: 'equation',
      part: BODY,
      mode: 'suggest',
      object: { space: object.space, offset: object.offset },
      equation: 0,
      linear: 'x=2',
    });
    decide(rejected, ids, 'rejected');
    expect(body(rejected)).toBe(before);
    // Accepted: the new equation alone
    decide(one.doc, one.ids, 'accepted');
    const accepted = body(one.doc);
    expect(accepted).not.toContain('x=1');
    expect(accepted).not.toContain('<w:ins');
    expect(accepted).toMatch(/<m:oMath[^>]*>[\s\S]*?x=2/);
  });

  test("the organiser's new form keeps the editor's ids, and its undo brings the old one back", () => {
    const { doc, base } = open();
    const object = objectOf(doc);
    const id = object.node?.id ?? '';
    const before = body(doc);
    const made = apply(
      doc,
      'olive-1:1',
      {
        kind: 'equation',
        part: BODY,
        mode: 'edit',
        object: { space: object.space, offset: object.offset },
        equation: 0,
        linear: '{a+b}/{2}',
      },
      'Olive'
    );
    const xml = body(doc);
    expect(xml).not.toContain('<w:ins');
    expect(xml).not.toContain('x=1');
    expect(xml).toContain('<m:f>');
    expect(objectOf(doc).node?.id).toBe(id);
    apply(
      doc,
      'olive-1:2',
      { kind: 'revert', suggestions: made, status: 'rejected' } as Body,
      'Olive'
    );
    expect(body(doc)).toBe(before);
    // Saved and opened again, the new form is the file's
    apply(
      doc,
      'olive-1:3',
      { kind: 'revert', suggestions: made, status: 'accepted' } as Body,
      'Olive'
    );
    const reopened = open(orderedPackageBytes(doc, base));
    expect(body(reopened.doc)).toContain('<m:f>');
    expect(reopened.doc.suggestions.size).toBe(0);
  });

  test('one equation of a display goes alone, and the other stays', () => {
    const { doc } = open(DISPLAY);
    const object = objectOf(doc);
    apply(
      doc,
      'olive-1:1',
      {
        kind: 'equation',
        part: BODY,
        mode: 'edit',
        object: { space: object.space, offset: object.offset },
        equation: 0,
        linear: null,
      },
      'Olive'
    );
    const xml = body(doc);
    expect(xml).not.toContain('a=1');
    expect(xml).toMatch(
      /<m:oMathPara[^>]*><m:oMath[^>]*>[\s\S]*?b=2[\s\S]*?<\/m:oMath><\/m:oMathPara>/
    );
  });

  test('a form that does not parse, and a lone equation removed by this kind, are refused', () => {
    const { doc } = open();
    const object = objectOf(doc);
    const at = { space: object.space, offset: object.offset };
    const before = snapshot(doc);
    for (const linear of ['{a+', null])
      expect(() =>
        apply(doc, 'ana-1:1', {
          kind: 'equation',
          part: BODY,
          mode: 'suggest',
          object: at,
          equation: 0,
          linear,
        })
      ).toThrow();
    expect(() =>
      apply(doc, 'ana-1:1', {
        kind: 'equation',
        part: BODY,
        mode: 'suggest',
        object: at,
        equation: 3,
        linear: 'y',
      })
    ).toThrow();
    expect(snapshot(doc)).toEqual(before);
  });

  test('an equation Word struck in each math run is one suggestion, decided both ways', () => {
    // As Word 16 wrote it (sample `b-equation-word.docx`): a `w:del` in each `m:r`, own ids
    const stamp = 'w:author="Ana" w:date="2026-10-29T23:18:00Z"';
    const run = (id: number, text: string, mark = 'del') =>
      `<m:r><w:${mark} w:id="${id}" ${stamp}><w:rPr><w:rFonts w:ascii="Cambria Math"/></w:rPr>` +
      `<m:t>${text}</m:t></w:${mark}></m:r>`;
    const struck = docx(
      p(t('Let ') + `<m:oMath xmlns:m="${M}">${run(0, 'x')}${run(1, '=1')}</m:oMath>` + t(' hold.'))
    );
    const pending = (doc: OrderedDocument) =>
      [...doc.suggestions.values()].filter((s) => s.status === 'pending').map((s) => s.id);
    for (const status of ['accepted', 'rejected'] as const) {
      const { doc } = open(struck);
      expect(pending(doc)).toHaveLength(1);
      const atom = orderedViewParagraph(doc, BODY, 0)?.chars.find((c) => c.item.bare)?.item;
      expect(atom?.dels).toEqual(pending(doc));
      // Written in the form Word reads as one deletion: the mark around the equation
      expect(body(doc)).toMatch(/<w:del [^>]*w:author="Ana"[^>]*w:id="0"[^>]*><m:oMath[ >]/);
      expect(body(doc)).not.toMatch(/<m:r><w:del /);
      decide(doc, pending(doc), status);
      if (status === 'accepted') expect(body(doc)).not.toContain('<m:oMath');
      else {
        expect(body(doc)).not.toContain('<w:del ');
        expect(body(doc)).toMatch(/<m:t>x<\/m:t>.*<m:t>=1<\/m:t>/);
      }
    }
    // Opened again from the room's own save: still one suggestion
    const { doc } = open(struck);
    expect(pending(open(orderedPackageBytes(doc, open(struck).base)).doc)).toHaveLength(1);
    // Marks on some runs only stay as the file has them: the list has no unit inside an equation
    const changed = docx(
      p(
        t('Let ') +
          `<m:oMath xmlns:m="${M}"><m:r><m:t>x=</m:t></m:r>${run(3, '1')}${run(4, '2', 'ins')}</m:oMath>` +
          t(' hold.')
      )
    );
    const partial = open(changed);
    expect(pending(partial.doc)).toHaveLength(0);
    expect(body(partial.doc)).toMatch(/<m:r><w:del [^>]*>.*<m:t>1<\/m:t><\/w:del><\/m:r>/);
  });
});
