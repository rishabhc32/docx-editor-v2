// The edits: deterministic from their id source, whole or refused by name with nothing changed,
// taken back exactly by their undo record, and the decided policies (design D5).

import { describe, expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  serializeOoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import {
  applyOrderedEdit,
  composeOrderedPart,
  decideOrdered,
  loadOrderedPackage,
  orderedCards,
  orderedPlaceAt,
  orderedViewParagraph,
  OrderedRefusal,
  takeBackOrdered,
  type OrderedDocument,
  type OrderedEdit,
  type OrderedPlace,
  type OrderedReplacedFormat,
} from '../index.ts';
import { snapshot } from './snapshot.ts';
import { testIds } from './test-ids.ts';
import { docx, MATRIX_EXTRAS, p, packageOf, t } from './word-fixtures.ts';
import { testAttribution } from './test-persons.ts';

const BODY = '/word/document.xml';
const DATE = '2026-10-18T00:00:00Z';
const BOLD = '<w:b/>';

function load(body: string): OrderedDocument {
  return loadOrderedPackage(packageOf(docx(body, MATRIX_EXTRAS)));
}

const text = (doc: OrderedDocument, index: number) =>
  orderedViewParagraph(doc, BODY, index)
    ?.chars.map((c) => c.ch)
    .join('') ?? '<none>';

/** The place before the n-th character of a paragraph (its end when n is its length). */
function at(doc: OrderedDocument, index: number, n: number): OrderedPlace {
  const view = orderedViewParagraph(doc, BODY, index);
  if (!view) throw new Error(`no paragraph ${index}`);
  return orderedPlaceAt(view, n);
}

function edit(
  doc: OrderedDocument,
  author: string,
  id: string,
  change: OrderedEdit,
  more: { direct?: boolean; onReplaced?: (r: OrderedReplacedFormat) => void } = {}
) {
  return applyOrderedEdit(
    { doc, story: BODY, attribution: testAttribution(author, DATE), ids: testIds(id), ...more },
    change
  );
}

const xml = (doc: OrderedDocument) => serializeOoxmlPart(composeOrderedPart(doc, BODY).part);

const SENTENCE = p(t('Alpha ') + t('beta', BOLD) + t(' gamma.')) + p(t('Second line.'));

describe('ordered edits', () => {
  test('the same edits with the same ids give the same document on two copies', () => {
    const run = () => {
      const doc = load(SENTENCE);
      edit(doc, 'Ana', 'c1', { kind: 'insert', at: at(doc, 0, 6), text: 'new ' });
      edit(doc, 'Ben', 'c2', { kind: 'delete', from: at(doc, 0, 10), to: at(doc, 0, 14) });
      edit(doc, 'Ana', 'c3', { kind: 'split', at: at(doc, 0, 3) });
      edit(doc, 'Ben', 'c4', {
        kind: 'format',
        from: at(doc, 1, 0),
        to: at(doc, 2, 3),
        properties: [{ localName: 'i' }],
      });
      edit(doc, 'Ana', 'c5', { kind: 'join', at: at(doc, 2, 0) });
      return doc;
    };
    const one = run();
    const two = run();
    expect(snapshot(one, false)).toEqual(snapshot(two, false));
    expect(canonicalOoxmlFingerprint(composeOrderedPart(one, BODY).part)).toBe(
      canonicalOoxmlFingerprint(composeOrderedPart(two, BODY).part)
    );
    expect([...one.suggestions.keys()]).toEqual([
      'S:c1:1.1',
      'S:c2:1.1',
      'S:c3:1.1',
      'S:c4:1.1',
      'S:c5:1.1',
    ]);
  });

  const kinds: Record<string, (doc: OrderedDocument) => OrderedEdit> = {
    insert: (doc) => ({ kind: 'insert', at: at(doc, 0, 3), text: 'xyz' }),
    delete: (doc) => ({ kind: 'delete', from: at(doc, 0, 2), to: at(doc, 1, 4) }),
    replace: (doc) => ({ kind: 'replace', from: at(doc, 0, 6), to: at(doc, 0, 10), text: 'B' }),
    format: (doc) => ({
      kind: 'format',
      from: at(doc, 0, 0),
      to: at(doc, 1, 6),
      properties: [{ localName: 'u', attributes: { val: 'double' } }],
    }),
    split: (doc) => ({ kind: 'split', at: at(doc, 0, 8) }),
    join: (doc) => ({ kind: 'join', at: at(doc, 1, 0) }),
    paragraph: (doc) => ({
      kind: 'paragraph',
      at: at(doc, 0, 0),
      ops: [
        {
          op: 'setParagraphProperties',
          properties: [{ localName: 'jc', attributes: { val: 'right' } }],
        },
      ],
    }),
    move: (doc) => ({ kind: 'move', from: at(doc, 0, 0), to: at(doc, 0, 5), at: at(doc, 1, 0) }),
  };

  for (const [name, make] of Object.entries(kinds)) {
    test(`${name}: its undo record takes it back exactly`, () => {
      const doc = load(SENTENCE);
      const before = snapshot(doc);
      const part = xml(doc);
      const { record } = edit(doc, 'Ana', 'u', make(doc));
      expect(snapshot(doc)).not.toEqual(before);
      takeBackOrdered(doc, record);
      expect(snapshot(doc)).toEqual(before);
      expect(xml(doc)).toBe(part);
    });

    test(`${name}: a direct edit is settled at once`, () => {
      const doc = load(SENTENCE);
      const { value } = edit(doc, 'Olive', 'd', make(doc), { direct: true });
      expect(value.suggestions.length).toBeGreaterThan(0);
      for (const id of value.suggestions) expect(doc.suggestions.get(id)?.status).toBe('accepted');
      expect(orderedCards(doc)).toEqual([]);
    });
  }

  test('a refused edit changes nothing and names its reason', () => {
    const doc = load(SENTENCE);
    const before = snapshot(doc);
    const refusals: [OrderedEdit, string][] = [
      [{ kind: 'insert', at: at(doc, 0, 1), text: 'a\tb' }, 'invalid-edit'],
      [{ kind: 'delete', from: at(doc, 0, 4), to: at(doc, 0, 4) }, 'empty-range'],
      [{ kind: 'join', at: at(doc, 0, 0) }, 'no-paragraph-before'],
      [{ kind: 'join', at: at(doc, 1, 3) }, 'not-paragraph-start'],
      [{ kind: 'insert', at: { space: 'b:nowhere', offset: 0 }, text: 'x' }, 'unknown-address'],
      [{ kind: 'insertRow', at: at(doc, 0, 0), where: 'below' }, 'not-in-table'],
      [{ kind: 'linkTarget', at: at(doc, 0, 1), url: 'https://example.com' }, 'not-in-link'],
      [{ kind: 'retract', suggestion: 'S9' }, 'unknown-suggestion'],
    ];
    for (const [change, reason] of refusals) {
      let caught: unknown;
      try {
        edit(doc, 'Ana', 'r', change);
      } catch (error) {
        caught = error;
      }
      expect({
        kind: change.kind,
        reason: caught instanceof OrderedRefusal ? caught.reason : caught,
      }).toEqual({
        kind: change.kind,
        reason,
      });
      expect(snapshot(doc)).toEqual(before);
    }
  });

  test('a link target with only one end of its part is refused, not applied to the whole link', () => {
    const doc = load(p(`<w:hyperlink w:anchor="mark">${t('Jump here')}</w:hyperlink>`));
    const before = snapshot(doc);
    const url = 'https://example.com';
    for (const end of [{ from: at(doc, 0, 1) }, { to: at(doc, 0, 4) }]) {
      let caught: unknown;
      try {
        edit(doc, 'Ana', 'r', { kind: 'linkTarget', at: at(doc, 0, 1), url, ...end });
      } catch (error) {
        caught = error;
      }
      expect(caught instanceof OrderedRefusal ? caught.reason : caught).toBe('invalid-edit');
      expect(snapshot(doc)).toEqual(before);
    }
  });

  test('typing beside your own pending text grows one item; each Backspace takes back its unit', () => {
    const doc = load(SENTENCE);
    const items = () => doc.stories.get(BODY)?.tree.size ?? 0;
    edit(doc, 'Ana', 'k1', { kind: 'insert', at: at(doc, 0, 6), text: 'n' });
    const after = items();
    for (const [k, ch] of [...'ew '].entries())
      edit(doc, 'Ana', `k${k + 2}`, { kind: 'insert', at: at(doc, 0, 7 + k), text: ch });
    expect(items()).toBe(after);
    expect(text(doc, 0)).toBe('Alpha new beta gamma.');
    expect(doc.suggestions.size).toBe(1);
    for (let k = 0; k < 3; k += 1)
      edit(doc, 'Ana', `b${k}`, {
        kind: 'delete',
        from: at(doc, 0, 9 - k),
        to: at(doc, 0, 10 - k),
      });
    expect(text(doc, 0)).toBe('Alpha nbeta gamma.');
    // Each taken-back unit names the change that took it, for that change's undo
    expect(items()).toBe(after + 3);
    expect([...doc.ownEdited.keys()].sort()).toEqual(['b0:1', 'b1:1', 'b2:1']);
  });
});

describe('policies at edit time', () => {
  test('policy 1: a second deletion of struck text keeps the first', () => {
    const doc = load(SENTENCE);
    const ana = edit(doc, 'Ana', 'a', { kind: 'delete', from: at(doc, 0, 6), to: at(doc, 0, 10) })
      .value.suggestions[0];
    const ben = edit(doc, 'Ben', 'b', { kind: 'delete', from: at(doc, 0, 4), to: at(doc, 0, 12) })
      .value.suggestions[0];
    const struck = doc.stories
      .get(BODY)
      ?.tree.toArray()
      .filter((item) => item.kind === 'text' && item.dels.length > 0);
    const beta = struck?.find((item) => item.text === 'beta');
    expect(beta?.dels).toEqual([ana]);
    expect(struck?.filter((item) => item.dels.includes(ben ?? '')).map((i) => i.text)).toEqual([
      'a ',
      ' g',
    ]);
  });

  test('policy 2: a second format replaces the first record and reports it', () => {
    const doc = load(SENTENCE);
    const replaced: OrderedReplacedFormat[] = [];
    const ana = edit(doc, 'Ana', 'a', {
      kind: 'format',
      from: at(doc, 0, 0),
      to: at(doc, 0, 5),
      properties: [{ localName: 'i' }],
    }).value.suggestions[0];
    const ben = edit(
      doc,
      'Ben',
      'b',
      {
        kind: 'format',
        from: at(doc, 0, 0),
        to: at(doc, 0, 3),
        properties: [{ localName: 'u', attributes: { val: 'single' } }],
      },
      { onReplaced: (r) => replaced.push(r) }
    ).value;
    expect(ben.replaced).toEqual([{ suggestion: ana ?? '', by: ben.suggestions[0] ?? '' }]);
    expect(replaced).toEqual([...ben.replaced]);
    const out = xml(doc);
    // One record for "Alp", under Ben, with both formats shown; "ha" keeps Ana's
    expect(out.match(/<w:rPrChange [^>]*w:author="Ben"/g)?.length).toBe(1);
    expect(out.match(/<w:rPrChange [^>]*w:author="Ana"/g)?.length).toBe(1);
    expect(out).toContain('<w:i/><w:u w:val="single"/>');
  });

  test("policy 2: a press on another author's new paragraph leaves that author's own press", () => {
    const doc = load(SENTENCE);
    const press = (val: string): OrderedEdit => ({
      kind: 'paragraph',
      at: at(doc, 0, 0),
      ops: [
        { op: 'setParagraphProperties', properties: [{ localName: 'jc', attributes: { val } }] },
      ],
    });
    const ana = edit(doc, 'Ana', 'a', { kind: 'split', at: at(doc, 0, 6) }).value.suggestions[0];
    // Ana's press on her own new paragraph is part of her suggestion, with no change record
    edit(doc, 'Ana', 'a2', press('center'));
    const ben = edit(doc, 'Ben', 'b', press('right')).value;
    expect(ben.replaced).toEqual([]);
    const mark = orderedViewParagraph(doc, BODY, 0)?.mark;
    expect(
      mark?.props.map((r) => [r.suggestion, 'replaced' in r ? r.replaced : undefined])
    ).toEqual([
      [ana, undefined],
      [ben.suggestions[0], undefined],
    ]);
    // Word's own form: Ana's centring under Ben's change record
    expect(xml(doc)).toMatch(
      /<w:jc w:val="right"\/>.*<w:pPrChange w:author="Ben"[^>]*><w:pPr><w:jc w:val="center"\/>/
    );
  });

  test('policy 3: a replacement is one card, and one decision applies to it whole', () => {
    const doc = load(SENTENCE);
    const { value } = edit(doc, 'Ana', 'r', {
      kind: 'replace',
      from: at(doc, 0, 6),
      to: at(doc, 0, 10),
      text: 'delta',
    });
    expect(value.suggestions).toHaveLength(1);
    // The typing goes before the struck words, as the engine types it: "Alpha deltabeta"
    expect(text(doc, 0)).toBe('Alpha deltabeta gamma.');
    // Ben strikes inside Ana's insertion: his change joins her card
    const ben = edit(doc, 'Ben', 'b', { kind: 'delete', from: at(doc, 0, 7), to: at(doc, 0, 9) })
      .value.suggestions[0];
    const cards = orderedCards(doc);
    expect(cards.map((c) => c.suggestions)).toEqual([[value.suggestions[0], ben]]);
    expect(decideOrdered(doc, ben ?? '', 'accept')).toEqual([value.suggestions[0], ben]);
    expect(text(doc, 0)).toBe('Alpha dta gamma.');
  });

  test('policy 3: a deletion and an insertion side by side by one author are one card, after reopen too', () => {
    const doc = load(SENTENCE);
    edit(doc, 'Ana', 'd', { kind: 'delete', from: at(doc, 0, 6), to: at(doc, 0, 10) });
    edit(doc, 'Ana', 'i', { kind: 'insert', at: at(doc, 0, 10), text: 'delta' });
    expect(orderedCards(doc).map((c) => c.suggestions.length)).toEqual([2]);
    const bytes = docx(/<w:body>([\s\S]*)<\/w:body>/.exec(xml(doc))?.[1] ?? '', MATRIX_EXTRAS);
    const again = loadOrderedPackage(packageOf(bytes));
    expect(orderedCards(again).map((c) => c.suggestions.length)).toEqual([2]);
  });

  test("policy 4: Enter in a single-line control's prompt is refused, and nothing changes", () => {
    const doc = load(
      p(
        t('Name: ') +
          '<w:sdt><w:sdtPr><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent>' +
          t('Click here', '<w:rStyle w:val="PlaceholderText"/>') +
          '</w:sdtContent></w:sdt>'
      )
    );
    const before = snapshot(doc);
    expect(() => edit(doc, 'Ana', 'e', { kind: 'split', at: at(doc, 0, 9) })).toThrow(
      'single-line-control'
    );
    expect(snapshot(doc)).toEqual(before);
  });

  test("policy 4: Enter in a block control's prompt removes the prompt, then breaks", () => {
    const doc = load(
      '<w:sdt><w:sdtPr><w:showingPlcHdr/></w:sdtPr><w:sdtContent>' +
        p(t('Type the clause here.', '<w:rStyle w:val="PlaceholderText"/>')) +
        '</w:sdtContent></w:sdt>' +
        p(t('After.'))
    );
    edit(doc, 'Ana', 'e', { kind: 'split', at: at(doc, 0, 4) });
    expect([text(doc, 0), text(doc, 1), text(doc, 2)]).toEqual(['', '', 'After.']);
    expect(xml(doc)).not.toContain('showingPlcHdr');
  });
});

describe('retract', () => {
  test('an author takes back their own pending suggestion, and only theirs', () => {
    const doc = load(SENTENCE);
    const before = xml(doc);
    const ana = edit(doc, 'Ana', 'a', {
      kind: 'replace',
      from: at(doc, 0, 0),
      to: at(doc, 0, 5),
      text: 'Omega',
    }).value.suggestions[0];
    expect(() => edit(doc, 'Ben', 'x', { kind: 'retract', suggestion: ana ?? '' })).toThrow(
      'not-own-suggestion'
    );
    edit(doc, 'Ana', 'y', { kind: 'retract', suggestion: ana ?? '' });
    expect(xml(doc)).toBe(before);
    expect(orderedCards(doc)).toEqual([]);
  });

  test('a retracted format gives the replaced record back to its suggestion', () => {
    const doc = load(SENTENCE);
    const ana = edit(doc, 'Ana', 'a', {
      kind: 'format',
      from: at(doc, 0, 0),
      to: at(doc, 0, 5),
      properties: [{ localName: 'i' }],
    }).value.suggestions[0];
    const anaOnly = xml(doc);
    const ben = edit(doc, 'Ben', 'b', {
      kind: 'format',
      from: at(doc, 0, 0),
      to: at(doc, 0, 5),
      properties: [{ localName: 'strike' }],
    }).value.suggestions[0];
    edit(doc, 'Ben', 'c', { kind: 'retract', suggestion: ben ?? '' });
    expect(xml(doc)).toBe(anaOnly);
    expect(orderedCards(doc).map((c) => c.suggestions)).toEqual([[ana]]);
  });
});

describe('what the first release does not cover is refused by name', () => {
  test('an edit kind the model does not know: a new picture, a crop, a paste with formatting', () => {
    const doc = load(SENTENCE);
    const before = snapshot(doc);
    for (const kind of ['insertPicture', 'cropPicture', 'resizePicture', 'pasteRich']) {
      // As an untrusted change from the wire names it
      const change = JSON.parse(JSON.stringify({ kind, at: at(doc, 0, 1) })) as OrderedEdit;
      expect(() => edit(doc, 'Ana', 'k', change)).toThrow('not-covered');
    }
    expect(snapshot(doc)).toEqual(before);
  });

  test('a new list definition', () => {
    const doc = load(SENTENCE);
    const before = snapshot(doc);
    expect(() =>
      edit(doc, 'Ana', 'l', {
        kind: 'paragraph',
        at: at(doc, 0, 0),
        ops: [{ op: 'setListNumbering', numId: '42' }],
      })
    ).toThrow('not-covered');
    expect(snapshot(doc)).toEqual(before);
    // A list the numbering defines is an ordinary press
    edit(doc, 'Ana', 'm', {
      kind: 'paragraph',
      at: at(doc, 0, 0),
      ops: [{ op: 'setListNumbering', numId: '1' }],
    });
    expect(xml(doc)).toContain('<w:numId w:val="1"/>');
  });
});
