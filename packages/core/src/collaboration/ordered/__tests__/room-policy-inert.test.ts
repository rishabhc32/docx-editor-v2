// A party's direct edit, or an edit in a header or a footer, that changes nothing (design D9).
// Today's check compares the Word file before and after it, so it allows one that changes no
// part, and refuses one that changes a part. The room judges it once applied: the edit made no
// suggestion. Found in the multi-party acceptance run (s2-short, x68-long).

import { describe, expect, test } from 'bun:test';
import { loadOrderedPackage } from '../load/load.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import { OrderedRoom } from '../room/room.ts';
import { suggestionsOnlyPolicy } from '../room/suggestions-only.ts';
import { actor, BODY, realApply } from './room-harness.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import { orderedPlaceAt } from '../edits/places.ts';
import { docx, M, MATRIX_EXTRAS, p, packageOf, pictureRun, R, t, W } from './word-fixtures.ts';

const OLGA = actor('Olga', 'olga-1', 'organiser');
const ANA = actor('Ana', 'ana-1', 'party');

const PAGE_FIELD =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${t('1')}` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const FORM_FIELD =
  '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Name"/><w:enabled/>' +
  '<w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>' +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${t('Bob')}` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const CONTROL =
  `<w:sdt><w:sdtPr><w:id w:val="5"/><w:text/></w:sdtPr><w:sdtContent>${t('Acme')}` +
  '</w:sdtContent></w:sdt>';
const EQUATION = `<m:oMath xmlns:m="${M}"><m:r><m:t>x=1</m:t></m:r></m:oMath>`;

/**
 * A picture between two words, a control, a page field, a form field, an equation and a list
 * item, in a section with one header.
 */
const FILE = docx(
  p(t('Alpha ') + pictureRun(5) + t(' Omega')) +
    p(t('Party: ') + CONTROL) +
    p(t('USD ') + PAGE_FIELD) +
    p(t('Name: ') + FORM_FIELD) +
    p(t('Sum ') + EQUATION) +
    p(t('First item'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/></w:sectPr>',
  {
    ...MATRIX_EXTRAS,
    parts: {
      ...MATRIX_EXTRAS.parts,
      'word/header1.xml': `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>`,
    },
    relationships: `${MATRIX_EXTRAS.relationships ?? ''}<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes: `${MATRIX_EXTRAS.contentTypes ?? ''}<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`,
  }
);

const TEXT_WATERMARK = {
  kind: 'text',
  text: 'DRAFT',
  fontFamily: 'Calibri',
  fontSizePt: null,
  bold: false,
  italic: false,
  colorHex: 'C0C0C0',
  opacity: 0.5,
  diagonal: true,
  widthPt: 400,
  heightPt: 100,
} as const;

function setup() {
  let tick = 0;
  const room = new OrderedRoom({
    document: loadOrderedPackage(packageOf(FILE)),
    applyChange: realApply,
    policy: suggestionsOnlyPolicy(),
    now: () => Date.UTC(2026, 10, 23) + 1000 * tick++,
  });
  const counters = new Map<string, number>();
  const send = (who: typeof ANA, change: Record<string, unknown>) => {
    const k = (counters.get(who.namespace) ?? 0) + 1;
    counters.set(who.namespace, k);
    const result = room.submit(who, [{ v: 1, id: `${who.namespace}:${k}`, ...change }]);
    const refusal = result.refused[0];
    return refusal ? `${refusal.code}${refusal.reason ? `/${refusal.reason}` : ''}` : 'accepted';
  };
  const doc = room.document;
  /** The place before the n-th character of shown paragraph `index`, now. */
  const at = (index: number, n: number) => {
    const view = orderedViewParagraph(doc, BODY, index);
    if (!view) throw new Error(`no paragraph ${index}`);
    return orderedPlaceAt(view, n);
  };
  const markOf = (index: number) => {
    const view = orderedViewParagraph(doc, BODY, index);
    if (!view) throw new Error(`no paragraph ${index}`);
    return { space: view.mark.space, offset: view.mark.offset };
  };
  return { room, doc, send, at, markOf };
}

/** The shown atom whose element is a drawing. */
function pictureOf(doc: OrderedDocument): OrderedItem {
  for (const item of doc.stories.get(BODY)?.tree.items() ?? []) {
    if (item.kind === 'atom' && item.node?.kind === 'drawing') {
      return item;
    }
  }

  throw new Error('no picture');
}

describe('a party’s edit that changes nothing', () => {
  test('a direct object property it has already passes; a new one is refused (s2-short)', () => {
    const { doc, send } = setup();
    const picture = pictureOf(doc);
    const wrap = (value: string) => ({
      kind: 'object-property',
      part: BODY,
      mode: 'edit',
      object: { space: picture.space, offset: picture.offset },
      ops: [{ op: 'wrap', wrap: value }],
    });
    expect(send(OLGA, wrap('topAndBottom'))).toBe('accepted');
    const suggestions = doc.suggestions.size;
    expect(send(ANA, wrap('topAndBottom'))).toBe('accepted');
    expect(doc.suggestions.size).toBe(suggestions);
    expect(send(ANA, wrap('square'))).toBe('policy-direct/package');
    expect(doc.suggestions.size).toBe(suggestions);
  });

  test('a direct page setup the sections have already passes; a new one is refused', () => {
    const { doc, send } = setup();
    const margin = (twips: number) => ({
      kind: 'section-property',
      part: BODY,
      mode: 'edit',
      sections: [],
      page: { marginLeftTwips: twips },
    });
    expect(send(OLGA, margin(1800))).toBe('accepted');
    const suggestions = doc.suggestions.size;
    expect(send(ANA, margin(1800))).toBe('accepted');
    expect(doc.suggestions.size).toBe(suggestions);
    expect(send(ANA, margin(1440))).toBe('policy-direct/package');
  });

  // Each kind sent with the value it gave already: the organiser's press, then the party's same
  // one, direct and as a suggestion where the kind has one; a new value stays refused
  test('a control value, a field code, a form field, an equation or a list level it has already passes', () => {
    const { doc, send, at, markOf } = setup();
    type Make = (value: string) => Record<string, unknown>;
    const control = () => {
      const open = [...(doc.stories.get(BODY)?.tree.items() ?? [])].find((i) => i.token === 'sdt');
      if (!open) throw new Error('fixture');
      return { space: open.space, offset: open.offset };
    };
    const KINDS: [string, Make, readonly string[], string][] = [
      [
        'control-value',
        (text) => ({ kind: 'control-value', control: control(), value: { kind: 'text', text } }),
        ['edit', 'suggest'],
        'policy-direct/text',
      ],
      [
        'field-code',
        (code) => ({ kind: 'field-code', field: at(2, 4), code }),
        ['edit', 'suggest'],
        'policy-direct/package',
      ],
      [
        'form-field',
        (text) => ({
          kind: 'form-field',
          field: at(3, 6),
          op: {
            op: 'setTextFormFieldDefault',
            text,
            options: { type: 'regular', maxLength: 20, format: 'Uppercase', enabled: true },
            locale: 'en-US',
          },
        }),
        ['edit'],
        'policy-direct/package',
      ],
      [
        'equation',
        (linear) => ({ kind: 'equation', object: at(4, 4), equation: 0, linear }),
        ['edit', 'suggest'],
        'policy-direct/text',
      ],
      [
        'list-level',
        (level) => ({
          kind: 'list-level',
          numId: '1',
          level: Number(level),
          format: 'bullet',
          marks: [markOf(5)],
        }),
        ['edit', 'suggest'],
        'policy-direct/package',
      ],
    ];
    const [first, second] = [
      {
        'control-value': 'Beta',
        'field-code': 'NUMPAGES',
        'form-field': 'Name',
        equation: 'x^2',
        'list-level': '1',
      },
      {
        'control-value': 'Gamma',
        'field-code': 'PAGE',
        'form-field': 'Other',
        equation: 'y^2',
        'list-level': '2',
      },
    ] as const;
    const verdicts: Record<string, string[]> = {};
    for (const [kind, make, modes, refused] of KINDS) {
      const value = first[kind as keyof typeof first];
      expect(send(OLGA, { ...make(value), part: BODY, mode: 'edit' })).toBe('accepted');
      const suggestions = doc.suggestions.size;
      verdicts[kind] = modes.map((mode) => send(ANA, { ...make(value), part: BODY, mode }));
      expect({ kind, suggestions: doc.suggestions.size }).toEqual({ kind, suggestions });
      const other = second[kind as keyof typeof second];
      expect({ kind, other: send(ANA, { ...make(other), part: BODY, mode: 'edit' }) }).toEqual({
        kind,
        other: refused,
      });
    }

    expect(verdicts).toEqual({
      'control-value': ['accepted', 'accepted'],
      'field-code': ['accepted', 'accepted'],
      'form-field': ['accepted'],
      equation: ['accepted', 'accepted'],
      'list-level': ['accepted', 'accepted'],
    });
  });

  test('a watermark removal with no watermark passes; one with a watermark is refused (x68-long)', () => {
    const { doc, send } = setup();
    const remove = { kind: 'watermark', part: BODY, mode: 'edit', watermark: null };
    expect(send(ANA, remove)).toBe('accepted');
    expect(doc.suggestions.size).toBe(0);
    expect(send(OLGA, { ...remove, watermark: TEXT_WATERMARK })).toBe('accepted');
    expect(send(ANA, remove)).toBe('policy-furniture/furniture');
    expect(send(ANA, { ...remove, watermark: TEXT_WATERMARK })).toBe('policy-furniture/furniture');
    expect(doc.suggestions.size).toBe(1);
  });
});
