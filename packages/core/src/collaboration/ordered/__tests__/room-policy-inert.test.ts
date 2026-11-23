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
import { docx, MATRIX_EXTRAS, p, packageOf, pictureRun, R, t, W } from './word-fixtures.ts';

const OLGA = actor('Olga', 'olga-1', 'organiser');
const ANA = actor('Ana', 'ana-1', 'party');

/** A picture between two words, in a section with one header. */
const FILE = docx(
  p(t('Alpha ') + pictureRun(5) + t(' Omega')) +
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
  return { room, doc: room.document, send };
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
