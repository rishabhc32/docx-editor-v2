// `setTableProperties` places the table: alignment (`w:tblPr/w:jc`) and preferred width
// (`w:tblPr/w:tblW`).
//
// Each value writes its leaf in `CT_TblPr` order, keeps everything else in the container, and
// `null` removes it. Tracked, the old `w:tblPr` goes in a `w:tblPrChange`. An alignment also
// records every cell, which is the form Word writes: Word reads a lone `w:tblPrChange` that only
// adds a `w:jc` as unused history, so on its own the proposal would offer nothing to review.

import { describe, expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPart,
  serializeOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { diffSemanticDigests, semanticDigest } from '../package/ooxml-digest.ts';
import { planRevisionBatch, revisionItemsOf } from '../index.ts';
import { reviewItemKey } from '../store/review-items.ts';
import { applyTreeOp, validateTreeOp } from '../store/tree-ops.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';
import { TreeDocumentStore } from '../store/tree-store.ts';

const W = WML_NAMESPACE_URI;
const BEA = { author: 'Bea', date: '2026-10-01T00:00:00Z' };
const NAME = '/word/document.xml';
const CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

function load(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:x="urn:x"><w:body>${body}</w:body></w:document>`,
    { name: NAME, contentType: CONTENT_TYPE }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function tableId(part: OoxmlPart): string {
  const visit = (node: OoxmlNode): OoxmlElement | undefined => {
    if (node.kind === 'textValue') return undefined;
    if (node.kind === 'table') return node;
    for (const child of node.children) {
      const found = visit(child);
      if (found) return found;
    }
    return undefined;
  };
  return visit(part.root)!.id;
}

const CELL = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const TABLE = (properties: string) =>
  `<w:tbl><w:tblPr>${properties}</w:tblPr>` +
  '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
  `<w:tr>${CELL('a1')}${CELL('a2')}</w:tr><w:tr>${CELL('b1')}${CELL('b2')}</w:tr></w:tbl><w:p/>`;
const WIDTH = '<w:tblW w:w="4800" w:type="dxa"/>';
const PLAIN = TABLE(WIDTH);

function apply(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`${op.op} refused: ${result.reason}`);
  return result.part;
}
const place = (
  part: OoxmlPart,
  fields: Partial<Extract<TreeDocOp, { op: 'setTableProperties' }>>
): TreeDocOp => ({ op: 'setTableProperties', tableId: tableId(part), ...fields });
const tblPr = (part: OoxmlPart): string =>
  serializeOoxmlPart(part).match(/<w:tblPr(\/>|>.*?<\/w:tblPr>)(?=<w:tblGrid>)/)![0];
const fingerprint = (part: OoxmlPart): string => canonicalOoxmlFingerprint(part.root);
const tableXml = (part: OoxmlPart): string =>
  serializeOoxmlPart(part).match(/<w:tbl>.*<\/w:tbl>/)![0];

describe('untracked placement', () => {
  test('alignment lands in CT_TblPr order, and the rest of the container stays', () => {
    const part = load(
      TABLE(
        '<w:tblStyle w:val="Grid"/><w:bidiVisual/>' +
          WIDTH +
          '<w:tblInd w:w="0" w:type="dxa"/><w:tblLook w:val="04A0"/>'
      )
    );
    expect(tblPr(apply(part, place(part, { justification: 'center' })))).toBe(
      '<w:tblPr><w:tblStyle w:val="Grid"/><w:bidiVisual/><w:tblW w:type="dxa" w:w="4800"/>' +
        '<w:jc w:val="center"/><w:tblInd w:type="dxa" w:w="0"/><w:tblLook w:val="04A0"/></w:tblPr>'
    );
  });

  test('a new value replaces only the attributes it names', () => {
    const part = load(TABLE(`${WIDTH}<w:jc w:val="start" x:keep="1"/>`));
    expect(tblPr(apply(part, place(part, { justification: 'right' })))).toBe(
      '<w:tblPr><w:tblW w:type="dxa" w:w="4800"/><w:jc w:val="right" x:keep="1"/></w:tblPr>'
    );
    expect(tblPr(apply(part, place(part, { preferredWidth: { type: 'pct', value: 2500 } })))).toBe(
      '<w:tblPr><w:tblW w:type="pct" w:w="2500"/><w:jc w:val="start" x:keep="1"/></w:tblPr>'
    );
  });

  test('null removes the leaf, and a table without one is left as it was', () => {
    const part = load(TABLE(`${WIDTH}<w:jc w:val="center"/>`));
    const cleared = apply(part, place(part, { justification: null, preferredWidth: null }));
    expect(tblPr(cleared)).toBe('<w:tblPr/>');
    const bare = load(TABLE(''));
    expect(tableXml(apply(bare, place(bare, { justification: null })))).toBe(tableXml(bare));
  });

  test('an auto width is written as Word writes it', () => {
    const part = load(PLAIN);
    expect(tblPr(apply(part, place(part, { preferredWidth: { type: 'auto', value: 0 } })))).toBe(
      '<w:tblPr><w:tblW w:type="auto" w:w="0"/></w:tblPr>'
    );
  });

  test('malformed or out-of-range placement is refused', () => {
    const part = load(PLAIN);
    const sound = { justification: 'left', preferredWidth: { type: 'dxa', value: 2400 } } as const;
    expect(validateTreeOp(part, place(part, sound))).toBeNull();
    const refused: unknown[] = [
      { justification: 'both' },
      { justification: 'start' },
      { preferredWidth: { type: 'dxa', value: 0 } },
      { preferredWidth: { type: 'dxa', value: 31_681 } },
      { preferredWidth: { type: 'dxa', value: 14.5 } },
      { preferredWidth: { type: 'pct', value: 5001 } },
      { preferredWidth: { type: 'auto', value: 10 } },
      { preferredWidth: { type: 'nil', value: 0 } },
      { preferredWidth: { type: 'dxa', value: 1440, extra: true } },
      { preferredWidth: 1440 },
      { preferredWidth: { type: 'dxa', value: 2400 }, columnWidthsTwips: [2400, 2400] },
    ];
    for (const fields of refused) {
      expect(validateTreeOp(part, place(part, fields as never))).toBe('invalidArgs');
    }
  });

  // An untracked op meets no tracked-table code, so the store's own check of the op refuses the
  // values; a tracked op meets that same check through its untracked form.
  test('the store refuses malformed placement, untracked and tracked alike', () => {
    const malformed = [{ justification: 'both' }, { preferredWidth: { type: 'dxa', value: 0 } }];
    for (const revision of [{}, { revision: BEA }]) {
      for (const fields of malformed) {
        const store = new TreeDocumentStore(load(PLAIN));
        const op = place(store.part, { ...fields, ...revision } as never);
        const result = store.transact((tx) => {
          tx.apply(op);
        });
        expect(result).toEqual({ ok: false, reason: 'invalidArgs' });
        expect(store.revision).toBe(0);
      }
    }
  });
});

const GESTURES: readonly (readonly [string, (part: OoxmlPart) => TreeDocOp])[] = [
  ['alignment', (part) => place(part, { justification: 'center', revision: BEA })],
  ['alignment cleared', (part) => place(part, { justification: null, revision: BEA })],
  [
    'preferred width',
    (part) => place(part, { preferredWidth: { type: 'pct', value: 2500 }, revision: BEA }),
  ],
];
const BODIES: Readonly<Record<string, string>> = {
  alignment: PLAIN,
  'alignment cleared': TABLE(`${WIDTH}<w:jc w:val="right"/>`),
  'preferred width': PLAIN,
};

describe.each(GESTURES)('tracked %s', (name, gesture) => {
  test('is one review card, which Accept and Reject answer whole', () => {
    const part = load(BODIES[name]!);
    const next = apply(part, gesture(part));
    expect(serializeOoxmlPart(next)).toContain(
      `<w:tblPrChange w:author="Bea" w:date="2026-10-01T00:00:00Z" w:id="0">${tblPr(part)}`
    );
    const items = revisionItemsOf(next);
    expect(items.map((item) => item.revisionKind)).toEqual(['format']);
    const resolve = (action: 'accept' | 'reject'): OoxmlPart => {
      const plan = planRevisionBatch(next, action, [reviewItemKey(items[0]!)]);
      expect(plan.result.skipped).toEqual([]);
      return plan.ops.reduce(apply, next);
    };
    const untracked = { ...gesture(part), revision: undefined };
    expect(fingerprint(resolve('accept'))).toBe(fingerprint(apply(part, untracked)));
    expect(tableXml(resolve('reject'))).toBe(tableXml(part));
  });

  test('a save keeps the proposal, and reopening reads the same tree', () => {
    const part = load(BODIES[name]!);
    const next = apply(part, gesture(part));
    const reopened = readOoxmlPart(serializeOoxmlPart(next), {
      name: NAME,
      contentType: CONTENT_TYPE,
    });
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(revisionItemsOf(reopened.part).map((item) => item.revisionKind)).toEqual(['format']);
    expect(fingerprint(reopened.part)).toBe(fingerprint(next));
    expect(diffSemanticDigests(semanticDigest([next]), semanticDigest([reopened.part]))).toEqual(
      []
    );
  });
});

describe('tracked alignment', () => {
  test('records every cell under the one revision, the form Word writes', () => {
    const part = load(PLAIN);
    const xml = serializeOoxmlPart(
      apply(part, place(part, { justification: 'right', revision: BEA }))
    );
    const snapshot =
      '<w:tcPr><w:tcW w:type="dxa" w:w="2400"/><w:tcPrChange w:author="Bea" ' +
      'w:date="2026-10-01T00:00:00Z" w:id="0"><w:tcPr><w:tcW w:type="dxa" w:w="2400"/></w:tcPr>' +
      '</w:tcPrChange></w:tcPr>';
    expect(xml.split(snapshot)).toHaveLength(5);
    expect(xml).not.toContain('<w:trPrChange');
  });

  test('a width change alone records no cell', () => {
    const part = load(PLAIN);
    const op = place(part, { preferredWidth: { type: 'pct', value: 2500 }, revision: BEA });
    expect(serializeOoxmlPart(apply(part, op))).not.toContain('<w:tcPrChange');
  });

  test('an alignment the table already has records nothing', () => {
    const part = load(TABLE(`${WIDTH}<w:jc w:val="center"/>`));
    const next = apply(part, place(part, { justification: 'center', revision: BEA }));
    expect(serializeOoxmlPart(next)).not.toMatch(/Change\b/);
    expect(revisionItemsOf(next)).toHaveLength(0);
  });

  test('a table the author proposed whole records nothing', () => {
    const own = `<w:trPr><w:ins w:id="5" w:author="Bea" w:date="2026-10-01T00:00:00Z"/></w:trPr>`;
    const part = load(PLAIN.replaceAll('<w:tr>', `<w:tr>${own}`));
    const next = apply(part, place(part, { justification: 'center', revision: BEA }));
    expect(serializeOoxmlPart(next)).not.toMatch(/PrChange/);
  });
});
