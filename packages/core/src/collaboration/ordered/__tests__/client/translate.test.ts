// `TreeDocOp` translation (task 6.1) for the ops the editor tests do not reach through a command:
// each op is translated against the list, applied as the client applies its own changes, and
// the paragraph or table it changed is compared with the engine's own op on the same part. An
// op the format cannot say is refused by name, and nothing applies.

import { describe, expect, test } from 'bun:test';
import { findNode } from '../../../../store/package/ooxml-edit.ts';
import {
  canonicalOoxmlFingerprint,
  serializeOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
} from '../../../../store/package/ooxml-tree.ts';
import { applyTreeOp } from '../../../../store/store/tree-op-apply.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import { validateOrderedChange } from '../../change/validate.ts';
import { OrderedClientRefusal } from '../../client/refusal.ts';
import { OrderedReplica } from '../../client/replica.ts';
import { translateOps } from '../../client/translate.ts';
import { composeOrderedPart, orderedViewParagraph } from '../../compose/compose.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { APPLY_FILE, BODY } from '../change-apply-fixture.ts';
import { testPerson } from '../test-persons.ts';

const DATE = '2026-10-28T09:30:00Z';
const NOW = Date.UTC(2026, 9, 28, 9, 30, 0);

function setup() {
  const { document: doc } = openOrderedWordFile(APPLY_FILE, { file: 1 });
  const replica = new OrderedReplica({
    doc,
    person: testPerson('Ana'),
    namespace: 'ana-1',
    now: () => NOW,
  });
  return { doc, replica };
}

const part = (doc: OrderedDocument): OoxmlPart => composeOrderedPart(doc, BODY).part;

/** The editor's text of a paragraph: the list's own, as after a publish. */
function textOf(doc: OrderedDocument) {
  return (nodeId: string): string | null => {
    const node = findNode(part(doc), nodeId);
    if (!node) return null;
    for (let k = 0; ; k += 1) {
      const view = orderedViewParagraph(doc, BODY, k);
      if (!view) return null;
      if (view.id === nodeId) return view.chars.map((c) => c.ch).join('');
    }
  };
}

/** Translate and apply as the gate does; the changes it made. */
function gate(doc: OrderedDocument, replica: OrderedReplica, ops: readonly TreeDocOp[]) {
  const batch = translateOps(ops, {
    doc,
    person: replica.person,
    editorText: textOf(doc),
    defaultParagraphStyle: 'Normal',
  });
  const made: string[] = [];
  let into: string | undefined;
  batch.steps.forEach((step, k) => {
    const mode = batch.modes[k] ?? batch.mode;
    for (const body of step()) {
      const checked = validateOrderedChange(
        {
          v: 1,
          id: replica.nextId(),
          mode,
          ...(into === undefined || mode !== 'suggest' ? {} : { into }),
          ...body,
        },
        { namespace: 'ana-1', parts: doc.stories }
      );
      if (!checked.ok) throw new Error(`${checked.error.code} ${checked.error.path}`);
      const { outcome } = replica.applyLocal(checked.value);
      if (mode === 'suggest') into ??= outcome.suggestions[0];
      made.push(checked.value.kind);
    }
  });
  return { made, mode: batch.mode };
}

/** The engine's own op on the same part, before the list's change. */
function engine(before: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(before, op);
  if (!result.ok) throw new Error(`the engine refused: ${result.reason}`);
  return result.part;
}

/** A node's fingerprint, without ids or revision ids and dates. */
function shape(root: OoxmlPart, id: string): string {
  const node = findNode(root, id);
  if (!node) return '<none>';
  const fp = canonicalOoxmlFingerprint({ ...root, root: node as never });
  // Revision ids and dates are each copy's own: the list mints ids, the engine counts them
  return fp.replace(
    /\["http:\/\/schemas\.openxmlformats\.org\/wordprocessingml\/2006\/main","(id|date)","[^"]*"\],?/g,
    ''
  );
}

const ids = (doc: OrderedDocument) => {
  const p = part(doc);
  const paragraphs: string[] = [];
  const tables: string[] = [];
  const rows: string[] = [];
  const cells: string[] = [];
  const visit = (node: OoxmlNode) => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'paragraph') paragraphs.push(node.id);
    if (node.kind === 'table') tables.push(node.id);
    if (node.kind === 'tableRow') rows.push(node.id);
    if (node.kind === 'tableCell') cells.push(node.id);
    for (const child of node.children) visit(child);
  };
  visit(p.root);
  return { paragraphs, tables, rows, cells };
};

const revision = { author: 'Ana', date: DATE };

describe('op translation', () => {
  test('tab stops, a list number and a list level, against the engine', () => {
    const { doc, replica } = setup();
    const [, , , item] = ids(doc).paragraphs;
    if (!item) throw new Error('fixture');
    const cases: TreeDocOp[] = [
      {
        op: 'setParagraphTabStops',
        paragraphId: item,
        stops: [{ positionTwips: 720, alignment: 'center', leader: 'dot' }],
        revision,
      },
      { op: 'setListLevel', paragraphId: item, level: 1, revision },
      { op: 'setListNumbering', paragraphId: item, numId: null, revision },
    ];
    for (const op of cases) {
      const before = part(doc);
      const expected = engine(before, op);
      expect(gate(doc, replica, [op]).made).toEqual(['paragraph-property']);
      expect({ op: op.op, shape: shape(part(doc), item) }).toEqual({
        op: op.op,
        shape: shape(expected, item),
      });
    }
  });

  test('table, row and cell properties name their containers, against the engine', () => {
    const cases = (table: string, row: string, cellIds: readonly string[]): TreeDocOp[] => [
      { op: 'setTableProperties', tableId: table, justification: 'center', revision },
      { op: 'setTableRowHeight', tableId: table, rowId: row, heightTwips: 600, revision },
      {
        op: 'setTableCellFill',
        tableId: table,
        cellIds,
        color: { kind: 'hex', value: 'FFEE00' },
        revision,
      },
      {
        op: 'setTableCellVerticalAlignment',
        tableId: table,
        cellIds,
        alignment: 'center',
        revision,
      },
    ];
    for (let k = 0; k < 4; k += 1) {
      // A second property change on one table is a documented limit (task 4.4): one each
      const { doc, replica } = setup();
      const { tables, rows, cells } = ids(doc);
      const [table] = tables;
      const [row] = rows;
      if (!table || !row || cells.length < 2) throw new Error('fixture');
      const op = cases(table, row, cells.slice(0, 2))[k];
      if (!op) throw new Error('no case');
      const expected = engine(part(doc), op);
      expect(gate(doc, replica, [op]).made).toEqual(['table-property']);
      expect({ op: op.op, shape: shape(part(doc), table) }).toEqual({
        op: op.op,
        shape: shape(expected, table),
      });
    }
  });

  test('a direct join, a struck mark, and Enter at a paragraph end with the next style', () => {
    const { doc, replica } = setup();
    const [first, second] = ids(doc).paragraphs;
    if (!first || !second) throw new Error('fixture');
    expect(
      gate(doc, replica, [
        { op: 'setParagraphMarkRevision', paragraphId: first, kind: 'del', revision },
      ])
    ).toEqual({
      made: ['join'],
      mode: 'suggest',
    });
    const last = ids(doc).paragraphs.at(-1);
    if (!last) throw new Error('fixture');
    const end = (textOf(doc)(last) ?? '').length;
    const enter = gate(doc, replica, [
      { op: 'splitParagraph', paragraphId: last, offset: end, tailStyleId: null },
      { op: 'setParagraphMarkRevision', paragraphId: last, kind: 'ins', revision },
    ]);
    expect(enter.made).toEqual(['split', 'paragraph-property']);
    // The tail keeps the paragraph's own mark: it takes the default style
    expect(shape(part(doc), last)).toContain('Normal');
    const direct = setup();
    const [a, b] = ids(direct.doc).paragraphs;
    if (!a || !b) throw new Error('fixture');
    expect(
      gate(direct.doc, direct.replica, [{ op: 'joinParagraphs', firstId: a, secondId: b }])
    ).toEqual({
      made: ['join'],
      mode: 'edit',
    });
  });

  test('what the format cannot say is refused by name, and nothing applies', () => {
    const { doc, replica } = setup();
    const [first] = ids(doc).paragraphs;
    const [table] = ids(doc).tables;
    if (!first || !table) throw new Error('fixture');
    const before = canonicalOoxmlFingerprint(part(doc));
    const refused = (ops: readonly TreeDocOp[]) => {
      try {
        gate(doc, replica, ops);
      } catch (error) {
        if (error instanceof OrderedClientRefusal) return `${error.code}:${error.reason}`;
        throw error;
      }
      return 'applied';
    };
    // A cell the list does not hold: refused by name before anything applies
    expect(
      refused([{ op: 'insertTableColumn', tableId: table, where: 'left', referenceCellId: 'x' }])
    ).toBe('ordered-refused:out-of-step');
    // An op that leaves out a property with no off value takes it off: a size the text authors
    const sized = setup();
    const [p0] = ids(sized.doc).paragraphs;
    if (!p0) throw new Error('fixture');
    const size = (properties: readonly { localName: string; attributes?: { val: string } }[]) =>
      gate(sized.doc, sized.replica, [
        { op: 'setRunProperties', paragraphId: p0, start: 0, end: 3, properties, revision },
      ]);
    size([{ localName: 'sz', attributes: { val: '28' } }]);
    expect(serializeOoxmlPart(part(sized.doc))).toContain('<w:sz w:val="28"/>');
    size([]);
    expect(serializeOoxmlPart(part(sized.doc))).not.toContain('<w:sz');
    // A batch of tracked and direct ops makes each change in its own op's mode
    const mixed = translateOps(
      [
        { op: 'insertText', paragraphId: first, offset: 0, text: 'a', revision },
        { op: 'insertText', paragraphId: first, offset: 0, text: 'b' },
      ],
      { doc, person: testPerson('Ana'), editorText: textOf(doc), defaultParagraphStyle: 'Normal' }
    );
    expect(mixed.modes).toEqual(['suggest', 'edit']);
    // An armed format of the typed text, where formatting is not tracked, follows the typing
    const armed = translateOps(
      [
        { op: 'insertText', paragraphId: first, offset: 0, text: 'ab', revision },
        {
          op: 'setRunProperties',
          paragraphId: first,
          start: 0,
          end: 2,
          properties: [{ localName: 'b' }],
        },
      ],
      { doc, person: testPerson('Ana'), editorText: textOf(doc), defaultParagraphStyle: 'Normal' }
    );
    expect(armed.modes).toEqual(['suggest', 'suggest']);
    expect(canonicalOoxmlFingerprint(part(doc))).toBe(before);
  });
});
