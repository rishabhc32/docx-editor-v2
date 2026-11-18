// The guard of task 9.1: every op the editor can make (`TREE_DOC_OP_KINDS`, which the engine
// keeps equal to the `TreeDocOp` union at compile time) has one entry in one group's table, and
// is translated, or planned on the list below that its group empties. A party is refused only by
// a party rule that today's check also applies. The options that change an op's meaning are
// translated or refused by name. And `design.md` lists every op with its state, so the table
// there cannot drift from the code.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import type { OoxmlNode, OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { TREE_DOC_OP_KINDS } from '../../../../store/store/tree-op-kinds.ts';
import type { TreeDocOp, TreeDocOpKind } from '../../../../store/store/tree-op-types.ts';
import { ORDERED_CHANGE_KINDS } from '../../change/kinds/registry.ts';
import { validateOrderedChange } from '../../change/validate.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import {
  ORDERED_CHANGES_WITHOUT_OP,
  ORDERED_OP_COVERAGE,
  ORDERED_OP_GROUPS,
  ORDERED_OP_OPTIONS,
  ORDERED_WORD_REFUSALS,
} from '../../client/ops/index.ts';
import { isEditBody } from '../../client/ops/types.ts';
import { OrderedClientRefusal } from '../../client/refusal.ts';
import { OrderedReplica } from '../../client/replica.ts';
import { translateOps } from '../../client/translate.ts';
import { composeOrderedPackage } from '../../compose/package.ts';
import { composeOrderedPart, orderedViewParagraph } from '../../compose/compose.ts';
import type { OrderedDocument } from '../../model/types.ts';
import type { OrderedPolicyCode, OrderedPolicyReason } from '../../room/policy.ts';
import { APPLY_FILE, BODY } from '../change-apply-fixture.ts';
import {
  bodyOf,
  docx,
  M,
  MATRIX_EXTRAS,
  MATRIX_ITEMS,
  p,
  pictureRun,
  t,
} from '../word-fixtures.ts';
import { testPerson } from '../test-persons.ts';

/**
 * The ops a later group builds. A group that translates an op takes it off its line; the final
 * state is every line empty.
 */
const STILL_PLANNED: Readonly<Record<string, readonly TreeDocOpKind[]>> = {
  '9.4': [],
  '9.5': [],
  '9.6': [],
  '9.7': [],
  '9.9': [],
  '9.10': [],
};

/** The options a later group builds, as `<op>.<option>`. The final state is empty. */
const STILL_PLANNED_OPTIONS: readonly string[] = [
  'insertFragment.object',
  'acceptRevision.equation-part',
];

/**
 * The party rules the host's earlier check applies, by the policy's code, with the reasons that
 * check gives (decisions are the organiser's alone).
 */
const PARTY_RULES: Readonly<
  Partial<Record<OrderedPolicyCode, readonly (OrderedPolicyReason | null)[]>>
> = {
  'policy-decide': [null],
  'policy-furniture': ['furniture'],
  'policy-nested': ['nested'],
  'policy-direct': ['text', 'package'],
  'policy-not-own': ['others_suggestions'],
  'policy-comment': ['resolve'],
};

const DATE = '2026-10-28T09:30:00Z';
const revision = { author: 'Ana', date: DATE };

/** An op of a kind with no fields but its name: a planned op is refused before its fields. */
function bare(kind: TreeDocOpKind): TreeDocOp {
  const op: TreeDocOp = JSON.parse(JSON.stringify({ op: kind }));
  return op;
}

/** Translate and apply as the gate does; the refusal as `<code>:<reason>`, or `applied`. */
function gate(doc: OrderedDocument, ops: readonly TreeDocOp[]): string {
  const replica = new OrderedReplica({
    doc,
    person: testPerson('Ana'),
    namespace: 'ana-1',
    now: () => 0,
  });
  const editorText = (nodeId: string): string | null => {
    for (const story of doc.stories.keys())
      for (let k = 0; ; k += 1) {
        const view = orderedViewParagraph(doc, story, k);
        if (!view) break;
        if (view.id === nodeId) return view.chars.map((c) => c.ch).join('');
      }
    return null;
  };
  try {
    const batch = translateOps(ops, {
      doc,
      person: replica.person,
      editorText,
      defaultParagraphStyle: 'Normal',
      shown: () => composeOrderedPackage(doc, base(doc)),
    });
    let into: string | undefined;
    for (const step of batch.steps)
      for (const body of step()) {
        const candidate = isEditBody(body)
          ? { v: 1, id: replica.nextId(), mode: batch.mode, ...(into ? { into } : {}), ...body }
          : { v: 1, id: replica.nextId(), ...body };
        const checked = validateOrderedChange(candidate, {
          namespace: 'ana-1',
          parts: doc.stories,
        });
        if (!checked.ok) throw new Error(`${checked.error.code} ${checked.error.path}`);
        const { outcome } = replica.applyLocal(checked.value);
        if (batch.mode === 'suggest') into ??= outcome.suggestions[0];
      }
  } catch (error) {
    if (error instanceof OrderedClientRefusal) return `${error.code}:${error.reason}`;
    throw error;
  }
  return 'applied';
}

const bases = new WeakMap<OrderedDocument, Parameters<typeof composeOrderedPackage>[1]>();
const base = (doc: OrderedDocument) => {
  const found = bases.get(doc);
  if (!found) throw new Error('no base');
  return found;
};

function open(bytes: Uint8Array): OrderedDocument {
  const opened = openOrderedWordFile(bytes, { file: 1 });
  bases.set(opened.document, opened.base);
  return opened.document;
}

/** Node ids of a part's paragraphs (nested ones too) and of its elements by local name. */
function nodes(part: OoxmlPart, localName?: string): string[] {
  const out: string[] = [];
  const visit = (node: OoxmlNode) => {
    if (node.kind === 'textValue') return;
    if (localName === undefined ? node.kind === 'paragraph' : node.localName === localName)
      out.push(node.id);
    for (const child of node.children) visit(child);
  };
  visit(part.root);
  return out;
}

/** The ops that exercise each planned option, and the document they run on. */
const OPTION_CASES: Readonly<Record<string, () => { doc: OrderedDocument; ops: TreeDocOp[] }>> = {
  // A drawing whose graphic is in a namespace the format does not hold, outside an extension
  // and a choice (a 3D model with no fallback): the change format carries none
  'insertFragment.object': () => {
    const doc = open(APPLY_FILE);
    const [first = ''] = nodes(composeOrderedPart(doc, BODY).part);
    const embedded =
      pictureRun(1).replace(
        /<pic:pic>.*<\/pic:pic>/,
        '<am3d:model3d xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"/>'
      ) + t('x');
    const pictures = composeOrderedPart(open(docx(p(embedded), MATRIX_EXTRAS)), BODY).part;
    const block =
      pictures.root.children.find((c) => c.kind !== 'textValue')?.kind === 'body'
        ? pictures.root.children.find((c) => c.kind !== 'textValue')
        : undefined;
    const paragraph =
      block && block.kind !== 'textValue'
        ? block.children.find((c) => c.kind === 'paragraph')
        : undefined;
    if (!paragraph) throw new Error('no picture paragraph');
    return {
      doc,
      ops: [{ op: 'insertFragment', paragraphId: first, offset: 0, blocks: [paragraph] }],
    };
  },
  // Word's mark inside one math run of an equation: the list keeps it as file content
  'acceptRevision.equation-part': () => {
    const stamp = { author: 'Olive', date: '2026-10-29T23:18:00Z' };
    const equation =
      `<m:oMath xmlns:m="${M}"><m:r><m:t>x=</m:t></m:r><m:r>` +
      `<w:del w:id="3" w:author="${stamp.author}" w:date="${stamp.date}"><m:t>1</m:t></w:del>` +
      '</m:r></m:oMath>';
    const doc = open(docx(p(t('Let ') + equation + t(' hold.'))));
    return { doc, ops: [{ op: 'acceptRevision', revision: { id: '3', ...stamp } }] };
  },
  'insertText.text-box': () => {
    const item = MATRIX_ITEMS.find((candidate) => candidate.key === 'textbox');
    if (!item) throw new Error('no text box item');
    const doc = open(docx(bodyOf(item), MATRIX_EXTRAS));
    const outer = new Set<string>();
    for (let k = 0; ; k += 1) {
      const view = orderedViewParagraph(doc, BODY, k);
      if (!view) break;
      outer.add(view.id);
    }
    const inner = nodes(composeOrderedPart(doc, BODY).part).find((id) => !outer.has(id)) ?? '';
    return { doc, ops: [{ op: 'insertText', paragraphId: inner, offset: 0, text: 'a', revision }] };
  },
};

/** `design.md`'s text. */
function design(): string {
  return readFileSync(
    fileURLToPath(
      new URL(
        '../../../../../../../openspec/changes/ordered-suggestion-collaboration/design.md',
        import.meta.url
      )
    ),
    'utf8'
  );
}

/** The rows of the op table in `design.md`: op name to its status cell. */
function designRows(): Map<string, string[]> {
  const source = design();
  const rows = new Map<string, string[]>();
  for (const match of source.matchAll(/^\| `([A-Za-z]+)` \|(.*)\|$/gm)) {
    const name = match[1] ?? '';
    const cells = (match[2] ?? '').split('|').map((cell) => cell.trim());
    rows.set(name, [...(rows.get(name) ?? []), cells.join('|')]);
  }
  return rows;
}

describe('every editor op (task 9.1)', () => {
  test('every op has one entry, in one group', () => {
    const kinds = [...TREE_DOC_OP_KINDS].sort();
    expect(Object.keys(ORDERED_OP_COVERAGE).sort()).toEqual(kinds);
    for (const kind of kinds) {
      const groups = ORDERED_OP_GROUPS.filter((group) => Object.hasOwn(group, kind));
      expect([kind, groups.length]).toEqual([kind, 1]);
    }
  });

  test('each op is translated, or planned on the list its group empties', () => {
    const changeKinds = new Set<string>(ORDERED_CHANGE_KINDS);
    const planned: Record<string, TreeDocOpKind[]> = {};
    for (const kind of TREE_DOC_OP_KINDS) {
      const entry = ORDERED_OP_COVERAGE[kind];
      if (entry.status === 'translated') {
        expect(typeof entry.translate).toBe('function');
        for (const change of entry.changes)
          expect([kind, changeKinds.has(change)]).toEqual([kind, true]);
        continue;
      }
      (planned[entry.group] ??= []).push(kind);
    }
    const sorted = (record: Readonly<Record<string, readonly string[]>>) =>
      Object.fromEntries(
        Object.entries(record)
          .filter(([, ops]) => ops.length > 0)
          .map(([group, ops]) => [group, [...ops].sort()])
      );
    expect(sorted(planned)).toEqual(sorted(STILL_PLANNED));
  });

  test("a party is refused only by a party rule that today's check also applies", () => {
    for (const kind of TREE_DOC_OP_KINDS) {
      const rule = ORDERED_OP_COVERAGE[kind].party;
      if (rule === null) continue;
      const reasons = PARTY_RULES[rule.code];
      expect([kind, reasons !== undefined && reasons.includes(rule.reason ?? null)]).toEqual([
        kind,
        true,
      ]);
      expect(rule.why.length).toBeGreaterThan(0);
    }
  });

  test('a planned op is refused by its name, and nothing applies', () => {
    const doc = open(APPLY_FILE);
    const before = JSON.stringify(composeOrderedPart(doc, BODY).part.root.children.length);
    for (const kind of TREE_DOC_OP_KINDS) {
      if (ORDERED_OP_COVERAGE[kind].status !== 'planned') continue;
      expect(gate(doc, [bare(kind)])).toBe(`ordered-unsupported-op:${kind}`);
    }
    expect(JSON.stringify(composeOrderedPart(doc, BODY).part.root.children.length)).toBe(before);
  });

  test("an option that changes an op's meaning is translated, or refused by name", () => {
    const planned = ORDERED_OP_OPTIONS.filter((option) => option.status === 'planned');
    expect(planned.map((option) => `${option.op}.${option.option}`).sort()).toEqual(
      [...STILL_PLANNED_OPTIONS].sort()
    );
    for (const option of planned) {
      const name = `${option.op}.${option.option}`;
      const make = OPTION_CASES[name];
      if (!make) throw new Error(`no case for ${name}`);
      const { doc, ops } = make();
      expect([name, gate(doc, ops)]).toEqual([name, option.refusal ?? 'translated']);
    }
  });

  test('design.md lists every op once, with its group and state', () => {
    const rows = designRows();
    for (const kind of TREE_DOC_OP_KINDS) {
      const entry = ORDERED_OP_COVERAGE[kind];
      const found = rows.get(kind) ?? [];
      expect([kind, found.length]).toEqual([kind, 1]);
      const cells = (found[0] ?? '').split('|');
      // | op | change kinds | group | party rule | state |
      expect([kind, cells[1], cells[3]]).toEqual([kind, entry.group, entry.status]);
    }
  });

  test('every change kind is made by an op, or by undo, redo or a host request', () => {
    const made = new Set<string>();
    for (const kind of TREE_DOC_OP_KINDS)
      for (const change of ORDERED_OP_COVERAGE[kind].changes) made.add(change);
    const text = design();
    for (const entry of ORDERED_CHANGES_WITHOUT_OP) {
      made.add(entry.kind);
      expect([entry.kind, text.includes(`| \`${entry.kind}\` |`)]).toEqual([entry.kind, true]);
      const rule = entry.party;
      if (rule !== null) expect(PARTY_RULES[rule.code]?.includes(rule.reason ?? null)).toBe(true);
    }
    for (const kind of ORDERED_CHANGE_KINDS) expect([kind, made.has(kind)]).toEqual([kind, true]);
  });

  test('an edit Word itself refuses stays refused by its name, and design.md lists it', () => {
    const text = design();
    for (const refusal of ORDERED_WORD_REFUSALS) {
      expect(Object.hasOwn(ORDERED_OP_COVERAGE, refusal.op)).toBe(true);
      const row = text
        .split('\n')
        .find(
          (line) => line.includes(`(\`${refusal.op}\`)`) && line.includes(`\`${refusal.reason}\``)
        );
      expect([refusal.reason, row !== undefined]).toEqual([refusal.reason, true]);
    }
    // Enter inside a plain text control that holds one line: Word does nothing
    const doc = open(APPLY_FILE);
    const view = orderedViewParagraph(doc, BODY, 2);
    if (!view) throw new Error('fixture');
    expect(view.chars.map((c) => c.ch).join('')).toBe('Party: Acme');
    expect(gate(doc, [{ op: 'splitParagraph', paragraphId: view.id, offset: 9 }])).toBe(
      'ordered-refused:single-line-control'
    );
  });
});
