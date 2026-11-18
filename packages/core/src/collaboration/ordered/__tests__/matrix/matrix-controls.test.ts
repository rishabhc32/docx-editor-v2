// The control edits of group 9.6 against the engine, over every matrix kind that holds a
// content control (inline, block, prompts of every type, a gallery, a group, a bound control, a
// repeating section). For each kind and each of the organiser's direct control edits:
//
//   V  the list's story equals the engine's after the edit;
//   U  the edit taken back (`revert`) gives the story before it;
//   W  the saved Word file reopens as the list's story;
//   D  after a party typed in the control first, each decision on the party's words gives the
//      engine's result for the same decision.
//
// A kind whose control the engine refuses the edit for is refused by the list by the same name
// (E alone is a difference). No difference is left.

import { describe, expect, test } from 'bun:test';
import type { OoxmlNode, OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import {
  applyOrderedChange,
  applyOrderedEdit,
  composeOrderedPart,
  loadOrderedPackage,
  orderedIdMinter,
  OrderedRefusal,
  setOrderedStatus,
  takeBackOrdered,
  withOrderedUndoRecord,
  type OrderedChange,
  type OrderedDocument,
} from '../../index.ts';
import { docx } from '../word-fixtures.ts';
import { ANA, type Action } from './actions.ts';
import { bodyOf, extrasOf, KINDS, paragraphOf, type MatrixKind } from './catalogue.ts';
import { agrees, firstDifference, grade, normalizedBody, semanticBody } from './compare.ts';
import { DATE, engineDecide, mount, performEngine, storyPart, triplesOf } from './engine.ts';
import { BODY, editFor } from './model.ts';
import { reopened, resolverOf, saved } from './word.ts';
import { testAttribution, testPerson } from '../test-persons.ts';

/** Kinds in the body that hold a content control. */
const CONTROL_KINDS: readonly MatrixKind[] = KINDS.filter(
  (kind) => kind.story === undefined && bodyOf(kind).includes('<w:sdt')
);

type ControlEdit = {
  readonly name: string;
  readonly op: (controlId: string) => TreeDocOp;
  readonly change: (control: { space: string; offset: number }) => Record<string, unknown>;
};

const EDITS: readonly ControlEdit[] = [
  {
    name: 'remove, keeping its content',
    op: (controlId) => ({ op: 'removeContentControl', controlId, keepContent: true }),
    change: (control) => ({ kind: 'remove-control', control }),
  },
  {
    name: 'remove with its content',
    op: (controlId) => ({ op: 'removeContentControl', controlId, keepContent: false }),
    change: (control) => ({ kind: 'delete-control', control }),
  },
  {
    name: 'a new tag and title',
    op: (controlId) => ({ op: 'setContentControlProperties', controlId, tag: 'mx', alias: 'MX' }),
    change: (control) => ({ kind: 'control-property', control, tag: 'mx', alias: 'MX' }),
  },
];

/** The first content control in the paragraph a kind's edits go to, or anywhere in the body. */
function firstControl(part: OoxmlPart): string | null {
  const visit = (node: OoxmlNode): string | null => {
    if (node.kind === 'textValue') return null;
    if (node.kind === 'contentControl') return node.id;
    for (const child of node.children) {
      const found = visit(child);
      if (found) return found;
    }
    return null;
  };
  return visit(part.root);
}

const label = (kind: MatrixKind, edit: ControlEdit, party: boolean) =>
  `${kind.key} | ${party ? 'after a party typed in it, ' : ''}${edit.name}`;

/** The list's direct edit, refused by name or applied; the suggestion it made. */
function listEdit(doc: OrderedDocument, change: Record<string, unknown>): string | null {
  const full = { v: 1, id: 'olive-1:1', part: BODY, mode: 'edit', ...change } as OrderedChange;
  try {
    const { value } = withOrderedUndoRecord(doc, () =>
      applyOrderedChange(doc, full, {
        person: testPerson('Olive'),
        date: DATE,
        ids: orderedIdMinter(full.id),
      })
    );
    return value.suggestions[0] ?? null;
  } catch (error) {
    if (error instanceof OrderedRefusal) return `refused ${error.reason}`;
    throw error;
  }
}

function run(kind: MatrixKind, edit: ControlEdit, party: boolean): string {
  const { surface, close } = mount(docx(bodyOf(kind), extrasOf(kind)), 'Ana');
  try {
    const p = paragraphOf(kind);
    const doc = loadOrderedPackage(surface.session.currentPackage());
    const anaTriples: ReturnType<typeof triplesOf> = [];
    if (party) {
      if (!kind.inside) return 'N';
      const action: Action = { ...ANA, p, kind: 'insert', at: kind.inside, text: 'X' };
      const edit = editFor(doc, BODY, action);
      if ('change' in edit) throw new Error('typing is an edit, not a change of a group');
      try {
        applyOrderedEdit(
          {
            doc,
            story: BODY,
            attribution: testAttribution('Ana', DATE),
            ids: orderedIdMinter('ana-1:1'),
          },
          edit
        );
      } catch (error) {
        if (error instanceof OrderedRefusal) return 'N';
        throw error;
      }
      performEngine(surface, undefined, action);
      anaTriples.push(...triplesOf(storyPart(surface, undefined)));
    }
    const before = storyPart(surface, undefined);
    const controlId = firstControl(before);
    if (!controlId) return 'N';
    surface.setEditingMode('edit');
    surface.setAuthor('Olive');
    const applied = surface.session.applyTreeOps([edit.op(controlId)]);
    const engineRefused = applied.rejected === true;
    // The list's address of the same control: a loaded control's open item
    const listControl = { space: `t:0:${controlId}`, offset: 0 };
    const made = listEdit(doc, edit.change(listControl));
    const listRefused = made !== null && made.startsWith('refused');
    if (engineRefused || listRefused)
      return engineRefused && listRefused ? 'R' : engineRefused ? `E` : `S:${made}`;
    const pkg = surface.session.currentPackage();
    const resolve = resolverOf(pkg, BODY, doc);
    const xml = (part: OoxmlPart) => normalizedBody(part, resolve);
    let codes = '';
    const after = storyPart(surface, undefined);
    if (!agrees(grade(xml(composeOrderedPart(doc, BODY).part), xml(after)))) codes += 'V';
    // Taken back
    if (made !== null) {
      const { record } = withOrderedUndoRecord(doc, () => setOrderedStatus(doc, made, 'rejected'));
      if (!agrees(grade(xml(composeOrderedPart(doc, BODY).part), xml(before)))) {
        codes += 'U';
        console.warn(
          label(kind, edit, party),
          'revert',
          firstDifference(
            semanticBody(xml(composeOrderedPart(doc, BODY).part)),
            semanticBody(xml(before))
          )
        );
      }
      takeBackOrdered(doc, record);
    }
    // Saved and reopened
    const file = saved(pkg, doc);
    const again = reopened(file);
    const filePart = file.parts.get(BODY);
    if (!filePart) throw new Error('no body');
    const fileResolve = resolverOf(file, BODY);
    if (
      !agrees(
        grade(
          normalizedBody(composeOrderedPart(again, BODY).part, fileResolve),
          normalizedBody(filePart, fileResolve)
        )
      )
    )
      codes += 'W';
    // The party's words decided, on the list and on the engine
    if (party) {
      const ana = [...doc.suggestions.values()].find((s) => s.attribution.author === 'Ana');
      if (ana)
        for (const status of ['accepted', 'rejected'] as const) {
          const { record } = withOrderedUndoRecord(doc, () =>
            setOrderedStatus(doc, ana.id, status)
          );
          const list = xml(composeOrderedPart(doc, BODY).part);
          const engine = xml(
            engineDecide(after, anaTriples, status === 'accepted' ? 'accept' : 'reject', pkg)
          );
          if (!agrees(grade(list, engine))) {
            codes += 'D';
            console.warn(
              label(kind, edit, party),
              status,
              firstDifference(semanticBody(list), semanticBody(engine))
            );
          }
          takeBackOrdered(doc, record);
        }
    }
    if (codes.includes('V'))
      console.warn(
        label(kind, edit, party),
        firstDifference(
          semanticBody(xml(composeOrderedPart(doc, BODY).part)),
          semanticBody(xml(after))
        )
      );
    return codes;
  } finally {
    close();
  }
}

describe('control edits against the engine (task 9.6)', () => {
  test('the matrix holds controls of many kinds', () => {
    expect(CONTROL_KINDS.length).toBeGreaterThan(8);
  });
  for (const kind of CONTROL_KINDS)
    for (const edit of EDITS)
      for (const party of [false, true]) {
        const name = label(kind, edit, party);
        test(name, () => {
          const codes = run(kind, edit, party);
          // N: the kind has no place to type in it; R: both refuse, by the engine's name
          expect([name, codes]).toEqual([name, codes === 'N' || codes === 'R' ? codes : '']);
        });
      }
});
