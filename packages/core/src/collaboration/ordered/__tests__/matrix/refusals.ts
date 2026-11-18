// Content and edits the first release does not cover, refused by name (`not-covered`) with
// nothing changed. Each case names the feature list entries it stands for.

import {
  applyOrderedEdit,
  loadOrderedPackage,
  orderedPlaceAt,
  orderedViewParagraph,
  OrderedRefusal,
  type OrderedDocument,
  type OrderedEdit,
} from '../../index.ts';
import { snapshot } from '../snapshot.ts';
import { testIds } from '../test-ids.ts';
import { docx, packageOf } from '../word-fixtures.ts';
import { bodyOf, extrasOf, KINDS, type MatrixKind } from './catalogue.ts';
import { BODY } from './model.ts';
import { testAttribution } from '../test-persons.ts';

export interface RefusalCase {
  readonly id: string;
  readonly name: string;
  /** The edit, on the kind's document. */
  readonly kind: string;
  readonly edit: (doc: OrderedDocument) => OrderedEdit;
}

function kind(key: string): MatrixKind {
  const found = KINDS.find((k) => k.key === key);
  if (!found) throw new Error(`no kind ${key}`);
  return found;
}

/** The place before the n-th character of the index-th body paragraph. */
function at(doc: OrderedDocument, index: number, n: number) {
  const view = orderedViewParagraph(doc, BODY, index);
  if (!view) throw new Error(`no paragraph ${index}`);
  return orderedPlaceAt(view, n);
}

/** An edit of a kind the model does not know, as an untrusted change from the wire names it. */
const unknown = (name: string) => (doc: OrderedDocument) =>
  JSON.parse(JSON.stringify({ kind: name, at: at(doc, 1, 1) })) as OrderedEdit;

export const REFUSALS: readonly RefusalCase[] = [
  {
    id: 'paste-with-formatting',
    name: 'a paste with formatting',
    kind: 'bold',
    edit: unknown('pasteRich'),
  },
  {
    id: 'new-list-definition',
    name: 'a new list definition',
    kind: 'list',
    edit: (doc) => ({
      kind: 'paragraph',
      at: at(doc, 1, 0),
      ops: [{ op: 'setListNumbering', numId: '99' }],
    }),
  },
  {
    id: 'bound-control',
    name: 'typing in a control bound to custom XML',
    kind: 'bound',
    edit: (doc) => ({ kind: 'insert', at: at(doc, 1, 8), text: 'x' }),
  },
  {
    id: 'table-of-contents',
    name: 'deleting in a table of contents result',
    kind: 'toc',
    edit: (doc) => ({ kind: 'delete', from: at(doc, 1, 1), to: at(doc, 1, 4) }),
  },
];

/** Apply the case's edit; the refusal's reason, and whether the document stayed the same. */
export function runRefusal(refusal: RefusalCase): { reason: string; unchanged: boolean } {
  const target = kind(refusal.kind);
  const doc = loadOrderedPackage(packageOf(docx(bodyOf(target), extrasOf(target))));
  const before = snapshot(doc);
  let reason = 'applied';
  try {
    applyOrderedEdit(
      {
        doc,
        story: BODY,
        attribution: testAttribution('Ana', '2026-10-18T00:00:00Z'),
        ids: testIds(refusal.id),
      },
      refusal.edit(doc)
    );
  } catch (error) {
    reason = error instanceof OrderedRefusal ? error.reason : `error: ${String(error)}`;
  }
  return { reason, unchanged: JSON.stringify(snapshot(doc)) === JSON.stringify(before) };
}
