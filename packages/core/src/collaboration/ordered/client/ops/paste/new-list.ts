import type { OoxmlPackage } from '../../../../../store/package/ooxml-package.ts';
import type { TreeDocOp } from '../../../../../store/store/tree-op-types.ts';
import type { OrderedFragmentList } from '../../../change/fragment.ts';
import { isWml } from '../../../ooxml-nodes.ts';
import { fragmentWire } from '../../fragment-wire.ts';
import { addressOf } from '../../paragraphs.ts';
import { currentListLevel } from '../../properties.ts';
import { OrderedClientRefusal } from '../../refusal.ts';
import type { OrderedChangeBody, OrderedTranslateBatch } from '../types.ts';
import { elements, listElementsOf } from './definitions.ts';

type SetListNumberingOp = Extract<TreeDocOp, { readonly op: 'setListNumbering' }>;

/** A list definition as a change carries it, read from a package by its number. */
export function listDefinitionIn(pkg: OoxmlPackage, numId: string): OrderedFragmentList {
  const list = listElementsOf(pkg, numId);
  if (!list) {
    throw new OrderedClientRefusal('ordered-refused', 'not-covered', `no list ${numId}`);
  }

  const { num, abstract } = list;
  const none = (what: string) => (): string => {
    throw new OrderedClientRefusal('ordered-refused', 'not-covered', `a list's ${what}`);
  };
  const refs = {
    numId: none('list'),
    relationship: none('link'),
    media: none('picture'),
    note: none('note'),
  };
  const [tree] = fragmentWire([abstract], refs);
  if (!tree) throw new OrderedClientRefusal('ordered-refused', 'not-covered', 'a list');
  const overrides = fragmentWire(
    elements(num).filter((n) => isWml(n, 'lvlOverride')),
    refs
  );
  return overrides.length > 0 ? { abstract: tree, overrides } : { abstract: tree };
}

/**
 * The `new-list` change of a batch's first op that names a new list; nothing for the others. The
 * editor writes the new definition before the batch: the change carries it, read from the package
 * the editor gave the gate, and every paragraph of the batch that names the number.
 */
export function newListChanges(
  op: SetListNumberingOp,
  batch: OrderedTranslateBatch
): OrderedChangeBody[] {
  const numId = op.numId;
  if (numId === null) return [];
  const same = batch.ops.filter(
    (other): other is SetListNumberingOp => other.op === 'setListNumbering' && other.numId === numId
  );
  if (same[0] !== op) return [];
  const pkg = batch.gated();
  if (pkg === null)
    throw new OrderedClientRefusal('ordered-refused', 'not-covered', 'no package with the list');
  const paragraphs = same.map((other) => {
    const p = batch.paragraph(other.paragraphId);
    const level = other.level ?? currentListLevel(batch.doc, p.story, p.mark) ?? 0;
    return { p, entry: { mark: addressOf(p.mark), level } };
  });
  const part = paragraphs[0]?.p.story.name;
  if (part === undefined || paragraphs.some(({ p }) => p.story.name !== part))
    throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'a list over two stories');
  return [
    {
      kind: 'new-list',
      part,
      list: listDefinitionIn(pkg, numId),
      paragraphs: paragraphs.map(({ entry }) => entry),
    },
  ];
}
