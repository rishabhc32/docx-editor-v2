import type { OoxmlNode } from '../package/ooxml-tree.ts';
import { WML_NAMESPACE_URI } from '../package/ooxml-tree.ts';

/** Property snapshots are historical copies, never live revision subtrees. */
export const REVISION_PROPERTY_CONTAINERS: ReadonlyMap<string, string> = new Map([
  ['rPrChange', 'rPr'],
  ['pPrChange', 'pPr'],
  ['trPrChange', 'trPr'],
  ['tcPrChange', 'tcPr'],
  ['tblPrChange', 'tblPr'],
  ['tblPrExChange', 'tblPrEx'],
  ['tblGridChange', 'tblGrid'],
  ['sectPrChange', 'sectPr'],
]);

export function preservedRevisionProperty(wrapperName: string, child: OoxmlNode): boolean {
  if (child.kind === 'textValue' || child.namespaceUri !== WML_NAMESPACE_URI) return false;
  const name = child.localName;
  switch (wrapperName) {
    case 'pPrChange':
      return name === 'rPr' || name === 'sectPr';
    case 'rPrChange':
      return ['ins', 'del', 'moveFrom', 'moveTo'].includes(name);
    case 'trPrChange':
      return name === 'ins' || name === 'del';
    case 'tcPrChange':
      return name === 'cellIns' || name === 'cellDel' || name === 'cellMerge';
    case 'sectPrChange':
      return name === 'headerReference' || name === 'footerReference';
    default:
      return false;
  }
}

/** Reject ambiguous duplicate live records before choosing a historical state. */
export function validRevisionPropertyRecord(node: OoxmlNode, parent: OoxmlNode | null): boolean {
  if (node.kind === 'textValue' || !parent || parent.kind === 'textValue') return false;
  const container = REVISION_PROPERTY_CONTAINERS.get(node.localName);
  if (parent.namespaceUri !== WML_NAMESPACE_URI || parent.localName !== container) return false;
  const records = parent.children.filter(
    (child) =>
      child.kind !== 'textValue' &&
      child.namespaceUri === WML_NAMESPACE_URI &&
      child.localName === node.localName
  );
  if (records.length !== 1) return false;
  return (
    node.children.filter(
      (child) =>
        child.kind !== 'textValue' &&
        child.namespaceUri === WML_NAMESPACE_URI &&
        child.localName === container
    ).length <= 1
  );
}

/** `CT_SectPrBase`'s sequence (ECMA-376 Part 1, §17.6.17): the order section properties go in. */
export const SECTION_PROPERTY_SEQUENCE = [
  'footnotePr',
  'endnotePr',
  'type',
  'pgSz',
  'pgMar',
  'paperSrc',
  'pgBorders',
  'lnNumType',
  'pgNumType',
  'cols',
  'formProt',
  'vAlign',
  'noEndnote',
  'titlePg',
  'textDirection',
  'bidi',
  'rtlGutter',
  'docGrid',
  'printerSettings',
] as const;

const sequenceRank = (node: OoxmlNode): number =>
  node.kind === 'textValue' || node.namespaceUri !== WML_NAMESPACE_URI
    ? -1
    : (SECTION_PROPERTY_SEQUENCE as readonly string[]).indexOf(node.localName);

/** `children` with each of `added` at its place in the sequence, before the first one after it. */
function inSequence(children: readonly OoxmlNode[], added: readonly OoxmlNode[]): OoxmlNode[] {
  const out = [...children];
  for (const node of added) {
    const rank = sequenceRank(node);
    const at = rank < 0 ? -1 : out.findIndex((child) => sequenceRank(child) > rank);
    out.splice(at < 0 ? out.length : at, 0, node);
  }
  return out;
}

/**
 * Word restores the recorded section fields, retaining fields absent from the record.
 *
 * Word states `w:orient="portrait"` in the record, but the section it restores leaves that
 * default out (`w:pgSz`, ECMA-376 Part 1, §17.6.13; Word for Mac, reject all, 2026-10-29), so
 * the restored page size does not state it either. Word records a different first page turned
 * on as `<w:titlePg w:val="0"/>` (2026-10-30); the restored section leaves the off value out, as
 * Word writes no `w:titlePg` for a section without one. A recorded field the section no longer
 * has (a `w:titlePg` taken off) goes back at its place in the sequence.
 */
export function restoredSectionProperties(
  current: readonly OoxmlNode[],
  recorded: readonly OoxmlNode[]
): OoxmlNode[] {
  const key = (n: OoxmlNode) =>
    n.kind === 'textValue' ? n.id : `${n.namespaceUri}:${n.localName}`;
  const oldByKey = new Map(recorded.map((n) => [key(n), n]));
  const seen = new Set<string>();
  const restored = current.map((node) => {
    const name = key(node);
    seen.add(name);
    const old = oldByKey.get(name);
    if (!old) return node;
    if (node.kind === 'textValue' || old.kind === 'textValue') return old;
    // An on/off field is its recorded value whole: Word records a different first page turned
    // off as `<w:titlePg/>` over `<w:titlePg w:val="0"/>`, and a reject turns it on again
    if (isOnOffField(old)) return old;
    const attrs = new Map(node.attributes.map((a) => [`${a.namespaceUri}:${a.localName}`, a]));
    for (const a of old.attributes) attrs.set(`${a.namespaceUri}:${a.localName}`, a);
    return withoutPortrait({ ...old, attributes: [...attrs.values()] } as typeof old);
  });
  return inSequence(
    restored.filter(isNotTitlePageOff),
    recorded
      .filter((n) => !seen.has(key(n)))
      .map(withoutPortrait)
      .filter(isNotTitlePageOff)
  );
}

/** The section fields of type `CT_OnOff` (ECMA-376 Part 1, §17.6). */
const ON_OFF_FIELDS = new Set(['formProt', 'noEndnote', 'titlePg', 'bidi', 'rtlGutter']);

function isOnOffField(node: OoxmlNode): boolean {
  return (
    node.kind !== 'textValue' &&
    node.namespaceUri === WML_NAMESPACE_URI &&
    ON_OFF_FIELDS.has(node.localName)
  );
}

/** False for a `w:titlePg` that says off (`CT_OnOff`), which the restored section leaves out. */
function isNotTitlePageOff(node: OoxmlNode): boolean {
  if (node.kind === 'textValue' || node.namespaceUri !== WML_NAMESPACE_URI) return true;
  if (node.localName !== 'titlePg') return true;
  const value = node.attributes.find(
    (a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === 'val'
  )?.value;
  return value === undefined || !['0', 'false', 'off'].includes(value);
}

function withoutPortrait(node: OoxmlNode): OoxmlNode {
  if (node.kind === 'textValue' || node.namespaceUri !== WML_NAMESPACE_URI) return node;
  if (node.localName !== 'pgSz') return node;
  const attributes = node.attributes.filter(
    (a) =>
      a.namespaceUri !== WML_NAMESPACE_URI || a.localName !== 'orient' || a.value !== 'portrait'
  );
  return attributes.length === node.attributes.length ? node : { ...node, attributes };
}

/**
 * Numbering-reference history is metadata; Word retains numPr in either direction. That holds
 * for a `w:numPr/w:ins` and for a first-edition `w:numberingChange` alike: Word for Mac counts
 * a `w:numberingChange` as one revision, and accepting or rejecting it removes the record and
 * keeps the list (2026-10-29).
 */
export function validNumberingRevision(node: OoxmlNode, parent: OoxmlNode | null): boolean {
  const named = (n: OoxmlNode, name: string) =>
    n.kind !== 'textValue' && n.namespaceUri === WML_NAMESPACE_URI && n.localName === name;
  if (
    !parent ||
    parent.kind === 'textValue' ||
    !named(parent, 'numPr') ||
    node.kind === 'textValue' ||
    !(named(node, 'ins') || named(node, 'numberingChange')) ||
    node.children.length !== 0
  )
    return false;
  if (parent.children.filter((n) => named(n, node.localName)).length !== 1) return false;
  const ids = parent.children.filter((n) => named(n, 'numId'));
  if (ids.length !== 1 || ids[0]!.kind === 'textValue') return false;
  const value = ids[0]!.attributes.find(
    (a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === 'val'
  )?.value;
  return /^\d+$/.test(value ?? '');
}

/** A paragraph's valid `w:numberingChange` is a revision; a malformed one stays file content. */
export function numberingChangeSite(
  node: OoxmlNode,
  parent: OoxmlNode | null,
  grandparentName: string | undefined
): boolean {
  return (
    node.kind !== 'textValue' &&
    node.localName === 'numberingChange' &&
    grandparentName === 'pPr' &&
    validNumberingRevision(node, parent)
  );
}
