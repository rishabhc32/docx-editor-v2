import type { OoxmlPackage } from '../../../../../store/package/ooxml-package.ts';
import type {
  OoxmlElement,
  OoxmlNode,
  OoxmlPart,
} from '../../../../../store/package/ooxml-tree.ts';
import type { OrderedDocument } from '../../../model/types.ts';
import { attributeOf, isWml } from '../../../ooxml-nodes.ts';

/** The element children of a node. */
export function elements(node: OoxmlNode): OoxmlElement[] {
  const out: OoxmlElement[] = [];
  if (node.kind === 'textValue') {
    return out;
  }

  for (const child of node.children) {
    if (child.kind !== 'textValue') {
      out.push(child);
    }
  }

  return out;
}

/** The part a relationship type points at from the main document, in a package. */
export function relatedPart(pkg: OoxmlPackage, type: string): OoxmlPart | null {
  const dir = pkg.mainDocumentPart.slice(0, pkg.mainDocumentPart.lastIndexOf('/') + 1);
  for (const record of pkg.relationships.get(pkg.mainDocumentPart) ?? []) {
    if (!record.type.endsWith(`/${type}`)) {
      continue;
    }

    const name = record.rawTarget.startsWith('/') ? record.rawTarget : `${dir}${record.rawTarget}`;
    const part = pkg.parts.get(name);
    if (part) {
      return part;
    }
  }

  return null;
}

/** Whether a list number is one the document does not define, nor a change added. */
export function isNewList(doc: OrderedDocument, numId: string | null): numId is string {
  if (numId === null || numId === '0' || doc.numbering === null || doc.numbering.has(numId)) {
    return false;
  }

  for (const resource of doc.resources.added.values()) {
    if (resource.kind === 'list' && resource.numId === numId) {
      return false;
    }
  }

  return true;
}

/** A list number's `w:num` and the `w:abstractNum` it names, in a package's numbering part. */
export function listElementsOf(
  pkg: OoxmlPackage,
  numId: string
): { readonly num: OoxmlElement; readonly abstract: OoxmlElement } | null {
  const numbering = relatedPart(pkg, 'numbering');
  const num = numbering
    ? elements(numbering.root).find((n) => isWml(n, 'num') && attributeOf(n, 'numId') === numId)
    : undefined;
  if (!numbering || !num) {
    return null;
  }

  const abstractId = attributeOf(
    elements(num).find((n) => isWml(n, 'abstractNumId')) ?? num,
    'val'
  );
  if (abstractId === undefined) {
    return null;
  }

  const abstract = elements(numbering.root).find(
    (n) => isWml(n, 'abstractNum') && attributeOf(n, 'abstractNumId') === abstractId
  );
  return abstract ? { num, abstract } : null;
}

/** Whether a change added a style with this id (a style no file of the room holds). */
export function addedStyleId(doc: OrderedDocument, styleId: string): boolean {
  for (const resource of doc.resources.added.values()) {
    if (resource.kind === 'style' && resource.styleId === styleId) {
      return true;
    }
  }

  return false;
}
