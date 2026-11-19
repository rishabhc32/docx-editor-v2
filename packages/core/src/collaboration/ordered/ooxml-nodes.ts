// Small node helpers that load and compose share.

import { RELATIONSHIPS_NAMESPACE_URI } from '../../store/package/ooxml-shared.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlAttribute,
  type OoxmlElement,
  type OoxmlNode,
} from '../../store/package/ooxml-tree.ts';

/**
 * Whether a node is the WordprocessingML element `w:<name>`. Not a type predicate on purpose:
 * a negated predicate would narrow the node to a text value.
 */
export function isWml(node: OoxmlNode, name: string): boolean {
  return (
    node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === name
  );
}

/** The node as the element `w:<name>`, or null. */
export function asWml(node: OoxmlNode, name: string): OoxmlElement | null {
  return node.kind !== 'textValue' && isWml(node, name) ? node : null;
}

/** An element's attribute value by local name, or undefined. */
export function attributeOf(node: OoxmlElement, name: string): string | undefined {
  return node.attributes.find((attribute) => attribute.localName === name)?.value;
}

/**
 * The same element with other children.
 *
 * The cast is confined here, as in the store's own `withChildren`: the typed node kinds narrow
 * their child lists, and the result is a copy of a node the reader made, with children of the
 * same kinds the reader gave it. Compose output is validated in the tests.
 */
export function withChildren(element: OoxmlElement, children: readonly OoxmlNode[]): OoxmlElement {
  return { ...element, children } as OoxmlElement;
}

/** The same element with another id (the same cast as {@link withChildren}). */
export function withId(element: OoxmlElement, id: string): OoxmlElement {
  return id === element.id ? element : ({ ...element, id } as OoxmlElement);
}

/** The same element with other attributes (the same cast as {@link withChildren}). */
export function withAttributes(
  element: OoxmlElement,
  attributes: readonly OoxmlAttribute[]
): OoxmlElement {
  return { ...element, attributes } as OoxmlElement;
}

/** The same element under another name (and kind), such as `w:t` as `w:delText`. */
export function renamed(
  element: OoxmlElement,
  kind: OoxmlElement['kind'],
  localName: string
): OoxmlElement {
  return { ...element, kind, localName } as OoxmlElement;
}

/** An element with only the children `keep` selects: the shell of a container or a run. */
export function shellOf(
  element: OoxmlElement,
  keep: (child: OoxmlNode) => boolean = () => false
): OoxmlElement {
  return withChildren(element, element.children.filter(keep));
}

/** The first child that is the element `w:<name>`. */
export function childNamed(element: OoxmlElement, name: string): OoxmlElement | undefined {
  for (const child of element.children) {
    const found = asWml(child, name);
    if (found) return found;
  }

  return undefined;
}

/** The first child that is an element. */
export function firstElement(element: OoxmlElement): OoxmlElement | undefined {
  for (const child of element.children) if (child.kind !== 'textValue') return child;
  return undefined;
}

export const isRunProperties = (node: OoxmlNode): boolean =>
  node.kind === 'runProperties' || isWml(node, 'rPr');

export const isParagraphProperties = (node: OoxmlNode): boolean =>
  node.kind === 'paragraphProperties' || isWml(node, 'pPr');

/** The first child that `test` selects, as an element, or undefined. */
export function childWhere(
  element: OoxmlElement,
  test: (child: OoxmlNode) => boolean
): OoxmlElement | undefined {
  for (const child of element.children) if (child.kind !== 'textValue' && test(child)) return child;
  return undefined;
}

/** Visit a node and every element under it, in no set order (a stack: a part can be deep). */
export function eachElement(root: OoxmlNode, visit: (element: OoxmlElement) => void): void {
  const stack: OoxmlNode[] = [root];
  for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
    if (node.kind === 'textValue') continue;
    visit(node);
    for (const child of node.children) if (child.kind !== 'textValue') stack.push(child);
  }
}

/** Field begins less field ends in a node. */
export function fieldCharBalance(node: OoxmlNode, depth = 0): number {
  if (node.kind === 'textValue' || depth > 8) {
    return 0;
  }

  if (node.localName === 'fldChar') {
    switch (node.attributes.find((a) => a.localName === 'fldCharType')?.value) {
      case 'begin':
        return 1;
      case 'end':
        return -1;
      default:
        return 0;
    }
  }

  let balance = 0;
  for (const child of node.children) {
    balance += fieldCharBalance(child, depth + 1);
  }

  return balance;
}

/** A new `w:<localName>` attribute. */
export const wmlAttribute = (localName: string, value: string): OoxmlAttribute => ({
  kind: 'genericExtension',
  namespaceUri: WML_NAMESPACE_URI,
  localName,
  prefix: 'w',
  value,
});

/** A new `r:id` attribute: the relationship a reference names. */
export const relationshipIdAttribute = (value: string): OoxmlAttribute => ({
  kind: 'genericExtension',
  namespaceUri: RELATIONSHIPS_NAMESPACE_URI,
  localName: 'id',
  prefix: 'r',
  value,
});

/** Whether a node is a control's content (`w:sdtContent`). */
export const isControlContent = (child: OoxmlNode): boolean =>
  child.kind === 'contentControlContent' || isWml(child, 'sdtContent');

/** Whether a control's properties say it shows its prompt (`w:sdtPr/w:showingPlcHdr`). */
export const showsPlaceholder = (control: OoxmlElement): boolean =>
  childNamed(control, 'sdtPr')?.children.some((c) => isWml(c, 'showingPlcHdr')) === true;

/** Revision markers inside a properties element: they become labels, not content. */
const MARKERS = new Set(['ins', 'del', 'rPrChange', 'moveFrom', 'moveTo', 'cellIns', 'cellDel']);

/** An element without its revision markers (`w:ins`, `w:del`, `w:rPrChange`, ...). */
export function withoutMarkers(element: OoxmlElement): OoxmlElement {
  if (!element.children.some((child) => child.kind !== 'textValue' && isMarker(child)))
    return element;
  return withChildren(
    element,
    element.children.filter((child) => child.kind === 'textValue' || !isMarker(child))
  );
}

function isMarker(node: OoxmlElement): boolean {
  return node.namespaceUri === WML_NAMESPACE_URI && MARKERS.has(node.localName);
}

/** A paragraph's `w:pPr` whose mark run properties are `face` (none when null). */
export function withMarkFace(pPr: OoxmlElement, face: OoxmlElement | null): OoxmlElement {
  const children: OoxmlNode[] = [];
  let placed = false;
  for (const child of pPr.children) {
    if (isRunProperties(child)) {
      if (face) children.push(face);
      placed = true;
      continue;
    }

    children.push(child);
  }

  if (!placed && face) {
    // The mark's run properties go before a section's properties and a change record
    const at = children.findIndex((child) => isWml(child, 'sectPr') || isWml(child, 'pPrChange'));
    children.splice(at < 0 ? children.length : at, 0, face);
  }

  return withChildren(pPr, children);
}

const NO_PREFIXES: ReadonlySet<string> = new Set();

/** The prefixes nodes use, with their namespaces, that the nodes do not declare themselves. */
function needed(nodes: readonly OoxmlNode[]): Map<string, string> {
  const out = new Map<string, string>();
  const need = (prefix: string | undefined, namespaceUri: string, local: ReadonlySet<string>) => {
    // `xml` is bound by XML itself
    if (
      prefix &&
      prefix !== 'xml' &&
      namespaceUri !== '' &&
      !local.has(prefix) &&
      !out.has(prefix)
    ) {
      out.set(prefix, namespaceUri);
    }
  };
  const stack: [OoxmlNode, ReadonlySet<string>][] = nodes.map((node) => [node, NO_PREFIXES]);
  for (let top = stack.pop(); top !== undefined; top = stack.pop()) {
    const [node, inherited] = top;
    if (node.kind === 'textValue') {
      continue;
    }

    const local =
      node.namespaceBindings.length === 0
        ? inherited
        : new Set([...inherited, ...node.namespaceBindings.map((binding) => binding.prefix)]);
    need(node.prefix, node.namespaceUri, local);
    for (const attribute of node.attributes) {
      need(attribute.prefix, attribute.namespaceUri, local);
    }

    for (const child of node.children) {
      if (child.kind !== 'textValue') {
        stack.push([child, local]);
      }
    }
  }

  return out;
}

/**
 * The bindings of the prefixes `nodes` use that neither they nor `root` declare, by prefix: what
 * a root that takes `nodes` in must declare.
 */
export function missingBindings(
  root: OoxmlElement,
  nodes: readonly OoxmlNode[]
): { prefix: string; namespaceUri: string }[] {
  const bound = new Set(root.namespaceBindings.map((binding) => binding.prefix));
  const added = [...needed(nodes)].filter(([prefix]) => !bound.has(prefix));
  added.sort(([a], [b]) => {
    if (a < b) {
      return -1;
    }

    return a > b ? 1 : 0;
  });
  return added.map(([prefix, namespaceUri]) => ({ prefix, namespaceUri }));
}

/** `root` declaring the prefixes `nodes` use that it lacks: `root` itself when it lacks none. */
export function withDeclaredPrefixes(
  root: OoxmlElement,
  nodes: readonly OoxmlNode[]
): OoxmlElement {
  const added = missingBindings(root, nodes);
  if (added.length === 0) {
    return root;
  }

  return { ...root, namespaceBindings: [...root.namespaceBindings, ...added] };
}
