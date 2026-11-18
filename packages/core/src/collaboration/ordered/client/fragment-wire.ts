import { namespaceScopeForNode } from '../../../store/package/drawing-projection-walk.ts';
import { W14_NAMESPACE_URI, WML_NAMESPACE_URI } from '../../../store/package/ooxml-shared.ts';
import type { OoxmlAttribute, OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { ORDERED_FRAGMENT_LIMITS, type OrderedFragmentElement } from '../change/fragment.ts';
import {
  FRAGMENT_LINK_OWNERS,
  FRAGMENT_OBJECT_ATTRIBUTES,
  FRAGMENT_OBJECT_ELEMENTS,
  fragmentAttributePrefix,
  fragmentElementNamespace,
  fragmentElementPrefix,
  isFragmentElement,
  isFragmentMediaReference,
  isFragmentPictureAddress,
  isFragmentTextElement,
  takesPlainAttributes,
  WP14_NAMESPACE_URI,
} from '../change/fragment-names.ts';
import { isPlainText } from '../change/read.ts';
import {
  DIAGRAM_DRAWING_REFERENCE,
  fragmentObjectReference,
  fragmentPartReference,
  type OrderedFragmentObjectUse,
  type OrderedFragmentPartType,
} from '../change/fragment-names-objects.ts';
import { OrderedClientRefusal } from './refusal.ts';

/** How the client names what a fragment uses outside the story text. */
export interface OrderedFragmentWireReferences {
  /** A `w:numId` value as the change says it: a decimal the document defines, or `#<k>`. */
  numId(value: string): string;
  /** A link's relationship id: `#<k>` into the change's links. */
  relationship(id: string): string;
  /** A picture's image relationship id: `#<k>` into the change's media. */
  media(id: string): string;
  /** A note reference's id: `#<k>` into the change's notes. */
  note(kind: 'footnote' | 'endnote', id: string): string;
  /** A chart's or a diagram's part, by the story's relationship id: `#<k>` into the parts. */
  part?(id: string, type: OrderedFragmentPartType): string;
  /** An embedded object's part or an ink part, by its relationship id: `#<k>` into the media. */
  object?(id: string, use: OrderedFragmentObjectUse): string;
  /** A list level's picture bullet, by its id: `#<j>` into the change's bullets. */
  bullet?(id: string): string;
  /** Set while a part's own tree is written (`fragmentPartWire`). */
  readonly own?: OrderedFragmentPartWireReferences;
}

/** How the client names what a chart's or a diagram's part uses. */
export interface OrderedFragmentPartWireReferences {
  /** One of the part's own relationship ids: `#<j>` into the part's relationships. */
  relationship(id: string): string;
  /** A diagram data part's drawing, by the story's relationship id: `#<k>` into the parts. */
  drawing(id: string): string;
}

/**
 * The element tree of a chart's or a diagram's part (its root), with its own relationships and
 * its diagram drawing named through `own`. Refused as `fragmentWire` refuses.
 */
export function fragmentPartWire(
  root: OoxmlElement,
  refs: OrderedFragmentWireReferences,
  own: OrderedFragmentPartWireReferences
): OrderedFragmentElement {
  return wire(root, { ...refs, own }, 1, new Map());
}

/**
 * Element trees for engine nodes (a pasted fragment's), the reverse of `edits/fragment-nodes.ts`.
 * Node ids, the paragraphs' `w14:paraId` and `w14:textId`, and Word's own drawing ids do not
 * travel (the apply mints them); a reference to a part outside the story text becomes `#<k>`
 * through `refs`. Throws `OrderedClientRefusal` for what the format cannot
 * carry (`insertFragment.object`: an ActiveX control, an element of a namespace the format
 * does not hold outside an extension and a choice, an object the refs carry no part for), and
 * for a picture by address with no embedded copy (`external-picture-link`).
 */
export function fragmentWire(
  nodes: readonly OoxmlNode[],
  refs: OrderedFragmentWireReferences
): OrderedFragmentElement[] {
  const out: OrderedFragmentElement[] = [];
  for (const node of nodes)
    if (node.kind !== 'textValue') out.push(wire(node, refs, 1, new Map()));
    else if (!INDENTATION.test(node.value))
      throw refused('fragment-text', WML_NAMESPACE_URI, 'body');
  return out;
}

/** Whitespace the reader keeps between elements: never content outside a text element. */
const INDENTATION = /^[ \t\r\n]*$/;

/**
 * Namespaces of objects the format does not carry: DrawingML's charts and diagrams, Office's
 * other drawing extensions, other VML namespaces, Word's drawing canvas and ink.
 */
function isObjectNamespace(uri: string): boolean {
  return (
    uri.startsWith('http://schemas.openxmlformats.org/drawingml/') ||
    uri.startsWith('http://purl.oclc.org/ooxml/drawingml/') ||
    uri.startsWith('http://schemas.microsoft.com/office/drawing/') ||
    uri.startsWith('urn:schemas-microsoft-com:') ||
    /^http:\/\/schemas\.microsoft\.com\/office\/word\/20[0-9]{2}\/wordprocessing[A-Za-z]{1,16}$/.test(
      uri
    )
  );
}

/** A pasted object the editor alone does not carry either. */
export const objectRefusal = (): OrderedClientRefusal =>
  new OrderedClientRefusal('ordered-unsupported-op', 'insertFragment.object');

/** A picture whose image only an address names: no room writes one (design D18). */
export const pictureByAddress = (): OrderedClientRefusal =>
  new OrderedClientRefusal('ordered-refused', 'external-picture-link', 'a picture by address');

/**
 * Whether a picture's address (`isFragmentPictureAddress`) is left out. Word loads the address
 * when it opens the file, with no click, so no room writes one (design D18): the picture keeps
 * its embedded copy, and a picture with no embedded copy is refused by name. Never fetched.
 */
function withoutAddress(node: OoxmlElement, owner: string, key: string): boolean {
  if (!isFragmentPictureAddress(owner, key)) return false;
  for (const attribute of node.attributes as readonly OoxmlAttribute[]) {
    const prefix = fragmentAttributePrefix(attribute.namespaceUri);
    if (prefix !== undefined && isFragmentMediaReference(owner, `${prefix}:${attribute.localName}`))
      return true;
  }

  throw pictureByAddress();
}

/** Word's own ids of a drawing, for its merging: the apply needs none. */
const DRAWING_IDS: ReadonlySet<string> = new Set(['anchorId', 'editId']);

/** Whether an attribute is one of Word's own drawing ids: a frame's, or a VML shape's. */
const isDrawingId = (attribute: OoxmlAttribute): boolean =>
  (attribute.namespaceUri === WP14_NAMESPACE_URI && DRAWING_IDS.has(attribute.localName)) ||
  (attribute.namespaceUri === W14_NAMESPACE_URI && attribute.localName === 'anchorId');

/**
 * The change's own prefixes of a choice's `Requires`: each through the declarations the blocks
 * carry, or as Word names it (the source document's root declarations do not travel). Undefined
 * for a namespace the format does not carry.
 */
function requiredPrefixes(
  choice: OoxmlElement,
  scope: ReadonlyMap<string, string>
): (string | undefined)[] {
  const value = choice.attributes.find(
    (attribute) => attribute.namespaceUri === '' && attribute.localName === 'Requires'
  )?.value;
  return (value ?? '')
    .trim()
    .split(/\s+/)
    .map((token) => {
      const uri = scope.get(token) ?? fragmentElementNamespace(token);
      return uri === undefined ? undefined : fragmentElementPrefix(uri);
    });
}

/** Whether a choice requires a namespace the format does not carry. */
const requiresUnknown = (choice: OoxmlElement, scope: ReadonlyMap<string, string>): boolean =>
  requiredPrefixes(choice, scope).includes(undefined);

/** A choice's `Requires` as the change says it. A namespace it does not carry is an object. */
function requires(choice: OoxmlElement, scope: ReadonlyMap<string, string>): string {
  const prefixes = requiredPrefixes(choice, scope);
  if (prefixes.includes(undefined)) throw objectRefusal();
  return [...new Set(prefixes)].join(' ');
}

const isNamed = (node: OoxmlElement, e: string): boolean =>
  `${fragmentElementPrefix(node.namespaceUri)}:${node.localName}` === e;

/** An element and the namespace scope it is read in. */
interface Scoped {
  readonly node: OoxmlElement;
  readonly scope: ReadonlyMap<string, string>;
}

/** Whether a choice the format holds is among a choice-holder's children. */
function holdsChoice(holder: OoxmlElement, scope: ReadonlyMap<string, string>): boolean {
  return holder.children.some(
    (c) =>
      c.kind !== 'textValue' &&
      isNamed(c, 'mc:Choice') &&
      !requiresUnknown(c, namespaceScopeForNode(scope, c))
  );
}

/**
 * What the change carries for one element child, read in `scope`. What an application that
 * does not know a namespace leaves out goes, as Word older than the namespace reads it: a
 * DrawingML extension (`a:ext`, ECMA-376 Part 1 §20.1.2.2.15) that holds an element of a
 * namespace the format does not hold, and a choice (`mc:Choice`, Part 3 §10.2) that requires
 * one; a choice-holder left with no choice is its fallback's content. The format names every
 * element it carries, so the room never takes content it cannot check.
 */
function carriedOf(child: OoxmlElement, scope: ReadonlyMap<string, string>): Scoped[] {
  const own = namespaceScopeForNode(scope, child);
  if (isNamed(child, 'a:ext'))
    return child.children.some(
      (c) => c.kind !== 'textValue' && fragmentElementPrefix(c.namespaceUri) === undefined
    )
      ? []
      : [{ node: child, scope }];
  if (isNamed(child, 'mc:Choice'))
    return requiresUnknown(child, own) ? [] : [{ node: child, scope }];
  if (!isNamed(child, 'mc:AlternateContent') || holdsChoice(child, own))
    return [{ node: child, scope }];
  const out: Scoped[] = [];
  for (const fallback of child.children) {
    if (fallback.kind === 'textValue' || !isNamed(fallback, 'mc:Fallback')) continue;
    const inner = namespaceScopeForNode(own, fallback);
    for (const c of fallback.children) if (c.kind !== 'textValue') out.push(...carriedOf(c, inner));
  }

  return out;
}

/** A refusal that names what the format does not hold, by a short and printable name. */
function refused(what: string, namespaceUri: string, localName: string): OrderedClientRefusal {
  const label = (text: string, max: number): string =>
    (text.length > max ? `${text.slice(0, max)}...` : text).replace(/[^\x21-\x7E]/g, '?');
  const segments = namespaceUri.replace(/^[A-Za-z]{1,16}:\/*/, '').split(/[/:]/);
  const namespace =
    fragmentAttributePrefix(namespaceUri) ??
    (namespaceUri === '' ? '(none)' : label(segments.slice(-2).join('/'), 48));
  return new OrderedClientRefusal(
    'ordered-refused',
    'invalid-change',
    `${what} ${namespace}:${label(localName, 64)}`
  );
}

function wire(
  node: OoxmlElement,
  refs: OrderedFragmentWireReferences,
  depth: number,
  inherited: ReadonlyMap<string, string>
): OrderedFragmentElement {
  if (depth > ORDERED_FRAGMENT_LIMITS.depth)
    throw new OrderedClientRefusal('ordered-refused', 'invalid-change', 'too-many fragment depth');
  const prefix = fragmentElementPrefix(node.namespaceUri);
  if (prefix === undefined) {
    if (isObjectNamespace(node.namespaceUri)) throw objectRefusal();
    throw refused('fragment-element', node.namespaceUri, node.localName);
  }

  const e = `${prefix}:${node.localName}`;
  if (FRAGMENT_OBJECT_ELEMENTS.has(e)) throw objectRefusal();
  if (!isFragmentElement(prefix, node.localName))
    throw refused('fragment-element', node.namespaceUri, node.localName);
  const scope = namespaceScopeForNode(inherited, node);
  const a = attributesOf(node, e, refs, scope);
  const c = childrenOf(node, e, refs, depth, scope);
  return {
    e,
    ...(a === undefined ? {} : { a }),
    ...(c.length === 0 ? {} : { c }),
  };
}

/**
 * A plain attribute's value as the change carries it: a choice's `Requires` read in its scope, a
 * diagram's drawing reference through `refs`, any other as it is.
 */
function plainValue(
  node: OoxmlElement,
  owner: string,
  attribute: OoxmlAttribute,
  refs: OrderedFragmentWireReferences,
  scope: ReadonlyMap<string, string>
): string {
  if (owner === 'mc:Choice' && attribute.localName === 'Requires') {
    return requires(node, scope);
  }

  if (refs.own !== undefined && `${owner}@${attribute.localName}` === DIAGRAM_DRAWING_REFERENCE) {
    return refs.own.drawing(attribute.value);
  }

  return attribute.value;
}

function attributesOf(
  node: OoxmlElement,
  owner: string,
  refs: OrderedFragmentWireReferences,
  scope: ReadonlyMap<string, string>
): Record<string, string> | undefined {
  let out: Record<string, string> | undefined;
  const plain = takesPlainAttributes(owner);
  for (const attribute of node.attributes as readonly OoxmlAttribute[]) {
    if (isDrawingId(attribute)) continue;
    if (attribute.namespaceUri === '' && plain) {
      const { localName } = attribute;
      if (withoutAddress(node, owner, localName)) continue;

      // An object by address: never fetched
      if (FRAGMENT_OBJECT_ATTRIBUTES.has(localName)) throw objectRefusal();
      out ??= {};
      out[localName] = plainValue(node, owner, attribute, refs, scope);
      continue;
    }

    const prefix = fragmentAttributePrefix(attribute.namespaceUri);
    if (prefix === undefined || (prefix === 'xml' && attribute.localName !== 'space'))
      throw refused('fragment-attribute', attribute.namespaceUri, attribute.localName);

    // The apply mints a paragraph's ids: two copies never share one.
    if (prefix === 'w14' && (attribute.localName === 'paraId' || attribute.localName === 'textId'))
      continue;
    const key = `${prefix}:${attribute.localName}`;
    if (withoutAddress(node, owner, key)) continue;
    out ??= {};
    out[key] = wireValue(owner, prefix, attribute, refs);
  }

  return out;
}

/** An attribute's value, with a reference outside the story text as the change names it. */
function wireValue(
  owner: string,
  prefix: string,
  attribute: OoxmlAttribute,
  refs: OrderedFragmentWireReferences
): string {
  const { localName, value } = attribute;
  const key = `${prefix}:${localName}`;
  if (prefix === 'r') {
    // In a chart's or a diagram's part, every relationship is the part's own (an empty one
    // names nothing)
    if (refs.own !== undefined) return value === '' ? value : refs.own.relationship(value);

    // A link's relationship, a picture's image, an embedded object's or ink's part and a
    // chart's or a diagram's parts travel as indexes; every other one names an object, never
    // fetched
    if (FRAGMENT_LINK_OWNERS.has(owner) && localName === 'id') return refs.relationship(value);
    if (isFragmentMediaReference(owner, key)) return refs.media(value);
    const use = fragmentObjectReference(owner, key);
    if (use !== undefined && refs.object !== undefined) return refs.object(value, use);
    const part = fragmentPartReference(owner, key);
    if (part !== undefined && refs.part !== undefined) return refs.part(value, part);
    throw objectRefusal();
  }

  // A VML picture's image by the Office relationship id older Word writes
  if (isFragmentMediaReference(owner, key)) return refs.media(value);

  // An object by address
  if (FRAGMENT_OBJECT_ATTRIBUTES.has(key)) throw objectRefusal();
  if (prefix !== 'w') return value;
  if (owner === 'w:numId' && localName === 'val') return refs.numId(value);
  if (owner === 'w:lvlPicBulletId' && localName === 'val') {
    if (refs.bullet === undefined) throw objectRefusal();
    return refs.bullet(value);
  }

  if (owner === 'w:footnoteReference' && localName === 'id') return refs.note('footnote', value);
  if (owner === 'w:endnoteReference' && localName === 'id') return refs.note('endnote', value);
  return value;
}

function childrenOf(
  node: OoxmlElement,
  owner: string,
  refs: OrderedFragmentWireReferences,
  depth: number,
  scope: ReadonlyMap<string, string>
): (OrderedFragmentElement | string)[] {
  const textual = isFragmentTextElement(owner);
  const out: (OrderedFragmentElement | string)[] = [];
  for (const child of node.children) {
    if (child.kind !== 'textValue') {
      for (const kept of carriedOf(child, scope))
        out.push(wire(kept.node, refs, depth + 1, kept.scope));
      continue;
    }

    if (!textual) {
      if (INDENTATION.test(child.value)) continue;
      throw refused('fragment-text', node.namespaceUri, node.localName);
    }

    if (child.value === '') continue;

    // Word holds a tab and a line break as elements; text holds neither.
    if (!isPlainText(child.value))
      throw refused('fragment-text', node.namespaceUri, node.localName);
    const last = out.length - 1;
    if (typeof out[last] === 'string') out[last] += child.value;
    else out.push(child.value);
  }

  return out;
}
