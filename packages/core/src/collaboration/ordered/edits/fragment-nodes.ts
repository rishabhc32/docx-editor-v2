// Engine nodes for a pasted fragment's checked element trees (group 9.10). The trees are
// written as XML text with every name taken from the allowlist and every value escaped, and
// read back with the engine's bounded part reader: the trust boundary every Word file passes.
// No XML text from a change is ever read: the change holds trees, and this module writes them.

import {
  readOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPartMetadata,
} from '../../../store/package/ooxml-tree.ts';
import { escapeXml, escapeXmlAttribute, isValidXmlText } from '../../../store/package/sinks.ts';
import { XML_HARD_MAX_BYTES } from '../../../store/package/xml-reader.ts';
import {
  ORDERED_FRAGMENT_LIMITS,
  type OrderedFragmentElement,
  type OrderedFragmentRoot,
} from '../change/fragment.ts';
import {
  FRAGMENT_LINK_OWNERS,
  FRAGMENT_NAMESPACE_DECLARATIONS,
  fragmentElementNamespace,
  fragmentElementPrefix,
  isFragmentMediaReference,
  isFragmentRequires,
} from '../change/fragment-names.ts';
import {
  DIAGRAM_DRAWING_REFERENCE,
  FRAGMENT_PART_ROOTS,
  fragmentObjectReference,
  fragmentPartReference,
  type OrderedFragmentObjectUse,
  type OrderedFragmentPartType,
} from '../change/fragment-names-objects.ts';
import { OrderedRefusal } from '../model/refusal.ts';

/** The Word names the apply minted for a fragment's `#<k>` references. */
export interface OrderedFragmentResolver {
  /** A `w:numId` value: `#<k>` gives the minted list number, a decimal stays as it is. */
  numId(value: string): string;
  /** A link's `r:id` value `#<k>`: the minted relationship id. */
  relationship(value: string): string;
  /** A note reference's `w:id` value `#<k>`: the minted note id. */
  note(value: string): string;
  /** A picture's image `#<k>` (`a:blip/@r:embed`, `v:imagedata/@r:id`): the minted id. */
  media(value: string): string;
  /** An embedded object's or ink's part `#<k>` (`o:OLEObject/@r:id`): the minted id. */
  object(value: string, use: OrderedFragmentObjectUse): string;
  /** A drawing's document-wide `wp:docPr/@id`: a new one. */
  drawingId(): string;
  /** A chart's or a diagram's part `#<k>` (`c:chart/@r:id`, `dgm:relIds/@r:dm`): the minted id. */
  part(value: string, type: OrderedFragmentPartType): string;
  /** A list level's picture bullet `#<j>` (`w:lvlPicBulletId/@w:val`): the minted bullet id. */
  bullet?(value: string): string;
}

/** The Word names the apply minted for the references of a chart's or a diagram's part. */
export interface OrderedFragmentPartResolver {
  /** An `r:` attribute's value `#<j>`, into the part's own relationships: the minted id. */
  relationship(value: string): string;
  /** A diagram data part's drawing `#<k>` (`dsp:dataModelExt/@relId`): the story's minted id. */
  drawing(value: string): string;
}

/**
 * The engine's root element of a chart's or a diagram's part, from its checked tree: written
 * with every fragment declaration on the root and read back as the part `name`, of `contentType`.
 */
export function fragmentPartNode(
  tree: OrderedFragmentElement,
  type: OrderedFragmentPartType,
  name: string,
  contentType: string,
  resolve: OrderedFragmentPartResolver
): OoxmlElement {
  if (tree.e !== FRAGMENT_PART_ROOTS[type]) refuse(`not a ${type} part`);
  const value = (owner: string, key: string, raw: string): string => {
    // An empty relationship names nothing (a diagram layout's `r:blip=""`)
    if (key.startsWith('r:')) return raw === '' ? raw : resolve.relationship(raw);
    if (`${owner}@${key}` === DIAGRAM_DRAWING_REFERENCE) return resolve.drawing(raw);
    return raw;
  };
  const out: string[] = [];
  write(tree, value, out, 1, ` ${FRAGMENT_NAMESPACE_DECLARATIONS}`);
  const root = readPart(out.join(''), { name, contentType });
  if (`${fragmentElementPrefix(root.namespaceUri) ?? ''}:${root.localName}` !== tree.e)
    refuse('the reader read another part');
  return root;
}

/**
 * Engine nodes for checked trees, in order, with every `#<k>` reference replaced through
 * `resolve`. Refused by name when the engine's reader refuses them (`engine-refused`).
 */
export function fragmentNodes(
  trees: readonly OrderedFragmentElement[],
  resolve: OrderedFragmentResolver
): OoxmlElement[] {
  const root = readPart(
    `<w:document ${FRAGMENT_NAMESPACE_DECLARATIONS}><w:body>${written(trees, resolve)}</w:body></w:document>`,
    DOCUMENT
  );
  for (const child of root.children as readonly OoxmlNode[])
    if (
      child.kind !== 'textValue' &&
      child.namespaceUri === WML_NAMESPACE_URI &&
      child.localName === 'body'
    )
      return elementsOf(child, trees);
  return refuse('no body');
}

/**
 * Engine nodes for checked resource trees (`w:abstractNum`, `w:lvlOverride`, `w:numPicBullet`,
 * `w:style`), in order: written inside a `w:numbering` or a `w:styles` root and read back as its
 * children.
 */
export function fragmentResourceNodes(
  trees: readonly OrderedFragmentElement[],
  root: Exclude<OrderedFragmentRoot, 'block'>,
  resolve: OrderedFragmentResolver
): OoxmlElement[] {
  const name = root === 'style' ? 'styles' : 'numbering';
  const read = readPart(
    `<w:${name} ${FRAGMENT_NAMESPACE_DECLARATIONS}>${written(trees, resolve)}</w:${name}>`,
    root === 'style' ? STYLES : NUMBERING
  );
  return elementsOf(read, trees);
}

const WML_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const DOCUMENT: OoxmlPartMetadata = {
  name: '/word/document.xml',
  contentType: `${WML_TYPE}.document.main+xml`,
};
const NUMBERING: OoxmlPartMetadata = {
  name: '/word/numbering.xml',
  contentType: `${WML_TYPE}.numbering+xml`,
};
const STYLES: OoxmlPartMetadata = {
  name: '/word/styles.xml',
  contentType: `${WML_TYPE}.styles+xml`,
};

/**
 * An element or attribute name the writer writes as it is: the checked form only. It allows
 * `xmlns:w`; `isDeclaration` refuses every namespace declaration.
 */
const NAME = /^([a-z][a-z0-9]{0,4}):([A-Za-z][A-Za-z0-9]{0,63})$/;

/** A drawing's attribute name, without a prefix (`isDeclaration` refuses `xmlns`). */
const PLAIN_NAME = /^[A-Za-z][A-Za-z0-9]{0,63}$/;

function refuse(reason: string): never {
  throw new OrderedRefusal('engine-refused', `a fragment: ${reason}`);
}

function readPart(xml: string, metadata: OoxmlPartMetadata): OoxmlElement {
  // The root and the body wrap the trees: two elements more than a change may hold.
  const read = readOoxmlPart(xml, metadata, {
    maxBytes: XML_HARD_MAX_BYTES,
    maxElements: ORDERED_FRAGMENT_LIMITS.nodes + 2,
  });
  if (!read.ok) refuse(read.reason);
  return read.part.root;
}

/** The element children of a read root: one for each tree, with the tree's `w` name. */
function elementsOf(
  parent: OoxmlElement,
  trees: readonly OrderedFragmentElement[]
): OoxmlElement[] {
  const out: OoxmlElement[] = [];
  for (const child of parent.children) if (child.kind !== 'textValue') out.push(child);
  if (out.length !== trees.length) refuse('the reader read another number of trees');
  out.forEach((node, k) => {
    if (node.namespaceUri !== WML_NAMESPACE_URI || `w:${node.localName}` !== trees[k]?.e)
      refuse('the reader read another tree');
  });
  return out;
}

/** The trees as XML text: names checked, values escaped, references resolved. */
function written(
  trees: readonly OrderedFragmentElement[],
  resolve: OrderedFragmentResolver
): string {
  const out: string[] = [];
  const value = (owner: string, key: string, raw: string) => resolved(owner, key, raw, resolve);
  for (const tree of trees) write(tree, value, out, 1);
  return out.join('');
}

/** `xmlns` and `xmlns:<prefix>`: a namespace declaration, which no tree may hold. */
const isDeclaration = (name: string): boolean => name === 'xmlns' || name.startsWith('xmlns:');

function write(
  tree: OrderedFragmentElement,
  value: (owner: string, key: string, raw: string) => string,
  out: string[],
  depth: number,
  declarations = ''
): void {
  if (depth > ORDERED_FRAGMENT_LIMITS.depth) refuse('too deep');
  if (!NAME.test(tree.e) || isDeclaration(tree.e)) refuse('an element name');
  out.push('<', tree.e, declarations);
  const attributes = tree.a;
  if (tree.e === 'mc:Choice') out.push(requiredDeclarations(attributes?.['Requires']));
  if (attributes !== undefined)
    for (const key of Object.keys(attributes)) {
      if ((!NAME.test(key) && !PLAIN_NAME.test(key)) || isDeclaration(key))
        refuse('an attribute name');
      const written = value(tree.e, key, checked(attributes[key]));
      out.push(' ', key, '="', escapeXmlAttribute(checked(written)), '"');
    }

  const children = tree.c;
  if (children === undefined || children.length === 0) {
    out.push('/>');
    return;
  }

  out.push('>');
  for (const child of children)
    if (typeof child === 'string') out.push(escapeXml(checked(child)));
    else write(child, value, out, depth + 1);
  out.push('</', tree.e, '>');
}

/**
 * The declarations of the prefixes a choice requires, on the choice itself: `Requires` names
 * prefixes, which resolve through the declarations in scope wherever the paste lands.
 */
function requiredDeclarations(value: string | undefined): string {
  if (value === undefined || !isFragmentRequires(value)) refuse('a choice');
  return value
    .split(' ')
    .map((prefix) => ` xmlns:${prefix}="${fragmentElementNamespace(prefix) ?? ''}"`)
    .join('');
}

/** A value with its reference outside the story replaced by the Word name the apply minted. */
function resolved(
  owner: string,
  key: string,
  value: string,
  resolve: OrderedFragmentResolver
): string {
  if (owner === 'w:numId' && key === 'w:val') return resolve.numId(value);
  if (owner === 'w:lvlPicBulletId' && key === 'w:val')
    return resolve.bullet ? resolve.bullet(value) : refuse('a picture bullet outside a list');
  if (key === 'r:id' && FRAGMENT_LINK_OWNERS.has(owner)) return resolve.relationship(value);
  if (isFragmentMediaReference(owner, key)) return resolve.media(value);
  const object = fragmentObjectReference(owner, key);
  if (object !== undefined) return resolve.object(value, object);
  const part = fragmentPartReference(owner, key);
  if (part !== undefined) return resolve.part(value, part);
  if (owner === 'wp:docPr' && key === 'id') return resolve.drawingId();
  if ((owner === 'w:footnoteReference' || owner === 'w:endnoteReference') && key === 'w:id')
    return resolve.note(value);
  return value;
}

/** A value the XML may hold: a string of XML 1.0 characters. */
function checked(value: unknown): string {
  if (typeof value !== 'string' || !isValidXmlText(value)) refuse('a value');
  return value;
}
