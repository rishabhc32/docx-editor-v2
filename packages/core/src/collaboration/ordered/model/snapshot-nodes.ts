// The Word XML nodes of a saved item list (checkpoint version 6, design D14), in the stream of
// `snapshot-strings.ts`. The base package's nodes have numbers: its parts by name, each part's
// nodes in document order. Any other node is written by value at its first use, and has the next
// number of the listed nodes, after its children; so two items that share one node share it again
// after an open. A node is written as:
// - `d`: the base node `d` after the last base node named;
// - `[d]`: the listed node `d` after the last listed node named;
// - `[k, children]`: a new node, the base element `k` (named as above) with other children;
// - `[id, value]`: a new text node;
// - `[id, kind, namespace, name, prefix, bindings, attributes, children]`: a new element, with
//   every field: its kind, names, authored prefix, namespace declarations, attributes, children
//   and id.

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { findNode } from '../../../store/package/ooxml-edit.ts';
import {
  deepFreezeNode,
  readOoxmlPart,
  serializeOoxmlPart,
  XML_NAMESPACE_URI,
  WML_NAMESPACE_URI,
  type OoxmlAttribute,
  type OoxmlElement,
  type OoxmlNamespaceBinding,
  type OoxmlNode,
  type OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import { isXmlNCName } from '../../../store/package/qname.ts';
import { MAX_XML_DEPTH, XML_HARD_MAX_BYTES } from '../../../store/package/xml-reader.ts';
import { isValidXmlText } from '../../../store/package/sinks.ts';
import { at, fieldsOf, Invalid, required } from '../change/read.ts';
import { withChildren } from '../ooxml-nodes.ts';
import {
  ID_MAX,
  SnapshotChain,
  SnapshotStringReader,
  SnapshotStringWriter,
} from './snapshot-strings.ts';

/** The namespace of the element that holds the nodes by value while the XML reader checks them. */
const HOLDER_NAMESPACE = 'https://docx-editor.dev/ordered-snapshot';

/**
 * The attribute kinds the engine writes, by their number: the typed ones, and `generic`, which
 * some edits give an attribute they make (its readers treat it as `genericExtension`). An
 * attribute an edit made may have no kind (null).
 */
const ATTRIBUTE_KINDS = ['xmlSpace', 'wmlVal', 'genericExtension', 'generic'] as const;

const KINDS: readonly string[] = ATTRIBUTE_KINDS;

/** The base package's nodes by number: its parts by name, each part's nodes in document order. */
function baseNodes(base: OoxmlPackage): OoxmlNode[] {
  const out: OoxmlNode[] = [];
  const walk = (node: OoxmlNode) => {
    out.push(node);
    if (node.kind !== 'textValue') {
      node.children.forEach(walk);
    }
  };
  const parts = [...base.parts.values()].sort((a, b) => (a.name < b.name ? -1 : 1));
  for (const part of parts) {
    walk(part.root);
  }

  return out;
}

/** The base package's node with `id`, or null. */
function baseNodeOf(base: OoxmlPackage, id: string): OoxmlNode | null {
  const at = id.indexOf('#');
  const part = at > 0 ? base.parts.get(id.slice(0, at)) : undefined;
  return part === undefined ? null : findNode(part, id);
}

/** Whether two elements are equal in every field but their children. */
const sameShell = (a: OoxmlElement, b: OoxmlElement): boolean =>
  a.id === b.id &&
  a.kind === b.kind &&
  a.namespaceUri === b.namespaceUri &&
  a.localName === b.localName &&
  a.prefix === b.prefix &&
  a.namespaceBindings.length === b.namespaceBindings.length &&
  a.namespaceBindings.every(
    (x, k) =>
      x.prefix === b.namespaceBindings[k]!.prefix &&
      x.namespaceUri === b.namespaceBindings[k]!.namespaceUri
  ) &&
  attributesDifference(a.attributes, b.attributes, '') === null;

/** Writes nodes: each base node by its number, each other node by value at its first use. */
export class SnapshotNodeWriter {
  readonly strings: SnapshotStringWriter;
  readonly #base: OoxmlPackage;
  readonly #numbers = new Map<OoxmlNode, number>();
  readonly #baseChain = new SnapshotChain();
  readonly #listed = new Map<OoxmlNode, number>();
  readonly #listChain = new SnapshotChain();

  constructor(base: OoxmlPackage) {
    this.#base = base;
    baseNodes(base).forEach((node, k) => {
      if (!this.#numbers.has(node)) {
        this.#numbers.set(node, k);
      }
    });
    this.strings = new SnapshotStringWriter((id) => this.#baseNumber(id), this.#baseChain);
  }

  /** How the checkpoint names `node`. */
  ref(node: OoxmlNode): unknown {
    const base = this.#numbers.get(node);
    if (base !== undefined) {
      return this.#baseChain.write(base);
    }

    const listed = this.#listed.get(node);
    if (listed !== undefined) {
      return [this.#listChain.write(listed)];
    }

    const value = this.#value(node);
    const place = this.#listed.size;
    this.#listed.set(node, place);
    this.#listChain.move(place);
    return value;
  }

  /** A part: the base package's by its name, any other by value. */
  part(part: OoxmlPart): unknown {
    if (this.#base.parts.get(part.name) === part) {
      return this.strings.string(part.name);
    }

    return {
      id: this.strings.string(part.id),
      name: this.strings.string(part.name),
      contentType: this.strings.string(part.contentType),
      root: this.ref(part.root),
    };
  }

  /** Attributes as the checkpoint writes them. */
  attributes(attributes: readonly OoxmlAttribute[]): unknown[] {
    return attributes.map((a) => [
      a.kind === undefined ? null : KINDS.indexOf(a.kind),
      this.strings.name(a.namespaceUri),
      this.strings.name(a.localName),
      a.prefix === undefined ? null : this.strings.name(a.prefix),
      this.strings.string(a.value),
    ]);
  }

  /** The number of the base node whose id is `id`. */
  #baseNumber(id: string): number | undefined {
    const node = baseNodeOf(this.#base, id);
    return node?.id === id ? this.#numbers.get(node) : undefined;
  }

  /** A new node by value: its fields in the order the reader reads them. */
  #value(node: OoxmlNode): unknown[] {
    if (node.kind === 'textValue') {
      return [this.strings.string(node.id), node.value];
    }

    const base = baseNodeOf(this.#base, node.id);
    if (base !== null && base.kind !== 'textValue' && sameShell(node, base)) {
      return [this.ref(base), node.children.map((child) => this.ref(child))];
    }

    const { strings } = this;
    return [
      strings.string(node.id),
      strings.name(node.kind),
      strings.name(node.namespaceUri),
      strings.name(node.localName),
      node.prefix === undefined ? null : strings.name(node.prefix),
      node.namespaceBindings.map((b) => [strings.name(b.prefix), strings.name(b.namespaceUri)]),
      this.attributes(node.attributes),
      node.children.map((child) => this.ref(child)),
    ];
  }
}

const PART_FIELDS = new Set(['id', 'name', 'contentType', 'root']);

/** Reads nodes against the base package, in the writer's order, and checks those by value. */
export class SnapshotNodeReader {
  readonly strings: SnapshotStringReader;
  readonly #base: OoxmlPackage;
  readonly #nodes: readonly OoxmlNode[];
  readonly #baseChain = new SnapshotChain();
  readonly #listed: OoxmlNode[] = [];
  readonly #listChain = new SnapshotChain();
  /** The nodes by value that no other node holds, for the bounded reader's check. */
  readonly #byValue: OoxmlNode[] = [];

  constructor(base: OoxmlPackage) {
    this.#base = base;
    this.#nodes = baseNodes(base);
    this.strings = new SnapshotStringReader(this.#nodes, this.#baseChain);
  }

  /**
   * The node `json` names.
   *
   * @throws Invalid when it names no node or the node does not read.
   */
  node(json: unknown, path: string): OoxmlNode {
    const count = this.#listed.length;
    const node = this.#node(json, path, 0);
    if (this.#listed.length > count && this.#listed.at(-1) === node) {
      this.#byValue.push(node);
    }

    return node;
  }

  /** An element `json` names. */
  element(json: unknown, path: string): OoxmlElement {
    const node = this.node(json, path);
    if (node.kind === 'textValue') {
      throw new Invalid('invalid-type', path, 'not an element');
    }

    return node;
  }

  /**
   * A part as the checkpoint names it.
   *
   * @throws Invalid when it does not read.
   */
  part(json: unknown, path: string): OoxmlPart {
    if (typeof json !== 'object' || json === null || Array.isArray(json)) {
      const part = this.#base.parts.get(this.strings.string(json, path));
      if (part === undefined) {
        throw new Invalid('unknown-part', path);
      }

      return part;
    }

    const f = fieldsOf(json, path, PART_FIELDS);
    return {
      id: this.#id(required(f, 'id'), at(path, 'id')),
      name: this.#id(required(f, 'name'), at(path, 'name')),
      contentType: this.#id(required(f, 'contentType'), at(path, 'contentType')),
      root: this.element(required(f, 'root'), at(path, 'root')),
    };
  }

  /**
   * Attributes as the checkpoint wrote them.
   *
   * @throws Invalid when one does not read.
   */
  attributes(json: unknown, path: string): OoxmlAttribute[] {
    return this.#array(json, path).map((entry, k) => this.#attribute(entry, at(path, k)));
  }

  /**
   * Check every node read by value with the engine's bounded XML reader, the trust boundary of
   * Word XML: each must read back as the same elements and text, and each element's kind must be
   * the kind the reader gives it, or generic.
   *
   * @throws Invalid when one does not.
   */
  check(): void {
    // The reader reads equal XML alike, so each shape is checked once
    const shapes = new ShapeKeys();
    const byShape = new Map<number, OoxmlNode>();
    for (const node of this.#byValue) {
      const key = shapes.of(node);
      if (!byShape.has(key)) {
        byShape.set(key, node);
      }
    }

    const nodes = [...byShape.values()];
    if (nodes.length === 0) {
      return;
    }

    const read = readHeld(nodes);
    nodes.forEach((node, k) => {
      const shape = read?.children[k];
      if (!(shape !== undefined && sameShape(node, shape)) && !readsInContext(node)) {
        throw new Invalid('invalid-value', 'nodes', 'not as the XML reader reads it');
      }
    });
  }

  #node(json: unknown, path: string, depth: number): OoxmlNode {
    if (typeof json === 'number') {
      return this.#nodes[this.#baseChain.read(json, this.#nodes.length, path)]!;
    }

    const fields = this.#array(json, path);
    if (fields.length === 1) {
      return this.#listed[this.#listChain.read(fields[0], this.#listed.length, path)]!;
    }

    if (depth > MAX_XML_DEPTH) {
      throw new Invalid('too-many', path, 'too deep');
    }

    const node = deepFreezeNode(this.#value(fields, path, depth));
    this.#listChain.move(this.#listed.push(node) - 1);
    return node;
  }

  /** A new node's fields, read in the writer's order. */
  #value(fields: readonly unknown[], path: string, depth: number): OoxmlNode {
    const children = (json: unknown, at7: string): OoxmlNode[] =>
      this.#array(json, at7).map((entry, k) => this.#node(entry, at(at7, k), depth + 1));
    if (fields.length === 2 && typeof fields[1] === 'string') {
      return {
        id: this.#id(fields[0], at(path, 0)),
        kind: 'textValue',
        value: this.#text(fields[1], at(path, 1)),
      };
    }

    if (fields.length === 2) {
      const base = typeof fields[0] === 'number' ? this.#node(fields[0], at(path, 0), depth) : null;
      if (base === null || base.kind === 'textValue') {
        throw new Invalid('invalid-id', at(path, 0), 'no base element');
      }

      return withChildren(base, children(fields[1], at(path, 1)));
    }

    if (fields.length !== 8) {
      throw new Invalid('invalid-value', path, 'not a node');
    }

    const [id, kind, ns, localName, prefix, bindings, attributes, nodes] = fields;
    const nodeId = this.#id(id, at(path, 0));
    const kindName = this.strings.name(kind, at(path, 1));
    if (kindName === 'textValue') {
      throw new Invalid('invalid-value', path, 'not a node');
    }

    return {
      id: nodeId,
      kind: kindName,
      namespaceUri: this.strings.name(ns, at(path, 2)),
      localName: this.#name(localName, at(path, 3)),
      ...(prefix === null ? {} : { prefix: this.#name(prefix, at(path, 4)) }),
      namespaceBindings: this.#array(bindings, at(path, 5)).map((entry, k) =>
        this.#binding(entry, at(at(path, 5), k))
      ),
      attributes: this.attributes(attributes, at(path, 6)),
      children: children(nodes, at(path, 7)),
    } as OoxmlElement;
  }

  #array(value: unknown, path: string): readonly unknown[] {
    if (!Array.isArray(value)) {
      throw new Invalid('invalid-type', path, 'not a list');
    }

    return value;
  }

  /** An id: a string of 1 to {@link ID_MAX} units. */
  #id(json: unknown, path: string): string {
    const value = this.strings.string(json, path);
    if (value.length < 1 || value.length > ID_MAX) {
      throw new Invalid('invalid-text', path);
    }

    return value;
  }

  /** A text node's value, written as it is. */
  #text(value: unknown, path: string): string {
    if (typeof value !== 'string' || value.length > XML_HARD_MAX_BYTES) {
      throw new Invalid('invalid-text', path);
    }

    if (!isValidXmlText(value)) {
      throw new Invalid('invalid-text', path);
    }

    return value;
  }

  #name(json: unknown, path: string): string {
    const value = this.strings.name(json, path);
    if (!isXmlNCName(value)) {
      throw new Invalid('invalid-value', path, 'not a name');
    }

    return value;
  }

  #attribute(json: unknown, path: string): OoxmlAttribute {
    const [kind, ns, localName, prefix, text, ...rest] = this.#array(json, path);
    const kindName = typeof kind === 'number' ? ATTRIBUTE_KINDS[kind] : undefined;
    if (rest.length > 0 || !(kindName !== undefined || kind === null)) {
      throw new Invalid('invalid-value', path, 'not an attribute');
    }

    const attribute = {
      ...(kindName === undefined ? {} : { kind: kindName }),
      namespaceUri: this.strings.name(ns, at(path, 1)),
      localName: this.#name(localName, at(path, 2)),
      ...(prefix === null ? {} : { prefix: this.#name(prefix, at(path, 3)) }),
      value: this.strings.string(text, at(path, 4)),
    } as OoxmlAttribute;
    if (!attributeFits(attribute)) {
      throw new Invalid('invalid-value', path, 'not an attribute of its kind');
    }

    return attribute;
  }

  #binding(json: unknown, path: string): OoxmlNamespaceBinding {
    const [prefix, ns, ...rest] = this.#array(json, path);
    const prefixName = this.strings.name(prefix, at(path, 0));
    if (rest.length > 0 || (prefixName !== '' && !isXmlNCName(prefixName))) {
      throw new Invalid('invalid-value', path, 'not a namespace binding');
    }

    return { prefix: prefixName, namespaceUri: this.strings.name(ns, at(path, 1)) };
  }
}

/**
 * A number for each shape of a node: every field the serializer writes or `sameShape` compares,
 * which is every field but the ids and the attributes' kinds. Equal shapes have equal numbers.
 */
class ShapeKeys {
  readonly #numbers = new Map<string, number>();
  readonly #ofNode = new WeakMap<OoxmlNode, number>();

  of(node: OoxmlNode): number {
    const known = this.#ofNode.get(node);
    if (known !== undefined) {
      return known;
    }

    const text =
      node.kind === 'textValue'
        ? JSON.stringify([node.value])
        : JSON.stringify([
            node.kind,
            node.namespaceUri,
            node.localName,
            node.prefix ?? null,
            node.namespaceBindings.map((b) => [b.prefix, b.namespaceUri]),
            node.attributes.map((a) => [a.namespaceUri, a.localName, a.prefix ?? null, a.value]),
            node.children.map((child) => this.of(child)),
          ]);
    let key = this.#numbers.get(text);
    if (key === undefined) {
      key = this.#numbers.size;
      this.#numbers.set(text, key);
    }

    this.#ofNode.set(node, key);
    return key;
  }
}

/**
 * Word elements a node by value may sit in: the reader types some elements only inside one (a
 * control's properties inside `w:sdt`, a drawing inside a run).
 */
const CONTEXTS = ['r', 'p', 'sdt', 'body'];

/** The nodes as the bounded XML reader reads them back, under one holder element, or null. */
function readHeld(nodes: readonly OoxmlNode[]): OoxmlElement | null {
  const holder: OoxmlElement = {
    id: 'snapshot',
    kind: 'generic',
    namespaceUri: HOLDER_NAMESPACE,
    localName: 'nodes',
    namespaceBindings: [],
    attributes: [],
    children: nodes,
  };
  const meta = { name: '/snapshot.xml', contentType: 'application/xml' };
  try {
    const xml = serializeOoxmlPart({ id: 'part:/snapshot.xml', ...meta, root: holder });
    const read = readOoxmlPart(xml, meta, { maxBytes: XML_HARD_MAX_BYTES });
    return read.ok && read.part.root.children.length === nodes.length ? read.part.root : null;
  } catch {
    return null;
  }
}

/** Whether the reader reads `node` back as it is inside one of the Word elements it may sit in. */
function readsInContext(node: OoxmlNode): boolean {
  return CONTEXTS.some((localName) => {
    const parent: OoxmlElement = {
      id: 'context',
      kind: 'generic',
      namespaceUri: WML_NAMESPACE_URI,
      localName,
      namespaceBindings: [],
      attributes: [],
      children: [node],
    };
    const read = readHeld([parent])?.children[0];
    const shape = read?.kind === 'textValue' ? undefined : read?.children[0];
    return shape !== undefined && sameShape(node, shape);
  });
}

/** Whether an attribute's kind fits its name: `xml:space` and `w:val` are typed. */
function attributeFits(a: OoxmlAttribute): boolean {
  switch (a.kind) {
    case 'xmlSpace':
      return (
        a.namespaceUri === XML_NAMESPACE_URI &&
        a.localName === 'space' &&
        a.prefix === 'xml' &&
        (a.value === 'default' || a.value === 'preserve')
      );
    case 'wmlVal':
      return a.namespaceUri === WML_NAMESPACE_URI && a.localName === 'val';
    default:
      return true;
  }
}

/**
 * Whether `read` (the XML reader's tree) has the elements and text of `node`, and each element of
 * `node` has the reader's kind or is generic.
 */
function sameShape(node: OoxmlNode, read: OoxmlNode): boolean {
  if (node.kind === 'textValue') {
    return read.kind === 'textValue' && node.value === read.value;
  }

  if (read.kind === 'textValue') {
    return false;
  }

  if (
    node.namespaceUri !== read.namespaceUri ||
    node.localName !== read.localName ||
    (node.kind !== 'generic' && node.kind !== read.kind) ||
    node.children.length !== read.children.length
  ) {
    return false;
  }

  return node.children.every((child, k) => sameShape(child, read.children[k]!));
}

/** The first difference between two nodes in any field, or null. */
export function nodeDifference(a: OoxmlNode, b: OoxmlNode, path: string): string | null {
  if (a === b) {
    return null;
  }

  if (a.id !== b.id || a.kind !== b.kind) {
    return at(path, 'id');
  }

  if (a.kind === 'textValue' || b.kind === 'textValue') {
    const same = a.kind === 'textValue' && b.kind === 'textValue' && a.value === b.value;
    return same ? null : at(path, 'value');
  }

  const bindings = (node: OoxmlElement) =>
    node.namespaceBindings.map((x) => `${x.prefix} ${x.namespaceUri}`).join('\n');
  if (
    a.namespaceUri !== b.namespaceUri ||
    a.localName !== b.localName ||
    a.prefix !== b.prefix ||
    bindings(a) !== bindings(b)
  ) {
    return at(path, a.localName);
  }

  const attributes = attributesDifference(a.attributes, b.attributes, at(path, 'attributes'));
  if (attributes !== null) {
    return attributes;
  }

  if (a.children.length !== b.children.length) {
    return at(path, 'children');
  }

  for (let k = 0; k < a.children.length; k += 1) {
    const found = nodeDifference(a.children[k]!, b.children[k]!, at(at(path, 'children'), k));
    if (found !== null) {
      return found;
    }
  }

  return null;
}

/** The first difference between two attribute lists, in order, or null. */
export function attributesDifference(
  a: readonly OoxmlAttribute[],
  b: readonly OoxmlAttribute[],
  path: string
): string | null {
  if (a.length !== b.length) {
    return path;
  }

  for (let k = 0; k < a.length; k += 1) {
    const x = a[k]!;
    const y = b[k]!;
    if (
      x.kind !== y.kind ||
      x.namespaceUri !== y.namespaceUri ||
      x.localName !== y.localName ||
      x.prefix !== y.prefix ||
      x.value !== y.value
    ) {
      return at(path, k);
    }
  }

  return null;
}
