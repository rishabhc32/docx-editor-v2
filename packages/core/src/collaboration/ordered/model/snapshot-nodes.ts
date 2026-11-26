// The Word XML nodes of a saved item list (checkpoint version 6, design D14). A node of the base
// package is saved as its id, which names its part and its place there. Any other node is saved
// once, by value, in the checkpoint's node list, so two items that share one node share it again
// after an open. A node by value keeps every field: its kind, names, authored prefix, namespace
// declarations, attributes, children and ids. A base element's shell (the element with no
// children, as load keeps each run and text element) is saved as its id alone.

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
import { isValidXmlText } from '../../../store/package/sinks.ts';
import { MAX_XML_DEPTH, XML_HARD_MAX_BYTES } from '../../../store/package/xml-reader.ts';
import { at, fieldsOf, Invalid, required } from '../change/read.ts';
import { shellOf } from '../ooxml-nodes.ts';

/** A node as the checkpoint names it: a base node's id, or its place in the node list. */
export type SnapshotNodeRef = string | number;

/** The namespace of the element that holds the nodes by value while the XML reader checks them. */
const HOLDER_NAMESPACE = 'https://docx-editor.dev/ordered-snapshot';

/** UTF-16 units of a node id. */
const ID_MAX = 4096;

/** The base package's node with `id`, when `node` is that node. */
function baseNodeOf(base: OoxmlPackage, id: string): OoxmlNode | null {
  const at = id.indexOf('#');
  const part = at > 0 ? base.parts.get(id.slice(0, at)) : undefined;
  return part === undefined ? null : findNode(part, id);
}

/** Writes nodes: each base node as its id, each other node once in the node list. */
export class SnapshotNodeWriter {
  readonly #base: OoxmlPackage;
  readonly #listed = new Map<OoxmlNode, number>();
  readonly #namespaces = new Map<string, number>();
  readonly list: unknown[] = [];

  constructor(base: OoxmlPackage) {
    this.#base = base;
  }

  /** The namespace URIs the listed nodes name, by their index. */
  get namespaces(): string[] {
    return [...this.#namespaces.keys()];
  }

  /** How the checkpoint names `node`. */
  ref(node: OoxmlNode): SnapshotNodeRef {
    if (baseNodeOf(this.#base, node.id) === node) {
      return node.id;
    }

    let index = this.#listed.get(node);
    if (index === undefined) {
      index = this.list.push(this.#value(node)) - 1;
      this.#listed.set(node, index);
    }

    return index;
  }

  /** A part: the base package's by its name, any other by value. */
  part(part: OoxmlPart): unknown {
    if (this.#base.parts.get(part.name) === part) {
      return part.name;
    }

    return {
      id: part.id,
      name: part.name,
      contentType: part.contentType,
      root: this.ref(part.root),
    };
  }

  /** Attributes as the checkpoint writes them. */
  attributes(attributes: readonly OoxmlAttribute[]): unknown[] {
    return attributes.map((a) => [
      a.kind ?? null,
      this.#namespace(a.namespaceUri),
      a.localName,
      a.prefix ?? null,
      a.value,
    ]);
  }

  #namespace(uri: string): number {
    let index = this.#namespaces.get(uri);
    if (index === undefined) {
      index = this.#namespaces.size;
      this.#namespaces.set(uri, index);
    }

    return index;
  }

  /** A node by value; a child that is a base node or a listed node by its reference. */
  #value(node: OoxmlNode): unknown[] {
    if (node.kind === 'textValue') {
      return [node.id, node.value];
    }

    const base = baseNodeOf(this.#base, node.id);
    if (base !== null && base.kind !== 'textValue' && isShellOf(node, base)) {
      return [node.id];
    }

    const child = (c: OoxmlNode): unknown =>
      baseNodeOf(this.#base, c.id) === c ? c.id : (this.#listed.get(c) ?? this.#value(c));
    return [
      node.id,
      node.kind,
      this.#namespace(node.namespaceUri),
      node.localName,
      node.prefix ?? null,
      node.namespaceBindings.map((b) => [b.prefix, this.#namespace(b.namespaceUri)]),
      this.attributes(node.attributes),
      node.children.map(child),
    ];
  }
}

const PART_FIELDS = new Set(['id', 'name', 'contentType', 'root']);

/**
 * The attribute kinds the engine writes: the typed ones, and `generic`, which some edits give an
 * attribute they make (its readers treat it as `genericExtension`).
 */
const ATTRIBUTE_KINDS = new Set(['xmlSpace', 'wmlVal', 'genericExtension', 'generic']);

/** Reads nodes against the base package, each listed node once, and checks those by value. */
export class SnapshotNodeReader {
  readonly #base: OoxmlPackage;
  readonly #list: readonly unknown[];
  readonly #namespaces: readonly string[];
  readonly #read = new Map<number, OoxmlNode>();
  readonly #reading = new Set<number>();
  /** The nodes by value, in the order they were read, for the bounded reader's check. */
  readonly #byValue: OoxmlNode[] = [];

  constructor(base: OoxmlPackage, list: readonly unknown[], namespaces: readonly string[]) {
    this.#base = base;
    this.#list = list;
    this.#namespaces = namespaces;
  }

  /**
   * The node a reference names.
   *
   * @throws Invalid when it names no node or the node does not read.
   */
  node(value: unknown, path: string): OoxmlNode {
    if (typeof value === 'string') {
      const node = value.length <= ID_MAX ? baseNodeOf(this.#base, value) : null;
      if (node === null) {
        throw new Invalid('invalid-id', path, 'no base node');
      }

      return node;
    }

    if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0) {
      throw new Invalid('invalid-number', path);
    }

    const known = this.#read.get(value);
    if (known !== undefined) {
      return known;
    }

    if (value >= this.#list.length || this.#reading.has(value)) {
      throw new Invalid('invalid-number', path, 'no listed node');
    }

    this.#reading.add(value);
    const node = deepFreezeNode(this.#value(this.#list[value], at('nodes', value), 0));
    this.#reading.delete(value);
    this.#read.set(value, node);
    this.#byValue.push(node);
    return node;
  }

  /** An element a reference names. */
  element(value: unknown, path: string): OoxmlElement {
    const node = this.node(value, path);
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
  part(value: unknown, path: string): OoxmlPart {
    if (typeof value === 'string') {
      const part = this.#base.parts.get(value);
      if (part === undefined) {
        throw new Invalid('unknown-part', path);
      }

      return part;
    }

    const f = fieldsOf(value, path, PART_FIELDS);
    return {
      id: this.#text(required(f, 'id'), at(path, 'id'), 1, ID_MAX),
      name: this.#text(required(f, 'name'), at(path, 'name'), 1, ID_MAX),
      contentType: this.#text(required(f, 'contentType'), at(path, 'contentType'), 1, ID_MAX),
      root: this.element(required(f, 'root'), at(path, 'root')),
    };
  }

  /**
   * Attributes as the checkpoint wrote them.
   *
   * @throws Invalid when one does not read.
   */
  attributes(value: unknown, path: string): OoxmlAttribute[] {
    return this.#array(value, path).map((entry, k) => this.#attribute(entry, at(path, k)));
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

  #array(value: unknown, path: string): readonly unknown[] {
    if (!Array.isArray(value)) {
      throw new Invalid('invalid-type', path, 'not a list');
    }

    return value;
  }

  #text(value: unknown, path: string, min: number, max: number): string {
    if (typeof value !== 'string' || value.length < min || value.length > max) {
      throw new Invalid('invalid-text', path);
    }

    if (!isValidXmlText(value)) {
      throw new Invalid('invalid-text', path);
    }

    return value;
  }

  #name(value: unknown, path: string): string {
    if (typeof value !== 'string' || !isXmlNCName(value)) {
      throw new Invalid('invalid-value', path, 'not a name');
    }

    return value;
  }

  #namespace(value: unknown, path: string): string {
    const uri = typeof value === 'number' ? this.#namespaces[value] : undefined;
    if (uri === undefined || !Number.isSafeInteger(value)) {
      throw new Invalid('invalid-number', path, 'no namespace');
    }

    return uri;
  }

  #attribute(value: unknown, path: string): OoxmlAttribute {
    const [kind, ns, localName, prefix, text, ...rest] = this.#array(value, path);
    const known = typeof kind === 'string' && ATTRIBUTE_KINDS.has(kind);
    if (rest.length > 0 || !(known || kind === null)) {
      throw new Invalid('invalid-value', path, 'not an attribute');
    }

    // An attribute an edit made may have no kind
    const attribute = {
      ...(kind === null ? {} : { kind }),
      namespaceUri: this.#namespace(ns, at(path, 1)),
      localName: this.#name(localName, at(path, 2)),
      ...(prefix === null ? {} : { prefix: this.#name(prefix, at(path, 3)) }),
      value: this.#text(text, at(path, 4), 0, XML_HARD_MAX_BYTES),
    } as OoxmlAttribute;
    if (!attributeFits(attribute)) {
      throw new Invalid('invalid-value', path, 'not an attribute of its kind');
    }

    return attribute;
  }

  #value(value: unknown, path: string, depth: number): OoxmlNode {
    if (depth > MAX_XML_DEPTH) {
      throw new Invalid('too-many', path, 'too deep');
    }

    const fields = this.#array(value, path);
    if (fields.length === 1) {
      return this.#shell(fields[0], at(path, 0));
    }

    if (fields.length === 2) {
      return {
        id: this.#text(fields[0], at(path, 0), 1, ID_MAX),
        kind: 'textValue',
        value: this.#text(fields[1], at(path, 1), 0, XML_HARD_MAX_BYTES),
      };
    }

    const [id, kind, ns, localName, prefix, bindings, attributes, children, ...rest] = fields;
    if (rest.length > 0 || typeof kind !== 'string' || kind === 'textValue') {
      throw new Invalid('invalid-value', path, 'not a node');
    }

    const child = (entry: unknown, k: number): OoxmlNode =>
      typeof entry === 'string' || typeof entry === 'number'
        ? this.node(entry, at(at(path, 7), k))
        : this.#value(entry, at(at(path, 7), k), depth + 1);
    return {
      id: this.#text(id, at(path, 0), 1, ID_MAX),
      kind,
      namespaceUri: this.#namespace(ns, at(path, 2)),
      localName: this.#name(localName, at(path, 3)),
      ...(prefix === null ? {} : { prefix: this.#name(prefix, at(path, 4)) }),
      namespaceBindings: this.#array(bindings, at(path, 5)).map((entry, k) =>
        this.#binding(entry, at(at(path, 5), k))
      ),
      attributes: this.attributes(attributes, at(path, 6)),
      children: this.#array(children, at(path, 7)).map(child),
    } as OoxmlElement;
  }

  /** The shell of the base element `value` names. */
  #shell(value: unknown, path: string): OoxmlElement {
    if (typeof value !== 'string') {
      throw new Invalid('invalid-id', path, 'no base element');
    }

    return shellOf(this.element(value, path));
  }

  #binding(value: unknown, path: string): OoxmlNamespaceBinding {
    const [prefix, ns, ...rest] = this.#array(value, path);
    if (rest.length > 0 || typeof prefix !== 'string' || (prefix !== '' && !isXmlNCName(prefix))) {
      throw new Invalid('invalid-value', path, 'not a namespace binding');
    }

    return { prefix, namespaceUri: this.#namespace(ns, at(path, 1)) };
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

/** Whether `node` is the shell of `base` in every field. */
const isShellOf = (node: OoxmlNode, base: OoxmlElement): boolean =>
  node.kind !== 'textValue' &&
  node.children.length === 0 &&
  nodeDifference(node, shellOf(base), '') === null;

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
