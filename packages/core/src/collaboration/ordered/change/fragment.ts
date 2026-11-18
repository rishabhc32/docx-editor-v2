// A pasted fragment in the change format (group 9.10): WordprocessingML blocks as element trees
// with bounded names and values, never as XML text.

import {
  FRAGMENT_LINK_OWNERS,
  FRAGMENT_OBJECT_ATTRIBUTES,
  FRAGMENT_OBJECT_ELEMENTS,
  fragmentAttributeNamespace,
  fragmentElementNamespace,
  isVmlPrefix,
  isFragmentElement,
  isFragmentMediaReference,
  isFragmentPictureAddress,
  isFragmentRequires,
  isFragmentTextElement,
  takesPlainAttributes,
} from './fragment-names.ts';
import {
  DIAGRAM_DRAWING_REFERENCE,
  FRAGMENT_PART_CHILDREN,
  FRAGMENT_PART_ROOTS,
  FRAGMENT_PART_TYPES,
  fragmentObjectReference,
  fragmentPartReference,
  prefixOf,
  type OrderedFragmentPartType,
} from './fragment-names-objects.ts';
import { at, fieldsOf, int, Invalid, list, oneOf, plainText, required, str } from './read.ts';
import { href } from './values.ts';

/**
 * One element of a pasted fragment.
 *
 * @public
 */
export interface OrderedFragmentElement {
  /** The qualified name: a fragment namespace prefix and a local name on its allowlist (`w:p`). */
  readonly e: string;
  /**
   * Attributes by qualified name (`w:val`, `xml:space`, `r:id`), or by local name alone on a
   * drawing's elements (`cx`), values plain text.
   */
  readonly a?: Readonly<Record<string, string>>;
  /** Children: elements, or text (the characters of a `w:t`, a `w:instrText`, an `m:t`). */
  readonly c?: readonly (OrderedFragmentElement | string)[];
}

/**
 * A list definition a paste adds: its `w:abstractNum` and the `w:lvlOverride`s of its
 * `w:num`, both without their ids, which the apply mints.
 *
 * @public
 */
export interface OrderedFragmentList {
  readonly abstract: OrderedFragmentElement;
  readonly overrides?: readonly OrderedFragmentElement[];
}

/**
 * A footnote or an endnote a paste adds: its kind and its blocks. A reference in the blocks
 * names it as `#<k>`.
 *
 * @public
 */
export interface OrderedFragmentNote {
  readonly kind: 'footnote' | 'endnote';
  readonly blocks: readonly OrderedFragmentElement[];
}

/**
 * The caps on a fragment in one change. Client and server share them; the client cuts a long
 * paste into several changes at paragraph edges.
 *
 * @public
 */
export const ORDERED_FRAGMENT_LIMITS = Object.freeze({
  /** Elements and text values in one change, every tree counted. */
  nodes: 16_384,
  /** Depth of one tree. */
  depth: 48,
  /** UTF-16 units of all text and attribute values in one change. */
  text: 262_144,
  /** UTF-16 units of one attribute value. */
  attributeValue: 2048,
  /** Attributes of one element. */
  attributes: 32,
  /** Blocks in one change. */
  blocks: 4096,
  /** List definitions, styles, links and notes in one change, each. */
  resources: 256,
});

/**
 * A relationship of a pasted chart's or diagram's part: to a later part of the change
 * (`parts[part]`), or to its media (`media[media]`: a picture, or a chart's embedded workbook).
 *
 * @public
 */
export type OrderedFragmentPartRelationship =
  | { readonly part: number }
  | { readonly media: number };

/**
 * A part a pasted chart or diagram needs: its type, its root as a checked tree, and its own
 * relationships, which the tree's `r:` attributes name as `#<j>`.
 *
 * @public
 */
export interface OrderedFragmentPart {
  readonly type: OrderedFragmentPartType;
  readonly root: OrderedFragmentElement;
  readonly rels?: readonly OrderedFragmentPartRelationship[];
}

/** What a fragment's trees are: blocks of a story, a resource's element, or an object's part. */
export type OrderedFragmentRoot =
  | 'block'
  | 'abstractNum'
  | 'lvlOverride'
  | 'numPicBullet'
  | 'style'
  | 'part';

/**
 * What the reader allows of a fragment's references, for one change. A reference outside the
 * story text is never a name of the sender's package: a list number, a link's relationship, a
 * picture's image, a note's id and an object's part are `#<k>`, an index into the change's own
 * lists, and the apply mints the Word names (design D14). Media travel by hash, never fetched.
 */
export interface OrderedFragmentReferences {
  readonly lists: number;
  readonly links: number;
  readonly notes: number;
  readonly media: number;
  /** The change's chart and diagram parts. */
  readonly parts?: number;
  /** The change's picture bullets, which a list level's `w:lvlPicBulletId` names. */
  readonly bullets?: number;
  /** In a part's tree: the number of the part's own relationships. */
  readonly own?: number;
}

/** A budget shared by every tree of one change. */
export interface OrderedFragmentBudget {
  nodes: number;
  text: number;
}

/** A new budget for one change. */
export function fragmentBudget(): OrderedFragmentBudget {
  return { nodes: 0, text: 0 };
}

const LIMITS = ORDERED_FRAGMENT_LIMITS;

/** What the trees of one root may be. */
const ROOTS: ReadonlyMap<OrderedFragmentRoot, ReadonlySet<string>> = new Map([
  ['block', new Set(['w:p', 'w:tbl', 'w:sdt'])],
  ['abstractNum', new Set(['w:abstractNum'])],
  ['lvlOverride', new Set(['w:lvlOverride'])],
  ['numPicBullet', new Set(['w:numPicBullet'])],
  ['style', new Set(['w:style'])],
  ['part', new Set(Object.values(FRAGMENT_PART_ROOTS))],
]);

const ELEMENT_FIELDS: ReadonlySet<string> = new Set(['e', 'a', 'c']);

/** `<prefix>:<localName>`: a short prefix, and a local name of letters and digits. */
const QUALIFIED_NAME = /^([a-z][a-z0-9]{0,4}):([A-Za-z][A-Za-z0-9]{0,63})$/;

/** A local name alone: an attribute of a drawing's element. */
const PLAIN_NAME = /^[A-Za-z][A-Za-z0-9]{0,63}$/;

/** `#<k>`: an index into one of the change's lists, without leading zeros. */
const REFERENCE = /^#(0|[1-9][0-9]{0,8})$/;

/** A list number the document defines. */
const DECIMAL = /^[0-9]{1,9}$/;

const NOTE_REFERENCES: ReadonlySet<string> = new Set(['w:footnoteReference', 'w:endnoteReference']);

interface Reading {
  readonly references: OrderedFragmentReferences;
  readonly budget: OrderedFragmentBudget;
}

/**
 * Read and bound fragment trees: a new tree with known fields only, every name on its
 * allowlist, every value plain text, every reference in range, within `budget`. Throws
 * `Invalid` by name. The apply reads the trees again through the engine's bounded part reader
 * (`edits/fragment-nodes.ts`), the trust boundary every Word file passes.
 */
export function fragmentElements(
  value: unknown,
  path: string,
  root: OrderedFragmentRoot,
  references: OrderedFragmentReferences,
  budget: OrderedFragmentBudget
): OrderedFragmentElement[] {
  const roots = ROOTS.get(root);
  if (roots === undefined) throw new Invalid('invalid-value', path, 'not a fragment root');
  const trees = list(value, path, 1, LIMITS.blocks);
  const reading: Reading = { references, budget };
  const out: OrderedFragmentElement[] = [];
  for (let k = 0; k < trees.length; k += 1) {
    const where = at(path, k);
    out.push(element(itemOf(trees, k, where), where, 1, reading, roots));
  }

  return out;
}

/** Read and bound one fragment tree at its own path, as {@link fragmentElements} reads each. */
export function fragmentElement(
  value: unknown,
  path: string,
  root: OrderedFragmentRoot,
  references: OrderedFragmentReferences,
  budget: OrderedFragmentBudget
): OrderedFragmentElement {
  const roots = ROOTS.get(root);
  if (roots === undefined) throw new Invalid('invalid-value', path, 'not a fragment root');
  return element(value, path, 1, { references, budget }, roots);
}

/** An array's own data item: never a getter, never a hole. */
function itemOf(items: readonly unknown[], k: number, path: string): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(items, k);
  if (descriptor === undefined || !('value' in descriptor))
    throw new Invalid('invalid-type', path, 'not a value');
  return descriptor.value;
}

/** One more element or text in the change. */
function spendNode(reading: Reading, path: string): void {
  reading.budget.nodes += 1;
  if (reading.budget.nodes > LIMITS.nodes)
    throw new Invalid('too-many', path, `over ${LIMITS.nodes} fragment nodes`);
}

/** `units` more of text in the change. */
function spendText(reading: Reading, units: number, path: string): void {
  reading.budget.text += units;
  if (reading.budget.text > LIMITS.text)
    throw new Invalid('too-long', path, `over ${LIMITS.text} units of fragment text`);
}

function element(
  value: unknown,
  path: string,
  depth: number,
  reading: Reading,
  roots?: ReadonlySet<string>
): OrderedFragmentElement {
  if (depth > LIMITS.depth) throw new Invalid('too-many', path, `deeper than ${LIMITS.depth}`);
  const fields = fieldsOf(value, path, ELEMENT_FIELDS);
  spendNode(reading, path);
  const e = elementName(required(fields, 'e'), at(path, 'e'));
  if (roots !== undefined && !roots.has(e))
    throw new Invalid('invalid-value', at(path, 'e'), `not ${[...roots].join(', ')}`);
  const out: {
    e: string;
    a?: Record<string, string>;
    c?: (OrderedFragmentElement | string)[];
  } = { e };
  if (fields.has('a')) out.a = attributes(fields.get('a'), at(path, 'a'), e, reading);
  if (fields.has('c')) out.c = children(fields.get('c'), at(path, 'c'), e, depth, reading);
  return out;
}

function elementName(value: unknown, path: string): string {
  const name = str(value, path, 80);
  const match = QUALIFIED_NAME.exec(name);
  if (match === null) throw new Invalid('invalid-value', path, 'not a qualified name');
  const prefix = match[1]!;
  const localName = match[2]!;
  if (fragmentElementNamespace(prefix) === undefined)
    throw new Invalid('invalid-namespace', path, 'not a fragment prefix');
  if (FRAGMENT_OBJECT_ELEMENTS.has(name)) throw new Invalid('invalid-value', path, 'object');
  if (!isFragmentElement(prefix, localName))
    throw new Invalid('invalid-value', path, `not an element of ${prefix}`);
  return name;
}

function attributes(
  value: unknown,
  path: string,
  owner: string,
  reading: Reading
): Record<string, string> {
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    throw new Invalid('invalid-type', path, 'not an object');
  const keys = Object.keys(value);
  if (keys.length === 0) throw new Invalid('empty', path);
  if (keys.length > LIMITS.attributes)
    throw new Invalid('too-many', path, `over ${LIMITS.attributes}`);

  // Checks the object's prototype and every key against `__proto__` and its kin.
  const fields = fieldsOf(value, path, new Set(keys));
  const vml = isVmlPrefix(prefixOf(owner));
  const out: Record<string, string> = {};
  for (const key of keys) {
    const where = at(path, key);
    const name = attributeName(key, where, owner);
    const raw = fields.get(key);
    if (typeof raw !== 'string') throw new Invalid('invalid-type', where, 'not a string');
    const text =
      key === 'o:gfxdata'
        ? base64Data(raw, where)
        : plainText(raw, where, 0, vml ? LIMITS.text : LIMITS.attributeValue);
    spendText(reading, text.length, where);
    attributeValue(owner, name, text, where, reading.references);
    out[key] = text;
  }

  return out;
}

/** An attribute's prefix (`''` for none) and local name. */
interface AttributeName {
  readonly prefix: string;
  readonly localName: string;
}

/**
 * An attribute's prefix and local name, from a key the writer may write as it is. A drawing's
 * element takes local names alone (prefix `''`), as DrawingML writes them.
 */
function attributeName(key: string, path: string, owner: string): AttributeName {
  if (key === 'xmlns' || key.startsWith('xmlns:'))
    throw new Invalid('invalid-attribute', path, 'a namespace declaration');
  if (PLAIN_NAME.test(key)) {
    const prefix = prefixOf(owner);
    if (!takesPlainAttributes(owner)) throw new Invalid('invalid-attribute', path, 'no prefix');

    // Markup compatibility has one attribute of its own: the prefixes a choice requires
    if (prefix === 'mc' && (owner !== 'mc:Choice' || key !== 'Requires'))
      throw new Invalid('invalid-attribute', path, 'Requires of mc:Choice only');
    return { prefix: '', localName: key };
  }

  const match = QUALIFIED_NAME.exec(key);
  if (match === null) throw new Invalid('invalid-attribute', path, 'not a qualified name');
  const prefix = match[1]!;
  const localName = match[2]!;
  if (fragmentAttributeNamespace(prefix) === undefined)
    throw new Invalid('invalid-namespace', path, 'not a fragment attribute prefix');
  if (prefix === 'xml' && localName !== 'space')
    throw new Invalid('invalid-attribute', path, 'xml:space only');
  return { prefix, localName };
}

/**
 * The rules on one attribute's value: `xml:space`, a picture by address, and every reference
 * outside the story.
 */
function attributeValue(
  owner: string,
  name: AttributeName,
  value: string,
  path: string,
  references: OrderedFragmentReferences
): void {
  if (name.prefix === 'xml') {
    if (value !== 'preserve' && value !== 'default') {
      throw new Invalid('invalid-value', path, 'not preserve or default');
    }

    return;
  }

  const key = name.prefix === '' ? name.localName : `${name.prefix}:${name.localName}`;

  // A picture by address: Word loads it when it opens the file, so no room writes one (D18)
  if (isFragmentPictureAddress(owner, key)) {
    throw new Invalid('invalid-value', path, 'a picture by address');
  }

  if (partAttribute(owner, name, key, value, path, references)) {
    return;
  }

  if (name.prefix === 'r') {
    relationshipAttribute(owner, name, key, value, path, references);
    return;
  }

  if (drawingAttribute(owner, name, key, value, path, references)) {
    return;
  }

  if (name.prefix === 'w') {
    wordAttribute(owner, name, value, path, references);
  }
}

/**
 * In a chart's or a diagram's part: a relationship is `#<j>` into the part's own list (an empty
 * one names nothing), and a diagram's drawing is `#<k>` into the change's parts. True when the
 * attribute is one of them.
 */
function partAttribute(
  owner: string,
  name: AttributeName,
  key: string,
  value: string,
  path: string,
  references: OrderedFragmentReferences
): boolean {
  if (references.own === undefined) {
    return false;
  }

  if (name.prefix === 'r') {
    if (value !== '') {
      reference(value, references.own, path, 'relationship');
    }

    return true;
  }

  if (`${owner}@${key}` === DIAGRAM_DRAWING_REFERENCE) {
    reference(value, references.parts ?? 0, path, 'part');
    return true;
  }

  return false;
}

/**
 * An `r:` attribute in a block: a link's relationship, a picture's image, an embedded object's
 * part and an ink part (media, whose type the change kind checks) and a chart's or a diagram's
 * part travel as indexes. Any other relationship names an object the change does not carry.
 */
function relationshipAttribute(
  owner: string,
  name: AttributeName,
  key: string,
  value: string,
  path: string,
  references: OrderedFragmentReferences
): void {
  if (FRAGMENT_LINK_OWNERS.has(owner) && name.localName === 'id') {
    reference(value, references.links, path, 'link');
  } else if (isFragmentMediaReference(owner, key)) {
    reference(value, references.media, path, 'picture');
  } else if (fragmentObjectReference(owner, key) !== undefined) {
    reference(value, references.media, path, 'object');
  } else if (fragmentPartReference(owner, key) !== undefined) {
    reference(value, references.parts ?? 0, path, 'part');
  } else {
    throw new Invalid('invalid-value', path, 'object');
  }
}

/**
 * The rules on an attribute of a drawing, VML or markup compatibility: a VML picture's image, an
 * object by address, the prefixes a choice requires, a VML shape's click link and a drawing's
 * document-wide id. True when the attribute is done with.
 */
function drawingAttribute(
  owner: string,
  name: AttributeName,
  key: string,
  value: string,
  path: string,
  references: OrderedFragmentReferences
): boolean {
  // A VML picture's image by the Office relationship id older Word writes
  if (isFragmentMediaReference(owner, key)) {
    reference(value, references.media, path, 'picture');
    return true;
  }

  // An object by address, never fetched
  if (FRAGMENT_OBJECT_ATTRIBUTES.has(key)) {
    throw new Invalid('invalid-value', path, 'object');
  }

  if (owner === 'mc:Choice' && key === 'Requires') {
    if (!isFragmentRequires(value)) {
      throw new Invalid('invalid-value', path, 'not fragment prefixes');
    }

    return true;
  }

  // A VML shape's click link: an address `sanitizeHref` keeps as it is
  if (key === 'href' && isVmlPrefix(prefixOf(owner))) {
    href(value, path, LIMITS.attributeValue);
    return true;
  }

  // A drawing's document-wide id: the apply mints its own
  if (owner === 'wp:docPr' && name.localName === 'id' && !DECIMAL.test(value)) {
    throw new Invalid('invalid-value', path, 'not a decimal');
  }

  return false;
}

/** A `w:` attribute that names a list, a picture bullet or a note of the change. */
function wordAttribute(
  owner: string,
  name: AttributeName,
  value: string,
  path: string,
  references: OrderedFragmentReferences
): void {
  if (owner === 'w:numId' && name.localName === 'val') {
    if (!DECIMAL.test(value)) {
      reference(value, references.lists, path, 'list');
    }
  } else if (owner === 'w:lvlPicBulletId' && name.localName === 'val') {
    reference(value, references.bullets ?? 0, path, 'picture bullet');
  } else if (owner === 'w:numPicBullet' && name.localName === 'numPicBulletId') {
    // A picture bullet's id: the apply mints it
    throw new Invalid('invalid-attribute', path, 'the apply gives the id');
  } else if (NOTE_REFERENCES.has(owner) && name.localName === 'id') {
    reference(value, references.notes, path, 'note');
  }
}

/** `o:gfxdata`: Word's DrawingML copy of a VML shape, base64 in lines, within the text cap. */
function base64Data(value: string, path: string): string {
  if (value.length > LIMITS.text) throw new Invalid('too-long', path, `over ${LIMITS.text}`);
  if (!/^[A-Za-z0-9+/=\r\n]*$/.test(value)) throw new Invalid('invalid-text', path, 'not base64');
  return value;
}

function reference(value: string, count: number, path: string, what: string): void {
  const match = REFERENCE.exec(value);
  if (match === null || Number(match[1]) >= count)
    throw new Invalid('invalid-value', path, `not a ${what} of the change`);
}

function children(
  value: unknown,
  path: string,
  owner: string,
  depth: number,
  reading: Reading
): (OrderedFragmentElement | string)[] {
  const items = list(value, path, 1, LIMITS.nodes);
  const textual = isFragmentTextElement(owner);
  const out: (OrderedFragmentElement | string)[] = [];
  for (let k = 0; k < items.length; k += 1) {
    const where = at(path, k);
    const item = itemOf(items, k, where);
    if (typeof item !== 'string') {
      out.push(element(item, where, depth + 1, reading));
      continue;
    }

    if (!textual) throw new Invalid('invalid-type', where, `text in ${owner}`);

    // One text per run of characters: the reader joins two texts in a row into one.
    if (typeof out[out.length - 1] === 'string')
      throw new Invalid('invalid-value', where, 'a second text in a row');
    spendNode(reading, where);
    const text = plainText(item, where, 1, LIMITS.text);
    spendText(reading, text.length, where);
    out.push(text);
  }

  return out;
}

const PART_FIELDS: ReadonlySet<string> = new Set(['type', 'root', 'rels']);
const PART_RELATIONSHIP_FIELDS = { part: new Set(['part']), media: new Set(['media']) };

/**
 * Part `k`'s own relationships: each to a later part of a type its own may name (so they never
 * loop), or to the change's media (`media` of them).
 */
function partRelationships(
  value: unknown,
  path: string,
  k: number,
  types: readonly OrderedFragmentPartType[],
  media: number
): OrderedFragmentPartRelationship[] {
  const type = types[k]!;
  const rels: OrderedFragmentPartRelationship[] = [];
  const entries = list(value, path, 1, LIMITS.attributes);
  entries.forEach((entry, j) => {
    const relPath = at(path, j);
    if (typeof entry === 'object' && entry !== null && 'part' in entry) {
      const target = int(
        required(fieldsOf(entry, relPath, PART_RELATIONSHIP_FIELDS.part), 'part'),
        at(relPath, 'part'),
        k + 1,
        types.length - 1
      );
      if (!FRAGMENT_PART_CHILDREN[type].includes(types[target]!)) {
        throw new Invalid('invalid-value', at(relPath, 'part'), `not a part of a ${type}`);
      }

      rels.push({ part: target });
    } else {
      const index = int(
        required(fieldsOf(entry, relPath, PART_RELATIONSHIP_FIELDS.media), 'media'),
        at(relPath, 'media'),
        0,
        media - 1
      );
      rels.push({ media: index });
    }
  });
  return rels;
}

/**
 * Read a change's chart and diagram parts (SmartArt too), which blocks name as `#<k>`
 * (`c:chart/@r:id`, `dgm:relIds/@r:dm`): each a checked tree of its type's root, whose
 * relationships name a LATER part of a type its own may name (so they never loop), or media
 * (a part's media are checked against the media's types by the change kind). `references`
 * gives the change's counts; `budget` is the change's.
 */
export function fragmentParts(
  value: unknown,
  path: string,
  references: OrderedFragmentReferences,
  budget: OrderedFragmentBudget
): OrderedFragmentPart[] {
  const items = list(value, path, 1, LIMITS.resources);

  // Every part's type first: a relationship names a later part of a type its own may name
  const read = items.map((item, k) => {
    const own = fieldsOf(item, at(path, k), PART_FIELDS);
    const type = oneOf(required(own, 'type'), at(at(path, k), 'type'), FRAGMENT_PART_TYPES);
    return { own, type };
  });
  const types = read.map((entry) => entry.type);
  return read.map(({ own, type }, k): OrderedFragmentPart => {
    const where = at(path, k);
    const rels = own.has('rels')
      ? partRelationships(own.get('rels'), at(where, 'rels'), k, types, references.media)
      : [];
    const reading: Reading = {
      references: { ...references, parts: types.length, own: rels.length },
      budget,
    };
    const root = element(required(own, 'root'), at(where, 'root'), 1, reading, ROOTS.get('part'));
    if (root.e !== FRAGMENT_PART_ROOTS[type]) {
      throw new Invalid('invalid-value', at(where, 'root'), `not the root of a ${type}`);
    }

    return rels.length > 0 ? { type, root, rels } : { type, root };
  });
}
