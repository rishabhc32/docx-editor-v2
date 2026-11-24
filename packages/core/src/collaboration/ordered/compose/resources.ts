// Resources outside the story text, written into a composed package (design D14) when their
// suggestion is pending or accepted and something written uses them. Every copy writes the same
// package: names come from change ids, and elements are built from typed values.

import { withPart, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { isRelationshipReference } from '../../../store/package/canonical-media-references.ts';
import { ensureNumberingLevel } from '../../../store/package/numbering-part.ts';
import {
  readOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlAttribute,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import { withBinaryPart } from '../../../store/package/drawing-package-edit.ts';
import { cloneOwnedRelationships } from '../../../store/package/hf-lifecycle-shell.ts';
import {
  relsPartNameFor,
  withContentTypeOverride,
  withNewPart,
  withRelationship,
  withRelationshipsPartFor,
} from '../../../store/package/package-edit.ts';
import { resolveRelationship } from '../../../store/package/relationships.ts';
import { writeNoteSeparatorReferences } from '../../../store/package/note-lifecycle-props.ts';
import {
  mergedPropertyChildren,
  PARAGRAPH_VOCABULARY,
  RUN_VOCABULARY,
} from '../../../store/store/tree-op-properties.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import { withRequiredNamespaceBindings } from '../../../store/store/tree-op-fragment.ts';
import {
  FRAGMENT_OBJECT_RELATIONSHIP_TYPES,
  FRAGMENT_PART_CONTENT_TYPES,
  FRAGMENT_PART_RELATIONSHIP_TYPES,
} from '../change/fragment-names-objects.ts';
import { customNodeBindingTarget } from '../../../store/package/custom-node-payloads.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import {
  orderedResourceShows,
  type OrderedListBullet,
  type OrderedListResource,
  type OrderedMediaReference,
  type OrderedMediaResource,
  type OrderedObjectPartResource,
  type OrderedRelationshipResource,
  type OrderedResource,
  type OrderedStyleResource,
} from '../model/resources.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  attributeOf,
  eachElement,
  isWml,
  withAttributes,
  withChildren,
  wmlAttribute,
} from '../ooxml-nodes.ts';
import { customXmlBindingKey, hasCustomXml, withCustomXml } from './custom-xml.ts';
import { withSettings } from './settings.ts';

export const OFFICE_RELATIONSHIPS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELATIONSHIPS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const WML_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const RELATIONSHIP_TYPES: Readonly<Record<OrderedRelationshipResource['type'], string>> = {
  image: `${OFFICE_RELATIONSHIPS}/image`,
  hyperlink: `${OFFICE_RELATIONSHIPS}/hyperlink`,
  header: `${OFFICE_RELATIONSHIPS}/header`,
  footer: `${OFFICE_RELATIONSHIPS}/footer`,
  footnotes: `${OFFICE_RELATIONSHIPS}/footnotes`,
  endnotes: `${OFFICE_RELATIONSHIPS}/endnotes`,
  oleObject: FRAGMENT_OBJECT_RELATIONSHIP_TYPES.oleObject,
  ink: FRAGMENT_OBJECT_RELATIONSHIP_TYPES.ink,
  ...FRAGMENT_PART_RELATIONSHIP_TYPES,
};

const STORY_TYPES = {
  header: `${WML_TYPE}.header+xml`,
  footer: `${WML_TYPE}.footer+xml`,
  footnotes: `${WML_TYPE}.footnotes+xml`,
  endnotes: `${WML_TYPE}.endnotes+xml`,
} as const;

/** How a composed package gets the bytes of media. @public */
export interface OrderedResourceOptions {
  /** The bytes of media by SHA-256, as the room's storage or the client's cache holds them. */
  readonly blobs?: (sha256: string) => Uint8Array | undefined;
  /**
   * Media whose bytes are missing: refuse the package (`missing-blob`, the default: a save
   * must hold every picture), or leave the media and its relationship out (a screen, until the
   * bytes arrive).
   */
  readonly missing?: 'refuse' | 'skip';
}

const plain = (localName: string, value: string): OoxmlAttribute => ({
  kind: 'genericExtension',
  namespaceUri: '',
  localName,
  value,
});

/** What a part uses: relationship ids, style ids, list numbers and custom nodes' bindings. */
interface PartUses {
  readonly relationships: Set<string>;
  readonly styles: Set<string>;
  readonly lists: Set<string>;
  /** The nodes its controls bind (`customXmlBindingKey`). */
  readonly bindings: Set<string>;
}

const STYLE_NAMES: ReadonlySet<string> = new Set(['pStyle', 'rStyle', 'tblStyle']);

/** What one node and everything under it uses, read in one bounded walk. */
function usesOfNode(root: OoxmlNode, into: PartUses): void {
  eachElement(root, (node) => {
    for (const entry of node.attributes)
      if (isRelationshipReference(entry)) into.relationships.add(entry.value);
    if (node.namespaceUri === WML_NAMESPACE_URI) {
      const value = attributeOf(node, 'val');
      if (value !== undefined && STYLE_NAMES.has(node.localName)) into.styles.add(value);
      if (value !== undefined && node.localName === 'numId') into.lists.add(value);
      if (node.localName === 'dataBinding') bindingOf(node, into.bindings);
    }
  });
}

/** The node a control's `w:dataBinding` names, when it has the engine's own shape. */
function bindingOf(node: OoxmlElement, into: Set<string>): void {
  const storeItemId = attributeOf(node, 'storeItemID');
  const xpath = attributeOf(node, 'xpath');
  const prefixMappings = attributeOf(node, 'prefixMappings');
  if (storeItemId === undefined || xpath === undefined || prefixMappings === undefined) return;
  const target = customNodeBindingTarget({ prefixMappings, xpath });
  if (target) into.add(customXmlBindingKey(storeItemId, target.nodeId));
}

const empty = (): PartUses => ({
  relationships: new Set(),
  styles: new Set(),
  lists: new Set(),
  bindings: new Set(),
});

/**
 * What each block uses, kept while the block is the same object: the publisher keeps every
 * block a key press does not touch, so a publish reads only the blocks that changed.
 */
const BLOCK_USES = new WeakMap<OoxmlNode, PartUses>();

/** What a part uses, read block by block. */
function usesOf(part: OoxmlPart): PartUses {
  const out = empty();
  const shell = withChildren(part.root, []);
  usesOfNode(shell, out);
  const body = part.root.children.find((child) => isWml(child, 'body'));
  const blocks: OoxmlNode[] = [];
  for (const child of part.root.children)
    if (child === body && child.kind !== 'textValue') {
      usesOfNode(withChildren(child, []), out);
      blocks.push(...child.children);
    } else blocks.push(child);
  for (const block of blocks) {
    let uses = BLOCK_USES.get(block);
    if (!uses) {
      uses = empty();
      usesOfNode(block, uses);
      BLOCK_USES.set(block, uses);
    }

    for (const id of uses.relationships) out.relationships.add(id);
    for (const id of uses.styles) out.styles.add(id);
    for (const id of uses.lists) out.lists.add(id);
    for (const key of uses.bindings) out.bindings.add(key);
  }

  return out;
}

/** The custom nodes the written stories' controls bind. */
function boundNodes(stories: readonly OoxmlPart[]): Set<string> {
  const bound = new Set<string>();
  for (const story of stories) for (const key of usesOf(story).bindings) bound.add(key);
  return bound;
}

/**
 * The resources by key: the same order on every copy, whatever order the changes arrived in.
 * Settings keep the room's order instead (`withSettings`).
 */
function byKey(doc: OrderedDocument): OrderedResource[] {
  return [...doc.resources.added.values()].sort(compareKeys);
}

/** Two resources by key, code unit by code unit (not `localeCompare`: every copy sorts alike). */
function compareKeys(a: OrderedResource, b: OrderedResource): number {
  if (a.key < b.key) return -1;
  if (a.key > b.key) return 1;
  return 0;
}

/** A relationship target from its owner: relative within the owner's folder, else absolute. */
export function targetFrom(owner: string, part: string): string {
  const folder = owner.slice(0, owner.lastIndexOf('/') + 1);
  return part.startsWith(folder) ? part.slice(folder.length) : part;
}

/** What a package writes of the resources: the keys that show and are used. */
interface Written {
  readonly keys: Set<string>;
  /** The parts the written stories are: the loaded ones and the new ones written. */
  readonly stories: OoxmlPart[];
}

const DIAGRAM_DRAWING_NS = 'http://schemas.microsoft.com/office/drawing/2008/diagram';

/** The relationship ids a diagram's data part names for its drawing (`dsp:dataModelExt`). */
function diagramDrawingIds(root: OoxmlNode, into: Set<string>): void {
  eachElement(root, (node) => {
    if (node.namespaceUri === DIAGRAM_DRAWING_NS && node.localName === 'dataModelExt')
      for (const entry of node.attributes)
        if (entry.namespaceUri === '' && entry.localName === 'relId') into.add(entry.value);
  });
}

/**
 * A link's external address: written while a written story names it (`isRelationshipUsed`),
 * whatever its suggestion's status, as Word keeps a rejected new link that still holds a
 * comment's start (`kept-links.ts`).
 */
function isLinkAddress(resource: OrderedResource): boolean {
  return (
    resource.kind === 'relationship' &&
    resource.type === 'hyperlink' &&
    'external' in resource.target
  );
}

/** A new notes part, or the main document's relationship to one (group 9.9). */
function isNotesResource(resource: OrderedResource): boolean {
  if (resource.kind === 'story')
    return resource.role === 'footnotes' || resource.role === 'endnotes';
  if (resource.kind === 'relationship')
    return resource.type === 'footnotes' || resource.type === 'endnotes';
  return false;
}

/**
 * Whether a relationship is used: its owner part is written and names its id (`r:id`, `r:embed`,
 * ...), or it points at a new notes part (Word finds a notes part by the relationship's type, not
 * by an id). A new notes part stays with its separators once its last note goes, as Word and the
 * editor alone keep it.
 */
function isRelationshipUsed(
  relationship: OrderedRelationshipResource,
  ids: ReadonlySet<string>,
  target: OrderedResource | undefined
): boolean {
  if (ids.has(relationship.id)) return true;
  return target?.kind === 'story' && isNotesResource(target);
}

/**
 * The chart and diagram parts that show, by part name, and the relationship ids their diagrams'
 * data parts name for their drawings (`dsp:dataModelExt`).
 */
function objectPartsOf(shown: Iterable<OrderedResource>): {
  readonly objects: Map<string, OrderedObjectPartResource>;
  readonly drawings: Set<string>;
} {
  const objects = new Map<string, OrderedObjectPartResource>();
  const drawings = new Set<string>();
  for (const resource of shown) {
    if (resource.kind !== 'part') continue;
    objects.set(resource.part, resource);
    if (resource.type === 'diagramData') diagramDrawingIds(resource.root, drawings);
  }

  return { objects, drawings };
}

/**
 * Which resources are written, from the loaded stories out along used relationships: a media
 * part, a new story part or an object part when a written relationship points at it, whose own
 * relationships are then read in turn.
 */
function reach(
  pkg: OoxmlPackage,
  doc: OrderedDocument,
  added: ReadonlyMap<string, OoxmlPart>,
  bytes: (media: OrderedMediaResource) => Uint8Array | undefined
): Written {
  const shown = new Map<string, OrderedResource>();

  // A new notes part is the notes' in it, not its first note's: it shows while the list holds it.
  // A link's address shows while a written link names it, whatever its suggestion's status.
  for (const resource of byKey(doc))
    if (orderedResourceShows(doc, resource) || isNotesResource(resource) || isLinkAddress(resource))
      shown.set(resource.key, resource);
  const relationships = [...shown.values()].filter(
    (resource): resource is OrderedRelationshipResource => resource.kind === 'relationship'
  );
  const keys = new Set<string>();
  const stories: OoxmlPart[] = [];

  // A chart's or a diagram's part names relationships of its own, read as a story's are. A
  // diagram's drawing is the story's, named by its data part's `relId`
  const { objects, drawings } = objectPartsOf(shown.values());
  const queue = [...doc.stories.keys()].filter((name) => !added.has(name));
  const seen = new Set(queue);
  for (let owner = queue.pop(); owner !== undefined; owner = queue.pop()) {
    const object = objects.get(owner);
    const part =
      object !== undefined
        ? { id: owner, name: owner, contentType: '', root: object.root }
        : (added.get(owner) ?? pkg.parts.get(owner));
    if (!part) continue;
    if (object === undefined) stories.push(part);

    // Every relationship of an object part is used (Word names a chart's style by type)
    const ids = object === undefined ? usesOf(part).relationships : null;
    for (const relationship of relationships) {
      if (relationship.owner !== owner) continue;
      const pointsAt = 'resource' in relationship.target ? relationship.target.resource : undefined;
      const target = pointsAt === undefined ? undefined : shown.get(pointsAt);
      const used = ids === null || isRelationshipUsed(relationship, ids, target);
      if (!used && !drawings.has(relationship.id)) continue;
      if (pointsAt !== undefined) {
        // A target that does not show takes its relationship with it
        if (!target) continue;
        if (target.kind === 'media' && bytes(target) === undefined) continue;
        keys.add(target.key);
        if ((target.kind === 'story' || target.kind === 'part') && !seen.has(target.part)) {
          seen.add(target.part);
          queue.push(target.part);
        }
      }

      keys.add(relationship.key);
    }
  }

  return { keys, stories };
}

/** The new story parts that are written, with their content types. */
function withStories(
  pkg: OoxmlPackage,
  doc: OrderedDocument,
  added: ReadonlyMap<string, OoxmlPart>,
  keys: ReadonlySet<string>
): OoxmlPackage {
  let next = pkg;
  for (const resource of byKey(doc)) {
    if (resource.kind !== 'story' || !keys.has(resource.key)) continue;
    const part = added.get(resource.part);
    if (!part) continue;
    next = withContentTypeOverride(
      withPart(next, { ...part, contentType: STORY_TYPES[resource.role] }),
      resource.part,
      STORY_TYPES[resource.role]
    );

    // A copy of a loaded part takes its relationships, as the engine's unlink does
    if (resource.relationshipsOf !== undefined && next.parts.has(resource.relationshipsOf)) {
      const copied = cloneOwnedRelationships(next, resource.relationshipsOf, resource.part);
      if (!copied)
        throw new OrderedRefusal('engine-refused', `no relationships for ${resource.part}`);
      next = copied;
    }
  }

  return next;
}

/** The chart and diagram parts that are written, with their content types. */
function withObjectParts(
  pkg: OoxmlPackage,
  doc: OrderedDocument,
  keys: ReadonlySet<string>
): OoxmlPackage {
  let next = pkg;
  for (const resource of byKey(doc)) {
    if (resource.kind !== 'part' || !keys.has(resource.key)) continue;
    const contentType = FRAGMENT_PART_CONTENT_TYPES[resource.type];
    next = withNewPart(next, resource.part, resource.root, contentType);
    if (!next.parts.has(resource.part))
      throw new OrderedRefusal('engine-refused', `no part ${resource.part}`);
  }

  return next;
}

/** The media parts that are written. */
function withMedia(
  pkg: OoxmlPackage,
  doc: OrderedDocument,
  keys: ReadonlySet<string>,
  bytes: (media: OrderedMediaResource) => Uint8Array | undefined
): OoxmlPackage {
  let next = pkg;
  for (const resource of byKey(doc)) {
    if (resource.kind !== 'media' || !keys.has(resource.key)) continue;
    const data = bytes(resource);
    if (data !== undefined) next = withBinaryPart(next, resource.part, data, resource.contentType);
  }

  return next;
}

/** One relationship to write: its node key, id, full type, target and whether it is external. */
interface RelationshipRow {
  readonly key: string;
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external: boolean;
}

/** Relationships added to an owner's relationship part, and to the package's index of them. */
function writeRelationships(
  pkg: OoxmlPackage,
  owner: string,
  rows: readonly RelationshipRow[]
): OoxmlPackage {
  if (rows.length === 0) return pkg;
  const next = withRelationshipsPartFor(pkg, owner);
  const relsName = relsPartNameFor(owner);
  const rels = next.parts.get(relsName);
  if (!rels)
    throw new OrderedRefusal('engine-refused', `${owner} keeps its relationships as bytes`);
  const records = rows.map(
    (row) =>
      ({
        id: `${relsName}#ordered-${row.key}`,
        kind: 'generic',
        namespaceUri: PACKAGE_RELATIONSHIPS,
        localName: 'Relationship',
        namespaceBindings: [],
        attributes: [
          plain('Id', row.id),
          plain('Type', row.type),
          plain('Target', row.target),
          ...(row.external ? [plain('TargetMode', 'External')] : []),
        ],
        children: [],
      }) as unknown as OoxmlElement
  );
  const internal = rows
    .filter((row) => !row.external)
    .map((row, k) => ({
      ownerPart: owner,
      id: row.id,
      type: row.type,
      rawTarget: row.target,
      targetMode: 'Internal' as const,
      order: (next.relationships.get(owner)?.length ?? 0) + k,
    }));
  const external = rows
    .filter((row) => row.external)
    .map((row) => ({
      ownerPart: owner,
      id: row.id,
      type: row.type,
      rawTarget: row.target,
      sinkSafe: true,
    }));
  const written = withPart(next, {
    ...rels,
    root: withChildren(rels.root, [...rels.root.children, ...records]),
  });
  const relationships = new Map(written.relationships);
  relationships.set(owner, [...(written.relationships.get(owner) ?? []), ...internal]);
  return Object.freeze({
    ...written,
    relationships,
    externalTargets: [...written.externalTargets, ...external],
  });
}

/** The relationships that are written, in their owners' relationship parts. */
function withRelationships(
  pkg: OoxmlPackage,
  doc: OrderedDocument,
  keys: ReadonlySet<string>
): OoxmlPackage {
  const byOwner = new Map<string, OrderedRelationshipResource[]>();
  for (const resource of byKey(doc)) {
    if (resource.kind !== 'relationship' || !keys.has(resource.key)) continue;
    byOwner.set(resource.owner, [...(byOwner.get(resource.owner) ?? []), resource]);
  }

  let next = pkg;
  for (const [owner, list] of byOwner)
    next = writeRelationships(
      next,
      owner,
      list.map((resource) => ({
        key: resource.key,
        id: resource.id,
        type: RELATIONSHIP_TYPES[resource.type],
        target: targetOf(doc, resource, owner),
        external: 'external' in resource.target,
      }))
    );
  return next;
}

function targetOf(doc: OrderedDocument, resource: OrderedRelationshipResource, owner: string) {
  if ('external' in resource.target) return resource.target.external;
  const target = doc.resources.added.get(resource.target.resource);
  if (!target || (target.kind !== 'media' && target.kind !== 'story' && target.kind !== 'part'))
    throw new OrderedRefusal('unknown-address', `no part for ${resource.key}`);
  return targetFrom(owner, target.part);
}

/** A part of the main document's by relationship type, or null. */
function relatedPart(pkg: OoxmlPackage, type: string): OoxmlPart | null {
  const main = pkg.mainDocumentPart;
  const record = pkg.relationships
    .get(main)
    ?.find((entry) => entry.type === `${OFFICE_RELATIONSHIPS}/${type}`);
  const resolved = record === undefined ? null : resolveRelationship(record);
  if (!resolved || resolved.mode !== 'Internal' || !resolved.target.ok) return null;
  return pkg.parts.get(resolved.target.partName) ?? null;
}

/** The part of the main document's of a type: its own, or a new one with its relationship. */
function ensurePart(pkg: OoxmlPackage, type: 'styles' | 'numbering'): [OoxmlPackage, OoxmlPart] {
  const found = relatedPart(pkg, type);
  if (found) return [pkg, found];
  const name = `/word/${type}.xml`;
  const read = readOoxmlPart(`<w:${type} xmlns:w="${WML_NAMESPACE_URI}"/>`, {
    name,
    contentType: `${WML_TYPE}.${type}+xml`,
  });
  if (!read.ok || pkg.parts.has(name))
    throw new OrderedRefusal('engine-refused', `no ${type} part can be made`);
  const made = withNewPart(pkg, name, read.part.root, `${WML_TYPE}.${type}+xml`);
  const related = withRelationship(
    withRelationshipsPartFor(made, made.mainDocumentPart),
    made.mainDocumentPart,
    `${OFFICE_RELATIONSHIPS}/${type}`,
    `${type}.xml`
  );
  const part = related.pkg.parts.get(name);
  if (!related.ok || !part) throw new OrderedRefusal('engine-refused', `no ${type} relationship`);
  return [related.pkg, part];
}

/**
 * A pasted tree under node ids of its own resource, numbered as the built ones are.
 *
 * The fragment reader gives every tree it reads the ids of a new part, from `#0.0` on. Kept,
 * two pasted list definitions, or one and the file's first `w:abstractNum`, held the same
 * ids, and the engine refused the composed numbering part (`duplicate-id`): the editor then
 * kept its old paragraphs. The key is unique and every copy composes the same ids.
 */
function keyedTree<T extends OoxmlNode>(node: T, id: () => string): T {
  if (node.kind === 'textValue') return { ...node, id: id() };
  return { ...node, id: id(), children: node.children.map((child) => keyedTree(child, id)) };
}

/** The node ids of a style resource's `w:style` and its children: `ordered-<key>-<n>`. */
const STYLE_NODE_ID = /^ordered-(.+)-\d+$/;

/** The key of the style resource compose drew a `w:style` from, if it drew it. */
export const styleResourceKeyOf = (element: OoxmlElement): string | undefined =>
  STYLE_NODE_ID.exec(element.id)?.[1];

/** A style resource as `w:style`, built from its typed values, or a pasted style as it came. */
function styleElement(style: OrderedStyleResource): OoxmlElement {
  let n = 0;
  const id = () => `ordered-${style.key}-${(n += 1)}`;
  if (style.element !== undefined) return keyedTree(style.element, id);
  const children: OoxmlNode[] = [
    build(id(), 'generic', 'name', [wmlAttribute('val', style.name)], []),
  ];
  for (const name of ['basedOn', 'next', 'link'] as const) {
    const value = style[name];
    if (value !== undefined)
      children.push(build(id(), 'generic', name, [wmlAttribute('val', value)], []));
  }

  // Word's order (CT_Style): name, basedOn, next, link, uiPriority, unhideWhenUsed, pPr, rPr
  if (style.uiPriority !== undefined)
    children.push(
      build(id(), 'generic', 'uiPriority', [wmlAttribute('val', String(style.uiPriority))], [])
    );
  if (style.unhideWhenUsed === true)
    children.push(build(id(), 'generic', 'unhideWhenUsed', [], []));
  if (style.paragraph.length > 0)
    children.push(
      build(
        id(),
        'paragraphProperties',
        'pPr',
        [],
        mergedPropertyChildren([], style.paragraph, PARAGRAPH_VOCABULARY, id)
      )
    );
  if (style.run.length > 0)
    children.push(
      build(
        id(),
        'runProperties',
        'rPr',
        [],
        mergedPropertyChildren([], style.run, RUN_VOCABULARY, id)
      )
    );
  return build(
    id(),
    'generic',
    'style',
    [wmlAttribute('type', style.type), wmlAttribute('styleId', style.styleId)],
    children
  );
}

/**
 * The styles written stories, styles and lists use, past the ones the file has: a style a written
 * story names, or one a written style draws on.
 */
function withStyles(pkg: OoxmlPackage, doc: OrderedDocument, stories: readonly OoxmlPart[]) {
  const styles = byKey(doc).filter(
    (resource): resource is OrderedStyleResource =>
      resource.kind === 'style' && orderedResourceShows(doc, resource)
  );
  if (styles.length === 0) return pkg;
  const used = new Set<string>();
  for (const part of stories) for (const id of usesOf(part).styles) used.add(id);
  const existing = relatedPart(pkg, 'styles');
  const present = new Set<string>();
  for (const child of existing?.root.children ?? [])
    if (child.kind !== 'textValue' && isWml(child, 'style')) {
      const id = attributeOf(child, 'styleId');
      if (id !== undefined) present.add(id);
    }

  // The first resource with an id wins; a style draws on its base, next and linked styles
  const chosen = new Map<string, OrderedStyleResource>();
  for (let grew = true; grew; ) {
    grew = false;
    for (const style of styles) {
      if (!used.has(style.styleId) || present.has(style.styleId) || chosen.has(style.styleId))
        continue;
      chosen.set(style.styleId, style);
      grew = true;
      for (const other of [style.basedOn, style.next, style.link])
        if (other !== undefined) used.add(other);
    }
  }

  if (chosen.size === 0) return pkg;
  const [next, part] = ensurePart(pkg, 'styles');
  const elements = [...chosen.values()].map(styleElement);
  return withPart(next, {
    ...part,
    root: withChildren(part.root, [...part.root.children, ...elements]),
  });
}

/** An element with only a `w:val`, under the next id `id` gives. */
const valueElement = (id: () => string, name: string, value: string): OoxmlElement =>
  build(id(), 'generic', name, [wmlAttribute('val', value)], []);

/** One level of a typed list as `w:lvl`, in Word's default form for it. */
function levelElement(
  level: OrderedListResource['levels'][number],
  id: () => string
): OoxmlElement {
  const children: OoxmlNode[] = [
    valueElement(id, 'start', String(level.start)),
    valueElement(id, 'numFmt', level.format),
    valueElement(id, 'lvlText', level.text),
    valueElement(id, 'lvlJc', level.justification),
  ];
  const indent = [
    wmlAttribute('left', String(level.indent.left)),
    wmlAttribute('hanging', String(level.indent.hanging)),
  ];
  children.push(
    build(id(), 'paragraphProperties', 'pPr', [], [build(id(), 'generic', 'ind', indent, [])])
  );
  if (level.font !== undefined) {
    const fonts = [
      wmlAttribute('ascii', level.font),
      wmlAttribute('hAnsi', level.font),
      wmlAttribute('hint', 'default'),
    ];
    children.push(
      build(id(), 'runProperties', 'rPr', [], [build(id(), 'generic', 'rFonts', fonts, [])])
    );
  }

  return build(id(), 'generic', 'lvl', [wmlAttribute('ilvl', String(level.level))], children);
}

/** A list resource as `w:abstractNum` and `w:num`. */
function listElements(list: OrderedListResource): [OoxmlElement, OoxmlElement] {
  let n = 0;
  const id = () => `ordered-${list.key}-${(n += 1)}`;
  const value = (name: string, val: string) => valueElement(id, name, val);
  const definition = list.definition;
  if (definition !== undefined) {
    // A pasted definition as it came, under the numbers the key gives
    const abstract = withAttributes(keyedTree(definition.abstract, id), [
      ...definition.abstract.attributes.filter((a) => a.localName !== 'abstractNumId'),
      wmlAttribute('abstractNumId', list.abstractNumId),
    ]);
    const num = build(
      id(),
      'generic',
      'num',
      [wmlAttribute('numId', list.numId)],
      [
        value('abstractNumId', list.abstractNumId),
        ...definition.overrides.map((override) => keyedTree(override, id)),
      ]
    );
    return [abstract, num];
  }

  const levels = list.levels.map((level) => levelElement(level, id));
  const abstract = build(
    id(),
    'generic',
    'abstractNum',
    [wmlAttribute('abstractNumId', list.abstractNumId)],
    [
      value('multiLevelType', list.levels.length > 1 ? 'hybridMultilevel' : 'singleLevel'),
      ...levels,
    ]
  );
  const num = build(
    id(),
    'generic',
    'num',
    [wmlAttribute('numId', list.numId)],
    [value('abstractNumId', list.abstractNumId)]
  );
  return [abstract, num];
}

/**
 * The picture bullets of the written lists, in the numbering part's relationships and media
 * (design D18): a bullet whose images have no bytes is left out (`missing: 'skip'`), and its
 * levels keep their text. Answers the bullets' elements and the ids left out.
 */
function withListBullets(
  pkg: OoxmlPackage,
  doc: OrderedDocument,
  numbering: string,
  lists: readonly OrderedListResource[],
  bytes: (media: OrderedMediaResource) => Uint8Array | undefined
): [OoxmlPackage, OoxmlElement[], ReadonlySet<string>] {
  let next = pkg;
  const elements: OoxmlElement[] = [];
  const leftOutBullets = new Set<string>();
  const rows: RelationshipRow[] = [];
  for (const list of lists)
    for (const bullet of list.bullets ?? []) {
      const media = bullet.images.map((image) => doc.resources.added.get(image.media));
      const data = media.map((m) => (m?.kind === 'media' ? bytes(m) : undefined));
      if (data.some((d) => d === undefined)) {
        leftOutBullets.add(bullet.id);
        continue;
      }

      media.forEach((m, k) => {
        if (m?.kind !== 'media' || next.partBytes.has(m.part)) return;
        next = withBinaryPart(next, m.part, data[k]!, m.contentType);
      });
      bullet.images.forEach((image, k) =>
        rows.push({
          key: `${list.key}-${image.id}`,
          id: image.id,
          type: RELATIONSHIP_TYPES.image,
          target: targetFrom(numbering, (media[k] as OrderedMediaResource).part),
          external: false,
        })
      );
      elements.push(bulletElement(list, bullet));
    }

  return [writeRelationships(next, numbering, rows), elements, leftOutBullets];
}

/** A list's picture bullet as `w:numPicBullet`, under the id the list's key gives it. */
function bulletElement(list: OrderedListResource, bullet: OrderedListBullet): OoxmlElement {
  let n = 0;
  const id = () => `ordered-${list.key}-b${bullet.id}-${(n += 1)}`;
  return withAttributes(keyedTree(bullet.element, id), [
    ...bullet.element.attributes.filter((a) => a.localName !== 'numPicBulletId'),
    wmlAttribute('numPicBulletId', bullet.id),
  ]);
}

/** The list definitions whose numbers written stories and styles name, in the numbering part. */
function withLists(
  pkg: OoxmlPackage,
  doc: OrderedDocument,
  stories: readonly OoxmlPart[],
  bytes: (media: OrderedMediaResource) => Uint8Array | undefined
) {
  const lists = byKey(doc).filter(
    (resource): resource is OrderedListResource =>
      resource.kind === 'list' && orderedResourceShows(doc, resource)
  );
  if (lists.length === 0) return pkg;
  const used = new Set<string>();
  const styles = relatedPart(pkg, 'styles');
  for (const part of styles ? [...stories, styles] : stories)
    for (const id of usesOf(part).lists) used.add(id);
  const chosen = lists.filter((list) => used.has(list.numId));
  if (chosen.length === 0) return pkg;
  const [made, found] = ensurePart(pkg, 'numbering');
  const [next, bullets, leftOutBullets] = withListBullets(made, doc, found.name, chosen, bytes);

  // A bullet's VML or DrawingML prefixes, declared on the numbering root
  const part = withRequiredNamespaceBindings(found, bullets);
  const children = [...part.root.children];

  // `CT_Numbering`: picture bullets, abstract definitions, then numbers
  let lastBullet = -1;
  let lastAbstract = -1;
  let lastNum = -1;
  children.forEach((child, k) => {
    if (isWml(child, 'numPicBullet')) lastBullet = k;
    if (isWml(child, 'abstractNum') || isWml(child, 'numPicBullet')) lastAbstract = k;
    if (isWml(child, 'num')) lastNum = k;
  });
  const elements = chosen.map((list) =>
    listElements(list).map((element) => withoutBullets(element, leftOutBullets))
  );
  const abstracts = elements.map(([abstract]) => abstract!);
  const nums = elements.map(([, num]) => num!);
  const numAt = lastNum >= 0 ? lastNum + 1 : children.length;
  children.splice(numAt, 0, ...nums);
  children.splice(lastAbstract + 1, 0, ...abstracts);
  children.splice(lastBullet + 1, 0, ...bullets);
  return withPart(next, { ...part, root: withChildren(part.root, children) });
}

/**
 * The list levels that changes declared, each in Word's default format for its depth (the
 * engine's `ensureNumberingLevel`), in the order changes declared them. A level already declared
 * leaves the part as it is; a definition the engine will not extend (a `w:numStyleLink`) too.
 */
function withListLevels(pkg: OoxmlPackage, doc: OrderedDocument): OoxmlPackage {
  let next = pkg;
  for (const resource of byKey(doc))
    if (resource.kind === 'list-level' && orderedResourceShows(doc, resource))
      next = ensureNumberingLevel(next, resource.numId, resource.level, resource.format) ?? next;
  return next;
}

/** A definition without the levels' picture bullets whose ids `leftOut` holds. */
function withoutBullets(element: OoxmlElement, leftOut: ReadonlySet<string>): OoxmlElement {
  if (leftOut.size === 0) return element;
  const children: OoxmlNode[] = [];
  for (const child of element.children) {
    if (child.kind === 'textValue') children.push(child);
    else if (isWml(child, 'lvlPicBulletId') && leftOut.has(attributeOf(child, 'val') ?? ''))
      continue;
    else children.push(withoutBullets(child, leftOut));
  }

  return withChildren(element, children);
}

/**
 * The settings' list of each written new notes part's separators (`w:footnotePr` and
 * `w:endnotePr`), as the engine's `ensureNotesPart` and Word write it: it goes with its part.
 */
function withNoteSeparators(
  pkg: OoxmlPackage,
  doc: OrderedDocument,
  keys: ReadonlySet<string>
): OoxmlPackage {
  let next = pkg;
  for (const resource of byKey(doc)) {
    if (resource.kind !== 'story' || !keys.has(resource.key)) continue;
    if (resource.role !== 'footnotes' && resource.role !== 'endnotes') continue;
    const kind = resource.role === 'footnotes' ? 'footnote' : 'endnote';
    const listed = writeNoteSeparatorReferences(next, kind);
    if (!listed)
      throw new OrderedRefusal('engine-refused', `the settings list of ${kind} separators`);
    next = listed;
  }

  return next;
}

/**
 * The bytes of used media (a resource a change added, or the file's media a checkpoint names), or
 * undefined to leave it out.
 */
export type OrderedMediaBytes = (media: OrderedMediaReference) => Uint8Array | undefined;

/**
 * The media bytes that `options` gives: bytes of another size are missing, and missing bytes
 * are refused (`missing-blob`) or left out, as `options.missing` says.
 */
export function mediaBytesOf(options: OrderedResourceOptions): OrderedMediaBytes {
  return (media) => {
    const data = options.blobs?.(media.sha256);
    if (data !== undefined && data.byteLength === media.size) {
      return data;
    }

    if ((options.missing ?? 'refuse') === 'skip') {
      return undefined;
    }

    throw new OrderedRefusal('missing-blob', media.sha256);
  };
}

/**
 * The package with every resource that shows and is used written, and the new story parts
 * (`added`, composed with the other stories) that a used relationship points at. A media part
 * whose `bytes` are undefined is left out with its relationship.
 *
 * @throws OrderedRefusal when `bytes` refuses missing media, or a part cannot be made.
 */
export function withOrderedResources(
  pkg: OoxmlPackage,
  doc: OrderedDocument,
  added: ReadonlyMap<string, OoxmlPart>,
  bytes: OrderedMediaBytes
): OoxmlPackage {
  if (doc.resources.added.size === 0 && !hasCustomXml(doc)) return pkg;
  const written = reach(pkg, doc, added, bytes);
  let next = withStories(pkg, doc, added, written.keys);
  next = withObjectParts(next, doc, written.keys);
  next = withMedia(next, doc, written.keys, bytes);
  next = withRelationships(next, doc, written.keys);
  next = withStyles(next, doc, written.stories);
  next = withListLevels(withLists(next, doc, written.stories, bytes), doc);
  if (hasCustomXml(doc)) next = withCustomXml(next, doc, boundNodes(written.stories));
  return withNoteSeparators(withSettings(next, doc), doc, written.keys);
}

/** The parts of the stories a change added, which compose holds apart from the package. */
export function isAddedStory(doc: OrderedDocument, name: string): boolean {
  for (const resource of doc.resources.added.values())
    if (resource.kind === 'story' && resource.part === name) return true;
  return false;
}
