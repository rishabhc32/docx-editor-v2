import type { OoxmlPackage } from '../../../../../store/package/ooxml-package.ts';
import { resolveInternalTarget } from '../../../../../store/package/opc-names.ts';
import { resolveContentTypeOf } from '../../../../../store/package/package-edit.ts';
import type { RelationshipRecord } from '../../../../../store/package/relationships.ts';
import type { OoxmlElement, OoxmlNode } from '../../../../../store/package/ooxml-tree.ts';
import {
  acceptedFragmentBlocks,
  withoutTrackedChanges,
} from '../../../../../store/store/fragment-final-view.ts';
import type { TreeDocOp } from '../../../../../store/store/tree-op-types.ts';
import type {
  OrderedFragmentElement,
  OrderedFragmentList,
  OrderedFragmentNote,
  OrderedFragmentPart,
  OrderedFragmentPartRelationship,
} from '../../../change/fragment.ts';
import {
  FRAGMENT_OBJECT_RELATIONSHIP_TYPES,
  FRAGMENT_PART_CHILDREN,
  FRAGMENT_PART_RELATIONSHIP_TYPES,
  FRAGMENT_PART_ROOTS,
  type OrderedFragmentObjectUse,
  type OrderedFragmentPartType,
} from '../../../change/fragment-names-objects.ts';
import { isOrderedMediaFor, ORDERED_WORKBOOK_TYPE } from '../../../model/resources.ts';
import { ORDERED_FRAGMENT_LIMITS } from '../../../change/fragment.ts';
import { attributeOf, isWml, withAttributes, withChildren } from '../../../ooxml-nodes.ts';
import {
  fragmentPartWire,
  fragmentWire,
  objectRefusal,
  pictureByAddress,
} from '../../fragment-wire.ts';
import { placeAt } from '../../paragraphs.ts';
import { OrderedClientRefusal } from '../../refusal.ts';
import type { OrderedChangeMedia } from '../../../change/resources.ts';
import type { OrderedDocument } from '../../../model/types.ts';
import type { OrderedChangeBody, OrderedTranslateBatch } from '../types.ts';
import { addedStyleId, elements, isNewList, listElementsOf, relatedPart } from './definitions.ts';

type InsertFragmentOp = Extract<TreeDocOp, { readonly op: 'insertFragment' }>;

/**
 * Blocks as their accepted final view: no other person's tracked change enters the room. The
 * room stamps each change with its sender, so the client sends the final view (design D18), as
 * Word copies from a document that tracks changes.
 */
function finalView(pkg: OoxmlPackage, blocks: readonly OoxmlNode[]): readonly OoxmlNode[] {
  const main = pkg.parts.get(pkg.mainDocumentPart);
  const accepted = main ? acceptedFragmentBlocks(main, blocks) : null;
  if (accepted === null)
    throw new OrderedClientRefusal('ordered-refused', 'engine-refused', 'a pasted tracked change');
  return accepted;
}

/** Element nodes of a tree, walked once, depth first. */
function* walk(nodes: readonly OoxmlNode[]): Generator<OoxmlElement> {
  const stack = [...nodes].reverse();
  for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
    if (node.kind === 'textValue') continue;
    yield node;
    for (let k = node.children.length - 1; k >= 0; k -= 1) {
      const child = node.children[k];
      if (child !== undefined) stack.push(child);
    }
  }
}

/**
 * What the earlier changes of one long paste made: the list numbers they minted, by the paste's
 * number, and the styles they carried. A later change names them and does not carry them again.
 */
class Earlier {
  readonly lists = new Map<string, string>();
  readonly styleIds = new Set<string>();
}

/** The numbers of the lists changes added, in the order they were added. */
function addedListNumbers(doc: OrderedDocument): string[] {
  const out: string[] = [];
  for (const resource of doc.resources.added.values()) {
    if (resource.kind === 'list') {
      out.push(resource.numId);
    }
  }

  return out;
}

/**
 * What one change of a paste carries besides its blocks, by index, collected as its blocks are
 * written: what the document does not hold yet. Bytes (pictures, an embedded object's or ink's
 * part, a chart's workbook) are held by hash, never fetched or opened.
 */
class Carried {
  readonly lists: OrderedFragmentList[] = [];
  readonly styles: OrderedFragmentElement[] = [];
  readonly links: string[] = [];
  readonly media: OrderedChangeMedia[] = [];
  readonly notes: OrderedFragmentNote[] = [];
  readonly parts: OrderedFragmentPart[] = [];
  readonly bullets: OrderedFragmentElement[] = [];
  readonly #parts = new Map<string, number>();
  readonly #bullets = new Map<string, string>();
  readonly #numbers = new Map<string, string>();
  readonly #rels = new Map<string, string>();
  readonly #images = new Map<string, number>();
  /** How many lists changes had added before this one. */
  readonly #listsBefore: number;
  /** The part whose relationships the nodes being written name: the story's, or a note's. */
  #owner: string;

  private readonly batch: OrderedTranslateBatch;
  private readonly pkg: OoxmlPackage;
  private readonly shown: OoxmlPackage | null;
  private readonly earlier: Earlier;

  constructor(
    batch: OrderedTranslateBatch,
    pkg: OoxmlPackage,
    shown: OoxmlPackage | null,
    earlier: Earlier
  ) {
    this.batch = batch;
    this.pkg = pkg;
    this.shown = shown;
    this.earlier = earlier;
    this.#listsBefore = addedListNumbers(batch.doc).length;
    this.#owner = pkg.mainDocumentPart;
  }

  /**
   * After the change applied: the numbers its lists got, for the later changes. The apply adds
   * the lists in the change's order, after every list added before it.
   */
  applied(): void {
    const minted = addedListNumbers(this.batch.doc).slice(this.#listsBefore);
    for (const [value, reference] of this.#numbers) {
      const numId = minted[Number(reference.slice(1))];
      if (numId !== undefined) {
        this.earlier.lists.set(value, numId);
      }
    }
  }

  /** Element trees for nodes, with what they use carried. */
  write(nodes: readonly OoxmlNode[]): OrderedFragmentElement[] {
    this.carryStyles(nodes);
    return fragmentWire(nodes, this.refs());
  }

  /** A list number the document holds stays; another is carried, once. */
  numId(value: string): string {
    if (!isNewList(this.batch.doc, value)) {
      return value;
    }

    const known = this.earlier.lists.get(value) ?? this.#numbers.get(value);
    if (known !== undefined) {
      return known;
    }

    const list = listElementsOf(this.pkg, value);
    if (!list) {
      throw new OrderedClientRefusal('ordered-refused', 'not-covered', `no list ${value}`);
    }

    const { num, abstract } = list;
    const reference = `#${this.lists.length}`;
    this.#numbers.set(value, reference);

    // The definition's own styles travel too (a list that a style links)
    this.carryStyles([abstract]);
    const [tree] = fragmentWire([withoutTrackedChanges(abstract)], this.refs());
    if (!tree) throw new OrderedClientRefusal('ordered-refused', 'not-covered', 'a list');
    const overrides = fragmentWire(
      elements(num)
        .filter((n) => isWml(n, 'lvlOverride'))
        .map(withoutTrackedChanges),
      this.refs()
    );
    this.lists.push(overrides.length > 0 ? { abstract: tree, overrides } : { abstract: tree });
    return reference;
  }

  refs() {
    return {
      numId: (value: string) => this.numId(value),
      relationship: (id: string) => this.link(id),
      media: (id: string) => this.image(id),
      note: (kind: 'footnote' | 'endnote', id: string) => this.note(kind, id),
      part: (id: string, type: OrderedFragmentPartType) => this.storyPart(id, type),
      object: (id: string, use: OrderedFragmentObjectUse) => this.object(id, use),
      bullet: (id: string) => this.bullet(id),
    };
  }

  /**
   * A list level's picture bullet, carried once without its id (the apply mints one), its
   * images held by hash as relationships of the numbering part (never fetched).
   */
  bullet(id: string): string {
    const known = this.#bullets.get(id);
    if (known !== undefined) return known;
    const numbering = relatedPart(this.pkg, 'numbering');
    const bullet = numbering
      ? elements(numbering.root).find(
          (n) => isWml(n, 'numPicBullet') && attributeOf(n, 'numPicBulletId') === id
        )
      : undefined;
    if (!numbering || !bullet)
      throw new OrderedClientRefusal('ordered-refused', 'not-covered', `no picture bullet ${id}`);
    const k = this.bullets.length;
    const reference = `#${k}`;
    this.#bullets.set(id, reference);
    this.bullets.push({ e: 'w:numPicBullet' });
    const owner = this.#owner;
    this.#owner = numbering.name;
    const [tree] = fragmentWire(
      [
        withAttributes(
          bullet,
          bullet.attributes.filter((a) => a.localName !== 'numPicBulletId')
        ),
      ],
      this.refs()
    );
    this.#owner = owner;
    if (!tree) throw new OrderedClientRefusal('ordered-refused', 'not-covered', 'a picture bullet');
    this.bullets[k] = tree;
    return reference;
  }

  /** Media held by hash, carried once: its index in the change's media. */
  held(at: string, take: () => OrderedChangeMedia | null): number {
    const known = this.#images.get(at);
    if (known !== undefined) {
      return known;
    }

    const media = take();
    if (media === null) {
      throw objectRefusal();
    }

    const k = this.media.length;
    this.#images.set(at, k);
    this.media.push(media);
    return k;
  }

  /** A chart's or a diagram's part a story's relationship names: `#<k>` into the parts. */
  storyPart(id: string, type: OrderedFragmentPartType): string {
    const record = this.pkg.relationships.get(this.#owner)?.find((entry) => entry.id === id);
    const target =
      record !== undefined &&
      record.targetMode !== 'External' &&
      record.type === FRAGMENT_PART_RELATIONSHIP_TYPES[type]
        ? resolveInternalTarget(this.#owner, record.rawTarget)
        : null;
    if (target === null || !target.ok) {
      throw objectRefusal();
    }

    return `#${this.objectPart(target.partName, type)}`;
  }

  /**
   * A part, carried once, with its own relationships: to the parts its type may name, to
   * pictures, and a chart's to its embedded workbook. An external one (a chart linked to a
   * workbook file) is refused by name: the address is never loaded.
   */
  objectPart(name: string, type: OrderedFragmentPartType): number {
    const known = this.#parts.get(name);
    if (known !== undefined) {
      return known;
    }

    const part = this.pkg.parts.get(name);
    if (part === undefined) {
      throw objectRefusal();
    }

    const k = this.parts.length;
    this.#parts.set(name, k);
    this.parts.push({ type, root: { e: FRAGMENT_PART_ROOTS[type] } });
    const rels: OrderedFragmentPartRelationship[] = [];
    const local = new Map<string, string>();
    for (const record of this.pkg.relationships.get(name) ?? []) {
      rels.push(this.#partRelationship(name, type, record));
      local.set(record.id, `#${rels.length - 1}`);
    }

    const root = fragmentPartWire(part.root, this.refs(), {
      relationship: (id) => {
        const reference = local.get(id);
        if (reference === undefined) {
          throw objectRefusal();
        }

        return reference;
      },
      drawing: (id) => this.storyPart(id, 'diagramDrawing'),
    });
    this.parts[k] = rels.length > 0 ? { type, root, rels } : { type, root };
    return k;
  }

  /**
   * One relationship of a carried part: to a part its type may name, to a picture, or a chart's
   * to its embedded workbook. Any other relationship is refused by name.
   */
  #partRelationship(
    name: string,
    type: OrderedFragmentPartType,
    record: RelationshipRecord
  ): OrderedFragmentPartRelationship {
    if (record.targetMode === 'External') {
      throw objectRefusal();
    }

    const target = resolveInternalTarget(name, record.rawTarget);
    if (!target.ok) {
      throw objectRefusal();
    }

    const child = FRAGMENT_PART_CHILDREN[type].find(
      (candidate) => FRAGMENT_PART_RELATIONSHIP_TYPES[candidate] === record.type
    );
    if (child !== undefined) {
      return { part: this.objectPart(target.partName, child) };
    }

    const workbook =
      record.type === FRAGMENT_PART_RELATIONSHIP_TYPES.package &&
      (type === 'chart' || type === 'chartEx');
    if (workbook) {
      return {
        media: this.held(`${name}|${record.id}`, () => {
          const bytes =
            this.pkg.partBytes.get(target.partName) ??
            this.pkg.partBytes.get(target.partName.slice(1));
          const isWorkbook =
            resolveContentTypeOf(this.pkg, target.partName) === ORDERED_WORKBOOK_TYPE;
          return isWorkbook
            ? this.batch.partMedia(target.partName, bytes, ORDERED_WORKBOOK_TYPE)
            : null;
        }),
      };
    }

    if (record.type.endsWith('/image')) {
      return { media: this.held(`${name}|${record.id}`, () => this.batch.media(name, record.id)) };
    }

    throw objectRefusal();
  }

  /**
   * An embedded object's part (an OLE part, or an embedded workbook, document or presentation)
   * or an ink part, held by its hash as bytes that nothing opens, and carried once. A linked
   * object (an address) never travels: the editor's copy leaves it out, as it never loads the
   * address, so a paste that still names one is refused by name. A part of another type is
   * one the editor alone does not carry either (`insertFragment.object`).
   */
  object(id: string, use: OrderedFragmentObjectUse): string {
    const owner = this.#owner;
    const record = this.pkg.relationships.get(owner)?.find((entry) => entry.id === id);
    if (record?.targetMode === 'External')
      throw new OrderedClientRefusal('ordered-refused', 'not-covered', 'a linked object');
    const types =
      use === 'ink'
        ? [FRAGMENT_OBJECT_RELATIONSHIP_TYPES.ink]
        : [
            FRAGMENT_OBJECT_RELATIONSHIP_TYPES.oleObject,
            FRAGMENT_OBJECT_RELATIONSHIP_TYPES.package,
          ];
    const k = this.held(`${owner}|${id}`, () => {
      const target =
        record !== undefined && types.includes(record.type)
          ? resolveInternalTarget(owner, record.rawTarget)
          : null;
      if (target === null || !target.ok) return null;
      const bytes =
        this.pkg.partBytes.get(target.partName) ?? this.pkg.partBytes.get(target.partName.slice(1));
      const type = resolveContentTypeOf(this.pkg, target.partName);
      return type !== null && isOrderedMediaFor(type, use)
        ? this.batch.partMedia(target.partName, bytes, type)
        : null;
    });
    return `#${k}`;
  }

  /**
   * A picture's image, held by its hash and carried once. An image by address is refused by
   * name (`external-picture-link`): no room writes one (design D18), and it is never fetched.
   */
  image(id: string): string {
    const owner = this.#owner;
    const record = this.pkg.relationships.get(owner)?.find((entry) => entry.id === id);
    if (record?.targetMode === 'External') throw pictureByAddress();
    return `#${this.held(`${owner}|${id}`, () => this.batch.media(owner, id))}`;
  }

  /** A link's address, carried once. */
  link(id: string): string {
    const at = `${this.#owner}|${id}`;
    const known = this.#rels.get(at);
    if (known !== undefined) return known;
    const target = this.pkg.externalTargets.find(
      (t) => t.ownerPart === this.#owner && t.id === id && t.type.endsWith('/hyperlink')
    );
    if (!target || !target.sinkSafe)
      throw new OrderedClientRefusal('ordered-refused', 'invalid-change', 'a link target');
    const reference = `#${this.links.length}`;
    this.#rels.set(at, reference);
    this.links.push(
      /^[\x21-\x7E]+$/.test(target.rawTarget) ? target.rawTarget : encodeURI(target.rawTarget)
    );
    return reference;
  }

  /** A note the paste moved into the notes part, carried with its blocks. */
  note(kind: 'footnote' | 'endnote', id: string): string {
    const part = relatedPart(this.pkg, kind === 'footnote' ? 'footnotes' : 'endnotes');
    const note = part
      ? elements(part.root).find((n) => isWml(n, kind) && attributeOf(n, 'id') === id)
      : undefined;
    if (!note) throw new OrderedClientRefusal('ordered-refused', 'not-covered', `no ${kind} ${id}`);
    const k = this.notes.length;
    this.notes.push({ kind, blocks: [] });

    // The note's links and pictures are relationships of its notes part
    const owner = this.#owner;
    this.#owner = part?.name ?? owner;
    const blocks = this.write(finalView(this.pkg, elements(note)));
    this.#owner = owner;
    this.notes[k] = { kind, blocks };
    return `#${k}`;
  }

  /** The styles nodes name that the document does not hold, with the styles they draw on. */
  carryStyles(nodes: readonly OoxmlNode[]): void {
    const wanted: string[] = [];
    for (const node of walk(nodes))
      if (isWml(node, 'pStyle') || isWml(node, 'rStyle') || isWml(node, 'tblStyle')) {
        const id = attributeOf(node, 'val');
        if (id !== undefined) wanted.push(id);
      }

    const styles = relatedPart(this.pkg, 'styles');
    const held = this.shown ? relatedPart(this.shown, 'styles') : null;
    const has = (part: typeof styles, id: string) =>
      part !== null &&
      elements(part.root).some((s) => isWml(s, 'style') && attributeOf(s, 'styleId') === id);
    for (let id = wanted.pop(); id !== undefined; id = wanted.pop()) {
      if (this.earlier.styleIds.has(id)) {
        continue;
      }

      this.earlier.styleIds.add(id);

      // A style the document already holds is the document's; the paste merged onto it
      if (has(held, id) && !addedStyleId(this.batch.doc, id)) continue;
      const style = styles
        ? elements(styles.root).find((s) => isWml(s, 'style') && attributeOf(s, 'styleId') === id)
        : undefined;
      if (!style) continue;
      for (const name of ['basedOn', 'next', 'link'])
        for (const child of elements(style))
          if (isWml(child, name)) {
            const other = attributeOf(child, 'val');
            if (other !== undefined) wanted.push(other);
          }

      const [tree] = fragmentWire([withoutTrackedChanges(style)], this.refs());
      if (tree) this.styles.push(tree);
    }
  }
}

/** Elements under a node, for the cut. */
function weight(node: OoxmlNode): number {
  let n = 0;
  const walker = walk([node]);
  while (!walker.next().done) n += 1;
  return n;
}

/**
 * The blocks cut into groups under the change caps, before a paragraph only (a table or a block
 * control first would split the paragraph it lands in). Each group goes as one change before the
 * same place, which lands the blocks as one landing does. The count is of elements against half
 * the node cap: each text value is inside an element, so a group's blocks with their text values
 * stay under the cap. The trees a change carries beside its blocks are not counted.
 */
function groups(blocks: readonly OoxmlNode[]): OoxmlNode[][] {
  const cap = Math.floor(ORDERED_FRAGMENT_LIMITS.nodes / 2);
  const out: OoxmlNode[][] = [];
  let group: OoxmlNode[] = [];
  let size = 0;
  for (const block of blocks) {
    const own = weight(block);
    const paragraph = block.kind === 'paragraph' || isWml(block, 'p');
    if (
      group.length > 0 &&
      paragraph &&
      (size + own > cap || group.length >= ORDERED_FRAGMENT_LIMITS.blocks)
    ) {
      out.push(group);
      group = [];
      size = 0;
    }

    group.push(block);
    size += own;
  }

  if (group.length > 0) out.push(group);
  return out;
}

/** Bookmarks the paste brings whose names the document holds already: the document's stay. */
function withoutTakenBookmarks(
  blocks: readonly OoxmlNode[],
  shown: OoxmlPackage | null
): OoxmlNode[] {
  const main = shown?.parts.get(shown.mainDocumentPart);
  if (!main) return [...blocks];
  const taken = new Set<string>();
  for (const node of walk([main.root]))
    if (isWml(node, 'bookmarkStart')) {
      const name = attributeOf(node, 'name');
      if (name !== undefined) taken.add(name);
    }

  const dropped = new Set<string>();
  for (const node of walk(blocks))
    if (isWml(node, 'bookmarkStart') && taken.has(attributeOf(node, 'name') ?? '')) {
      const id = attributeOf(node, 'id');
      if (id !== undefined) dropped.add(id);
    }

  if (dropped.size === 0) return [...blocks];
  const keep = (node: OoxmlNode): OoxmlNode | null => {
    if (node.kind === 'textValue') return node;
    if (
      (isWml(node, 'bookmarkStart') || isWml(node, 'bookmarkEnd')) &&
      dropped.has(attributeOf(node, 'id') ?? '')
    )
      return null;
    const children = node.children.map(keep).filter((c): c is OoxmlNode => c !== null);
    return children.length === node.children.length ? node : withChildren(node, children);
  };
  return blocks.map(keep).filter((b): b is OoxmlNode => b !== null);
}

/**
 * The changes of one paste with formatting, each read after the one before it applied. The op's
 * list numbers, style ids, link relationships and note ids are the ones the paste's merge gave
 * them in the package the editor hands the gate (`gated`).
 */
export function* insertFragmentChanges(
  op: InsertFragmentOp,
  batch: OrderedTranslateBatch
): Generator<OrderedChangeBody> {
  const p = batch.paragraph(op.paragraphId);
  const pkg = batch.gated() ?? batch.shown();
  if (pkg === null)
    throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'nothing published yet');
  const place = placeAt(p, op.offset);
  const cut = groups(withoutTakenBookmarks(finalView(pkg, op.blocks), batch.shown()));
  const earlier = new Earlier();
  for (const [k, group] of cut.entries()) {
    const carried = new Carried(batch, pkg, batch.shown(), earlier);
    const blocks = carried.write(group);
    const whole = k < cut.length - 1 || op.lastMarkCovered === true;
    yield {
      kind: 'insert-fragment',
      part: p.story.name,
      place,
      blocks,
      // A group before another keeps its last mark: the next group starts a paragraph
      ...(whole ? { whole: true } : {}),
      ...(carried.lists.length > 0 ? { lists: carried.lists } : {}),
      ...(carried.styles.length > 0 ? { styles: carried.styles } : {}),
      ...(carried.links.length > 0 ? { links: carried.links } : {}),
      ...(carried.media.length > 0 ? { media: carried.media } : {}),
      ...(carried.notes.length > 0 ? { notes: carried.notes } : {}),
      ...(carried.parts.length > 0 ? { parts: carried.parts } : {}),
      ...(carried.bullets.length > 0 ? { bullets: carried.bullets } : {}),
    };
    carried.applied();
  }
}
