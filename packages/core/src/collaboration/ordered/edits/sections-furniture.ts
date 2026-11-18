// Header and footer edits (task 9.8): create, delete, link and unlink a section's header or
// footer, as the engine's lifecycle ops do (`store/package/hf-lifecycle.ts`), on the sections and
// the headers and footers they show now.

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { cloneWithNewIds } from '../../../store/store/tree-op-nodes.ts';
import type { OrderedMintedParagraph } from '../change/ids.ts';
import type { OrderedSectionView } from '../compose/sections.ts';
import { isAddedStory } from '../compose/resources.ts';
import { addOrderedStory, createOrderedItem, labelUnindex } from '../model/document.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import {
  addOrderedResource,
  orderedResourceShows,
  type OrderedRelationshipResource,
  type OrderedStoryResource,
} from '../model/resources.ts';
import type { OrderedFurnitureKind, OrderedFurnitureVariant } from '../model/sections.ts';
import type {
  OrderedDocument,
  OrderedItem,
  OrderedItemAddress,
  OrderedStory,
  OrderedSuggestion,
} from '../model/types.ts';
import { recordUndo } from '../model/undo-record.ts';
import { attributeOf, isWml, withChildren, withId } from '../ooxml-nodes.ts';
import { loadedFurnitureOf } from '../load/sections.ts';
import type { Ctx } from './context.ts';
import { mintParagraph, paragraphIdAttributes, withParagraphIdentity } from './para-ids.ts';
import { addRelationship, addStory, addStoryResource } from './resources.ts';
import { addSectionWrite, sectionAt, sectionsNow } from './sections.ts';

/**
 * A header or footer a section shows: its story, the relationship that names it, whether it is
 * its own.
 */
export interface OrderedFurnitureSlot {
  readonly story: string;
  readonly id: string;
  readonly inherited: boolean;
}

/** What one section shows: its headers and footers by variant, and its different first page. */
export interface OrderedSectionFurniture {
  readonly headers: ReadonlyMap<OrderedFurnitureVariant, OrderedFurnitureSlot>;
  readonly footers: ReadonlyMap<OrderedFurnitureVariant, OrderedFurnitureSlot>;
  readonly titlePage: boolean;
}

const VARIANTS = new Set<string>(['default', 'first', 'even']);
const ON = new Set(['true', '1', 'on']);

type FurnitureSlots = Map<OrderedFurnitureVariant, OrderedFurnitureSlot>;

/** The story a relationship of the main part names, of a kind; null when none shows. */
function storyOfReference(
  doc: OrderedDocument,
  main: OrderedStory,
  id: string,
  kind: OrderedFurnitureKind
): string | null {
  const loaded = loadedFurnitureOf(doc).parts.get(id);
  if (loaded && loaded.kind === kind) return doc.stories.has(loaded.part) ? loaded.part : null;
  for (const resource of doc.resources.added.values()) {
    if (resource.kind !== 'relationship' || resource.owner !== main.name) continue;
    if (resource.id !== id || resource.type !== kind || !('resource' in resource.target)) continue;
    if (!orderedResourceShows(doc, resource)) return null;
    const target = doc.resources.added.get(resource.target.resource);
    return target?.kind === 'story' && orderedResourceShows(doc, target) ? target.part : null;
  }

  return null;
}

/** The story kind a child of a section's properties references, or null. */
function referenceKind(child: OoxmlElement): OrderedFurnitureKind | null {
  if (isWml(child, 'headerReference')) {
    return 'header';
  }

  if (isWml(child, 'footerReference')) {
    return 'footer';
  }

  return null;
}

/**
 * The headers and footers a section declares itself, the first reference of each variant, and
 * its different first page.
 */
function declaredFurniture(
  doc: OrderedDocument,
  main: OrderedStory,
  sectPr: OoxmlElement | null | undefined
): { readonly header: FurnitureSlots; readonly footer: FurnitureSlots; titlePage: boolean } {
  const declared = {
    header: new Map<OrderedFurnitureVariant, OrderedFurnitureSlot>(),
    footer: new Map<OrderedFurnitureVariant, OrderedFurnitureSlot>(),
    titlePage: false,
  };
  for (const child of sectPr?.children ?? []) {
    if (child.kind === 'textValue') continue;
    if (isWml(child, 'titlePg')) {
      const value = attributeOf(child, 'val');
      declared.titlePage = value === undefined || ON.has(value);
      continue;
    }

    const kind = referenceKind(child);
    const variant = attributeOf(child, 'type');
    const id = child.attributes.find((a) => a.localName === 'id' && a.prefix === 'r')?.value;
    if (kind === null || variant === undefined || !VARIANTS.has(variant) || id === undefined) {
      continue;
    }

    const story = storyOfReference(doc, main, id, kind);
    const slots = declared[kind];
    const name = variant as OrderedFurnitureVariant;
    if (story !== null && !slots.has(name)) slots.set(name, { story, id, inherited: false });
  }

  return declared;
}

/** A section's slots: the previous section's, as inherited ones, under the section's own. */
function inherit(
  before: ReadonlyMap<OrderedFurnitureVariant, OrderedFurnitureSlot> | undefined,
  own: FurnitureSlots
): FurnitureSlots {
  const slots: FurnitureSlots = new Map();
  for (const [variant, slot] of before ?? []) slots.set(variant, { ...slot, inherited: true });
  for (const [variant, slot] of own) slots.set(variant, slot);
  return slots;
}

/**
 * The headers and footers each section shows, as the engine resolves them
 * (`resolveHeaderFooterResolutionBySection`): a section's own references, the first of each
 * variant, then the previous section's for the variants it does not declare. A section names a
 * story through a relationship of the main part: the file's own (`loadedFurnitureOf`), or one a
 * change added (design D14).
 */
export function furnitureOf(
  doc: OrderedDocument,
  main: OrderedStory,
  sections: readonly OrderedSectionView[]
): OrderedSectionFurniture[] {
  const out: OrderedSectionFurniture[] = [];
  let previous: OrderedSectionFurniture | undefined;
  for (const section of sections.length > 0 ? sections : [null]) {
    const declared = declaredFurniture(doc, main, section?.sectPr);
    const resolved: OrderedSectionFurniture = {
      headers: inherit(previous?.headers, declared.header),
      footers: inherit(previous?.footers, declared.footer),
      titlePage: declared.titlePage,
    };
    out.push(resolved);
    previous = resolved;
  }

  return out;
}

/** What a header or footer action names: the section's holder, the story kind and variant. */
export interface OrderedFurnitureTarget {
  readonly section: OrderedItemAddress;
  readonly story: OrderedFurnitureKind;
  readonly variant: OrderedFurnitureVariant;
}

const refuse = (op: string, detail: string): never => {
  throw new OrderedRefusal('engine-refused', `${op}: ${detail}`);
};

function slotOf(ctx: Ctx, target: OrderedFurnitureTarget) {
  const sections = sectionsNow(ctx);
  const { index, section } = sectionAt(ctx, target.section);
  const furniture = furnitureOf(ctx.doc, ctx.story, sections)[index];
  const slots = target.story === 'header' ? furniture?.headers : furniture?.footers;
  return { index, section, slot: slots?.get(target.variant) };
}

/**
 * A new, empty header or footer for a section's variant (`createHeaderFooter`): a story part
 * and its relationship from the main part (D14), and the section's reference to it. With
 * `titlePage` the section also takes a different first page.
 *
 * Word has no change record for a reference (`CT_SectPrChange` holds `CT_SectPrBase`, which has
 * no `w:headerReference` or `w:footerReference`, ECMA-376 Part 1, §17.13.5.32). So this edit,
 * `removeFurniture` and `unlinkFurniture` are direct edits only: the organiser's, accepted at
 * once and taken back with `revert`.
 */
export function createFurniture(
  ctx: Ctx,
  target: OrderedFurnitureTarget,
  suggestion: OrderedSuggestion,
  titlePage: boolean
): void {
  const { section, slot } = slotOf(ctx, target);
  if (slot) refuse('createHeaderFooter', 'slot-occupied');
  const { resource } = addStory(ctx, suggestion.id, target.story);
  const relationship = addRelationship(ctx, suggestion.id, target.story, {
    resource: resource.key,
  });
  addSectionWrite(
    ctx,
    section.holder,
    { op: 'reference', kind: target.story, variant: target.variant, id: relationship.id },
    suggestion
  );
  if (titlePage)
    addSectionWrite(ctx, section.holder, { op: 'furniture', titlePage: true }, suggestion);
}

/**
 * A section's own header or footer reference goes: `deleteHeaderFooter`, or `linkToPrevious`,
 * after which the section shows the previous section's. A loaded part nothing names any more is
 * left out of the package, as the engine removes it. A direct edit only (see `createFurniture`).
 */
export function removeFurniture(
  ctx: Ctx,
  target: OrderedFurnitureTarget,
  suggestion: OrderedSuggestion,
  link: boolean
): void {
  const op = link ? 'linkToPrevious' : 'deleteHeaderFooter';
  const { index, section, slot } = slotOf(ctx, target);
  if (link && index === 0) refuse(op, 'first-section');
  if (!slot || slot.inherited) refuse(op, 'not-declared');
  addSectionWrite(
    ctx,
    section.holder,
    { op: 'reference', kind: target.story, variant: target.variant, id: null },
    suggestion
  );
}

/**
 * The section gets its own copy of the header or footer it showed from the section before
 * (`unlinkFromPrevious`): a new story whose items copy the shown story's, under new addresses,
 * a relationship to it, and the section's reference. A direct edit only (see `createFurniture`).
 */
export function unlinkFurniture(
  ctx: Ctx,
  target: OrderedFurnitureTarget,
  suggestion: OrderedSuggestion
): void {
  const { section, slot } = slotOf(ctx, target);
  if (!slot) return refuse('unlinkFromPrevious', 'nothing-to-unlink');
  if (!slot.inherited) refuse('unlinkFromPrevious', 'already-declared');
  const source = ctx.doc.stories.get(slot.story);
  if (!source) return refuse('unlinkFromPrevious', 'missing-source');
  const resource = copyStory(ctx, suggestion, source, target.story);
  const relationship = addRelationship(ctx, suggestion.id, target.story, {
    resource: resource.key,
  });
  addSectionWrite(
    ctx,
    section.holder,
    { op: 'reference', kind: target.story, variant: target.variant, id: relationship.id },
    suggestion
  );
}

// ---- a copy of a story ----

/**
 * The relationships another change added from the source part go with the copy, under their
 * ids.
 */
function copyRelationships(
  ctx: Ctx,
  suggestion: OrderedSuggestion,
  from: string,
  to: string
): void {
  for (const resource of [...ctx.doc.resources.added.values()]) {
    if (resource.kind !== 'relationship' || resource.owner !== from) continue;
    if (!orderedResourceShows(ctx.doc, resource)) continue;
    const copy: OrderedRelationshipResource = {
      ...resource,
      key: ctx.ids.resource(),
      suggestion: suggestion.id,
      owner: to,
    };
    addOrderedResource(ctx.doc, copy);
  }
}

/** New spaces for a copy's items, minted from the change in the order the copy meets them. */
interface SpaceMinter {
  /** The copy's space of each source space met so far. */
  readonly spaces: ReadonlyMap<string, string>;
  /** The Word ids minted for each new paragraph mark's space. */
  readonly identities: ReadonlyMap<string, OrderedMintedParagraph>;
  /** The copy's space of an item's space, minted the first time. */
  readonly spaceOf: (item: OrderedItem) => string;
}

function spaceMinter(ctx: Ctx): SpaceMinter {
  const spaces = new Map<string, string>();
  const identities = new Map<string, OrderedMintedParagraph>();
  const spaceOf = (item: OrderedItem): string => {
    const found = spaces.get(item.space);
    if (found !== undefined) {
      return found;
    }

    let next: string;
    if (item.kind === 'mark') {
      const paragraph = mintParagraph(ctx);
      next = paragraph.space;
      identities.set(next, paragraph);
    } else if (item.kind === 'open' || item.kind === 'close') {
      next = ctx.ids.container();
    } else if (item.kind === 'zero') {
      next = ctx.ids.marker();
    } else {
      next = ctx.ids.units();
    }

    spaces.set(item.space, next);
    return next;
  };
  return { spaces, identities, spaceOf };
}

/** A copy of a node with ids from the change; text and no node stay as they are. */
function copyOf<T extends OoxmlNode | null | undefined>(ctx: Ctx, value: T): T {
  return value && value.kind !== 'textValue'
    ? (cloneWithNewIds(value, () => ctx.ids.node()) as T)
    : value;
}

/** A field slot under the copy's spaces: it names the copy of its field, at the same offset. */
function copiedFieldSlot(
  slot: NonNullable<OrderedItem['fieldSlot']>,
  spaces: ReadonlyMap<string, string>
): NonNullable<OrderedItem['fieldSlot']> {
  const [fieldSpace = '', offset = '0'] = slot.field.split('|');
  return { field: `${spaces.get(fieldSpace) ?? fieldSpace}|${offset}`, before: slot.before };
}

/**
 * An item of the source again under a new address, with its labels, inside the copy's own
 * insertion (`within`). Node ids are new; a paragraph takes new Word ids.
 */
function copiedItem(
  ctx: Ctx,
  suggestion: OrderedSuggestion,
  minter: SpaceMinter,
  item: OrderedItem
): OrderedItem {
  const space = minter.spaceOf(item);
  const pair =
    item.pair === undefined
      ? undefined
      : (minter.spaces.get(item.pair) ?? minter.spaceOf({ ...item, space: item.pair }));
  let shell = copyOf(ctx, item.node);
  const identity = minter.identities.get(space);
  if (item.kind === 'mark' && identity && shell && shell.kind !== 'textValue') {
    shell = withId(withParagraphIdentity(shell, paragraphIdAttributes('w14', identity)), space);
  }

  const fieldSlot = item.fieldSlot ? copiedFieldSlot(item.fieldSlot, minter.spaces) : undefined;
  return createOrderedItem({
    ...item,
    space,
    node: shell,
    face: copyOf(ctx, item.face),
    ...(item.textShell ? { textShell: copyOf(ctx, item.textShell) } : {}),
    ...(item.pieces
      ? { pieces: item.pieces.map((piece) => ({ ...piece, node: copyOf(ctx, piece.node) })) }
      : {}),
    ...(pair === undefined ? {} : { pair }),
    ...(fieldSlot === undefined ? {} : { fieldSlot }),
    home: undefined,
    unitCount: undefined,
    within: { by: [...item.within.by, suggestion.id], dels: item.within.dels },
  });
}

/**
 * A new story that copies `source`: every item again under new addresses minted from the change,
 * so a take-back of the copy hides it all.
 */
function copyStory(
  ctx: Ctx,
  suggestion: OrderedSuggestion,
  source: OrderedStory,
  role: OrderedFurnitureKind
): OrderedStoryResource {
  const doc = ctx.doc;

  // A loaded source's relationships are the file's: the copy takes them as the engine does
  const loaded = !isAddedStory(doc, source.name);
  const resource = addStoryResource(ctx, suggestion.id, role, loaded ? source.name : undefined);
  const name = resource.part;
  copyRelationships(ctx, suggestion, source.name, name);
  const minter = spaceMinter(ctx);
  const items = [...source.tree.items()].map((item) => copiedItem(ctx, suggestion, minter, item));
  addOrderedStory(
    doc,
    name,
    { ...source.part, name, root: withChildren(source.part.root, []) },
    items
  );
  recordUndo(doc, () => {
    doc.stories.delete(name);
    for (const item of items) labelUnindex(doc, item);
  });
  return resource;
}
