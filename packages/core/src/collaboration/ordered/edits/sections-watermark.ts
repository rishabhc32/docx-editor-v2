// The watermark (task 9.8), as the engine's `setWatermark` writes it
// (`store/package/watermark-package.ts`).

import { headerFooterVariantCanPaint } from '../../../store/package/hf-references.ts';
import { isLegacyVmlAtom } from '../../../store/package/legacy-vml-projection.ts';
import { ooxmlTreesEqual } from '../../../store/package/ooxml-serialize.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import {
  watermarkRun,
  watermarkShape,
  type PackageWatermark,
} from '../../../store/package/watermark-package.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import type { OrderedChangeWatermark } from '../change/kinds/sections/watermark.ts';
import { orderedParagraphCount, orderedViewParagraph } from '../compose/compose.ts';
import { loadedFurnitureOf } from '../load/sections.ts';
import {
  createOrderedItem,
  isHidden,
  isPending,
  statusOf,
  storyNamed,
  suggestionsOf,
} from '../model/document.ts';
import { insertItems, relabel } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedFurnitureVariant } from '../model/sections.ts';
import type {
  OrderedDocument,
  OrderedItem,
  OrderedStory,
  OrderedSuggestion,
} from '../model/types.ts';
import { childWhere, isRunProperties, isWml, withChildren } from '../ooxml-nodes.ts';
import { suggestionFor, type Ctx } from './context.ts';
import { mintParagraph, paragraphIdAttributes } from './para-ids.ts';
import { addMedia, addRelationship } from './resources.ts';
import {
  createFurniture,
  furnitureOf,
  type OrderedSectionFurniture,
} from './sections-furniture.ts';
import { sectionsNow, sectionsOfStory } from './sections.ts';

const VARIANTS: readonly OrderedFurnitureVariant[] = ['default', 'first', 'even'];

/**
 * A picture of a run as the list holds it: one unit when the engine draws it (a VML shape it
 * projects), else no width, as the engine's model text counts it (`tree-op-segments.ts`).
 */
const pictureItem = (node: OoxmlElement) =>
  isLegacyVmlAtom(node) ? { kind: 'atom' as const, text: '\uFFFC' } : { kind: 'zero' as const };

/** Whether the document separates even and odd pages now: the file's setting, then the writes. */
function evenAndOddHeadersNow(doc: OrderedDocument): boolean {
  let on = loadedFurnitureOf(doc).evenAndOddHeaders;

  // Settings keep the room's order, and the last accepted write wins (design D14)
  for (const resource of doc.resources.added.values()) {
    if (
      resource.kind === 'setting' &&
      resource.op.op === 'setSectionFurnitureOptions' &&
      statusOf(doc, resource.suggestion) === 'accepted'
    ) {
      on = resource.op.evenAndOddHeaders;
    }
  }

  return on;
}

const canPaint = (
  section: OrderedSectionFurniture,
  variant: OrderedFurnitureVariant,
  evenAndOddHeaders: boolean
) => headerFooterVariantCanPaint({ titlePage: section.titlePage, evenAndOddHeaders }, variant);

/** A new header for each variant a section can paint and shows none of (`withShownHeaders`). */
function createShownHeaders(ctx: Ctx, suggestion: OrderedSuggestion, evenAndOdd: boolean): void {
  let sections = sectionsNow(ctx);
  let headers = furnitureOf(ctx.doc, ctx.story, sections);
  for (let index = 0; index < sections.length; index += 1)
    for (const variant of VARIANTS) {
      const section = headers[index];
      const holder = sections[index]?.holder;
      if (!section || !holder || !canPaint(section, variant, evenAndOdd)) continue;
      if (section.headers.has(variant)) continue;
      const target = {
        section: { space: holder.space, offset: holder.offset },
        story: 'header' as const,
        variant,
      };
      createFurniture(ctx, target, suggestion, false);

      // A new header is what the sections after it show too
      sections = sectionsNow(ctx);
      headers = furnitureOf(ctx.doc, ctx.story, sections);
    }
}

const isVml = (uri: string) => uri === 'urn:schemas-microsoft-com:vml';

/**
 * The node without its watermark shapes; null when nothing else is left in it: a `w:pict` that
 * holds only shape types goes whole, as the engine and Word remove a watermark's own run.
 */
function withoutWatermark(node: OoxmlElement): OoxmlElement | null {
  const strip = (element: OoxmlElement): OoxmlElement =>
    withChildren(
      element,
      element.children
        .filter((child) => child.kind === 'textValue' || watermarkShape(child) === null)
        .map((child) => (child.kind === 'textValue' ? child : strip(child)))
    );
  const stripped = strip(node);
  const rest = stripped.children.filter(
    (child) =>
      child.kind !== 'textValue' && !(child.localName === 'shapetype' && isVml(child.namespaceUri))
  );
  return isWml(node, 'pict') && rest.length === 0 ? null : stripped;
}

/** Whether a node holds a watermark shape anywhere in it. */
function holdsWatermark(node: OoxmlNode, depth = 0): boolean {
  if (node.kind === 'textValue' || depth > 64) return false;
  if (watermarkShape(node) !== null) return true;
  return node.children.some((child) => holdsWatermark(child, depth + 1));
}

/** The shown items of a story that hold a watermark shape. */
const watermarkItems = (doc: OrderedDocument, story: OrderedStory): OrderedItem[] =>
  [...story.tree.items()].filter(
    (item) => item.node !== null && !isHidden(doc, item) && holdsWatermark(item.node)
  );

/** Every header a section names, its own or the one it takes from the section before. */
const namedHeaders = (sections: readonly OrderedSectionFurniture[]): Set<string> =>
  new Set(sections.flatMap((section) => [...section.headers.values()].map((slot) => slot.story)));

/** Every header a section can paint, once, with the first section that shows it. */
function paintedHeaders(
  sections: readonly OrderedSectionFurniture[],
  evenAndOdd: boolean
): Map<string, number> {
  const painted = new Map<string, number>();
  sections.forEach((section, index) => {
    for (const variant of VARIANTS) {
      const slot = section.headers.get(variant);
      if (slot && canPaint(section, variant, evenAndOdd) && !painted.has(slot.story)) {
        painted.set(slot.story, index);
      }
    }
  });
  return painted;
}

/** Whether a header a section of the story names shows a watermark. */
function showsWatermark(doc: OrderedDocument, story: OrderedStory): boolean {
  const sections = furnitureOf(doc, story, sectionsOfStory(doc, story));
  return [...namedHeaders(sections)].some(
    (name) => watermarkItems(doc, storyNamed(doc, name)).length > 0
  );
}

/**
 * Take the watermark out of one header: each item that shows one is hidden by the suggestion,
 * and what else its node held (another shape in the same `w:pict`) comes back as a new item in
 * the same run, as the engine removes the watermark's shape alone.
 */
function removeWatermark(ctx: Ctx, story: OrderedStory, suggestion: OrderedSuggestion): void {
  for (const item of watermarkItems(ctx.doc, story)) {
    const node = item.node;
    if (node === null || node.kind === 'textValue') continue;
    const rest = withoutWatermark(node);
    if (rest !== null)
      insertItems(ctx.doc, story, item, [
        createOrderedItem({
          space: ctx.ids.marker(),
          ...pictureItem(rest),
          node: rest,
          face: item.face,
          run: item.run,
          ...(item.inRun ? { inRun: true } : {}),
          ...(item.bare ? { bare: true } : {}),
          within: item.within,
          by: suggestion.id,
        }),
      ]);
    relabel(ctx.doc, story, item, { dels: [...item.dels, suggestion.id] });
  }
}

/**
 * The engine's watermark as it writes it, from the change's typed values, sized for the first
 * section that shows its header (`sizedForSection`).
 */
function packageWatermark(write: OrderedChangeWatermark, sectionIndex: number): PackageWatermark {
  const size = write.sectionSizes?.[sectionIndex] ?? null;
  if (write.kind === 'picture') {
    const { media: _media, sectionSizes: _sizes, ...picture } = write;
    return size === null ? picture : { ...picture, ...size };
  }

  const { sectionSizes: _sizes, ...text } = write;
  return size === null ? text : { ...text, ...size };
}

/**
 * The watermark's run as the engine builds it, as the list holds it: its `w:pict`, which keeps
 * the VML namespaces the run declares, and its `w:rPr`.
 */
function watermarkPicture(
  part: OoxmlPart,
  watermark: PackageWatermark,
  ordinal: number,
  relationshipId: string | null,
  mint: () => string
): { readonly pict: OoxmlElement; readonly face: OoxmlElement | null } {
  const run = watermarkRun(part, watermark, ordinal, relationshipId, mint);
  if (run === null || run.kind === 'textValue') {
    throw new OrderedRefusal('engine-refused', 'setWatermark: watermark-markup');
  }

  const shape = childWhere(run, (child) => isWml(child, 'pict'));
  if (!shape) {
    throw new OrderedRefusal('engine-refused', 'setWatermark: watermark-markup');
  }

  return {
    pict: { ...shape, namespaceBindings: [...run.namespaceBindings, ...shape.namespaceBindings] },
    face: childWhere(run, isRunProperties) ?? null,
  };
}

/**
 * The first paragraph of a story outside every container, as the engine's `firstParagraph`
 * finds it; null when the story starts with none.
 */
function firstParagraphItems(
  doc: OrderedDocument,
  story: OrderedStory
): readonly OrderedItem[] | null {
  for (let index = 0; index < orderedParagraphCount(doc, story.name); index += 1) {
    const view = orderedViewParagraph(doc, story.name, index);
    if (view && view.blocks.length === 0) {
      return [...view.items, view.mark];
    }
  }

  return null;
}

/** A new empty paragraph at the start of a story, inserted by the suggestion. */
function newParagraph(ctx: Ctx, story: OrderedStory, suggestion: OrderedSuggestion): OrderedItem {
  const minted = mintParagraph(ctx);
  const mark = createOrderedItem({
    space: minted.space,
    kind: 'mark',
    node: build(minted.space, 'paragraph', 'p', paragraphIdAttributes('w14', minted), []),
    by: suggestion.id,
  });
  insertItems(ctx.doc, story, null, [mark]);
  return mark;
}

/**
 * Add the watermark at the start of one header's first paragraph (a new paragraph when it has
 * none), as the run the engine builds: its `w:pict` as a zero-width item, its `w:rPr` as the
 * item's face.
 */
function addWatermark(
  ctx: Ctx,
  story: OrderedStory,
  suggestion: OrderedSuggestion,
  watermark: PackageWatermark,
  ordinal: number,
  relationshipId: string | null
): void {
  const { pict, face } = watermarkPicture(story.part, watermark, ordinal, relationshipId, () =>
    ctx.ids.node()
  );
  const first = (firstParagraphItems(ctx.doc, story) ?? [newParagraph(ctx, story, suggestion)])[0];
  const space = ctx.ids.marker();
  const item = createOrderedItem({
    space,
    ...pictureItem(pict),
    node: pict,
    face,
    run: space,
    inRun: true,
    by: suggestion.id,
  });
  insertItems(ctx.doc, story, first ? story.tree.prev(first) : null, [item]);
}

/**
 * Whether the header a section paints shows `write` already, as the edit would write it: alone,
 * at its first paragraph's start, with nothing pending on it.
 */
function showsAlready(
  doc: OrderedDocument,
  header: OrderedStory,
  write: OrderedChangeWatermark,
  sectionIndex: number,
  ordinal: number
): boolean {
  const items = watermarkItems(doc, header);
  const item = items.length === 1 ? items[0] : undefined;
  if (item === undefined || item.node === null || item !== firstParagraphItems(doc, header)?.[0]) {
    return false;
  }

  if ([...suggestionsOf(item)].some((id) => isPending(doc, id))) {
    return false;
  }

  let scratch = 0;
  const watermark = packageWatermark(write, sectionIndex);
  const { pict, face } = watermarkPicture(header.part, watermark, ordinal, null, () => {
    scratch += 1;
    return `check#${scratch}`;
  });
  const sameFace =
    face === null || item.face === null ? face === item.face : ooxmlTreesEqual(face, item.face);
  return sameFace && ooxmlTreesEqual(pict, item.node);
}

/**
 * Whether the watermark write changes a header: a removal when one shows a watermark, a new
 * watermark unless the edit would make no header and every header it paints shows the same one
 * already (`showsAlready`), and no other named header holds one. A picture always changes them:
 * the edit adds its relationship again. A write that changes none makes no suggestion.
 */
export function watermarkChanges(
  doc: OrderedDocument,
  story: OrderedStory,
  write: OrderedChangeWatermark | null
): boolean {
  if (write === null) {
    return showsWatermark(doc, story);
  }

  if (write.kind === 'picture') {
    return true;
  }

  const evenAndOdd = evenAndOddHeadersNow(doc);
  const sections = furnitureOf(doc, story, sectionsOfStory(doc, story));
  const painted = paintedHeaders(sections, evenAndOdd);
  const missing = sections.some((section) =>
    VARIANTS.some(
      (variant) => canPaint(section, variant, evenAndOdd) && !section.headers.has(variant)
    )
  );
  if (missing) {
    return true;
  }

  for (const name of namedHeaders(sections)) {
    if (!painted.has(name) && watermarkItems(doc, storyNamed(doc, name)).length > 0) {
      return true;
    }
  }

  return ![...painted].every(([name, sectionIndex], ordinal) =>
    showsAlready(doc, storyNamed(doc, name), write, sectionIndex, ordinal)
  );
}

/**
 * Set the document's watermark, or remove it (`null`), as one direct edit of a new suggestion,
 * which it returns; a write that changes no header makes none (`watermarkChanges`). The change's
 * story is the main document's. A section that shows no header of a variant that can paint gets
 * a new one first; the old watermark goes from every header a section names; then every header a
 * section shows takes the new shape at the start of its first paragraph. In the list a watermark
 * is the item of its `w:pict`, so the edit hides old items and adds new ones, and the organiser's
 * undo (`revert`) shows the old ones again exactly.
 */
export function setWatermark(ctx: Ctx, write: OrderedChangeWatermark | null): string | null {
  if (ctx.story.part.root.localName !== 'document')
    throw new OrderedRefusal('invalid-edit', "a watermark is the main document's");
  if (!watermarkChanges(ctx.doc, ctx.story, write)) {
    return null;
  }

  const suggestion = suggestionFor(ctx);
  const evenAndOdd = evenAndOddHeadersNow(ctx.doc);
  if (write !== null) {
    createShownHeaders(ctx, suggestion, evenAndOdd);
  }

  const sections = furnitureOf(ctx.doc, ctx.story, sectionsNow(ctx));
  for (const story of namedHeaders(sections)) {
    removeWatermark(ctx, storyNamed(ctx.doc, story), suggestion);
  }

  if (write === null) {
    return suggestion.id;
  }

  const media = write.kind === 'picture' ? addMedia(ctx, suggestion.id, write.media) : null;

  // Every header a section can paint, once, sized for the first section that shows it
  let ordinal = 0;
  for (const [name, sectionIndex] of paintedHeaders(sections, evenAndOdd)) {
    const relationship =
      media === null
        ? null
        : addRelationship(ctx, suggestion.id, 'image', { resource: media.key }, name);

    addWatermark(
      ctx,
      storyNamed(ctx.doc, name),
      suggestion,
      packageWatermark(write, sectionIndex),
      ordinal,
      relationship?.id ?? null
    );
    ordinal += 1;
  }

  return suggestion.id;
}
