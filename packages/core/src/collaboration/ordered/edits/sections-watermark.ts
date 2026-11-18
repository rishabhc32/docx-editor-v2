// The watermark (task 9.8), as the engine's `setWatermark` writes it
// (`store/package/watermark-package.ts`).

import { headerFooterVariantCanPaint } from '../../../store/package/hf-references.ts';
import { isLegacyVmlAtom } from '../../../store/package/legacy-vml-projection.ts';
import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import {
  watermarkRun,
  watermarkShape,
  type PackageWatermark,
} from '../../../store/package/watermark-package.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import type { OrderedChangeWatermark } from '../change/kinds/sections/watermark.ts';
import { orderedParagraphCount, orderedViewParagraph } from '../compose/compose.ts';
import { loadedFurnitureOf } from '../load/sections.ts';
import { createOrderedItem, isHidden, statusOf, storyNamed } from '../model/document.ts';
import { insertItems, relabel } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedFurnitureVariant } from '../model/sections.ts';
import type { OrderedItem, OrderedStory, OrderedSuggestion } from '../model/types.ts';
import { childWhere, isRunProperties, isWml, withChildren } from '../ooxml-nodes.ts';
import type { Ctx } from './context.ts';
import { mintParagraph, paragraphIdAttributes } from './para-ids.ts';
import { addMedia, addRelationship } from './resources.ts';
import {
  createFurniture,
  furnitureOf,
  type OrderedSectionFurniture,
} from './sections-furniture.ts';
import { sectionsNow } from './sections.ts';

const VARIANTS: readonly OrderedFurnitureVariant[] = ['default', 'first', 'even'];

/**
 * A picture of a run as the list holds it: one unit when the engine draws it (a VML shape it
 * projects), else no width, as the engine's model text counts it (`tree-op-segments.ts`).
 */
const pictureItem = (node: OoxmlElement) =>
  isLegacyVmlAtom(node) ? { kind: 'atom' as const, text: '\uFFFC' } : { kind: 'zero' as const };

/** Whether the document separates even and odd pages now: the file's setting, then the writes. */
function evenAndOddHeadersNow(ctx: Ctx): boolean {
  let on = loadedFurnitureOf(ctx.doc).evenAndOddHeaders;

  // Settings keep the room's order, and the last accepted write wins (design D14)
  for (const resource of ctx.doc.resources.added.values())
    if (
      resource.kind === 'setting' &&
      resource.op.op === 'setSectionFurnitureOptions' &&
      statusOf(ctx.doc, resource.suggestion) === 'accepted'
    )
      on = resource.op.evenAndOddHeaders;
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

/**
 * Take the watermark out of one header: each item that shows one is hidden by the suggestion,
 * and what else its node held (another shape in the same `w:pict`) comes back as a new item in
 * the same run, as the engine removes the watermark's shape alone.
 */
function removeWatermark(ctx: Ctx, story: OrderedStory, suggestion: OrderedSuggestion): void {
  const found = [...story.tree.items()].filter(
    (item) => item.node !== null && !isHidden(ctx.doc, item) && holdsWatermark(item.node)
  );
  for (const item of found) {
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

/** The engine's watermark as it writes it, from the change's typed values. */
function packageWatermark(write: OrderedChangeWatermark): PackageWatermark {
  if (write.kind === 'picture') {
    const { media: _media, sectionSizes: _sizes, ...picture } = write;
    return picture;
  }

  const { sectionSizes: _sizes, ...text } = write;
  return text;
}

/**
 * The first paragraph of a story outside every container, as the engine's `firstParagraph`
 * finds it; null when the story starts with none.
 */
function firstParagraphItems(ctx: Ctx, story: OrderedStory): readonly OrderedItem[] | null {
  for (let index = 0; index < orderedParagraphCount(ctx.doc, story.name); index += 1) {
    const view = orderedViewParagraph(ctx.doc, story.name, index);
    if (view && view.blocks.length === 0) return [...view.items, view.mark];
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
  const run = watermarkRun(story.part, watermark, ordinal, relationshipId, () => ctx.ids.node());
  if (run === null || run.kind === 'textValue')
    throw new OrderedRefusal('engine-refused', 'setWatermark: watermark-markup');
  const face = childWhere(run, isRunProperties) ?? null;
  const shape = childWhere(run, (child) => isWml(child, 'pict'));
  if (!shape) throw new OrderedRefusal('engine-refused', 'setWatermark: watermark-markup');

  // The run declares the VML namespaces its picture uses: the picture keeps them without it
  const pict: OoxmlElement = {
    ...shape,
    namespaceBindings: [...run.namespaceBindings, ...shape.namespaceBindings],
  };
  const first = (firstParagraphItems(ctx, story) ?? [newParagraph(ctx, story, suggestion)])[0];
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
 * Set the document's watermark, or remove it (`null`), as one direct edit of the suggestion.
 * The change's story is the main document's. A section that shows no header of a variant that
 * can paint gets a new one first; the old watermark goes from every header a section names; then
 * every header a section shows takes the new shape at the start of its first paragraph. In the
 * list a watermark is the item of its `w:pict`, so the edit hides old items and adds new ones,
 * and the organiser's undo (`revert`) shows the old ones again exactly.
 */
export function setWatermark(
  ctx: Ctx,
  suggestion: OrderedSuggestion,
  write: OrderedChangeWatermark | null
): void {
  if (ctx.story.part.root.localName !== 'document')
    throw new OrderedRefusal('invalid-edit', "a watermark is the main document's");
  const evenAndOdd = evenAndOddHeadersNow(ctx);
  if (write !== null) createShownHeaders(ctx, suggestion, evenAndOdd);
  const sections = furnitureOf(ctx.doc, ctx.story, sectionsNow(ctx));

  // Every header a section names, its own or the one it takes from the section before
  const named = new Set(sections.flatMap((section) => [...section.headers.values()]));
  for (const story of new Set([...named].map((slot) => slot.story)))
    removeWatermark(ctx, storyNamed(ctx.doc, story), suggestion);
  if (write === null) return;
  const media = write.kind === 'picture' ? addMedia(ctx, suggestion.id, write.media) : null;
  const watermark = packageWatermark(write);

  // Every header a section can paint, once, sized for the first section that shows it
  const shown = new Map<string, number>();
  sections.forEach((section, index) => {
    for (const variant of VARIANTS) {
      const slot = section.headers.get(variant);
      if (slot && canPaint(section, variant, evenAndOdd) && !shown.has(slot.story))
        shown.set(slot.story, index);
    }
  });
  let ordinal = 0;
  for (const [name, sectionIndex] of shown) {
    const relationship =
      media === null
        ? null
        : addRelationship(ctx, suggestion.id, 'image', { resource: media.key }, name);

    // A header takes the size of the first section that shows it (`sizedForSection`)
    const size = write.sectionSizes?.[sectionIndex] ?? null;
    const sized = size === null ? watermark : { ...watermark, ...size };
    addWatermark(
      ctx,
      storyNamed(ctx.doc, name),
      suggestion,
      sized,
      ordinal,
      relationship?.id ?? null
    );
    ordinal += 1;
  }
}
