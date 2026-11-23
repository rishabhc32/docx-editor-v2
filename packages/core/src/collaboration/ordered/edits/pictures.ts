// The model edits of group 9.4: a new picture or text box at a place, and a direct change of a
// picture's or a shape's properties.

import { buildInlinePictureDrawing } from '../../../store/package/drawing-package-edit.ts';
import type {
  DrawingLocksInput,
  DrawingPositionInput,
  ImageWrapTarget,
  SourceCrop,
} from '../../../store/package/drawing-projection.ts';
import { findNode } from '../../../store/package/ooxml-edit.ts';
import type { OoxmlDrawingNode, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { EMU_PER_POINT } from '../../../store/package/drawing-package-edit.ts';
import { buildTextboxDrawing } from '../../../store/package/textbox-drawing-builder.ts';
import { applyTreeOp } from '../../../store/store/tree-op-apply.ts';
import { cloneWithNewIds } from '../../../store/store/tree-op-nodes.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import type { OrderedChangeMedia } from '../change/resources.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import { objectAfterOps, objectNodeOf } from '../compose/objects.ts';
import { withStoryBlocks } from '../compose/paragraph-ops.ts';
import { isHidden } from '../model/document.ts';
import { newSuggestion, nextSeq, relabel } from '../model/mutations.ts';
import type { OrderedObjectOp, OrderedObjectRecord } from '../model/objects.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { orderedRelationshipId } from '../model/resources.ts';
import type { OrderedItem } from '../model/types.ts';
import { refuseLockedAt } from '../placement.ts';
import { scratchCompose, type Ctx } from './context.ts';
import { directInsert, settle } from './direct.ts';
import { resolveOrderedPlace, type OrderedPlace } from './places.ts';
import { addMedia, addRelationshipResource } from './resources.ts';
import { insertText } from './text.ts';

/**
 * A new object, as a change names it: an inline picture (the editor's Insert Picture), or a
 * text box (the editor's Insert Text Box). Every value is typed and bounded by the validator.
 *
 * @public
 */
export type OrderedNewObject =
  | {
      readonly type: 'picture';
      /** The image, named by its SHA-256: the room stores the bytes (design D14). */
      readonly media: OrderedChangeMedia;
      /** The size in EMU. */
      readonly extent: { readonly cx: number; readonly cy: number };
      /** `wp:docPr/@title`. */
      readonly title?: string;
      /** `wp:docPr/@descr`: the alt text. */
      readonly description?: string;
      /** The address the picture's link opens, kept as `sanitizeHref` keeps it. */
      readonly link?: string;
      /** The aspect ratio is locked (`a:graphicFrameLocks/@noChangeAspect`). */
      readonly lockAspect?: boolean;
    }
  | {
      readonly type: 'textbox';
      /** The width in EMU; the height is the engine's default. */
      readonly extent: { readonly cx: number };
    };

/**
 * One property change of an object as a change names it: the model's op, with a new link as its
 * address and a new image as its media in place of resource keys.
 *
 * @public
 */
export type OrderedObjectChangeOp =
  | { readonly op: 'resize'; readonly cx: number; readonly cy: number }
  | { readonly op: 'crop'; readonly crop: SourceCrop }
  | { readonly op: 'position'; readonly position: DrawingPositionInput }
  | { readonly op: 'wrap'; readonly wrap: ImageWrapTarget }
  | {
      readonly op: 'metadata';
      readonly title: string;
      readonly description: string;
      /** A new link's address; `null` removes the link; absent keeps it. */
      readonly link?: string | null;
    }
  | { readonly op: 'locks'; readonly locks: DrawingLocksInput }
  | { readonly op: 'transform'; readonly action: 'rotateCW' | 'rotateCCW' | 'flipH' | 'flipV' }
  | { readonly op: 'image'; readonly media: OrderedChangeMedia };

/** Resource keys minted before the element is built: it names them by id. */
interface ObjectKeys {
  readonly media: string;
  readonly image: string;
  readonly link: string | null;
}

/**
 * The engine's own element for a new drawing (its `insertDrawing` run on a scratch part), under
 * ids from the change: a change carries typed values, never XML.
 */
function engineElement(ctx: Ctx, drawing: OoxmlNode): OoxmlNode {
  let scratchIds = 0;
  const mint = () => `ordered-object#${(scratchIds += 1)}`;
  const paragraph = build(mint(), 'paragraph', 'p', [], []);
  const scratch = withStoryBlocks(ctx.story.part, [paragraph]);
  const result = applyTreeOp(
    scratch,
    {
      op: 'insertDrawing',
      paragraphId: paragraph.id,
      offset: 0,
      drawing: drawing as OoxmlDrawingNode,
    },
    { deferValidation: true }
  );
  if (!result.ok) throw new OrderedRefusal('engine-refused', `insertDrawing: ${result.reason}`);
  const made = result.effect.created
    .map((id) => findNode(result.part, id))
    .find((node) => node?.kind === 'drawing');
  if (!made) throw new OrderedRefusal('engine-refused', 'insertDrawing made no drawing');
  return cloneWithNewIds(made, () => ctx.ids.node());
}

/** The element of a new object, before its ids are the change's. */
function builtElement(ctx: Ctx, object: OrderedNewObject, keys: ObjectKeys | null): OoxmlNode {
  // `wp:docPr/@id` is document-wide: a decimal from the change, above any file's dense ids
  const docPrId = ctx.ids.decimal();
  if (object.type === 'textbox')
    return buildTextboxDrawing(docPrId, object.extent.cx / EMU_PER_POINT);
  if (keys === null) throw new OrderedRefusal('invalid-edit', 'a picture needs its media');
  return buildInlinePictureDrawing({
    docPrId,
    relationshipId: orderedRelationshipId(keys.image),
    extentEmu: object.extent,
    ...(object.title === undefined ? {} : { title: object.title }),
    ...(object.description === undefined ? {} : { description: object.description }),
    ...(keys.link === null ? {} : { hyperlinkRelationshipId: orderedRelationshipId(keys.link) }),
    ...(object.lockAspect === true ? { lockAspectRatio: true } : {}),
  });
}

/** The media and image relationship resources of a picture, under keys minted earlier. */
function addImage(
  ctx: Ctx,
  suggestion: string,
  media: OrderedChangeMedia,
  keys: { readonly media: string; readonly image: string }
): void {
  addMedia(ctx, suggestion, media, keys.media);
  addRelationshipResource(ctx, keys.image, suggestion, ctx.story.name, 'image', {
    resource: keys.media,
  });
}

/**
 * The suggestion that inserted the object whose element is `node`, in the shown paragraph at
 * `index`.
 */
function objectSuggestion(ctx: Ctx, index: number, node: OoxmlNode): string {
  const view = orderedViewParagraph(ctx.doc, ctx.story.name, index);
  const atom = view?.chars.find((char) => char.item.node === node)?.item;
  if (!atom || atom.by === null)
    throw new OrderedRefusal('malformed-story', 'the new object is not in its paragraph');
  return atom.by;
}

/**
 * A new picture or text box before `place`, in a run with no run properties of its own, as the
 * engine's `insertDrawing` writes it. Its one atom is placed by the rules typed text follows.
 * Word tracks it as an insertion of the run that holds it (`w:ins` around `w:r` with
 * `w:drawing`), so a party makes it as a suggestion. Returns the suggestion that holds it, or
 * null for a direct edit (settled at once).
 */
export function insertObject(
  ctx: Ctx,
  place: OrderedPlace,
  object: OrderedNewObject
): string | null {
  const at = resolveOrderedPlace(ctx.doc, ctx.story.name, place);
  const keys: ObjectKeys | null =
    object.type === 'picture'
      ? {
          media: ctx.ids.resource(),
          image: ctx.ids.resource(),
          link: object.link === undefined ? null : ctx.ids.resource(),
        }
      : null;
  const node = engineElement(ctx, builtElement(ctx, object, keys));

  // A run of its own with no run properties, as the engine writes a picture: tracked, a new
  // run; untracked, beside its neighbours' faces
  const piece = { text: '\uFFFC', content: node, face: null };
  if (ctx.direct) directInsert(ctx, at.paragraph, at.offset, '', { pieces: [piece] });
  else insertText(ctx, at.paragraph, at.offset, '', { pieces: [piece], outsidePrompt: true });

  // The media part and its relationships (design D14) belong to the suggestion that holds the
  // atom, so a rejected or taken-back picture leaves no part behind: a direct edit's settled
  // one, or (typed into another person's pending insertion) that insertion
  const suggestion = objectSuggestion(ctx, at.index, node);
  if (object.type === 'picture' && keys !== null) {
    addImage(ctx, suggestion, object.media, keys);

    // The validator kept the link's address as `sanitizeHref` keeps it
    if (keys.link !== null && object.link !== undefined) {
      addRelationshipResource(ctx, keys.link, suggestion, ctx.story.name, 'hyperlink', {
        external: object.link,
      });
    }
  }

  return ctx.direct ? null : suggestion;
}

/** The object's element as it shows now: its own under its accepted records. */
function currentElement(ctx: Ctx, atom: OrderedItem): OoxmlNode {
  const node = objectNodeOf(scratchCompose(ctx.doc, ctx.story, 'object'), atom);
  if (!node || node.kind === 'textValue')
    throw new OrderedRefusal('invalid-edit', 'the target is not a picture or a shape');
  return node;
}

/** The model op of a change's op: a new link or image becomes resources of the suggestion. */
function modelOp(ctx: Ctx, suggestion: string, op: OrderedObjectChangeOp): OrderedObjectOp {
  if (op.op === 'image') {
    const keys = { media: ctx.ids.resource(), image: ctx.ids.resource() };
    addImage(ctx, suggestion, op.media, keys);
    return { op: 'image', relationship: keys.image };
  }

  if (op.op === 'metadata' && typeof op.link === 'string') {
    const key = ctx.ids.resource();
    addRelationshipResource(ctx, key, suggestion, ctx.story.name, 'hyperlink', {
      external: op.link,
    });
    return { op: 'metadata', title: op.title, description: op.description, link: key };
  }

  return op;
}

/**
 * A direct change of an object's properties: one suggestion, accepted at once, with one record
 * for each op, in order, that compose runs with the engine's own op. Refused by name when the
 * edit is a suggestion (Word has no tracked form for it, `model/objects.ts`), the object no
 * longer shows, or the engine refuses an op on the object as it shows (a lock, a position on an
 * inline picture). Returns the direct suggestion.
 */
export function changeObject(
  ctx: Ctx,
  atom: OrderedItem,
  ops: readonly OrderedObjectChangeOp[]
): string {
  if (!ctx.direct)
    throw new OrderedRefusal('untracked-only', 'a picture property has no tracked form');
  if (atom.kind !== 'atom') throw new OrderedRefusal('invalid-edit', 'the target is no object');
  if (isHidden(ctx.doc, atom)) throw new OrderedRefusal('hidden-target', atom.space.slice(0, 160));
  if (ops.length === 0) throw new OrderedRefusal('invalid-edit', 'a property change needs an op');
  refuseLockedAt(ctx, atom);
  const before = currentElement(ctx, atom);
  const suggestion = newSuggestion(ctx.doc, ctx.attribution, ctx.ids.suggestion());
  const made = ops.map((op) => modelOp(ctx, suggestion.id, op));

  // The engine checks each op on the object as it shows, as the editor's own commit does
  let scratch = 0;
  objectAfterOps(
    ctx.doc,
    ctx.story.part,
    before,
    made,
    () => `object-check#${(scratch += 1)}`,
    (reason) => {
      throw new OrderedRefusal('engine-refused', `a picture property: ${reason}`);
    }
  );
  const records: OrderedObjectRecord[] = made.map((op) => ({
    kind: 'object',
    suggestion: suggestion.id,
    seq: nextSeq(ctx.doc),
    target: 'object',
    op,
  }));
  relabel(ctx.doc, ctx.story, atom, { props: [...atom.props, ...records] });
  settle(ctx.doc, suggestion.id);
  return suggestion.id;
}
