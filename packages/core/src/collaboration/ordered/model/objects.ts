// Property records on a picture's or a shape's unit (group 9.4, design D13).

import type {
  DrawingLocksInput,
  DrawingPositionInput,
  ImageWrapTarget,
  SourceCrop,
} from '../../../store/package/drawing-projection.ts';

/**
 * One property change of a picture or a shape: the engine's drawing op without its node. A link
 * and an image are resources the change added, named by their keys (design D14). A record keeps
 * the op, not its result, so each record keeps its own meaning: compose runs the accepted ops
 * in creation order on the unit's loaded or inserted element.
 *
 * Word has no tracked form of these properties: ECMA-376 has no revision element that can hold
 * an earlier DrawingML property (`wml.xsd` lines 803 to 955 list every revision type, and
 * `CT_Drawing`, `CT_Inline`, `CT_Anchor` and `CT_NonVisualDrawingProps` have no change child).
 * So only a direct edit makes one: the organiser's, a suggestion accepted at once, which
 * `revert` turns off exactly.
 *
 * @public
 */
export type OrderedObjectOp =
  /** A new size, in EMU (`resizeDrawing`). */
  | { readonly op: 'resize'; readonly cx: number; readonly cy: number }
  /** A new crop, in thousandths of a percent of each edge (`cropDrawing`). */
  | { readonly op: 'crop'; readonly crop: SourceCrop }
  /** A floating object's position (`positionDrawing`). */
  | { readonly op: 'position'; readonly position: DrawingPositionInput }
  /** Inline, or a wrap around text (`setDrawingWrap`). */
  | { readonly op: 'wrap'; readonly wrap: ImageWrapTarget }
  /**
   * Title and description (`setDrawingMetadata`). `link` is the key of the external
   * relationship the picture's link names, `null` removes the link, and no `link` keeps it.
   */
  | {
      readonly op: 'metadata';
      readonly title: string;
      readonly description: string;
      readonly link?: string | null;
    }
  /** What a reader may change (`setDrawingLocks`). */
  | { readonly op: 'locks'; readonly locks: DrawingLocksInput }
  /** A quarter turn or a flip (`transformDrawing`). */
  | { readonly op: 'transform'; readonly action: 'rotateCW' | 'rotateCCW' | 'flipH' | 'flipV' }
  /** A new image: the key of the image relationship the change added (`replaceDrawingResource`). */
  | { readonly op: 'image'; readonly relationship: string };

/**
 * A property record on an object's unit: one op of a direct edit.
 *
 * @public
 */
export interface OrderedObjectRecord {
  readonly kind: 'object';
  readonly suggestion: string;
  readonly seq: number;
  readonly target: 'object';
  readonly op: OrderedObjectOp;
  /** Never set: a direct edit takes over no record (policy 2). */
  readonly replaced?: string;
}
