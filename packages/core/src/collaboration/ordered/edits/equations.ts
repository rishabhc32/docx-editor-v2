// Equations (group 9.9): a new linear form for one equation, or one equation of a display
// removed. An equation is one object (`m:oMath`, or a display `m:oMathPara`).

import { linearMathReplacementOmml } from '../../../store/package/omml-equation.ts';
import { equationsOfAtom, isOmmlEquationAtom } from '../../../store/package/omml-display.ts';
import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import type { OrderedViewParagraph } from '../compose/context.ts';
import { isHidden } from '../model/document.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedItem } from '../model/types.ts';
import { withChildren, withId } from '../ooxml-nodes.ts';
import { refuseLockedAt } from '../placement.ts';
import { withNewIds, type Ctx } from './context.ts';
import { directDelete, directInsert } from './direct.ts';
import { deleteText, insertText, type Piece } from './text.ts';

/** The shown paragraph at an index. */
function paragraphAt(ctx: Ctx, index: number): OrderedViewParagraph {
  const p = orderedViewParagraph(ctx.doc, ctx.story.name, index);
  if (!p) throw new OrderedRefusal('malformed-story', `no paragraph ${index}`);
  return p;
}

/**
 * The object's node with one equation replaced or removed. A direct edit keeps the ids the
 * editor names (the object's and each equation's, as the editor keeps them): only one of the
 * two objects ever shows. A suggestion shows both, so its new object takes the change's ids.
 */
function nextObject(
  ctx: Ctx,
  object: OoxmlElement,
  equation: number,
  linear: string | null
): OoxmlElement {
  const equations = equationsOfAtom(object);
  const target = equations[equation];
  if (!target) throw new OrderedRefusal('unknown-address', `no equation ${equation} there`);
  const fresh = !ctx.direct;
  let replacement: OoxmlElement | null = null;
  if (linear !== null) {
    const built = linearMathReplacementOmml(linear, () => ctx.ids.node(), target);
    if (!built.ok) throw new OrderedRefusal('invalid-edit', 'the linear form does not parse');
    replacement = fresh ? built.equation : withId(built.equation, target.id);
  } else if (equations.length < 2)
    throw new OrderedRefusal('invalid-edit', 'a lone equation goes by a deletion');
  if (object === target) {
    if (replacement === null)
      throw new OrderedRefusal('invalid-edit', 'an equation needs its new form');
    return replacement;
  }

  // A display: the other equations stay
  const children = object.children.flatMap((child) => {
    if (child === target) {
      return replacement ? [replacement] : [];
    }

    return [fresh && child.kind !== 'textValue' ? withNewIds(ctx, child) : child];
  });
  const shell = fresh ? withNewIds(ctx, withChildren(object, [])) : object;
  return withChildren(shell, children);
}

/**
 * Give the equation object `item` a new form for its `equation`-th equation (`linear`), or
 * remove that equation from its display (`linear` null). Returns the suggestion that holds it.
 * The old object is struck and the new one inserted at its place, in one suggestion (a direct
 * edit settles it at once), so a decision or an undo never shows a mix. The new OMML comes from
 * the engine's own builder with ids from the change: every copy writes the same equation.
 */
export function replaceEquation(
  ctx: Ctx,
  item: OrderedItem,
  equation: number,
  linear: string | null
): string | null {
  const node = item.node;
  if (item.kind !== 'atom' || !item.bare || !node || !isOmmlEquationAtom(node))
    throw new OrderedRefusal('invalid-edit', 'not an equation');
  if (isHidden(ctx.doc, item)) throw new OrderedRefusal('hidden-target', 'the equation');
  refuseLockedAt(ctx, item);
  const piece: Piece = {
    text: item.text || '￼',
    face: null,
    content: nextObject(ctx, node, equation, linear),
    bare: true,
  };

  // The index of the shown paragraph that holds the item: the shown marks before it
  const index = ctx.story.tree.before(item).marks;
  const p = paragraphAt(ctx, index);
  const offset = p.chars.findIndex((char) => char.item === item);
  if (offset < 0) throw new OrderedRefusal('hidden-target', 'the equation does not show');
  if (ctx.direct) {
    // The struck object hides: the new one goes where it was
    directDelete(ctx, p, offset, offset + 1);
    directInsert(ctx, paragraphAt(ctx, index), offset, '', [piece]);
    return null;
  }

  // The struck object stays in view and the new one follows it, as Word writes a replacement:
  // the deletion, then the insertion. One's own new object is taken back instead, and goes.
  const struck = deleteText(ctx, p, offset, offset + 1, null);
  const view = paragraphAt(ctx, index);
  const still = view.chars.findIndex((char) => char.item === item);
  return insertText(
    struck === null ? ctx : { ...ctx, join: struck },
    view,
    still >= 0 ? still + 1 : offset,
    '',
    { pieces: [piece], outsidePrompt: true }
  );
}
