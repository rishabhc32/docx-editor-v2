// A checkpoint's run and paragraph properties drawn only for change records (design D14).

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { createOrderedItem } from '../model/document.ts';
import { replaceItems } from '../model/mutations.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import { childWhere, isParagraphProperties, withChildren, withMarkFace } from '../ooxml-nodes.ts';

/**
 * Take off each run's or mark's `w:rPr` and each mark's `w:pPr` that a checkpoint marks as drawn
 * only for change records (`oc:drawn`, by node id in `drawn`) and that holds nothing once the
 * load took its records out. The room drew it where the run or the paragraph has none, so a
 * decision that takes the records away draws none, after a fold as in the room. A file's own
 * empty properties carry no mark and stay, as the engine keeps them.
 */
export function withoutDrawnProperties(doc: OrderedDocument, drawn: ReadonlySet<string>): void {
  if (drawn.size === 0) {
    return;
  }

  const empty = (element: OoxmlElement | null | undefined): boolean =>
    element !== null &&
    element !== undefined &&
    drawn.has(element.id) &&
    element.children.length === 0;
  for (const story of doc.stories.values()) {
    for (let item = story.tree.first(); item !== null; item = story.tree.next(item)) {
      const face = empty(item.face) ? null : item.face;
      const node = item.kind === 'mark' ? markShell(item, face, empty) : item.node;
      if (face !== item.face || node !== item.node) {
        const next = createOrderedItem({ ...item, face, node });
        replaceItems(doc, story, [item], [next]);
        item = next;
      }
    }
  }
}

/** A mark's paragraph with `face` in its `w:pPr`, and no `w:pPr` where a drawn one is empty. */
function markShell(
  mark: OrderedItem,
  face: OoxmlElement | null,
  empty: (element: OoxmlElement | undefined) => boolean
): OrderedItem['node'] {
  const shell = mark.node;
  if (shell === null || shell.kind === 'textValue') {
    return shell;
  }

  const pPr = childWhere(shell, isParagraphProperties);
  if (pPr === undefined || (face === mark.face && !empty(pPr))) {
    return shell;
  }

  const own = face === mark.face ? pPr : withMarkFace(pPr, face);
  const others = shell.children.filter((child) => child !== pPr);
  return withChildren(shell, empty(own) ? others : [own, ...others]);
}
