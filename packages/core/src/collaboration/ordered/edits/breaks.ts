// A tab and the breaks: one object each, typed at a place as the engine types them.

import { hardBreakText } from '../../../store/package/hard-break.ts';
import type { OoxmlHardBreakNode, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { simpleElement } from '../../../store/store/tree-op-inline-elements.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { wmlAttribute } from '../ooxml-nodes.ts';
import type { Ctx } from './context.ts';
import type { Piece } from './text.ts';

/** What one object breaks: nothing (a tab), the line, the page, or the column. @public */
export type OrderedBreak = 'tab' | 'line' | 'page' | 'column';

/** The floating objects a line break restarts below (`w:clear`). @public */
export type OrderedBreakClear = 'left' | 'right' | 'all';

/**
 * The tab's or the break's element, with ids from the edit's id source: the engine's builder
 * (`insertTab`, `insertHardBreak`, `insertPageBreak`), and the column and text wrapping breaks
 * the engine reads (`w:br w:type="column"`, `w:br w:clear`).
 */
function breakNode(ctx: Ctx, kind: OrderedBreak, clear: OrderedBreakClear | undefined): OoxmlNode {
  const mint = () => ctx.ids.node();
  if (kind === 'tab') return simpleElement(mint, 'tab');
  if (kind === 'line' || kind === 'page') {
    const node = simpleElement(mint, 'br', kind);
    if (clear === undefined || node.kind === 'textValue') return node;

    // Word writes a text wrapping break as both: the type, and the side it clears
    return {
      ...node,
      attributes: [wmlAttribute('type', 'textWrapping'), wmlAttribute('clear', clear)],
    } as OoxmlNode;
  }

  const node = simpleElement(mint, 'br', 'line');
  return node.kind === 'textValue'
    ? node
    : ({ ...node, attributes: [wmlAttribute('type', 'column')] } as OoxmlNode);
}

/**
 * The piece an insertion types for a tab or a break: its element, and the character the
 * engine's model text reads for it (`\t`, `\n`, or `\f` for a page break).
 */
export function breakPiece(
  ctx: Ctx,
  kind: OrderedBreak,
  clear: OrderedBreakClear | undefined
): Piece {
  if (clear !== undefined && kind !== 'line')
    throw new OrderedRefusal('invalid-edit', 'only a line break clears floating objects');
  const content = breakNode(ctx, kind, clear);
  const text = content.kind === 'hardBreak' ? hardBreakText(content as OoxmlHardBreakNode) : '\t';
  return { text, content };
}
