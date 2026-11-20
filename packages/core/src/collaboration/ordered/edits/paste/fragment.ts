// A paste with formatting in the list (group 9.10).

import type { OoxmlElement, OoxmlNode } from '../../../../store/package/ooxml-tree.ts';
import { acrossMarks } from '../../../../store/store/fragment-landing-inside.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import type { OrderedViewParagraph } from '../../compose/context.ts';
import {
  innermostSegment,
  isControlSegment,
  isLinkSegment,
  openOfSegment,
  segmentsOf,
  sharedDepth,
} from '../../containers.ts';
import { OrderedRefusal } from '../../model/refusal.ts';
import { BLOCK_TOKENS, type OrderedItem } from '../../model/types.ts';
import { isParagraphProperties, withChildren } from '../../ooxml-nodes.ts';
import {
  breakSite,
  insertAtSite,
  promptSite,
  refuseLocked,
  siteAt,
  type Site,
} from '../../placement.ts';
import { nestedIn, type Ctx } from '../context.ts';
import { holdsLines, promptBreakSite, splitContainers } from '../marks.ts';
import { newParagraphIdentity, withParagraphIdentity } from '../para-ids.ts';
import {
  engineScratch,
  scratchBlocks,
  templateItems,
  TEMPLATE_REVISION,
  viewNode,
  viewPart,
} from '../templates.ts';

/** A new mark's paragraph shell with paragraph ids of its own, past every id in use. */
function withOwnParagraphIds(ctx: Ctx, node: OoxmlNode | null): OoxmlNode | null {
  if (!node || node.kind === 'textValue') return node;
  return withParagraphIdentity(node, newParagraphIdentity(ctx).attributes);
}

/**
 * The template's items as the paste's: every unit, mark and inline container is its; a direct
 * paste's tables, rows, cells and block controls too (they go when its undo rejects it).
 */
function asPasted(ctx: Ctx, p: OrderedViewParagraph, item: OrderedItem, suggestion: string) {
  const block = item.token !== undefined && BLOCK_TOKENS.has(item.token);
  const own = item.by === null && (ctx.direct || !block);

  // The rows and cells around the host paragraph hold the paste too
  const out = nestedIn({ ...item, ...(own ? { by: suggestion } : {}) }, p.mark.within);
  return item.kind === 'mark' ? { ...out, node: withOwnParagraphIds(ctx, item.node) } : out;
}

/**
 * Where the pasted items go. Content with no paragraph break goes where typed text goes: inside
 * the link or control around the caret, and into a control at its prompt (the prompt shows
 * again when the paste goes, rule 7), as Word pastes into a link's or a control's text.
 * Paragraphs at a block control's prompt replace it the same way, as the engine pastes over a
 * prompt. Other paragraphs go where Enter splits the paragraph, with the links and controls
 * around it split; in an inline control's prompt they meet Enter's refusals.
 */
function siteOf(
  ctx: Ctx,
  p: OrderedViewParagraph,
  offset: number,
  inline: boolean
): { readonly site: Site; readonly split: readonly OrderedItem[] } {
  if (inline) return { site: promptSite(ctx, p, offset) ?? siteAt(ctx, p, offset), split: [] };
  const prev = offset > 0 ? p.chars[offset - 1] : undefined;
  const next = p.chars[offset];
  const prompt = [next, prev].find((char) => char?.item.prompt === true);
  const inControl = isControlSegment(innermostSegment(prompt?.containers ?? ''));
  const blockPrompt = prompt && !inControl ? promptSite(ctx, p, offset) : null;
  if (blockPrompt) return { site: blockPrompt, split: [] };
  return prev?.item.prompt && next?.item.prompt
    ? { site: promptBreakSite(ctx, p, prev.item), split: [] }
    : breakSite(ctx, p, offset);
}

/**
 * The link or the control that strictly holds a view offset (the characters on both sides of
 * it are inside, a control's prompt too), a link first; null at a container's edge.
 */
function holderAt(ctx: Ctx, p: OrderedViewParagraph, offset: number): 'link' | 'control' | null {
  const prev = offset > 0 ? p.chars[offset - 1] : undefined;
  const next = p.chars[offset];
  if (!prev || !next) return null;
  const before = segmentsOf(prev.containers);
  const shared = before.slice(0, sharedDepth(before, segmentsOf(next.containers)));
  if (shared.some(isLinkSegment)) return 'link';
  for (const segment of shared)
    if (isControlSegment(segment) && openOfSegment(ctx.story, segment)?.token === 'sdt')
      return 'control';
  return null;
}

/**
 * The blocks as Word lands them inside a link or a control (paste evidence 2): paragraphs in a
 * control become one line, each pasted mark a space (`acrossMarks`, as the engine lands them);
 * paragraphs in a link are refused by name, as the engine refuses them.
 */
function insideBlocks(
  ctx: Ctx,
  p: OrderedViewParagraph,
  offset: number,
  blocks: readonly OoxmlElement[],
  whole: boolean
): { readonly blocks: readonly OoxmlElement[]; readonly whole: boolean } {
  const holder = holderAt(ctx, p, offset);
  if (holder === null || (blocks.length === 1 && !whole)) return { blocks, whole };
  if (holder === 'link')
    throw new OrderedRefusal('paragraphs-in-link', 'paragraphs pasted inside a link');
  if (!blocks.every((block) => block.kind === 'paragraph'))
    throw new OrderedRefusal('break-in-container', 'a table pasted inside an inline control');
  let minted = 0;
  const line = acrossMarks(blocks, whole, () => `across#${(minted += 1)}`);
  return {
    blocks: line.filter((node): node is OoxmlElement => node.kind !== 'textValue'),
    whole: false,
  };
}

/**
 * Paste `fragment` (engine nodes, every reference resolved) at a view offset, as one insertion
 * of `suggestion`. `fragmentWhole` is the engine's `lastMarkCovered`: the fragment's last
 * paragraph keeps its own mark. The engine's own `insertFragment` runs on a scratch part that
 * holds one empty paragraph with the host paragraph's properties, so its result loads as items
 * with the labels a reopened Word file gives them: runs in `w:ins`, inserted marks, inserted rows
 * (ECMA-376 Part 1, §17.13.5.17, §17.13.5.18, §17.13.5.20).
 */
export function pasteFragment(
  ctx: Ctx,
  p: OrderedViewParagraph,
  offset: number,
  fragment: readonly OoxmlElement[],
  fragmentWhole: boolean,
  suggestion: string
): void {
  refuseLocked(ctx, p, offset, offset);
  const { blocks, whole } = insideBlocks(ctx, p, offset, fragment, fragmentWhole);
  const view = viewPart(ctx);
  const host = viewNode(view, p.id);
  const scratch = withChildren(
    host,
    host.children.filter((child) => child.kind !== 'textValue' && isParagraphProperties(child))
  );
  const op: TreeDocOp = {
    op: 'insertFragment',
    paragraphId: scratch.id,
    offset: 0,
    blocks,
    lastMarkCovered: whole,
    // A suggestion lands as Word tracks a paste; a direct paste lands as the engine pastes
    // (equations and block controls have no tracked form), every item the paste's settled one
    ...(ctx.direct ? {} : { revision: TEMPLATE_REVISION }),
  };
  const result = engineScratch(view, [scratch], op);
  const items: OrderedItem[] = [];
  for (const node of scratchBlocks(result)) {
    if (node.kind !== 'textValue') items.push(...templateItems(ctx, node, suggestion));
  }

  // The scratch paragraph's own mark ends the result: the host's mark ends the last part, as
  // the engine's merge keeps the host's mark on the tail
  const last = items.pop();
  if (!last || last.kind !== 'mark')
    throw new OrderedRefusal('engine-refused', 'insertFragment: no paragraph at the end');
  const pasted = items.map((item) => asPasted(ctx, p, item, suggestion));
  const { site, split } = siteOf(ctx, p, offset, !pasted.some((item) => item.kind === 'mark'));
  for (const open of split)
    if (open.token === 'sdt' && !holdsLines(open))
      throw new OrderedRefusal('single-line-control', 'a paste in a control that holds one line');
  const { closes, opens } = splitContainers(ctx, split, suggestion, p.mark);
  insertAtSite(ctx, site, [...closes, ...pasted, ...opens]);
}
