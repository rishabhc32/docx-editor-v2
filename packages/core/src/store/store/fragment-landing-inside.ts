// A paste whose place is INSIDE a link, an inline content control or a tracked insertion, landed
// the way Word lands it (Word 16.114 for Mac, paste evidence 2, samples `~$paste-out-e2-*.docx`).

import type {
  OoxmlElement,
  OoxmlNode,
  OoxmlParagraphNode,
  OoxmlPart,
} from '../package/ooxml-tree.ts';
import { WML_NAMESPACE_URI } from '../package/ooxml-tree.ts';
import {
  createNodeIdAllocator,
  findNode,
  replaceChildren,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import { contentControlPropertiesOf } from '../package/content-control-nodes.ts';
import { markerSiteAt } from './tree-op-comments.ts';
import { cloneWithFreshIds, textRun } from './tree-op-content-controls.ts';
import { paragraphPropertiesNodeOf, runPropertiesNodeOf, TEXT_DEPS } from './tree-op-nodes.ts';
import { paragraphOffsetIndex } from './tree-op-segments.ts';
import { liftedOutOfForeignInsertions } from './tree-op-tracked-foreign-split.ts';
import type { TreeOpRejection, TreeOpResult } from './tree-op-types.ts';

/** The containers that strictly hold a paste's place, as Word's evidence covers them. */
export interface InsideContainers {
  readonly link: boolean;
  readonly control: boolean;
  readonly insertion: boolean;
}

/** Control types that hold formatted content: rich text, building blocks, or no type. */
const RICH_CONTROLS: ReadonlySet<string> = new Set([
  'richText',
  'docPartList',
  'docPartObj',
  'untyped',
]);

/** A control that takes formatted content. */
function isRichControl(node: OoxmlNode): boolean {
  return node.kind === 'contentControl' && RICH_CONTROLS.has(contentControlPropertiesOf(node).type);
}

/** Whether `node` holds the node `id`, itself included. */
function holdsNode(node: OoxmlNode, id: string): boolean {
  if (node.id === id) return true;
  return node.kind !== 'textValue' && node.children.some((child) => holdsNode(child, id));
}

/**
 * A paste at the prompt of an inline control in `paragraph`, as the one line Word lands in the
 * control (its marks spaces); null for a block control's prompt, or for a table.
 */
export function inlinePromptBlocks(
  paragraph: OoxmlNode | null,
  control: OoxmlNode,
  blocks: readonly OoxmlNode[],
  lastMarkCovered: boolean,
  nextId: () => string
): readonly OoxmlNode[] | null {
  if (!paragraph || !holdsNode(paragraph, control.id)) return null;
  if (!blocks.every((block) => block.kind === 'paragraph')) return null;
  return acrossMarks(blocks, lastMarkCovered, nextId);
}

/**
 * The element containers whose span strictly holds `offset`, outermost first. Runs are not
 * containers: an offset inside a run splits it.
 */
function containersAt(paragraph: OoxmlParagraphNode, offset: number): OoxmlElement[] {
  const index = paragraphOffsetIndex(paragraph);
  const out: OoxmlElement[] = [];
  let level: readonly OoxmlNode[] = paragraph.children;
  for (let depth = 0; depth < 64; depth += 1) {
    const holder = level.find((node) => {
      if (node.kind === 'textValue' || node.kind === 'run') return false;
      const span = index.spanOf(node);
      return span !== null && span.start < offset && offset < span.end;
    });
    if (!holder || holder.kind === 'textValue') break;
    if (holder.kind !== 'contentControlContent') out.push(holder);
    level = holder.children;
  }

  return out;
}

/**
 * The link, control and insertion around a place that a paste may land inside, or null when
 * something else holds the place (an atom, a deletion, a plain text control, an insertion around
 * a link or a control, whose split would duplicate their identity).
 */
export function insideContainersAt(
  paragraph: OoxmlParagraphNode,
  offset: number
): InsideContainers | null {
  // An atom (a field's result, a note reference, a picture) is one unit: nothing lands inside
  for (const segment of paragraphOffsetIndex(paragraph).segments) {
    if (segment.removeNodeIds && segment.start < offset && offset < segment.end) return null;
  }

  let link = false;
  let control = false;
  let insertion = false;
  for (const node of containersAt(paragraph, offset)) {
    if (node.kind === 'revisionInsert') {
      insertion = true;
    } else if (insertion) {
      return null;
    } else if (node.kind === 'hyperlink') {
      link = true;
    } else if (isRichControl(node)) {
      control = true;
    } else {
      return null;
    }
  }

  return { link, control, insertion };
}

const holdsLink = (node: OoxmlNode): boolean =>
  node.kind !== 'textValue' && (node.kind === 'hyperlink' || node.children.some(holdsLink));

/**
 * Why a fragment cannot land inside the containers at its place, or null when it can.
 * `inside` is {@link insideContainersAt} of the place.
 *
 * Inside a link, Word puts a paste without a paragraph mark into the link, and its runs keep their
 * own formatting (they do not take the link's character style). Pasted paragraphs inside a link
 * make Word turn the link into a complex `HYPERLINK` field that spans them; that paste is refused
 * by name (`fragment-paragraphs-in-link`) rather than approximated.
 */
export function insideLandingRefusal(
  inside: InsideContainers | null,
  blocks: readonly OoxmlNode[],
  lastMarkCovered: boolean
): TreeOpRejection | null {
  if (inside === null) return 'indivisible-content';
  if (!inside.link && !inside.control) return null;

  // A table or a block control has no inline form: Word's evidence covers paragraphs only
  if (!blocks.every((block) => block.kind === 'paragraph')) return 'indivisible-content';
  if (inside.link) {
    if (blocks.length > 1 || lastMarkCovered) return 'fragment-paragraphs-in-link';

    // A link inside a link: Word's evidence does not cover it
    if (blocks.some(holdsLink)) return 'indivisible-content';
  }

  return null;
}

/** The last run under `nodes`, depth first from the end. */
function lastRunOf(nodes: readonly OoxmlNode[]): OoxmlNode | null {
  for (let k = nodes.length - 1; k >= 0; k -= 1) {
    const node = nodes[k]!;
    if (node.kind === 'textValue') continue;
    if (node.kind === 'run') return node;
    const inner = lastRunOf(node.children);
    if (inner) return inner;
  }

  return null;
}

/** A run's properties for a space after it, without another author's format record. */
export function spaceProperties(
  run: OoxmlNode | null,
  nextId: () => string
): OoxmlNode | undefined {
  const rPr = run ? runPropertiesNodeOf(run) : undefined;
  if (!rPr) return undefined;
  const own = {
    ...rPr,
    children: rPr.children.filter(
      (child) => child.kind === 'textValue' || child.localName !== 'rPrChange'
    ),
  } as OoxmlElement;
  return cloneWithFreshIds(own, nextId);
}

/**
 * Pasted paragraphs as the one paragraph Word lands in a rich text control, which a paste does
 * not split: each pasted paragraph mark is a space in the format of the text before it (the last
 * mark only when it was copied), and the first paragraph's properties stay, as the landing reads
 * none of them.
 */
export function acrossMarks(
  blocks: readonly OoxmlNode[],
  lastMarkCovered: boolean,
  nextId: () => string
): OoxmlNode[] {
  const first = blocks[0];
  if (blocks.length === 1 && !lastMarkCovered) return [...blocks];
  if (!first || first.kind === 'textValue') return [...blocks];
  const children: OoxmlNode[] = [];
  const pPr = paragraphPropertiesNodeOf(first);
  if (pPr) children.push(pPr);
  blocks.forEach((block, index) => {
    if (block.kind === 'textValue') return;
    const inline = block.children.filter((child) => child !== paragraphPropertiesNodeOf(block));
    for (const child of inline) {
      children.push(child);
    }

    if (index < blocks.length - 1 || lastMarkCovered) {
      children.push(textRun(nextId, ' ', spaceProperties(lastRunOf(inline), nextId)));
    }
  });
  return [{ ...first, children } as OoxmlNode];
}

const isWmlElement = (node: OoxmlNode): node is OoxmlElement =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI;

/** A landing's result: `paragraphId`'s flow changed, and no node was created or deleted. */
function landedResult(part: OoxmlPart, paragraphId: string): TreeOpResult {
  return {
    ok: true,
    part,
    effect: {
      dirty: [paragraphId],
      created: [],
      deleted: [],
      dependencyKeys: TEXT_DEPS,
      impact: 'flow-structural',
    },
  };
}

/**
 * Land inline content at `offset` inside the link or control that holds it: the run there
 * splits, the content goes between its halves, and an insertion around it splits so the
 * content sits beside it, untracked or as the paster's own insertion (Word never nests a paste
 * in an insertion). Pasted paragraphs land beside an insertion through Enter's split of the
 * paragraph, which already splits it so.
 */
export function landInside(
  part: OoxmlPart,
  paragraph: OoxmlParagraphNode,
  offset: number,
  inline: readonly OoxmlNode[],
  options?: EditOptions
): TreeOpResult {
  const site = markerSiteAt(part, paragraph, offset, options);
  if (!site.ok) return site;
  const container = findNode(site.part, site.containerId);
  if (!container || container.kind === 'textValue') return { ok: false, reason: 'tree-invariant' };
  const placed = replaceChildren(
    site.part,
    container.id,
    [
      ...container.children.slice(0, site.index),
      ...inline,
      ...container.children.slice(site.index),
    ],
    { ...options, deferValidation: true }
  );
  if (!placed.ok) return { ok: false, reason: 'tree-invariant' };
  const host = findNode(placed.part, paragraph.id);
  if (!host || host.kind !== 'paragraph') return { ok: false, reason: 'tree-invariant' };

  // No author is the paster's here: every insertion around the paste splits, a tracked paste
  // sitting beside the halves as its own insertion
  const lifted = liftedOutOfForeignInsertions(
    host.children,
    new Set(inline.filter(isWmlElement).map((node) => node.id)),
    '',
    createNodeIdAllocator(placed.part, 'paste')
  );
  if (!lifted) return { ok: false, reason: 'indivisible-content' };
  const landed = replaceChildren(placed.part, host.id, lifted, options);
  if (!landed.ok) return { ok: false, reason: 'tree-invariant' };
  return landedResult(landed.part, host.id);
}

/**
 * Land inline content as the content of a control whose prompt the paste cleared: the empty
 * run the clearing leaves for typing goes, and the content takes its place.
 */
export function landIntoControl(
  part: OoxmlPart,
  paragraphId: string,
  controlId: string,
  inline: readonly OoxmlNode[],
  options?: EditOptions
): TreeOpResult {
  const control = findNode(part, controlId);
  const paragraph = findNode(part, paragraphId);
  if (!control || control.kind === 'textValue' || !paragraph || paragraph.kind !== 'paragraph') {
    return { ok: false, reason: 'tree-invariant' };
  }

  let holder: OoxmlElement = control;
  for (const child of control.children) {
    if (child.kind === 'contentControlContent') holder = child;
  }

  const index = paragraphOffsetIndex(paragraph);
  const kept = holder.children.filter((child) => child.kind !== 'run' || index.lengthOf(child) > 0);
  const landed = replaceChildren(part, holder.id, [...kept, ...inline], options);
  if (!landed.ok) return { ok: false, reason: 'tree-invariant' };
  return landedResult(landed.part, paragraph.id);
}
