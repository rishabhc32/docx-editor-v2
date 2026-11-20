// A text box's story drawn back into its object (group 9.9, `model/text-boxes.ts`).

import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
} from '../../../store/package/ooxml-tree.ts';
import { withTextboxFallbackInStep } from '../../../store/package/textbox-fallback-export.ts';
import { canonicalTreeDifference } from '../../../store/package/ooxml-difference.ts';
import { MAX_XML_DEPTH } from '../../../store/package/xml-reader.ts';
import { textBoxesOfAtom, textBoxHostOf } from '../model/text-boxes.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from '../model/types.ts';
import { withChildren } from '../ooxml-nodes.ts';
import { composeBlocks } from './blocks.ts';
import { ComposeContext } from './context.ts';
import { objectNodeOf } from './objects.ts';

/** The `w:txbxContent` elements below a node, in document order. Bounded by the XML depth cap. */
function textBoxContents(node: OoxmlNode, out: OoxmlElement[] = [], depth = 0): OoxmlElement[] {
  if (node.kind === 'textValue' || depth > MAX_XML_DEPTH) return out;
  if (node.localName === 'txbxContent' && node.namespaceUri === WML_NAMESPACE_URI) out.push(node);
  for (const child of node.children) textBoxContents(child, out, depth + 1);
  return out;
}

/** `node` with the element `id` below it replaced; the same node when nothing changed. */
function replaceBelow(node: OoxmlNode, id: string, next: OoxmlElement, depth = 0): OoxmlNode {
  if (node.kind === 'textValue' || depth > MAX_XML_DEPTH) return node;
  if (node.id === id) return next;
  let changed = false;
  const children = node.children.map((child) => {
    const replaced = replaceBelow(child, id, next, depth + 1);
    if (replaced !== child) changed = true;
    return replaced;
  });
  return changed ? withChildren(node, children) : node;
}

/**
 * One text box's story composed, with what its compose leaves for the part (prompts, takeovers),
 * and whether a change record in it holds records by takeover.
 */
function composedBlocks(
  ctx: ComposeContext,
  name: string
): { readonly blocks: readonly OoxmlNode[]; readonly takeovers: boolean } | null {
  const story = ctx.doc.stories.get(name);
  if (!story) return null;
  const inner = new ComposeContext(ctx.doc, story, ctx.options, ctx.mint);
  const { children } = composeBlocks(inner);
  ctx.emptied.push(...inner.emptied);
  for (const id of inner.decisionEmptied) ctx.decisionEmptied.add(id);
  ctx.reached.push(...inner.reached);
  ctx.prompts.push(...inner.prompts);
  for (const [id, taken] of inner.takeovers) {
    ctx.takeovers.set(id, taken);
  }

  return { blocks: children, takeovers: inner.takeovers.size > 0 };
}

/**
 * Whether an element's children are these blocks, node by node: the same nodes, or nodes Word
 * reads the same (compose draws some unedited paragraphs anew, a field inside a loaded insertion,
 * and an untouched text box keeps the file's form).
 */
const sameBlocks = (root: OoxmlElement, blocks: readonly OoxmlNode[]): boolean =>
  root.children.length === blocks.length &&
  blocks.every((block, k) => {
    const loaded = root.children[k];
    return (
      block === loaded || (loaded !== undefined && canonicalTreeDifference(loaded, block).equal)
    );
  });

/**
 * An object's node with each of its text boxes holding its story as it is now: its
 * `w:txbxContent` takes the story's composed blocks, its VML fallback copy too. `node` is the
 * object as compose draws it (under its object records), or the loaded one. Object records give
 * the inner nodes new ids but keep every text box, so each is found by its place among the
 * object's text boxes. The node itself when every text box draws as the file had it (or the
 * object has none).
 */
export function withTextBoxContent(
  ctx: ComposeContext,
  item: OrderedItem,
  node: OoxmlNode | null = item.node
): OoxmlNode | null {
  if (!node || !item.node) return node;
  const names = textBoxesOfAtom(ctx.doc, item);
  if (names.length === 0) return node;
  const loaded = textBoxContents(item.node);
  const drawn = node === item.node ? loaded : textBoxContents(node);
  let out: OoxmlNode = node;
  for (const name of names) {
    const host = textBoxHostOf(ctx.doc, name);
    const root = host ? drawn[loaded.findIndex((content) => content.id === host.root)] : undefined;
    const composed = composedBlocks(ctx, name);

    // The file's form when it reads the same, but for a takeover: a checkpoint names the change
    // record compose drew, which holds the records taken
    if (!root || !composed || (!composed.takeovers && sameBlocks(root, composed.blocks))) {
      continue;
    }

    out = replaceBelow(out, root.id, withChildren(root, composed.blocks));

    // The VML fallback copy follows, as a save writes it and Word keeps it (same paragraph ids)
    out = withTextboxFallbackInStep(out, root.id, ctx.mint);
  }

  return out;
}

/** Whether every object of a paragraph draws its text boxes as the file had them. */
export function textBoxesUntouched(ctx: ComposeContext, buffer: readonly OrderedItem[]): boolean {
  for (const item of buffer) {
    if (item.kind !== 'atom' || textBoxesOfAtom(ctx.doc, item).length === 0) continue;

    // A scratch context: the paragraph is drawn again when this answers no
    const probe = new ComposeContext(ctx.doc, ctx.story, ctx.options, ctx.mint);
    if (withTextBoxContent(probe, item) !== item.node) return false;
  }

  return true;
}

/**
 * An object's node as it shows now, under its object records, its text boxes holding their
 * stories: what a copy or a move of the object takes.
 */
export function currentObjectNode(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem
): OoxmlNode | null {
  let scratch = 0;
  const ctx = new ComposeContext(doc, story, {}, () => `object#${(scratch += 1)}`);
  return withTextBoxContent(ctx, item, objectNodeOf(ctx, item));
}
