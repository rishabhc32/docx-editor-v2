// A composed part holds each node id once.

import { createNodeIdAllocator } from '../../../store/package/ooxml-edit.ts';
import type { OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { cloneWithNewIds } from '../../../store/store/tree-op-nodes.ts';
import { withChildren } from '../ooxml-nodes.ts';

/**
 * The part with every node id once: a node whose id came earlier is a copy with new ids. Compose
 * reuses the loaded nodes of untouched paragraphs and copies a node again when it is placed
 * twice (`ComposeContext.fresh`), but a node an edit shares with a paragraph compose reuses whole
 * can still reach a part twice. The package writer finds nodes by id (it brings a text box's
 * legacy copy in step with its DrawingML copy), and refuses a part where one id names two nodes.
 */
export function withUniqueIds(part: OoxmlPart): OoxmlPart {
  const seen = new Set<string>();
  let mint: (() => string) | null = null;
  const note = (node: OoxmlNode): void => {
    // An explicit stack: a copy can be as deep as the part allows
    const stack: OoxmlNode[] = [node];
    for (let next = stack.pop(); next !== undefined; next = stack.pop()) {
      seen.add(next.id);
      if (next.kind !== 'textValue') for (const child of next.children) stack.push(child);
    }
  };
  const visit = (node: OoxmlNode): OoxmlNode => {
    if (seen.has(node.id)) {
      mint ??= createNodeIdAllocator(part);
      const copy = cloneWithNewIds(node, mint);
      note(copy);
      return copy;
    }

    seen.add(node.id);
    if (node.kind === 'textValue') return node;
    let changed = false;
    const children = node.children.map((child) => {
      const kept = visit(child);
      changed ||= kept !== child;
      return kept;
    });
    return changed ? withChildren(node, children) : node;
  };
  const root = visit(part.root);
  return root === part.root || root.kind === 'textValue' ? part : { ...part, root };
}
