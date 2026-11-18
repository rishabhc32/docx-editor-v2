// The prefixes pasted content uses that its story's root does not declare (group 9.10).

import { carryIndexToRebuiltRoot } from '../../../store/package/ooxml-edit.ts';
import type { OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import type { OrderedDocument, OrderedStory } from '../model/types.ts';
import { recordUndo } from '../model/undo-record.ts';
import { withDeclaredPrefixes } from '../ooxml-nodes.ts';

/**
 * Declare on a story's root the prefixes `nodes` use that nothing in scope declares (a pasted
 * shape's `mc`, `wps`, `v`, `o`, `asvg`), as the editor alone does when its paste lands
 * (`withRequiredNamespaceBindings` in `store/store/tree-op-fragment.ts`). The undo record takes
 * a declaration back with its edit; a later take-back or reject keeps it, as the editor does.
 */
export function declarePrefixes(
  doc: OrderedDocument,
  story: OrderedStory,
  nodes: readonly OoxmlNode[]
): void {
  const before = story.part;
  const root = withDeclaredPrefixes(before.root, nodes);
  if (root === before.root) {
    return;
  }

  // Same id, same children: the part's node index goes with the root, as an engine edit's does
  carryIndexToRebuiltRoot(before.root, root);
  story.part = { ...before, root };
  recordUndo(doc, () => {
    carryIndexToRebuiltRoot(root, before.root);
    story.part = before;
  });
}
