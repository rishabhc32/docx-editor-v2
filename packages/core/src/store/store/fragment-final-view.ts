// A fragment's accepted final view: its blocks with every tracked change in them accepted.

import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { WML_NAMESPACE_URI } from '../package/ooxml-tree.ts';
import { applyTreeOp } from './tree-op-apply.ts';
import { holdsRevisionMarkup, REVISION_MARKUP } from './tracked-fragment.ts';

/** The `w:body` of a main document part's root. */
function bodyOf(root: OoxmlElement): OoxmlElement | null {
  for (const node of root.children) {
    if (
      node.kind !== 'textValue' &&
      node.namespaceUri === WML_NAMESPACE_URI &&
      node.localName === 'body'
    ) {
      return node;
    }
  }

  return null;
}

/**
 * `blocks` with every tracked change accepted, or null when the engine cannot accept one of
 * them. `part` is a main document part: the scratch part takes its root and replaces its body's
 * blocks. Blocks with no tracked change come back as they are.
 *
 * Word copies this view from a document that tracks changes (Word 16.114 for Mac, paste evidence
 * 1d, `~$paste-out-e1e.docx`): an insertion is plain text, a deletion is gone, and a format
 * change keeps the new format with no record. Only the source's Track Changes setting at copy
 * time decides it, not the view a reader picks. The engine's own `acceptAllRevisions` on the
 * scratch part joins a deleted paragraph mark's paragraphs and removes a deleted row exactly as
 * Accept All does in the editor.
 */
export function acceptedFragmentBlocks(
  part: OoxmlPart,
  blocks: readonly OoxmlNode[]
): readonly OoxmlNode[] | null {
  if (!holdsRevisionMarkup(blocks)) return blocks;
  const body = bodyOf(part.root);
  if (!body) return null;
  const scratch: OoxmlPart = {
    ...part,
    root: {
      ...part.root,
      children: part.root.children.map((child) =>
        child === body ? ({ ...body, children: [...blocks] } as OoxmlElement) : child
      ),
    } as OoxmlElement,
  };
  const accepted = applyTreeOp(scratch, { op: 'acceptAllRevisions' }, { deferValidation: true });
  if (!accepted.ok) return null;
  return bodyOf(accepted.part.root)?.children ?? null;
}

/** A note with its tracked changes accepted, or null when the engine cannot accept one. */
export function acceptedNote(part: OoxmlPart, note: OoxmlElement): OoxmlElement | null {
  const accepted = acceptedFragmentBlocks(part, note.children);
  if (accepted === null) return null;
  return accepted === note.children ? note : ({ ...note, children: [...accepted] } as OoxmlElement);
}

/**
 * A style or a list definition with its tracked property changes accepted: each record of the
 * earlier properties (`w:rPrChange`, `w:pPrChange` and the table ones) is dropped, and the
 * properties it records a change of stay. Such a resource holds no other tracked change.
 */
export function withoutTrackedChanges(node: OoxmlElement): OoxmlElement {
  const kept: OoxmlNode[] = [];
  let changed = false;
  for (const child of node.children) {
    if (child.kind === 'textValue') {
      kept.push(child);
      continue;
    }

    if (child.namespaceUri === WML_NAMESPACE_URI && REVISION_MARKUP.has(child.localName)) {
      changed = true;
      continue;
    }

    const own = withoutTrackedChanges(child);
    if (own !== child) changed = true;
    kept.push(own);
  }

  return changed ? ({ ...node, children: kept } as OoxmlElement) : node;
}
