// Enter in a block content control's prompt.

import { findNode, type EditOptions } from '../package/ooxml-edit.ts';
import type { OoxmlPart } from '../package/ooxml-tree.ts';
import { isInlineControl } from './content-control-checkbox.ts';
import { withParagraphDiff } from './content-control-value-content.ts';
import { applyTreeOp } from './tree-op-apply.ts';
import { clearPlaceholder, placeholderControlForInsertion } from './tree-op-content-controls.ts';
import { withoutEmptyRuns } from './tree-op-fragment.ts';
import type { TreeDocOp, TreeOpResult } from './tree-op-types.ts';

type Split = Extract<TreeDocOp, { op: 'splitParagraph' }>;

/**
 * The split at a block control's prompt: the prompt emptied into the caret's paragraph, then
 * that paragraph split at its start. Null when the split is not at a block control's prompt.
 *
 * Word 16.114 (2026-10-30, `~$p4-enter.docx` -> `~$p4-enter-tracked.docx`, Track Changes on)
 * cannot hold a collapsed caret in a prompt: it selects the whole prompt, and Return REPLACES
 * it. `w:showingPlcHdr` goes, the prompt's text goes with no `w:del`, and the control holds two
 * empty paragraphs, the first mark a tracked `w:ins` (the editor's companion
 * `setParagraphMarkRevision` on the caret's paragraph). Reject All (`~$p4-enter-rejected.docx`)
 * leaves Word's blank prompt: no glossary block keeps the prompt, as typing over it does. A
 * `w:placeholder` the control already names stays.
 */
export function applySplitAtBlockPrompt(
  part: OoxmlPart,
  op: Split,
  options?: EditOptions
): TreeOpResult | null {
  const prompt = placeholderControlForInsertion(part, op.paragraphId, op.offset);
  if (!prompt || isInlineControl(part, prompt.control.id)) return null;
  const cleared = clearPlaceholder(part, prompt.control.id, options, op.paragraphId);
  if (!cleared) return { ok: false, reason: 'unsupported' };
  const emptied = withoutEmptyRuns(part, cleared, op.paragraphId, options);
  const result = applyTreeOp(emptied, { ...op, offset: 0 }, options);

  // Emptying the prompt may drop its other paragraphs; the split's effect must say so
  const control = findNode(emptied, prompt.control.id);
  if (!result.ok || !control) return result;
  return { ...result, effect: withParagraphDiff(result.effect, prompt.control, control) };
}
