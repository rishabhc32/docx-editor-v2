import type { CollaborationFailureCode } from '@docx-editor.dev/core/collaboration';
import type { TreeDocOp } from '@docx-editor.dev/core/store';

/**
 * Refuse an insert whose own text is longer than one shared text admits, before the editor
 * commits it. Shared state refuses it too, but after the commit, where the edit only disappears
 * in the realign and its caller never learns why. Reads the operations alone, so a keystroke
 * costs nothing more. A short insert into a text near the limit passes here and is still
 * refused after the commit: the commit decides which shared text takes the insert, and the
 * document port shows paragraph text only.
 */
export function textLimitFailure(
  ops: readonly TreeDocOp[],
  maxTextLength: number
): CollaborationFailureCode | null {
  for (const op of ops) {
    if (op.op === 'insertText' && op.text.length > maxTextLength) return 'collaboration-text-limit';
  }

  return null;
}
