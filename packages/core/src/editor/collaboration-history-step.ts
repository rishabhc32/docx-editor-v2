import type { EditorCollaborationSession } from '../collaboration/index.ts';

/**
 * One undo or redo step of a collaboration session's shared history. Returns whether it ran.
 * A local step the session turns down is a refusal, not an empty history. A session that
 * cannot write holds no step to take; its own reason says why.
 */
export function stepCollaborationHistory(
  session: EditorCollaborationSession,
  direction: 'undo' | 'redo',
  publishRefusal: (reason: string) => void
): boolean {
  const held = direction === 'undo' ? session.canUndo() : session.canRedo();
  const done = session[direction]();
  if (done) {
    return true;
  }

  if (held) {
    publishRefusal(`the collaboration session refused this ${direction}`);
    return false;
  }

  const reason = session.writeRefusal?.() ?? null;
  if (reason !== null) {
    publishRefusal(reason);
  }

  return false;
}
