// Tells a refused `exec` write from one that had nothing to change.

import type { ExecResult } from '../contracts/editor.ts';
import type { PaginatedSurface } from './paginated-surface-contract.ts';

/** The fallback when the surface has no reason left to show. */
export const REFUSED_WRITE_REASON = 'the edit was refused';

export interface ExecRefusalWatch {
  /** The surface hooks, with `onRejection` also counted. */
  hooks<T extends { readonly onRejection?: (reason: string) => void }>(hooks: T): T;
  /**
   * How many refusals the surface has reported so far. The surface reports each refusal once,
   * through `onRejection`, so a count that moved during one command means it was refused.
   */
  mark(): number;
  /**
   * The answer for one write that `execEditorCommand` finished.
   *
   * A verb that answered for itself with a refusal or a change keeps its answer. Otherwise a
   * mutating command that moved no revision while the surface refused something is that
   * refusal; with no refusal it is a write that had nothing to change, `changed: false`.
   * The count decides whether the write was refused; the reason prefers the surface's readable
   * `lastRejection`, then the raw reason the hook saw.
   */
  settle(input: {
    readonly surface: PaginatedSurface;
    readonly mutating: boolean;
    readonly revisionBefore: number;
    readonly mark: number;
    readonly result: ExecResult | null;
  }): ExecResult;
}

export function createExecRefusalWatch(): ExecRefusalWatch {
  let count = 0;
  let lastRaw: string | null = null;
  return {
    hooks(hooks) {
      const forward = hooks.onRejection;
      return {
        ...hooks,
        onRejection: (reason: string) => {
          count += 1;
          lastRaw = reason;
          forward?.(reason);
        },
      };
    },
    mark: () => count,
    settle({ surface, mutating, revisionBefore, mark, result }) {
      if (result && (!result.ok || result.changed)) return result;
      const changed = surface.session.packageRevision() !== revisionBefore;
      if (!changed && mutating && count !== mark) {
        // The surface's own sentence first: `refuse` stores the readable form, while the hook
        // saw the raw one (a store code such as `content-locked`).
        const reason = surface.state().lastRejection ?? lastRaw ?? REFUSED_WRITE_REASON;
        return { ok: false, code: 'unsupported', reason };
      }

      return result ?? { ok: true, changed };
    },
  };
}
