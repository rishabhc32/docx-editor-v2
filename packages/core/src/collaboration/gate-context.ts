// What a collaboration write gate is told besides the ops (`CollaborationGateContext`).

import type { CollaborationGateContext, CollaborationPackageWrite } from './index.ts';
import { runWithTransactionActor } from '../store/package/actor-scoped-ids.ts';
import type { OoxmlPackage } from '../store/package/ooxml-package.ts';
import type { HistoryGroup } from '../store/store/history-group.ts';
import { isPackageEditRefusal } from '../store/store/package-edit-refusal.ts';

/**
 * A package edit a write carries, applied by the store before the write's ops; `write` says
 * what it does, for a session that carries it by its meaning (`describedPackageEdit`).
 */
export type GatePackageEdit = ((pkg: OoxmlPackage) => OoxmlPackage) & {
  readonly write?: CollaborationPackageWrite;
};

/** A package edit that says what it does. */
export function describedPackageEdit(
  write: CollaborationPackageWrite,
  edit: (pkg: OoxmlPackage) => OoxmlPackage
): GatePackageEdit {
  return Object.assign((pkg: OoxmlPackage) => edit(pkg), { write });
}

/**
 * The context for a write over `current()` with `packageEdits` applied in order, under the
 * actor the store transaction binds, so an edit that mints ids mints the ones the commit will.
 *
 * `package` is undefined when an edit refuses the package: the store refuses that write
 * with the same reason, so the gate has nothing to judge beyond the ops. `historyGroup` is
 * the undo group the commit will carry, when it has one.
 */
export function collaborationGateContext(
  current: () => OoxmlPackage,
  packageEdits: readonly GatePackageEdit[] = [],
  actorId?: string,
  historyGroup?: HistoryGroup
): CollaborationGateContext {
  // Built lazily, once: most sessions never read the package, and a gate that ignores it must
  // not pay for a package build on every keystroke.
  let built = false;
  let pkg: OoxmlPackage | undefined;
  const packageWrites = packageEdits.flatMap((edit) => (edit.write ? [edit.write] : []));
  const described = packageEdits.length > 0 && packageWrites.length === packageEdits.length;
  return {
    ...(historyGroup === undefined ? {} : { historyGroup }),
    ...(described ? { packageWrites } : {}),
    get package(): OoxmlPackage | undefined {
      if (built) return pkg;
      built = true;
      try {
        pkg = runWithTransactionActor(actorId, () =>
          packageEdits.reduce((edited, edit) => edit(edited), current())
        );
      } catch (error) {
        if (!isPackageEditRefusal(error)) throw error;
        pkg = undefined;
      }

      return pkg;
    },
  };
}

/**
 * The context factory one editor uses: its live package and its collaboration actor. With no
 * session attached the gate is never asked, so the factory is never called.
 */
export function gateContextsFor(
  session: { currentPackage(): OoxmlPackage },
  collaboration: { readonly identity: { readonly actorId: string } } | undefined
): (
  packageEdits?: readonly GatePackageEdit[],
  historyGroup?: HistoryGroup
) => CollaborationGateContext {
  return (packageEdits, historyGroup) =>
    collaborationGateContext(
      () => session.currentPackage(),
      packageEdits,
      collaboration?.identity.actorId,
      historyGroup
    );
}
