// A checkpoint's typers back in the document (design D14).

import type { OrderedPerson } from '../model/person.ts';
import { isUnit, type OrderedDocument } from '../model/types.ts';
import { keepLoadedTypers, type TypedRun } from '../model/typers.ts';

/**
 * Keep who typed each unit a checkpoint names, for `orderedUnitTyper`: the typer each run (or
 * unit not in a run) of the file carries, by the run's node id, which names the units the load
 * read from it, and the person of each listed namespace. False when a run names a namespace the
 * manifest does not list.
 */
export function withCheckpointTypers(
  doc: OrderedDocument,
  typers: ReadonlyMap<string, string>,
  persons: ReadonlyMap<string, OrderedPerson>
): boolean {
  for (const namespace of typers.values()) {
    if (!persons.has(namespace)) {
      return false;
    }
  }

  for (const [namespace, person] of persons) {
    doc.namespaces.set(namespace, person);
  }

  if (typers.size === 0) {
    return true;
  }

  // A loaded unit's `run` is the node id of the run it was read from, or its own when bare
  const kept = new Map<string, TypedRun[]>();
  for (const story of doc.stories.values()) {
    for (const item of story.tree.items()) {
      const typer = isUnit(item) ? typers.get(item.run) : undefined;
      if (typer === undefined) {
        continue;
      }

      const runs = kept.get(item.space) ?? [];
      const last = runs[runs.length - 1];
      const to = item.offset + item.length;
      if (last !== undefined && last.to === item.offset && last.typer === typer) {
        runs[runs.length - 1] = { ...last, to };
      } else {
        runs.push({ from: item.offset, to, typer });
      }

      kept.set(item.space, runs);
    }
  }

  keepLoadedTypers(doc, kept);
  return true;
}
