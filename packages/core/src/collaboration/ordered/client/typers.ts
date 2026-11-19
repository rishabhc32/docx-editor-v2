// A checkpoint's typers back in the document (design D14).

import type { CheckpointTyperRuns } from '../model/checkpoint-suggestions.ts';
import type { OrderedDocument } from '../model/types.ts';
import { eachInsertedUnit, keepLoadedTypers, type TypedRun } from '../model/typers.ts';

/**
 * Keep who typed each unit of an insertion a checkpoint names, for `orderedUnitTyper`: the runs
 * of the listed suggestion's units, in the order the file writes them. False when the runs do
 * not fit the units the file holds for the suggestion.
 */
export function withCheckpointTypers(
  doc: OrderedDocument,
  typers: readonly CheckpointTyperRuns[],
  listedIds: readonly string[]
): boolean {
  if (typers.length === 0) {
    return true;
  }

  // Each suggestion's runs, and how far into them its units are
  const cursors = new Map<
    string,
    { units: CheckpointTyperRuns['units']; index: number; left: number }
  >();
  for (const { suggestion, units } of typers) {
    const id = listedIds[suggestion];
    if (id === undefined) {
      return false;
    }

    cursors.set(id, { units, index: 0, left: units[0]?.[0] ?? 0 });
  }

  const kept = new Map<string, TypedRun[]>();
  let fits = true;
  eachInsertedUnit(doc, (suggestion, item, offset) => {
    const cursor = cursors.get(suggestion);
    if (cursor === undefined || !fits) {
      return;
    }

    const typer = cursor.units[cursor.index]?.[1];
    if (typer === undefined) {
      fits = false;
      return;
    }

    if (typer !== null) {
      const runs = kept.get(item.space) ?? [];
      const last = runs[runs.length - 1];
      if (last !== undefined && last.to === offset && last.typer === typer) {
        runs[runs.length - 1] = { ...last, to: offset + 1 };
      } else {
        runs.push({ from: offset, to: offset + 1, typer });
      }

      kept.set(item.space, runs);
    }

    cursor.left -= 1;
    if (cursor.left === 0) {
      cursor.index += 1;
      cursor.left = cursor.units[cursor.index]?.[0] ?? 0;
    }
  });
  if (!fits || [...cursors.values()].some((cursor) => cursor.index < cursor.units.length)) {
    return false;
  }

  keepLoadedTypers(doc, kept);
  return true;
}
