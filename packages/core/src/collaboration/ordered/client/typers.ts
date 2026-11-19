// A checkpoint's typers back in the document (design D14).

import type { CheckpointTypers } from '../model/checkpoint-suggestions.ts';
import type { OrderedDocument } from '../model/types.ts';
import { eachInsertedUnit, keepLoadedTypers, typedUnitKey } from '../model/typers.ts';

/**
 * Keep who typed each unit of an insertion a checkpoint names, for `orderedUnitTyper`: the runs
 * of the listed suggestion's units, in the order the file writes them. False when the runs do
 * not fit the units the file holds for the suggestion.
 */
export function withCheckpointTypers(
  doc: OrderedDocument,
  typers: readonly CheckpointTypers[],
  listedIds: readonly string[]
): boolean {
  if (typers.length === 0) {
    return true;
  }

  // Each suggestion's runs, and how far into them its units are
  const runs = new Map<string, { units: CheckpointTypers['units']; run: number; left: number }>();
  for (const { suggestion, units } of typers) {
    const id = listedIds[suggestion];
    if (id === undefined) {
      return false;
    }

    runs.set(id, { units, run: 0, left: units[0]?.[0] ?? 0 });
  }

  const kept = new Map<string, string>();
  let fits = true;
  eachInsertedUnit(doc, (suggestion, item, offset) => {
    const at = runs.get(suggestion);
    if (at === undefined || !fits) {
      return;
    }

    const run = at.units[at.run];
    if (run === undefined) {
      fits = false;
      return;
    }

    if (run[1] !== null) {
      kept.set(typedUnitKey(item.space, offset), run[1]);
    }

    at.left -= 1;
    if (at.left === 0) {
      at.run += 1;
      at.left = at.units[at.run]?.[0] ?? 0;
    }
  });
  if (!fits || [...runs.values()].some((at) => at.run < at.units.length)) {
    return false;
  }

  keepLoadedTypers(doc, kept);
  return true;
}
