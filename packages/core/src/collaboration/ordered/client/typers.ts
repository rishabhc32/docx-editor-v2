// A checkpoint's typers back in the document (design D14).

import type { CheckpointTyperRun } from '../model/checkpoint-suggestions.ts';
import type { OrderedDocument } from '../model/types.ts';
import { eachFileUnit, keepLoadedTypers, type TypedRun } from '../model/typers.ts';

/**
 * Keep who typed each unit a checkpoint names, for `orderedUnitTyper`: runs over the units the
 * file writes, in the order a load reads them. False when the runs name more units than the file
 * holds.
 */
export function withCheckpointTypers(
  doc: OrderedDocument,
  typers: readonly CheckpointTyperRun[]
): boolean {
  if (typers.length === 0) {
    return true;
  }

  // The run the next unit is in, and how many of its units are left
  let index = 0;
  let left = typers[0]?.[0] ?? 0;
  const kept = new Map<string, TypedRun[]>();
  eachFileUnit(doc, (item, offset) => {
    const run = typers[index];
    if (run === undefined) {
      return;
    }

    const typer = run[1];
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

    left -= 1;
    if (left === 0) {
      index += 1;
      left = typers[index]?.[0] ?? 0;
    }
  });
  if (index < typers.length) {
    return false;
  }

  keepLoadedTypers(doc, kept);
  return true;
}
