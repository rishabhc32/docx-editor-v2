// The room facts a checkpoint's manifest keeps, by where the composed file writes them (design
// D14): a model address holds the file number, so none is kept.

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { textboxFallbackExportPackage } from '../../../store/package/textbox-fallback-export.ts';
import { orderedStoryParts } from '../load/load.ts';
import type { CheckpointFacts } from '../model/checkpoint.ts';
import type { CheckpointTypers } from '../model/checkpoint-suggestions.ts';
import {
  checkpointPlacesOf,
  isCheckpointContainer,
  takenProperties,
  type CheckpointTakeover,
} from '../model/checkpoint-takeovers.ts';
import type { CheckpointUntracked } from '../model/checkpoint-untracked.ts';
import type { OrderedPerson } from '../model/person.ts';
import { revisionIdsOf } from '../model/revision-ids.ts';
import type {
  OrderedDocument,
  OrderedElementRecord,
  OrderedPropertyTarget,
  OrderedStory,
} from '../model/types.ts';
import { eachInsertedUnit, orderedUnitTyper } from '../model/typers.ts';
import { attributeOf, eachElement, isWml } from '../ooxml-nodes.ts';
import type { OrderedPartTakeovers } from './package.ts';

/** A container of the composed file: its part, local name and place among those of its name. */
interface ContainerPlace {
  readonly part: string;
  readonly container: string;
  readonly occurrence: number;
}

/**
 * The room facts of a document composed as `pkg`, whose change records hold `held` by takeover.
 * Each pending suggestion the file writes marks for, or another fact names, is listed once, in
 * the order the room made them, so a load makes them in that order.
 */
export function checkpointFacts(
  doc: OrderedDocument,
  pkg: OoxmlPackage,
  held: OrderedPartTakeovers
): CheckpointFacts {
  const written = writtenRevisionIds(pkg);
  const order: string[] = [];
  const listed = new Map<string, number>();
  const indexOf = (id: string): number => {
    let k = listed.get(id);
    if (k === undefined) {
      k = order.length;
      listed.set(id, k);
      order.push(id);
    }

    return k;
  };
  const revisions = new Map<string, string[]>();
  const writtenOf = (id: string): string[] => {
    let ids = revisions.get(id);
    if (ids === undefined) {
      ids = [...revisionIdsOf(doc, id)].filter((r) => written.has(r));
      revisions.set(id, ids);
    }

    return ids;
  };
  for (const { id, status } of doc.suggestions.values()) {
    if (status === 'pending' && !doc.partRevisions.has(id) && writtenOf(id).length > 0) {
      indexOf(id);
    }
  }

  // As a load reads the file: with each text box's legacy copy in step
  const containers = new Map<string, ContainerPlace>();
  const takeovers: CheckpointTakeover[] = [];
  for (const [part, story] of orderedStoryParts(textboxFallbackExportPackage(pkg))) {
    const records = held.get(part);
    for (const { element, occurrence } of checkpointPlacesOf(part, story)) {
      const container = element.localName;
      if (isCheckpointContainer(container) && !containers.has(element.id)) {
        containers.set(element.id, { part, container, occurrence });
      }

      const found = records?.get(element.id);
      const id = attributeOf(element, 'id');
      if (found !== undefined && id !== undefined) {
        const taken = found.taken.flatMap(({ replaced, after }) =>
          doc.suggestions.has(replaced)
            ? [{ suggestion: indexOf(replaced), properties: takenProperties(found.record, after) }]
            : []
        );
        takeovers.push({ part, record: container, id, occurrence, taken });
      }
    }
  }

  const untracked = checkpointUntracked(doc, containers, indexOf);
  const typers = checkpointTypers(doc, listed);

  // In the order the room made them: the takeovers and the untracked records may name more
  const sorted = order
    .map((id, k) => ({ id, k, seq: doc.suggestions.get(id)?.seq ?? 0 }))
    .sort((a, b) => a.seq - b.seq);
  const place = new Map(sorted.map(({ k }, n) => [k, n]));
  const at = (k: number): number => place.get(k) ?? k;
  return {
    suggestions: sorted.map(({ id }) => {
      const suggestion = doc.suggestions.get(id);
      return {
        attribution: suggestion?.attribution ?? { author: '', person: null },
        revisions: writtenOf(id),
        status: suggestion?.status ?? 'pending',
      };
    }),
    comments: commentPersons(doc),
    takeovers: takeovers.map((t) => ({
      ...t,
      taken: t.taken.map((taken) => ({ ...taken, suggestion: at(taken.suggestion) })),
    })),
    typers: typers.map((t) => ({ ...t, suggestion: at(t.suggestion) })),
    untracked: untracked.map((u) => ({
      ...u,
      records: u.records.map((r) => ({ ...r, suggestion: at(r.suggestion) })),
    })),
  };
}

/** The person of each room comment, by its `w:id`. */
function commentPersons(doc: OrderedDocument): Map<string, OrderedPerson> {
  const persons = new Map<string, OrderedPerson>();
  for (const comment of doc.comments.values()) {
    if (comment.person !== null) {
      persons.set(comment.id, comment.person);
    }
  }

  return persons;
}

/**
 * For each properties element of a container the file writes that a pending record changes with
 * no change record: the container's own properties, and every record of them, decided or not,
 * that no takeover holds (a takeover's records are kept by their change record).
 */
function checkpointUntracked(
  doc: OrderedDocument,
  containers: ReadonlyMap<string, ContainerPlace>,
  indexOf: (id: string) => number
): CheckpointUntracked[] {
  const out: CheckpointUntracked[] = [];
  for (const story of doc.stories.values()) {
    for (const open of story.tree.items()) {
      const shell = open.node;
      const place = shell === null ? undefined : containers.get(shell.id);
      if (open.kind !== 'open' || place === undefined || shell?.kind === 'textValue') {
        continue;
      }

      const byTarget = new Map<OrderedPropertyTarget, OrderedElementRecord[]>();
      for (const record of open.props) {
        if (record.kind === 'element' && record.replaced === undefined) {
          byTarget.set(record.target, [...(byTarget.get(record.target) ?? []), record]);
        }
      }

      for (const [target, records] of byTarget) {
        const pending = (id: string) => doc.suggestions.get(id)?.status === 'pending';
        if (!records.some((record) => record.untracked === true && pending(record.suggestion))) {
          continue;
        }

        const properties = shell?.children.find((c) => isWml(c, target));
        out.push({
          ...place,
          target,
          ...(properties === undefined || properties.kind === 'textValue' ? {} : { properties }),
          records: records.map((record) => ({
            suggestion: indexOf(record.suggestion),
            element: record.element,
            ...(record.untracked === undefined ? {} : { untracked: record.untracked }),
            ...(record.merge === undefined ? {} : { merge: record.merge }),
            ...(record.recordUnchanged === undefined
              ? {}
              : { recordUnchanged: record.recordUnchanged }),
            ...(record.tableCells === undefined
              ? {}
              : { tableCells: cellPlaces(story, record.tableCells, containers) }),
          })),
        });
      }
    }
  }

  return out;
}

/** The places of the cells of a record's table that the file writes. */
function cellPlaces(
  story: OrderedStory,
  cells: readonly string[],
  containers: ReadonlyMap<string, ContainerPlace>
): number[] {
  return cells.flatMap((space) => {
    const shell = story.spaces.get(space)?.[0]?.node;
    const place = shell === null || shell === undefined ? undefined : containers.get(shell.id);
    return place === undefined ? [] : [place.occurrence];
  });
}

/**
 * Who typed the units of each listed suggestion's insertion, as runs in the order the file
 * writes them; none for a suggestion no room change typed in.
 */
function checkpointTypers(
  doc: OrderedDocument,
  listed: ReadonlyMap<string, number>
): CheckpointTypers[] {
  const runs = new Map<number, [number, string | null][]>();
  eachInsertedUnit(doc, (suggestion, item, offset) => {
    const k = listed.get(suggestion);
    if (k === undefined) {
      return;
    }

    const typer = orderedUnitTyper(doc, item, offset);
    const units = runs.get(k) ?? [];
    const last = units[units.length - 1];
    if (last !== undefined && last[1] === typer) {
      last[0] += 1;
    } else {
      units.push([1, typer]);
    }

    runs.set(k, units);
  });
  return [...runs]
    .filter(([, units]) => units.some(([, typer]) => typer !== null))
    .map(([suggestion, units]) => ({ suggestion, units }));
}

/** The `w:id` of every revision mark in a package's stories. */
function writtenRevisionIds(pkg: OoxmlPackage): Set<string> {
  const ids = new Set<string>();
  for (const part of orderedStoryParts(pkg).values()) {
    eachElement(part.root, (element: OoxmlElement) => {
      const id = attributeOf(element, 'id');
      if (id !== undefined && attributeOf(element, 'author') !== undefined) {
        ids.add(id);
      }
    });
  }

  return ids;
}
