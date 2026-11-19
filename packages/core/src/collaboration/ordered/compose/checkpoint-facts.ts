// The room facts a checkpoint's manifest keeps, by where the composed file writes them (design
// D14): a model address holds the file number, so none is kept.

import { glossaryPartOf } from '../../../store/package/building-blocks.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { textboxFallbackExportPackage } from '../../../store/package/textbox-fallback-export.ts';
import { orderedStoryParts } from '../load/load.ts';
import type { CheckpointFacts } from '../model/checkpoint.ts';
import type { CheckpointTyperRuns } from '../model/checkpoint-suggestions.ts';
import {
  checkpointPlacesOf,
  isCheckpointContainer,
  takenProperties,
  type CheckpointTaken,
  type CheckpointTakeover,
} from '../model/checkpoint-takeovers.ts';
import type {
  CheckpointUntracked,
  CheckpointUntrackedRecord,
} from '../model/checkpoint-untracked.ts';
import type { OrderedPerson } from '../model/person.ts';
import { revisionIdsOf } from '../model/revision-ids.ts';
import type {
  OrderedDocument,
  OrderedElementRecord,
  OrderedPropertyTarget,
  OrderedStory,
  OrderedSuggestion,
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
 * The suggestions a manifest lists, each once, in the order they are first named, with the
 * `w:id`s the file writes for their marks.
 */
class SuggestionList {
  readonly listed: OrderedSuggestion[] = [];
  private readonly doc: OrderedDocument;
  private readonly written: ReadonlySet<string>;
  private readonly indexOfId = new Map<string, number>();
  private readonly revisions = new Map<string, string[]>();

  /** A list of the suggestions of `doc`, whose file writes marks under the `w:id`s `written`. */
  constructor(doc: OrderedDocument, written: ReadonlySet<string>) {
    this.doc = doc;
    this.written = written;
  }

  /** The place of a suggestion in the list: at its end when it is new. */
  indexOf(suggestion: OrderedSuggestion): number {
    let k = this.indexOfId.get(suggestion.id);
    if (k === undefined) {
      k = this.listed.length;
      this.indexOfId.set(suggestion.id, k);
      this.listed.push(suggestion);
    }

    return k;
  }

  /** The place of a listed suggestion, or undefined for one not listed. */
  placeOf(id: string): number | undefined {
    return this.indexOfId.get(id);
  }

  /** The `w:id`s the file writes for a suggestion's marks, every role. */
  writtenOf(id: string): string[] {
    let ids = this.revisions.get(id);
    if (ids === undefined) {
      ids = [...revisionIdsOf(this.doc, id)].filter((r) => this.written.has(r));
      this.revisions.set(id, ids);
    }

    return ids;
  }
}

/** A container's open item, with each of its properties that a record changes untracked. */
interface UntrackedOpen {
  readonly story: OrderedStory;
  readonly shell: OoxmlElement;
  /** Every record of each such properties element, decided or not, that no takeover holds. */
  readonly byTarget: ReadonlyMap<OrderedPropertyTarget, readonly OrderedElementRecord[]>;
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
  const list = new SuggestionList(doc, writtenRevisionIds(pkg));
  for (const suggestion of doc.suggestions.values()) {
    const { id, status } = suggestion;
    if (status === 'pending' && !doc.partRevisions.has(id) && list.writtenOf(id).length > 0) {
      list.indexOf(suggestion);
    }
  }

  // The file is read only for what a takeover or an untracked record names
  const opens = untrackedOpens(doc);
  const named = opens.length > 0 || [...held.values()].some((records) => records.size > 0);
  const { containers, takeovers } = named
    ? placesAndTakeovers(doc, pkg, held, list)
    : { containers: new Map<string, ContainerPlace>(), takeovers: [] };
  return inRoomOrder(list, {
    commentPersons: commentPersons(doc),
    takeovers,
    typers: checkpointTypers(doc, list),
    untracked: checkpointUntracked(doc, opens, containers, list),
  });
}

/**
 * The place of each container the file writes, by its node id, and each change record that holds
 * records by takeover, as a load reads the file: with each text box's legacy copy in step.
 */
function placesAndTakeovers(
  doc: OrderedDocument,
  pkg: OoxmlPackage,
  held: OrderedPartTakeovers,
  list: SuggestionList
): { containers: Map<string, ContainerPlace>; takeovers: CheckpointTakeover[] } {
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
        const taken: CheckpointTaken[] = [];
        for (const { replaced, after } of found.taken) {
          const suggestion = doc.suggestions.get(replaced);
          if (suggestion !== undefined) {
            const properties = takenProperties(found.record, after);
            taken.push({ suggestion: list.indexOf(suggestion), properties });
          }
        }

        takeovers.push({ part, record: container, id, occurrence, taken });
      }
    }
  }

  return { containers, takeovers };
}

/**
 * The facts with the listed suggestions in the order the room made them, and every section's
 * suggestion renumbered to that order: the takeovers and the untracked records may name
 * suggestions made before those listed first.
 */
function inRoomOrder(
  list: SuggestionList,
  facts: Omit<CheckpointFacts, 'suggestions'>
): CheckpointFacts {
  const sorted = list.listed
    .map((suggestion, k) => ({ suggestion, k }))
    .sort((a, b) => a.suggestion.seq - b.suggestion.seq);
  const place = new Map(sorted.map(({ k }, n) => [k, n]));
  const renumbered = (k: number): number => place.get(k) ?? k;
  return {
    suggestions: sorted.map(({ suggestion: { id, attribution, status } }) => ({
      attribution,
      revisions: list.writtenOf(id),
      status,
    })),
    commentPersons: facts.commentPersons,
    takeovers: facts.takeovers.map((t) => ({
      ...t,
      taken: t.taken.map((taken) => ({ ...taken, suggestion: renumbered(taken.suggestion) })),
    })),
    typers: facts.typers.map((t) => ({ ...t, suggestion: renumbered(t.suggestion) })),
    untracked: facts.untracked.map((u) => ({
      ...u,
      records: u.records.map((r) => ({ ...r, suggestion: renumbered(r.suggestion) })),
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

/** Each container's open item that holds a pending record of some properties, untracked. */
function untrackedOpens(doc: OrderedDocument): UntrackedOpen[] {
  const pending = (id: string) => doc.suggestions.get(id)?.status === 'pending';
  const out: UntrackedOpen[] = [];
  for (const story of doc.stories.values()) {
    for (const open of story.tree.items()) {
      const shell = open.node;
      if (open.kind !== 'open' || shell === null || shell.kind === 'textValue') {
        continue;
      }

      const byTarget = new Map<OrderedPropertyTarget, OrderedElementRecord[]>();
      for (const record of open.props) {
        if (record.kind === 'element' && record.replaced === undefined) {
          const records = byTarget.get(record.target) ?? [];
          records.push(record);
          byTarget.set(record.target, records);
        }
      }

      for (const [target, records] of byTarget) {
        if (!records.some((record) => record.untracked === true && pending(record.suggestion))) {
          byTarget.delete(target);
        }
      }

      if (byTarget.size > 0) {
        out.push({ story, shell, byTarget });
      }
    }
  }

  return out;
}

/**
 * For each properties element of a container the file writes that a pending record changes with
 * no change record: the container's own properties, and every record of them, decided or not,
 * that no takeover holds (a takeover's records are kept by their change record).
 */
function checkpointUntracked(
  doc: OrderedDocument,
  opens: readonly UntrackedOpen[],
  containers: ReadonlyMap<string, ContainerPlace>,
  list: SuggestionList
): CheckpointUntracked[] {
  const out: CheckpointUntracked[] = [];
  for (const { story, shell, byTarget } of opens) {
    const place = containers.get(shell.id);
    if (place === undefined) {
      continue;
    }

    for (const [target, records] of byTarget) {
      const listed: CheckpointUntrackedRecord[] = [];
      for (const {
        suggestion,
        element,
        untracked,
        merge,
        recordUnchanged,
        tableCells,
      } of records) {
        const made = doc.suggestions.get(suggestion);
        if (made !== undefined) {
          listed.push({
            suggestion: list.indexOf(made),
            element,
            untracked,
            merge,
            recordUnchanged,
            tableCells: tableCells && cellPlaces(story, tableCells, containers),
          });
        }
      }

      const properties = shell.children.find((c) => isWml(c, target));
      out.push({
        ...place,
        target,
        properties: properties?.kind === 'textValue' ? undefined : properties,
        records: listed,
      });
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
function checkpointTypers(doc: OrderedDocument, list: SuggestionList): CheckpointTyperRuns[] {
  const runs = new Map<number, [number, string | null][]>();
  eachInsertedUnit(doc, (suggestion, item, offset) => {
    const k = list.placeOf(suggestion);
    if (k === undefined) {
      return;
    }

    const typer = orderedUnitTyper(doc, item, offset);
    let units = runs.get(k);
    if (units === undefined) {
      units = [];
      runs.set(k, units);
    }

    const last = units[units.length - 1];
    if (last !== undefined && last[1] === typer) {
      last[0] += 1;
    } else {
      units.push([1, typer]);
    }
  });
  return [...runs]
    .filter(([, units]) => units.some(([, typer]) => typer !== null))
    .map(([suggestion, units]) => ({ suggestion, units }));
}

/**
 * The `w:id` of every revision mark in a package's stories, and in its glossary: a prompt that
 * content took the place of keeps its records there.
 */
function writtenRevisionIds(pkg: OoxmlPackage): Set<string> {
  const ids = new Set<string>();
  const glossary = glossaryPartOf(pkg);
  for (const part of [...orderedStoryParts(pkg).values(), ...(glossary ? [glossary] : [])]) {
    eachElement(part.root, (element: OoxmlElement) => {
      const id = attributeOf(element, 'id');
      if (id !== undefined && attributeOf(element, 'author') !== undefined) {
        ids.add(id);
      }
    });
  }

  return ids;
}
