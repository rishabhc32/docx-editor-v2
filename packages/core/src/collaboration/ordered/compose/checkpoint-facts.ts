// The room facts a checkpoint's manifest keeps, by the labels the composed file carries (design
// D14): a model address holds the file number, and a place counted in the file holds the engine's
// rules of the day, so neither is kept.

import { glossaryPartOf } from '../../../store/package/building-blocks.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { orderedStoryParts } from '../load/load.ts';
import type { CheckpointFacts } from '../model/checkpoint.ts';
import type { CheckpointList } from '../model/checkpoint-lists.ts';
import type { CheckpointSuggestion } from '../model/checkpoint-suggestions.ts';
import { withCheckpointLabels } from '../model/checkpoint-labels.ts';
import {
  takenProperties,
  type CheckpointTaken,
  type CheckpointTakeover,
} from '../model/checkpoint-takeovers.ts';
import type {
  CheckpointUntracked,
  CheckpointUntrackedRecord,
} from '../model/checkpoint-untracked.ts';
import type { OrderedPerson } from '../model/person.ts';
import { orderedResourceShows } from '../model/resources.ts';
import { revisionIdsOf } from '../model/revision-ids.ts';
import type {
  OrderedDocument,
  OrderedElementRecord,
  OrderedPropertyTarget,
  OrderedStory,
  OrderedSuggestion,
} from '../model/types.ts';
import { attributeOf, eachElement, isWml } from '../ooxml-nodes.ts';
import type { RoomNumber } from './context.ts';
import type { OrderedPartTakeovers } from './package.ts';

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

  /**
   * A listed suggestion as the manifest lists it; a file's change record outside the stories by
   * its local name, with its own `w:id`.
   */
  entryOf({ id, attribution, status }: OrderedSuggestion): CheckpointSuggestion {
    const record = this.doc.partRevisions.get(id);
    if (record === undefined) {
      return { id, attribution, revisions: this.writtenOf(id), status };
    }

    return { id, attribution, revisions: [record.revision.id], record: record.element, status };
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
  /**
   * Every record of each such properties element, decided or not, but those a change record holds
   * by takeover: a record taken by an untracked one is listed, as no change record holds it.
   */
  readonly byTarget: ReadonlyMap<OrderedPropertyTarget, readonly OrderedElementRecord[]>;
}

/**
 * The room facts of a document composed as `pkg`, whose change records hold `held` by takeover,
 * whose units the namespaces of `typers` typed and whose runs the room runs of `runs` number, and
 * the package with a label on each element they name, and a room run's number only where it joins
 * pieces. Each pending suggestion the file writes
 * marks for, or another fact names, is listed once, in the order the room made them, so a load
 * makes them in that order.
 */
export function checkpointFacts(
  doc: OrderedDocument,
  pkg: OoxmlPackage,
  held: OrderedPartTakeovers,
  typers: ReadonlyMap<string, OrderedPerson>,
  runs: ReadonlyMap<string, RoomNumber>
): { readonly pkg: OoxmlPackage; readonly facts: CheckpointFacts } {
  const list = new SuggestionList(doc, writtenRevisionIds(pkg));
  for (const suggestion of doc.suggestions.values()) {
    const { id, status } = suggestion;
    if (status === 'pending' && (doc.partRevisions.has(id) || list.writtenOf(id).length > 0)) {
      list.indexOf(suggestion);
    }
  }

  // A label on each change record that holds a takeover, and on each container and cell an
  // untracked record names
  const opens = untrackedOpens(doc);
  const wanted = new Set<string>();
  for (const records of held.values()) {
    for (const id of records.keys()) {
      wanted.add(id);
    }
  }

  for (const { story, shell, byTarget } of opens) {
    wanted.add(shell.id);
    for (const records of byTarget.values()) {
      for (const space of records.flatMap((record) => record.tableCells ?? [])) {
        const cell = cellShell(story, space);
        if (cell !== undefined) {
          wanted.add(cell);
        }
      }
    }
  }

  // A number joins pieces only when more than one element carries it
  const joining = new Set<string>();
  for (const { number, uses } of runs.values()) {
    if (uses > 1) {
      joining.add(number);
    }
  }

  const labelled = withCheckpointLabels(pkg, wanted, typers.size > 0 || runs.size > 0, joining);
  const { labels } = labelled;
  const facts = inRoomOrder(list, {
    commentPersons: commentPersons(doc),
    takeovers: checkpointTakeovers(doc, held, labels, list),
    namespacePersons: typers,
    untracked: checkpointUntracked(doc, opens, labels, list),
    lists: checkpointLists(doc, pkg, list),
  });
  return { pkg: labelled.pkg, facts };
}

/**
 * Each list definition a change added that shows (its suggestion pending or accepted) and the file
 * holds, with no picture bullet: all of them, so an open gives them back in the room's order.
 */
function checkpointLists(
  doc: OrderedDocument,
  pkg: OoxmlPackage,
  list: SuggestionList
): CheckpointList[] {
  const written = new Set<string>();
  for (const part of pkg.parts.values()) {
    if (isWml(part.root, 'numbering')) {
      for (const child of part.root.children) {
        const numId = child.kind === 'textValue' ? undefined : attributeOf(child, 'numId');
        if (numId !== undefined && isWml(child, 'num')) {
          written.add(numId);
        }
      }
    }
  }

  const out: CheckpointList[] = [];
  for (const resource of doc.resources.added.values()) {
    const suggestion = doc.suggestions.get(resource.suggestion);
    if (
      resource.kind === 'list' &&
      resource.bullets === undefined &&
      written.has(resource.numId) &&
      suggestion !== undefined &&
      orderedResourceShows(doc, resource)
    ) {
      out.push({ suggestion: list.indexOf(suggestion), key: resource.key });
    }
  }

  return out;
}

/** Each change record of the file that holds records by takeover, by its label. */
function checkpointTakeovers(
  doc: OrderedDocument,
  held: OrderedPartTakeovers,
  labels: ReadonlyMap<string, number>,
  list: SuggestionList
): CheckpointTakeover[] {
  const takeovers: CheckpointTakeover[] = [];
  for (const records of held.values()) {
    for (const [id, found] of records) {
      const label = labels.get(id);
      if (label === undefined) {
        continue;
      }

      const taken: CheckpointTaken[] = [];
      for (const { replaced, after } of found.taken) {
        const suggestion = doc.suggestions.get(replaced);
        if (suggestion !== undefined) {
          const properties = takenProperties(found.record, after);
          taken.push({ suggestion: list.indexOf(suggestion), properties });
        }
      }

      takeovers.push({ label, taken });
    }
  }

  return takeovers;
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
    suggestions: sorted.map(({ suggestion }) => list.entryOf(suggestion)),
    commentPersons: facts.commentPersons,
    takeovers: facts.takeovers.map((t) => ({
      ...t,
      taken: t.taken.map((taken) => ({ ...taken, suggestion: renumbered(taken.suggestion) })),
    })),
    namespacePersons: facts.namespacePersons,
    untracked: facts.untracked.map((u) => ({
      ...u,
      records: u.records.map((r) => ({ ...r, suggestion: renumbered(r.suggestion) })),
    })),
    lists: facts.lists.map((l) => ({ ...l, suggestion: renumbered(l.suggestion) })),
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
        if (record.kind === 'element' && (record.replaced === undefined || record.untracked)) {
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
 * that no change record holds by takeover (those records are kept by their change record).
 */
function checkpointUntracked(
  doc: OrderedDocument,
  opens: readonly UntrackedOpen[],
  labels: ReadonlyMap<string, number>,
  list: SuggestionList
): CheckpointUntracked[] {
  const out: CheckpointUntracked[] = [];
  for (const { story, shell, byTarget } of opens) {
    const label = labels.get(shell.id);
    if (label === undefined) {
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
        replaced,
      } of records) {
        const made = doc.suggestions.get(suggestion);
        const taken = replaced === undefined ? undefined : doc.suggestions.get(replaced);
        if (made !== undefined && (replaced === undefined || taken !== undefined)) {
          listed.push({
            suggestion: list.indexOf(made),
            element,
            untracked,
            merge,
            recordUnchanged,
            tableCells: tableCells && cellLabels(story, tableCells, labels),
            replaced: taken && list.indexOf(taken),
          });
        }
      }

      const properties = shell.children.find((c) => isWml(c, target));
      out.push({
        label,
        target,
        properties: properties?.kind === 'textValue' ? undefined : properties,
        records: listed,
      });
    }
  }

  return out;
}

/** The node id of a cell's shell, by the cell's space. */
const cellShell = (story: OrderedStory, space: string): string | undefined =>
  story.spaces.get(space)?.[0]?.node?.id;

/** The labels of the cells of a record's table that the file writes. */
function cellLabels(
  story: OrderedStory,
  cells: readonly string[],
  labels: ReadonlyMap<string, number>
): number[] {
  return cells.flatMap((space) => {
    const shell = cellShell(story, space);
    const label = shell === undefined ? undefined : labels.get(shell);
    return label === undefined ? [] : [label];
  });
}

/**
 * The `w:id` of every revision mark in a package's stories, and in its glossary: a prompt that
 * content took the place of keeps its records there. A grid record (`w:tblGridChange`) writes
 * its `w:id` alone (`CT_Markup`): the manifest names its suggestion, or a load makes it nobody's.
 */
function writtenRevisionIds(pkg: OoxmlPackage): Set<string> {
  const ids = new Set<string>();
  const glossary = glossaryPartOf(pkg);
  for (const part of [...orderedStoryParts(pkg).values(), ...(glossary ? [glossary] : [])]) {
    eachElement(part.root, (element: OoxmlElement) => {
      // Most elements hold no `w:id`: only those are read further
      const id = attributeOf(element, 'id');
      if (id === undefined) {
        return;
      }

      if (attributeOf(element, 'author') !== undefined || isWml(element, 'tblGridChange')) {
        ids.add(id);
      }
    });
  }

  return ids;
}
