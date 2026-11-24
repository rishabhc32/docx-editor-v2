// The room facts a checkpoint's manifest keeps, by the labels the composed file carries (design
// D14): a model address holds the file number, and a place counted in the file holds the engine's
// rules of the day, so neither is kept.

import { glossaryPartOf } from '../../../store/package/building-blocks.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { orderedStoryParts } from '../load/load.ts';
import type { CheckpointFacts } from '../model/checkpoint.ts';
import type { CheckpointList, CheckpointStyle } from '../model/checkpoint-lists.ts';
import type { CheckpointSuggestion } from '../model/checkpoint-suggestions.ts';
import { withCheckpointLabels } from '../model/checkpoint-labels.ts';
import {
  takenProperties,
  type CheckpointTaken,
  type CheckpointTakeover,
} from '../model/checkpoint-takeovers.ts';
import {
  UNTRACKED_CONTAINERS,
  type CheckpointChangeRecord,
  type CheckpointChangeRecords,
  type CheckpointUntracked,
  type CheckpointUntrackedRecord,
} from '../model/checkpoint-untracked.ts';
import type { OrderedPerson } from '../model/person.ts';
import { orderedResourceShows } from '../model/resources.ts';
import { revisionIdsOf } from '../model/revision-ids.ts';
import type {
  OrderedDocument,
  OrderedElementRecord,
  OrderedItem,
  OrderedPropertyTarget,
  OrderedStory,
  OrderedSuggestion,
} from '../model/types.ts';
import { attributeOf, eachElement, isWml } from '../ooxml-nodes.ts';
import type { RoomNumber } from './context.ts';
import { styleResourceKeyOf } from './resources.ts';
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
  entryOf({ id, attribution, status, version }: OrderedSuggestion): CheckpointSuggestion {
    const record = this.doc.partRevisions.get(id);
    if (record === undefined) {
      return { id, attribution, revisions: this.writtenOf(id), status, version };
    }

    const { revision, element } = record;
    return { id, attribution, revisions: [revision.id], record: element, status, version };
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

/**
 * A container's open item, with each of its properties that a record changes untracked, and the
 * other pending records whose room facts a change record does not say.
 */
interface RecordsOpen {
  readonly story: OrderedStory;
  readonly shell: OoxmlElement;
  /**
   * Every record of each such properties element, decided or not, but those a change record holds
   * by takeover: a record taken by an untracked one is listed, as no change record holds it.
   */
  readonly byTarget: ReadonlyMap<OrderedPropertyTarget, readonly OrderedElementRecord[]>;
  /** Each other pending record that sets some children or holds a table's cells. */
  readonly recorded: readonly OrderedElementRecord[];
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
  // untracked or table record names
  const opens = recordsOpens(doc);
  const wanted = new Set<string>();
  for (const records of held.values()) {
    for (const id of records.keys()) {
      wanted.add(id);
    }
  }

  for (const { story, shell, byTarget, recorded } of opens) {
    wanted.add(shell.id);
    for (const records of [...byTarget.values(), recorded]) {
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
    changeRecords: checkpointChangeRecords(doc, opens, labels, list),
    lists: checkpointLists(doc, pkg, list),
    styles: checkpointStyles(doc, pkg, list),
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

/**
 * Each style a change added that the file holds as compose drew it (its suggestion pending or
 * accepted): all of them, so an open gives them back in the room's order.
 */
function checkpointStyles(
  doc: OrderedDocument,
  pkg: OoxmlPackage,
  list: SuggestionList
): CheckpointStyle[] {
  const written = new Set<string>();
  for (const part of pkg.parts.values()) {
    if (isWml(part.root, 'styles')) {
      for (const child of part.root.children) {
        const key = child.kind === 'textValue' ? undefined : styleResourceKeyOf(child);
        if (key !== undefined && isWml(child, 'style')) {
          written.add(key);
        }
      }
    }
  }

  const out: CheckpointStyle[] = [];
  for (const resource of doc.resources.added.values()) {
    const suggestion = doc.suggestions.get(resource.suggestion);
    if (resource.kind === 'style' && written.has(resource.key) && suggestion !== undefined) {
      const { key, styleId } = resource;
      out.push({ suggestion: list.indexOf(suggestion), key, styleId });
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
      records: u.records.map((r) => ({
        ...r,
        suggestion: renumbered(r.suggestion),
        replaced: r.replaced === undefined ? undefined : renumbered(r.replaced),
      })),
    })),
    changeRecords: facts.changeRecords.map((c) => ({
      ...c,
      records: c.records.map((r) => ({
        ...r,
        suggestion: renumbered(r.suggestion),
        replaced: r.replaced === undefined ? undefined : renumbered(r.replaced),
      })),
    })),
    lists: facts.lists.map((l) => ({ ...l, suggestion: renumbered(l.suggestion) })),
    styles: facts.styles.map((s) => ({ ...s, suggestion: renumbered(s.suggestion) })),
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
 * Each container's open item that holds a pending record of some properties, untracked, or a
 * pending record with a change record, which does not say the children the record sets or the
 * cells it holds. A row's cells are one decision (`resetUnrecordedCells` reads the records of
 * each): where a decided record of a row or a cell is kept, the row and each of its cells keep
 * every record.
 */
function recordsOpens(doc: OrderedDocument): RecordsOpen[] {
  const pending = (id: string) => doc.suggestions.get(id)?.status === 'pending';
  const kept = (records: readonly OrderedElementRecord[]) =>
    records.some((record) => record.untracked === true && pending(record.suggestion));
  const out: RecordsOpen[] = [];
  for (const story of doc.stories.values()) {
    // The containers with records, each with the row it is or is a cell of: the innermost row
    const found: {
      readonly shell: OoxmlElement;
      readonly open: OrderedItem;
      readonly row: OrderedItem | undefined;
      readonly byTarget: Map<OrderedPropertyTarget, OrderedElementRecord[]>;
    }[] = [];
    const rows: OrderedItem[] = [];
    const whole = new Set<OrderedItem>();
    for (const open of story.tree.items()) {
      if (open.token === 'tr' && open.kind === 'close') {
        rows.pop();
        continue;
      }

      if (open.token === 'tr') {
        rows.push(open);
      }

      const shell = open.node;
      if (
        open.kind !== 'open' ||
        open.props.length === 0 ||
        shell === null ||
        shell.kind === 'textValue'
      ) {
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

      const row = open.token === 'tr' || open.token === 'tc' ? rows[rows.length - 1] : undefined;
      const keptRecords = [...byTarget.values()].filter(kept).flat();
      if (row !== undefined && keptRecords.some((record) => !pending(record.suggestion))) {
        whole.add(row);
      }

      found.push({ shell, open, row, byTarget });
    }

    for (const { shell, open, row, byTarget } of found) {
      const keepsAll = row !== undefined && whole.has(row);
      for (const [target, records] of byTarget) {
        if (!kept(records) && !(keepsAll && UNTRACKED_CONTAINERS.has(target))) {
          byTarget.delete(target);
        }
      }

      const recorded = open.props.filter(
        (record): record is OrderedElementRecord =>
          record.kind === 'element' &&
          (record.merge !== undefined || record.tableCells !== undefined) &&
          record.untracked !== true &&
          UNTRACKED_CONTAINERS.has(record.target) &&
          pending(record.suggestion) &&
          !byTarget.get(record.target)?.includes(record)
      );
      if (byTarget.size > 0 || recorded.length > 0) {
        out.push({ story, shell, byTarget, recorded });
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
  opens: readonly RecordsOpen[],
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
      for (const record of records) {
        const kept = changeRecordOf(doc, story, record, labels, list);
        if (kept !== undefined) {
          const { element, untracked, recordUnchanged } = record;
          listed.push({ ...kept, element, untracked, recordUnchanged });
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

/** What the room holds of each pending record with a change record, by the label on its container. */
function checkpointChangeRecords(
  doc: OrderedDocument,
  opens: readonly RecordsOpen[],
  labels: ReadonlyMap<string, number>,
  list: SuggestionList
): CheckpointChangeRecords[] {
  const out: CheckpointChangeRecords[] = [];
  for (const { story, shell, recorded } of opens) {
    const label = labels.get(shell.id);
    if (label === undefined) {
      continue;
    }

    const byTarget = new Map<OrderedPropertyTarget, CheckpointChangeRecord[]>();
    for (const record of recorded) {
      const kept = changeRecordOf(doc, story, record, labels, list);
      if (kept !== undefined) {
        byTarget.set(record.target, [...(byTarget.get(record.target) ?? []), kept]);
      }
    }

    for (const [target, records] of byTarget) {
      out.push({ label, target, records });
    }
  }

  return out;
}

/**
 * What the manifest keeps of a record that no change record says; undefined when its suggestion,
 * or the one it was taken from, is not the document's.
 */
function changeRecordOf(
  doc: OrderedDocument,
  story: OrderedStory,
  { suggestion, replaced, merge, tableCells }: OrderedElementRecord,
  labels: ReadonlyMap<string, number>,
  list: SuggestionList
): CheckpointChangeRecord | undefined {
  const made = doc.suggestions.get(suggestion);
  const taken = replaced === undefined ? undefined : doc.suggestions.get(replaced);
  if (made === undefined || (replaced !== undefined && taken === undefined)) {
    return undefined;
  }

  return {
    suggestion: list.indexOf(made),
    replaced: taken && list.indexOf(taken),
    merge,
    tableCells: tableCells && cellLabels(story, tableCells, labels),
  };
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
