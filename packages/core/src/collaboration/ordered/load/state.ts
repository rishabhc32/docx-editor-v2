// What a story load carries while it walks a part: the items so far, the revision reader, the
// limits, and the labels the revision wrappers around the current node give.

import {
  WML_NAMESPACE_URI,
  type OoxmlAttribute,
  type OoxmlElement,
} from '../../../store/package/ooxml-tree.ts';
import { DEFAULT_LIMITS } from '../../../store/runtime/limits.ts';
import { MAX_XML_DEPTH } from '../../../store/package/xml-reader.ts';
import type { CheckpointSuggestion } from '../model/checkpoint-suggestions.ts';
import { noteCommentMarker } from '../model/comments.ts';
import {
  createOrderedItem,
  isLabelled,
  NOT_WITHIN,
  type OrderedItemInit,
} from '../model/document.ts';
import { newSuggestion, setStatus } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { REVISION_ROLE, seedRevisionId, suggestionOfRevision } from '../model/revision-ids.ts';
import { recordUndo } from '../model/undo-record.ts';
import {
  tokenSpace,
  zeroSpace,
  type OrderedAttribution,
  type OrderedContainerToken,
  type OrderedDocument,
  type OrderedItem,
  type OrderedPropertyRecord,
  type OrderedWithin,
} from '../model/types.ts';
import { attributeOf, isWml } from '../ooxml-nodes.ts';

/** The labels the revision wrappers around a node give it. */
export interface Labels {
  readonly by: string | null;
  readonly dels: readonly string[];
  readonly move: string | null;
  readonly within: OrderedWithin;
}

export const NO_LABELS: Labels = { by: null, dels: [], move: null, within: NOT_WITHIN };

/**
 * Labels for the content of a row or cell whose own markers are `own`. A suggestion is named
 * once: a row and its cell struck by one suggestion strike its content once.
 */
export function contentLabels(labels: Labels, own: Labels): Labels {
  if (own.by === null && own.dels.length === 0) return labels;
  const { by, dels } = labels.within;
  return {
    ...labels,
    within: {
      by: own.by === null || by.includes(own.by) ? by : [...by, own.by],
      dels: [...dels, ...own.dels.filter((id) => !dels.includes(id))],
    },
  };
}

/** Limits shared by every story of one load. */
export interface LoadBudget {
  readonly maxItems: number;
  readonly maxDepth: number;
  items: number;
}

/** A load's budget: the limits a caller asks for, never past the engine's own ceilings. */
export function newLoadBudget(maxItems?: number, maxDepth?: number): LoadBudget {
  return {
    maxItems: bounded(maxItems, DEFAULT_LIMITS.maxElementCount),
    maxDepth: bounded(maxDepth, MAX_XML_DEPTH),
    items: 0,
  };
}

function bounded(value: number | undefined, ceiling: number): number {
  return value !== undefined && Number.isFinite(value) && value > 0
    ? Math.min(Math.floor(value), ceiling)
    : ceiling;
}

/**
 * File revisions as suggestions: one suggestion for each revision triple (`w:id`, `w:author`,
 * `w:date`), the identity the engine decides a revision by, or for each suggestion a room
 * checkpoint lists, whatever the triples of its marks. Each one keeps the file's `w:id` for the
 * role it was read in, so compose writes it back.
 */
export class RevisionReader {
  private readonly doc: OrderedDocument;
  private readonly byTriple = new Map<string, string>();
  private readonly byMove = new Map<string, string>();
  /** Every revision is this suggestion (an engine-built template), and no `w:id` is kept. */
  private readonly fixed: string | undefined;
  /** The suggestions a room checkpoint lists, by the triple of each of their marks. */
  private readonly listedByTriple = new Map<string, string>();
  /** A prompt the glossary kept: its marks read as `suggestionOfRevision` reads them. */
  private readonly prompt: boolean;

  /**
   * A reader of a document's revisions. The suggestions a room checkpoint lists are made first,
   * in its order and under their own ids, and `doc.revisionIds.listed` keeps their ids in that
   * order.
   */
  constructor(
    doc: OrderedDocument,
    fixed?: string,
    listed: readonly CheckpointSuggestion[] = [],
    prompt = false
  ) {
    this.doc = doc;
    this.fixed = fixed;
    this.prompt = prompt;
    for (const { id, attribution, revisions, record, status } of listed) {
      newSuggestion(doc, attribution, id);
      if (status !== undefined) {
        setStatus(doc, id, status);
      }

      // A record outside the stories is no mark of them (`loadPartRevisions` reads it)
      doc.revisionIds.listed.set(id, revisions);
      if (record === undefined) {
        this.addListed(id, attribution, revisions);
      }
    }
  }

  /**
   * A reader of a prompt the glossary kept: a mark a load read, a room checkpoint listed or the
   * room wrote for a suggestion's role is that suggestion's, as a reopen reads it, also with the
   * attribution the load read when a change gave the suggestion another (`suggestionOfRevision`),
   * and every other one is `fixed`.
   */
  static forPrompt(doc: OrderedDocument, fixed: string): RevisionReader {
    return new RevisionReader(doc, fixed, [], true);
  }

  /**
   * Read the marks a checkpoint listed for a suggestion as that suggestion's: by their triple, and
   * by their `w:id` alone for a mark that writes nothing else (`CT_Markup`: a `w:tblGridChange`),
   * so a fold keeps its owner.
   */
  private addListed(
    id: string,
    { author, date }: OrderedAttribution,
    revisions: readonly string[]
  ): void {
    for (const revision of revisions) {
      // An id names one suggestion: two share one only when two minted ids meet
      for (const triple of [
        revisionTriple(revision, author, date),
        revisionTriple(revision, '', undefined),
      ]) {
        if (!this.listedByTriple.has(triple)) {
          this.listedByTriple.set(triple, id);
        }
      }
    }
  }

  /**
   * The suggestion for one revision element, its `w:id` kept for `role`. Both halves of a named
   * move (`move`) by one author are one suggestion: the engine decides them together.
   */
  suggestionFor(revision: OoxmlElement, role: string, move?: string): string {
    const triple = tripleOf(revision);
    const listed =
      this.listedByTriple.get(triple) ??
      (this.prompt ? suggestionOfRevision(this.doc, revision) : undefined);
    if (this.fixed !== undefined && listed === undefined) {
      return this.fixed;
    }

    const author = attributeOf(revision, 'author') ?? '';
    const date = attributeOf(revision, 'date');
    const moveKey = move === undefined ? undefined : `${move}\u0000${author}`;
    let suggestion =
      this.byTriple.get(triple) ??
      listed ??
      (moveKey === undefined ? undefined : this.byMove.get(moveKey));
    if (suggestion === undefined) {
      const attributes = otherAttributes(revision);
      suggestion = newSuggestion(this.doc, {
        author,
        person: null,
        ...(date === undefined ? {} : { date }),
        ...(attributes.length > 0 ? { attributes } : {}),
      }).id;
    }

    this.byTriple.set(triple, suggestion);
    if (moveKey !== undefined && !this.byMove.has(moveKey)) this.byMove.set(moveKey, suggestion);
    seedRevisionId(this.doc, suggestion, role, revision);
    return suggestion;
  }

  /**
   * Keep a file mark's `w:id` for a suggestion's role, as `suggestionFor` does. A template keeps
   * none: its marks are the view's, not the file's.
   */
  keepId(suggestion: string, role: string, revision: OoxmlElement): void {
    if (this.fixed === undefined) seedRevisionId(this.doc, suggestion, role, revision);
  }

  /**
   * One suggestion for several marks of one author and date (the per-run marks of a struck
   * equation): the first mark's `w:id` is kept for `role`, and the others name the same one.
   */
  suggestionForAll(revisions: readonly OoxmlElement[], role: string): string {
    const suggestion = this.suggestionFor(revisions[0]!, role);
    if (this.fixed !== undefined) return suggestion;
    for (const revision of revisions.slice(1)) {
      const triple = tripleOf(revision);
      if (!this.byTriple.has(triple)) this.byTriple.set(triple, suggestion);
    }

    return suggestion;
  }

  /** The suggestion's creation order. */
  seqOf(suggestion: string): number {
    return this.doc.suggestions.get(suggestion)?.seq ?? 0;
  }
}

/** A revision mark's identity: its `w:id`, author and date. */
function tripleOf(revision: OoxmlElement): string {
  return revisionTriple(
    attributeOf(revision, 'id') ?? '',
    attributeOf(revision, 'author') ?? '',
    attributeOf(revision, 'date')
  );
}

const revisionTriple = (id: string, author: string, date: string | undefined): string =>
  `${id}\u0000${author}\u0000${date ?? '\u0001'}`;

/** Most other attributes kept from one revision mark. */
export const MAX_OTHER_ATTRIBUTES = 8;
const TRIPLE = new Set(['id', 'author', 'date']);

/** Whether an attribute is of a revision mark's triple: `w:id`, `w:author` or `w:date`. */
export const isTripleAttribute = ({ namespaceUri, localName }: OoxmlAttribute): boolean =>
  namespaceUri === WML_NAMESPACE_URI && TRIPLE.has(localName);

/** A revision mark's attributes past its triple, such as `w16du:dateUtc`. */
function otherAttributes(revision: OoxmlElement): OoxmlAttribute[] {
  return revision.attributes.filter((a) => !isTripleAttribute(a)).slice(0, MAX_OTHER_ATTRIBUTES);
}

/** A move range a file opened and the wrappers after it belong to. */
export interface MoveRange {
  name: string | null;
  /**
   * The range start that waits for the move's first wrapper, its `w:id` and attributes kept for
   * that wrapper's suggestion.
   */
  start: OoxmlElement | null;
  /** The range start's item, labelled once the first wrapper names the move's suggestion. */
  startIndex: number | null;
  /** The suggestion of the move's last wrapper: the range end's when its `w:id` names no start. */
  last: string | null;
  /** Each range's suggestion, by its start's `w:id`: a range end names its start by that id. */
  readonly byId: Map<string, string>;
}

/** The move's name: its range's, or one per author when the file names none. */
export function moveName(range: MoveRange, node: OoxmlElement): string {
  return range.name ?? `mv:${attributeOf(node, 'author') ?? ''}`;
}

/** One story's load. */
export interface LoadState {
  readonly doc: OrderedDocument;
  readonly story: string;
  readonly items: OrderedItem[];
  readonly revisions: RevisionReader;
  readonly budget: LoadBudget;
  /** Item indexes of a control's prompt units: marked once the story is loaded. */
  readonly prompts: Set<number>;
  readonly moveFrom: MoveRange;
  readonly moveTo: MoveRange;
  /** The paragraph being loaded: its items keep its id. */
  home: string | null;
  /**
   * The roles of the story's paragraph marks, and of its inserted runs (`run`); a notes part has
   * its own.
   */
  readonly markRoles: { readonly insert: string; readonly delete: string; readonly run: string };
}

export function newLoadState(
  doc: OrderedDocument,
  story: string,
  revisions: RevisionReader,
  budget: LoadBudget,
  markRoles: LoadState['markRoles'] = MARK_ROLES
): LoadState {
  return {
    doc,
    story,
    markRoles,
    items: [],
    revisions,
    budget,
    prompts: new Set(),
    moveFrom: { name: null, start: null, startIndex: null, last: null, byId: new Map() },
    moveTo: { name: null, start: null, startIndex: null, last: null, byId: new Map() },
    home: null,
  };
}

/**
 * Run a load inside an edit (an engine-built template, a new text box): the edit's undo record
 * takes back what the load adds to the document's tables past the items (run shells, moves with
 * range markers, revision ids), so a copy that takes the edit back, to apply it again after the
 * room's changes, holds what it held before. Outside an edit (a file's load) it only loads.
 */
export function loadInEdit<T>(doc: OrderedDocument, load: () => T): T {
  if (doc.undo === null) return load();
  const { byRole, attributesByRole, fileIds } = doc.revisionIds;
  const tables: readonly { keys(): Iterable<string>; delete(key: string): boolean }[] = [
    doc.runShells,
    doc.movesWithRanges,
    byRole,
    attributesByRole,
    fileIds,
  ];
  const before = tables.map((table) => new Set(table.keys()));
  const value = load();
  const added = tables.map((table, k) => [...table.keys()].filter((key) => !before[k]?.has(key)));
  if (added.some((keys) => keys.length > 0))
    recordUndo(doc, () => {
      tables.forEach((table, k) => {
        for (const key of added[k] ?? []) table.delete(key);
      });
    });
  return value;
}

/**
 * Mark a control's prompt once the control is loaded: its open item (at `first - 1`) and each
 * item from `first` on that `inPrompt` takes. A control inside the prompt keeps its own items,
 * as a control the room puts in a prompt keeps them (`ownItems`): only its open item is marked,
 * and its content stays its own, shown or not by its own rule. The edges of a link or another
 * container in the prompt are not marked, as the room makes a new link round prompt words: only
 * a control's open item reads as a control that holds a prompt.
 */
export function markPrompt(
  state: LoadState,
  first: number,
  inPrompt: (item: OrderedItem) => boolean
): void {
  state.prompts.add(first - 1);
  let inner: string | undefined;
  for (let k = first; k < state.items.length; k += 1) {
    const item = state.items[k];
    if (!item) {
      continue;
    }

    if (inner !== undefined) {
      if (item.kind === 'close' && item.pair === inner) {
        inner = undefined;
      }

      continue;
    }

    if (item.kind === 'open' && (item.token === 'sdt' || item.token === 'sdtBlock')) {
      state.prompts.add(k);
      inner = item.pair;
    } else if (item.kind !== 'open' && item.kind !== 'close' && inPrompt(item)) {
      state.prompts.add(k);
    }
  }
}

/** The story's items, each of a control's prompt units marked as such. */
export function itemsWithPrompts(state: LoadState): OrderedItem[] {
  return state.items.map((item, k) => (state.prompts.has(k) ? { ...item, prompt: true } : item));
}

/**
 * A story's loaded items, each one the file labelled marked `rebuilt`: a take-back takes its
 * labels off, but its loaded paragraph still shows the file's record.
 */
export function withFileLabels(items: readonly OrderedItem[]): OrderedItem[] {
  return items.map((item) => (isLabelled(item) ? { ...item, rebuilt: true } : item));
}

/** Refuse a walk past the depth limit. */
export function checkDepth(state: LoadState, depth: number): void {
  if (depth > state.budget.maxDepth)
    throw new OrderedRefusal('too-deep', `${state.story} nests past ${state.budget.maxDepth}`);
}

/** Count one more item against the load's item limit. */
export function chargeItem(budget: LoadBudget): void {
  if (budget.items >= budget.maxItems) {
    throw new OrderedRefusal('too-many-items', `more than ${budget.maxItems}`);
  }

  budget.items += 1;
}

/** A new item at the end of the story, within the item limit. Returns its index. */
export function pushItem(state: LoadState, init: OrderedItemInit): number {
  chargeItem(state.budget);
  state.items.push(createOrderedItem(state.home === null ? init : { ...init, home: state.home }));
  return state.items.length - 1;
}

/** A marker with no width, kept whole. */
export function pushZero(
  state: LoadState,
  node: OoxmlElement,
  labels: Labels,
  more: Partial<OrderedItemInit> = {}
): number {
  const at = pushItem(state, {
    space: zeroSpace(state.doc.file, node.id),
    kind: 'zero',
    node,
    ...labels,
    ...more,
  });
  noteCommentMarker(state.doc, state.story, state.items[at]!);
  return at;
}

/** A container's open item. */
export function pushOpen(
  state: LoadState,
  element: OoxmlElement,
  token: OrderedContainerToken,
  shell: OoxmlElement,
  labels: Labels,
  props: readonly OrderedPropertyRecord[] = []
): OrderedItem {
  const space = tokenSpace(state.doc.file, element.id);
  const at = pushItem(state, {
    space,
    offset: 0,
    kind: 'open',
    token,
    pair: space,
    node: shell,
    ...labels,
    props,
  });
  return state.items[at]!;
}

/** The close item of an open one. */
export function pushClose(state: LoadState, open: OrderedItem, labels: Labels): void {
  pushItem(state, {
    space: open.space,
    offset: 1,
    kind: 'close',
    ...(open.token === undefined ? {} : { token: open.token }),
    ...(open.pair === undefined ? {} : { pair: open.pair }),
    ...labels,
  });
}

/**
 * Labels a mark or a row gets from the `w:ins`, `w:del`, `w:moveFrom` and `w:moveTo` inside its
 * properties (a paragraph mark's `w:rPr`, a row's `w:trPr`).
 */
export function markerLabels(
  state: LoadState,
  properties: OoxmlElement | undefined,
  outer: Labels,
  roles: { readonly insert: string; readonly delete: string }
): Labels {
  if (!properties) return outer;
  let by = outer.by;
  let move = outer.move;
  const dels = [...outer.dels];
  for (const child of properties.children) {
    if (child.kind === 'textValue') continue;
    if (isWml(child, 'ins') || isWml(child, 'cellIns')) {
      by = state.revisions.suggestionFor(child, roles.insert);
    } else if (isWml(child, 'del') || isWml(child, 'cellDel')) {
      dels.push(state.revisions.suggestionFor(child, roles.delete));
    } else if (isWml(child, 'moveTo')) {
      by = state.revisions.suggestionFor(child, roles.insert);
      move = moveName(state.moveTo, child);
    } else if (isWml(child, 'moveFrom')) {
      dels.push(state.revisions.suggestionFor(child, roles.delete));
      move = moveName(state.moveFrom, child);
    }
  }

  return { by, dels, move, within: outer.within };
}

/** The roles of a row's or a cell's markers. */
export const ROW_ROLES = { insert: REVISION_ROLE.rowInsert, delete: REVISION_ROLE.rowDelete };

/** The roles of a paragraph mark's markers. */
export const MARK_ROLES = {
  insert: REVISION_ROLE.markInsert,
  delete: REVISION_ROLE.markDelete,
  run: REVISION_ROLE.insert,
};

/** The roles of a paragraph mark's markers in a notes part. */
export const NOTE_MARK_ROLES = {
  insert: REVISION_ROLE.noteMarkInsert,
  delete: REVISION_ROLE.noteMarkDelete,
  run: REVISION_ROLE.noteInsert,
};
