// What one compose knows and keeps: the document, the story, the options, the ids it placed,
// and how each item reads in this drawing (gone, struck, inserted).

import { inlineTextOf } from '../../../store/package/content-control-nodes.ts';
import type { OoxmlAttribute, OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { revisionAttributes } from '../../../store/store/tree-op-tracked.ts';
import { cloneWithNewIds } from '../../../store/store/tree-op-nodes.ts';
import {
  insertedPending,
  isHidden,
  isSplitEdge,
  itemAt,
  pendingDeleters,
  statusOf,
} from '../model/document.ts';
import { checkpointAttribute } from '../model/checkpoint.ts';
import type { OrderedPerson } from '../model/person.ts';
import { textBoxHostOf } from '../model/text-boxes.ts';
import { revisionIdFor, seededAttributes } from '../model/revision-ids.ts';
import { orderedUnitTyper } from '../model/typers.ts';
import { withAttributes } from '../ooxml-nodes.ts';
import type {
  OrderedAttribution,
  OrderedDocument,
  OrderedItem,
  OrderedItemAddress,
  OrderedStory,
  OrderedSuggestionStatus,
} from '../model/types.ts';
import type { RemovedColumns } from './table-columns.ts';

/** How to draw. @public */
export interface OrderedComposeOptions {
  /** Draw every paragraph and block again from its items, even untouched ones (for checks). */
  readonly rebuild?: boolean;
  /** `show` (default): pending suggestions as Word revisions. `reject`: as if all rejected. */
  readonly pending?: 'show' | 'reject';
}

/** How one compose draws: the public options, and what only a checkpoint's compose does. */
export interface ComposeOptions extends OrderedComposeOptions {
  /**
   * Set for a checkpoint: each element a unit is drawn in (its run, each run of a field, a unit
   * not in a run) carries the namespace of the room change that typed the unit inside a pending
   * insertion (`oc:typer`), and the person of each namespace written is kept here.
   */
  readonly typers?: Map<string, OrderedPerson>;
  /**
   * Set for a checkpoint: each run drawn from units carries the number of the run the room draws
   * them in, and each text element the number of the room's text element (`oc:run`), kept here
   * by what no file holds, with how many elements carry it. A pending suggestion may split a run
   * of the room into pieces; an open joins them again. A number on one element joins nothing, and
   * the checkpoint's label walk takes it off.
   */
  readonly runs?: Map<string, RoomNumber>;
}

/** The number of a room run or text element in a checkpoint's drawing, and its elements. */
export interface RoomNumber {
  readonly number: string;
  uses: number;
}

/** One shown unit of a paragraph's view. @public */
export interface OrderedViewChar {
  readonly address: OrderedItemAddress;
  readonly ch: string;
  /** The inline containers around it, outermost first: `h:<space>` links, `c:<space>` others. */
  readonly containers: string;
  readonly face: OoxmlElement | null;
  readonly deleted: boolean;
  readonly item: OrderedItem;
}

/** One paragraph of the view. @public */
export interface OrderedViewParagraph {
  /** The paragraph node's id in the composed part. */
  readonly id: string;
  readonly mark: OrderedItem;
  readonly chars: readonly OrderedViewChar[];
  /** The paragraph's inline items in order, shown or not. */
  readonly items: readonly OrderedItem[];
  /** The block containers it sits in (open items), outermost first. */
  readonly blocks: readonly OrderedItem[];
}

/**
 * A control's prompt that the composed part no longer shows, to keep in the glossary as the
 * placeholder block `name`, which the control's `w:placeholder` names. The shape matches the
 * engine's retained prompt: the caller writes it with the package's glossary.
 *
 * @public
 */
export interface OrderedRetainedPrompt {
  /** The `w:docPartPr/w:name` of the block. */
  readonly name: string;
  /** Whether the control sits in a paragraph, so the block wraps `nodes` in one. */
  readonly inline: boolean;
  /** The control's content as it showed the prompt. */
  readonly nodes: readonly OoxmlNode[];
  /** The story part that showed the prompt, whose relationships its nodes name. */
  readonly owner: string;
}

/**
 * A record a suggestion holds by takeover (policy 2), as a drawing draws it under the holder's
 * change record: the children of the properties before it and with it, without their own change
 * record (a nested one, such as `w:pPr/w:rPr/w:rPrChange`, stays).
 */
export interface DrawnTakeover {
  /** The suggestion the record was taken from. */
  readonly replaced: string;
  readonly before: readonly OoxmlNode[];
  readonly after: readonly OoxmlNode[];
}

/** A change record a drawing wrote that holds records by takeover. */
export interface HeldTakeovers {
  /** The change record's local name: `rPrChange`, `pPrChange`, `tcPrChange`, ... */
  readonly record: string;
  readonly taken: readonly DrawnTakeover[];
}

/** Bookkeeping of a revision wrapper compose wrote. */
export interface WrapperInfo {
  /** The whole chain of wrappers, outer first: a move's range covers one chain. */
  readonly key: string;
  /** The outer wrapper's kind, id and suggestion: neighbours with one outer wrapper share it. */
  readonly outer: string;
  readonly suggestion: string;
  /** A move's name, for its range markers. */
  readonly move: string | null;
}

/** The run the room draws a unit in, by what no file holds: its `run`, face and format records. */
const roomRunOf = (unit: OrderedItem): string =>
  [unit.run, unit.face?.id ?? '', ...unit.fmts.map((record) => record.id)].join(' ');

/** The number of a room run or text element in a checkpoint's drawing, the first time a new one. */
function roomNumber(runs: Map<string, RoomNumber>, key: string): string {
  const known = runs.get(key);
  if (known !== undefined) {
    known.uses += 1;
    return known.number;
  }

  const number = String(runs.size);
  runs.set(key, { number, uses: 1 });
  return number;
}

export class ComposeContext {
  readonly doc: OrderedDocument;
  readonly story: OrderedStory;
  readonly options: ComposeOptions;
  /** Ids of nodes this drawing placed: a second use of one node gets a copy with new ids. */
  readonly placed = new Set<string>();
  /** Run node ids this drawing gave (`runId`). */
  readonly #usedRunIds = new Set<string>();
  /** Controls a decision emptied: they show their prompt again. */
  readonly emptied: string[] = [];
  /**
   * The controls in `emptied` that a decision on a suggestion emptied, not a direct edit: they
   * show Word's blank prompt when they name none (`withDecisionPrompt`), as the engine's
   * decision leaves them.
   */
  readonly decisionEmptied = new Set<string>();
  /** Controls a decision reached: their value record follows the text they show. */
  readonly reached: string[] = [];
  /** Controls whose value record the words a reject took away stated: words of no value clear it. */
  readonly stale = new Set<string>();
  /** Whether this drawing put a checkpoint's attribute on an element: its part declares it. */
  markedAny = false;
  /** Prompt units of controls that hold content now, and the controls those prompts held. */
  readonly hiddenPrompts = new Set<OrderedItem>();
  /** Open item spaces of the controls whose prompt is hidden. */
  readonly unprompted = new Set<string>();
  /**
   * Open item spaces of the block controls whose prompt Enter removed for good (policy 4): no
   * glossary block keeps it, and the control, once empty, shows a prompt as any control does.
   */
  readonly promptRemoved = new Set<string>();
  /**
   * Open item spaces of the prompt controls a suggestion ever filled: they name the glossary
   * block that keeps their prompt, even once it shows again.
   */
  readonly filledPrompts = new Set<string>();
  /** Glossary blocks the caller must write for the prompts this drawing hides. */
  readonly prompts: OrderedRetainedPrompt[] = [];
  /** Tables whose removed cells this drawing took out, before and after (`table-columns.ts`). */
  readonly removedColumns: RemovedColumns[] = [];
  /** The takeovers each change record this drawing wrote holds, by the record's node id. */
  readonly takeovers = new Map<string, HeldTakeovers>();
  readonly wrappers = new WeakMap<OoxmlNode, WrapperInfo>();
  /**
   * Insertions whose inline control is being drawn inside its own `w:ins` (as the engine and
   * Word write a new control): their units in it take no second wrapper.
   */
  readonly covered = new Set<string>();
  /** The mark space of the paragraph being drawn: struck runs take one id per paragraph. */
  paragraphKey = '';
  /**
   * The node id of the paragraph being drawn. A loaded run's nodes (its shell, text forms and
   * face) keep their ids only in the paragraph they were loaded in: after a split both halves
   * hold that run, and the editor gets each paragraph composed on its own, so a second copy
   * under the same ids would give its tree two nodes with one id.
   */
  home: string | undefined = undefined;

  /** A new node id, unique across every compose of the document (or from `mint` when given). */
  readonly mint: () => string;

  /**
   * A text box's story whose object shows nowhere in this drawing (a rejected or struck text
   * box, or one in a rejected row): everything in it goes with the object, also what someone
   * typed there under a suggestion of their own.
   */
  readonly #objectGone: boolean;

  constructor(
    doc: OrderedDocument,
    story: OrderedStory,
    options: ComposeOptions,
    mint?: () => string
  ) {
    this.doc = doc;
    this.story = story;
    this.options = options;
    this.mint =
      mint ??
      ((): string => {
        const id = `ordered#${this.doc.composeIds.next}`;
        this.doc.composeIds.next += 1;
        return id;
      });
    const host = textBoxHostOf(doc, story.name);
    const hostStory = host ? doc.stories.get(host.host) : undefined;
    const object = host && hostStory ? itemAt(hostStory, host.atom) : null;
    this.#objectGone = host !== null && (object === null || this.unshown(object));
  }

  /**
   * The element units are drawn in, with the typer of `unit` on it when a checkpoint is drawn: one
   * room change typed every unit of an item, and of a run, whose key holds the item's `run`.
   */
  typed(element: OoxmlElement, unit: OrderedItem | undefined): OoxmlElement {
    const typers = this.options.typers;
    const typer =
      typers === undefined || unit === undefined
        ? null
        : orderedUnitTyper(this.doc, unit, unit.offset);
    if (typers === undefined || typer === null) {
      return element;
    }

    this.markedAny = true;
    typers.set(typer.namespace, typer.person);
    const attribute = checkpointAttribute('typer', typer.namespace);
    return withAttributes(element, [...element.attributes, attribute]);
  }

  /** A run drawn from `unit` and its neighbours, numbered by its room run for a checkpoint. */
  ofRoomRun(run: OoxmlElement, unit: OrderedItem): OoxmlElement {
    const runs = this.options.runs;
    if (runs === undefined) {
      return run;
    }

    this.markedAny = true;
    const number = roomNumber(runs, roomRunOf(unit));
    return withAttributes(run, [...run.attributes, checkpointAttribute('run', number)]);
  }

  /**
   * A text element of `unit`'s run, numbered for a checkpoint by the room's text element: the
   * text form the room draws it from (`shell`). A piece a pending paragraph break moved out of its
   * own paragraph takes the same number, so a reject draws the pieces as one element again.
   */
  ofRoomText(text: OoxmlNode, unit: OrderedItem, shell: OoxmlElement | undefined): OoxmlNode {
    const runs = this.options.runs;
    if (runs === undefined || text.kind === 'textValue') {
      return text;
    }

    const key = `text ${roomRunOf(unit)} ${shell?.id ?? ''}`;
    const number = roomNumber(runs, key);
    return withAttributes(text, [...text.attributes, checkpointAttribute('run', number)]);
  }

  /**
   * Run or paragraph properties drawn only for change records, where the run or the paragraph
   * has none, marked so when a checkpoint is drawn: an open reads them as none once a decision
   * leaves them empty, as the room draws none then.
   */
  drawn(properties: OoxmlElement): OoxmlElement {
    if (this.options.runs === undefined) {
      return properties;
    }

    this.markedAny = true;
    return withAttributes(properties, [
      ...properties.attributes,
      checkpointAttribute('drawn', '1'),
    ]);
  }

  /** Whether an item shows nowhere in this drawing, by its labels and the drawing's options. */
  unshown(item: OrderedItem): boolean {
    if (isHidden(this.doc, item)) return true;
    return this.options.pending === 'reject' && insertedPending(this.doc, item);
  }

  /** Whether a suggestion is rejected, or pending and drawn as rejected. */
  rejectedOrDrawnSo(suggestion: string): boolean {
    const status = statusOf(this.doc, suggestion);
    return status === 'rejected' || (status === 'pending' && this.options.pending === 'reject');
  }

  /**
   * Whether a container around the item went, by the decisions alone: its insertion rejected or
   * its deletion accepted.
   */
  containerWent(item: OrderedItem): boolean {
    if (item.within.by.some((id) => statusOf(this.doc, id) === 'rejected')) return true;
    return item.within.dels.some((id) => statusOf(this.doc, id) === 'accepted');
  }

  /** A run's node id: `id` when no run of this drawing has it yet, else a new one. */
  runId(id: string | undefined): string {
    const out = id !== undefined && !this.#usedRunIds.has(id) ? id : this.mint();
    this.#usedRunIds.add(out);
    return out;
  }

  /** Whether an item's loaded nodes may keep their ids here: it is drawn in its own paragraph. */
  homed(item: OrderedItem): boolean {
    return item.home !== undefined && item.home === this.home;
  }

  /**
   * `node` itself the first time, or a copy with new ids when one of its ids is placed, or
   * when it is not `homed` (drawn outside the paragraph it was loaded in).
   */
  fresh<T extends OoxmlNode | null>(node: T, homed = true): T {
    if (node === null || node.kind === 'textValue') return node;

    // Two faces can share children, so every id in the node counts, not only its own
    const ids: string[] = [];
    const visit = (n: OoxmlNode): void => {
      ids.push(n.id);
      if (n.kind !== 'textValue') for (const child of n.children) visit(child);
    };
    visit(node);
    if (homed && !ids.some((id) => this.placed.has(id))) {
      for (const id of ids) this.placed.add(id);
      return node;
    }

    const copy = cloneWithNewIds(node, this.mint);
    const originals = this.takeovers.size > 0 ? ids.slice() : [];
    ids.length = 0;
    visit(copy);
    for (const id of ids) this.placed.add(id);

    // A copy's change records hold the takeovers their originals did
    for (const [k, original] of originals.entries()) {
      const held = this.takeovers.get(original);
      const id = ids[k];
      if (held !== undefined && id !== undefined) {
        this.takeovers.set(id, held);
      }
    }

    // cloneWithNewIds keeps the node's kind, so the copy is a T as the node is
    return copy as T;
  }

  /** Whether the item shows nowhere in this drawing. */
  gone(item: OrderedItem): boolean {
    if (this.#objectGone || isHidden(this.doc, item)) return true;
    if (this.hiddenPrompts.has(item)) {
      return true;
    }

    // A rejected Enter keeps the link or control it split in two (`isSplitEdge`)
    if (isSplitEdge(item)) return false;
    return this.options.pending === 'reject' && insertedPending(this.doc, item);
  }

  /**
   * Whether a decision on a suggestion took the unit away: a rejected insertion or an accepted
   * deletion, not a direct edit (whose suggestion is accepted when made). A drawing as if every
   * suggestion were rejected rejects the pending insertions.
   */
  decidedAway(item: OrderedItem): boolean {
    if (item.gone || item.retracted || !this.gone(item)) return false;
    if (this.options.pending === 'reject' && insertedPending(this.doc, item)) return true;
    const decision = (suggestion: string, status: OrderedSuggestionStatus): boolean => {
      const found = this.doc.suggestions.get(suggestion);
      return found?.status === status && found.direct !== true;
    };
    if (item.by !== null && decision(item.by, 'rejected')) return true;
    return item.dels.some((deleter) => decision(deleter, 'accepted'));
  }

  /**
   * Whether the decisions so far change whether a unit reads as text, against the view with
   * every suggestion pending. The engine updates a control's value record when they do. A
   * prompt hides and shows with the content over it, which is no decision: typing over it
   * writes no value record, in the engine too.
   */
  decided(item: OrderedItem): boolean {
    if (item.gone || item.retracted || item.prompt === true) {
      return false;
    }

    // A direct edit taken back (`revert`) decides nothing: the unit is as it was before it
    if (item.by !== null && this.revertedDirect(item.by)) return false;
    const before = !item.dels.some((deleter) => !this.revertedDirect(deleter));
    const now = !this.gone(item) && this.struckBy(item) === null;
    if (before === now) {
      return false;
    }

    // A field's code, a field with no result or a picture reads as no text, as the engine reads it
    if (item.kind === 'text') {
      return item.instr !== true;
    }

    const nodes =
      item.pieces?.map((piece) => piece.node) ?? (item.node === null ? [] : [item.node]);
    return inlineTextOf(nodes) !== '';
  }

  /** A direct edit's suggestion its author took back. */
  revertedDirect(suggestion: string): boolean {
    const found = this.doc.suggestions.get(suggestion);
    return found?.direct === true && found.status === 'rejected';
  }

  /** The suggestion that shows as striking the item: the first pending one (policy 1). */
  struckBy(item: OrderedItem): string | null {
    if (this.options.pending === 'reject') return null;
    return pendingDeleters(this.doc, item)[0] ?? null;
  }

  /** The words of a unit whose insertion a reject took away, as a control reads its text, or ''. */
  rejectedWords(item: OrderedItem): string {
    const found = item.by === null ? undefined : this.doc.suggestions.get(item.by);
    const rejected = found?.status === 'rejected' && found.direct !== true;
    return rejected && !item.gone && item.kind === 'text' && item.instr !== true ? item.text : '';
  }

  /** Whether a shown unit is a control's own word: neither struck nor a suggestion's insertion. */
  ownWord(item: OrderedItem): boolean {
    return (
      this.struckBy(item) === null &&
      !this.acceptedInsertion(item) &&
      this.insertedBy(item) === null
    );
  }

  /** Whether an accept decided the item's insertion: a suggestion's, not a direct edit's. */
  acceptedInsertion(item: OrderedItem): boolean {
    const found = item.by === null ? undefined : this.doc.suggestions.get(item.by);
    return found?.status === 'accepted' && found.direct !== true;
  }

  /** The pending suggestion that shows as inserting the item, if its container does not. */
  insertedBy(item: OrderedItem): string | null {
    if (this.options.pending === 'reject' || item.quiet) return null;
    if (item.by !== null && this.covered.has(item.by)) return null;
    return insertedPending(this.doc, item) ? item.by : null;
  }

  attribution(suggestion: string): OrderedAttribution {
    return this.doc.suggestions.get(suggestion)?.attribution ?? { author: '', person: null };
  }

  /**
   * A suggestion's revision mark attributes for a role, its `w:id` first: the file's own mark
   * for that role when it had one, else the author, date and the file's other attributes.
   */
  revisionMark(suggestion: string, role: string): OoxmlAttribute[] {
    const id = this.revisionId(suggestion, role);
    const [idAttribute] = revisionAttributes(id, { author: '' });
    const seeded = seededAttributes(this.doc, suggestion, role);
    if (idAttribute && seeded) return [idAttribute, ...seeded];
    const attribution = this.attribution(suggestion);
    return [...revisionAttributes(id, attribution), ...(attribution.attributes ?? [])];
  }

  /** Change records the engine wrote (`w:rPrChange`, ...) with the suggestion's mark for `role`. */
  withRecordAttributes(
    children: readonly OoxmlNode[],
    suggestion: string,
    role: string
  ): readonly OoxmlNode[] {
    return children.map((child) =>
      child.kind !== 'textValue' && child.localName.endsWith('Change')
        ? withAttributes(child, this.revisionMark(suggestion, role))
        : child
    );
  }

  /**
   * Keep the takeovers a drawn change record holds: `children` hold the record (a properties
   * element's children), and `taken` the records it took over, as drawn.
   */
  keepTakeovers(children: readonly OoxmlNode[], taken: readonly DrawnTakeover[]): void {
    if (taken.length === 0) {
      return;
    }

    const record = children.find(isChangeRecord);
    if (record !== undefined) {
      this.takeovers.set(record.id, { record: record.localName, taken });
    }
  }

  revisionId(suggestion: string, role: string): string {
    return revisionIdFor(this.doc, suggestion, role);
  }
}

/** Whether a node is a change record: `w:rPrChange`, `w:pPrChange`, `w:tcPrChange`, ... */
export const isChangeRecord = (node: OoxmlNode): node is OoxmlElement =>
  node.kind !== 'textValue' && node.localName.endsWith('Change');
