// The ordered suggestion model's types: one ordered list of items per story, held in a summed
// B+ tree, where a suggestion is a label on items and a decision only changes a status.

import type {
  OoxmlAttribute,
  OoxmlElement,
  OoxmlNode,
  OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import type { OoxmlProperty, TabStopWrite } from '../../../store/store/tree-op-types.ts';
import type { SumTree, SumTreeItem } from '../tree/sum-tree.ts';
import type { OrderedComment, OrderedCommentRange } from './comments.ts';
import type { OrderedObjectRecord } from './objects.ts';
import type { OrderedPerson } from './person.ts';
import type { OrderedResources } from './resources.ts';
import type { OrderedSectionRecord } from './sections.ts';

/**
 * What one item is: a run of units of one kind. Characters (`text`), one object (`atom`: a
 * picture, a tab, a break, a note reference, a whole field, an equation), one marker with no
 * width (`zero`: a bookmark, a comment range, a positional tab), one paragraph mark (`mark`), or
 * one edge of a container (`open` and `close`: a link, a control, a table, a row, a cell).
 *
 * @public
 */
export type OrderedItemKind = 'text' | 'atom' | 'zero' | 'mark' | 'open' | 'close';

/** Whether an item is a unit the view draws: characters or an object. */
export const isUnit = (item: OrderedItem): boolean => item.kind === 'text' || item.kind === 'atom';

/**
 * Whether a unit can hold a run format: a unit in a run, or a bare unit with runs in it (a simple
 * field, whose runs take the format). Any other bare unit (an equation, an element kept whole)
 * has no run for the format's change record, so a format leaves it as it is, as Word does.
 */
export const holdsRunFormat = (item: Pick<OrderedItem, 'bare' | 'pieces'>): boolean =>
  item.bare !== true || (item.pieces ?? []).some((piece) => piece.node.kind === 'run');

/** A suggestion's decision state. @public */
export type OrderedSuggestionStatus = 'pending' | 'accepted' | 'rejected';

/** Who made a suggestion, and when (ISO 8601, written as `w:date`). @public */
export interface OrderedAttribution {
  /** The name written as `w:author`: a room item's person's Word name, or the name Word saved. */
  readonly author: string;
  /** The person the room stamped; null for an item loaded from an uploaded file. */
  readonly person: OrderedPerson | null;
  readonly date?: string;
  /**
   * Other attributes a file's revision carried (such as `w16du:dateUtc`), written back on every
   * revision mark of the suggestion as the file had them.
   */
  readonly attributes?: readonly OoxmlAttribute[];
}

/**
 * A unit's address: the space it was inserted in and its offset there. It never changes during
 * a room session, whatever the edits and decisions.
 *
 * @public
 */
export interface OrderedItemAddress {
  readonly space: string;
  readonly offset: number;
}

/** The kind of container an `open` and its `close` item stand for. @public */
export type OrderedContainerToken =
  | 'hyperlink'
  | 'sdt'
  | 'smartTag'
  | 'customXml'
  | 'dir'
  | 'bdo'
  | 'fldSimple'
  | 'tbl'
  | 'tr'
  | 'tc'
  | 'sdtBlock'
  | 'customXmlBlock'
  | 'note';

/** Container tokens that hold blocks (paragraphs and tables), not runs. */
export const BLOCK_TOKENS: ReadonlySet<OrderedContainerToken> = new Set([
  'tbl',
  'tr',
  'tc',
  'sdtBlock',
  'customXmlBlock',
  'note',
]);

/**
 * A run format press: the run properties it sets, and those it takes off, as the engine's
 * `setRunProperties` takes off a property it leaves out.
 *
 * @public
 */
export interface OrderedFormatWrite {
  readonly properties: readonly OoxmlProperty[];
  readonly remove?: readonly string[];
}

/**
 * A format suggestion on text, atoms or a paragraph mark: run properties it sets (an edit), or
 * the whole face it gives (a file's `w:rPrChange`, whose current `w:rPr` is the suggestion's).
 *
 * @public
 */
export type OrderedFormatRecord =
  | {
      readonly kind: 'properties';
      /** The record's own id. It stays when another suggestion takes the record over. */
      readonly id: string;
      readonly suggestion: string;
      readonly seq: number;
      readonly properties: readonly OoxmlProperty[];
      /** The run properties the record takes off (Clear All Formatting, a toggle off). */
      readonly remove?: readonly string[];
      /**
       * The character style (`w:rStyle`) the record gives the units, or takes off them: a new
       * link's words take the link style and an unlinked link's words lose it.
       */
      readonly style?: OrderedStyleWrite;
      /** The suggestion a later format took this record from (policy 2), if any. */
      readonly replaced?: string;
      /**
       * A direct format made over pending format records: their suggestions. Word keeps one
       * record, the face before the first pending format, so a reject of one of them takes this
       * format away too, as the engine's reject does.
       */
      readonly over?: readonly string[];
      /**
       * A format on its author's own pending words: part of their insertion, shown with no
       * change record, as the engine formats its own insertion.
       */
      readonly untracked?: boolean;
      /** The change that made an `untracked` format: its undo takes the format off. */
      readonly change?: string;
    }
  | {
      readonly kind: 'face';
      readonly id: string;
      readonly suggestion: string;
      readonly seq: number;
      readonly face: OoxmlElement | null;
      /** The file's own `w:rPrChange`, written back while it is the only record and pending. */
      readonly fileRecord?: OoxmlElement;
      /** The suggestion a later format took this record from (policy 2), if any. */
      readonly replaced?: string;
    };

/**
 * A character style write: `set` puts `w:rStyle` first in the run properties, in place of any
 * other; `remove` takes off a `w:rStyle` that names the style.
 *
 * @public
 */
export type OrderedStyleWrite = { readonly set: string } | { readonly remove: string };

/**
 * The properties element a property record targets. Two are targets only for a file's own
 * record, which no edit writes: a row's table property exceptions (`w:tblPrEx`, its
 * `w:tblPrExChange`) and a paragraph's numbering reference (`w:numPr`, its `w:numberingChange`).
 *
 * @public
 */
export type OrderedPropertyTarget =
  | 'pPr'
  | 'numPr'
  | 'tblPr'
  | 'tblGrid'
  | 'tblPrEx'
  | 'trPr'
  | 'tcPr'
  | 'sdtPr'
  | 'sectPr';

/**
 * A paragraph property press, as the engine's own op without its paragraph and revision. A
 * record keeps the op, not its result, so each suggestion keeps its own meaning: compose runs
 * the accepted ops in creation order, then the pending ones with Word's change record.
 *
 * @public
 */
export type OrderedParagraphOp =
  | {
      readonly op: 'setParagraphProperties';
      readonly properties: readonly OoxmlProperty[];
      /** The paragraph properties the press takes off (Clear All Formatting). */
      readonly remove?: readonly string[];
    }
  | { readonly op: 'setListLevel'; readonly level: number }
  | { readonly op: 'setListNumbering'; readonly numId: string | null; readonly level?: number }
  | {
      readonly op: 'setParagraphTabStops';
      readonly stops: readonly TabStopWrite[];
      readonly inForcePositionsTwips?: readonly number[];
    };

/**
 * A property suggestion on a paragraph mark or a container's open item: the properties element
 * the suggestion gives (a file's `w:pPrChange`, `w:tblPrChange`, ...), a paragraph press, or a
 * section's note properties.
 *
 * @public
 */
export type OrderedPropertyRecord =
  | OrderedElementRecord
  | OrderedParagraphOpRecord
  | OrderedSectionRecord
  | OrderedObjectRecord;

/**
 * A property suggestion that gives a whole properties element (a file's change record, a
 * table op's result, a control's value).
 *
 * @public
 */
export interface OrderedElementRecord {
  readonly kind: 'element';
  readonly suggestion: string;
  readonly seq: number;
  readonly target: OrderedPropertyTarget;
  readonly element: OoxmlElement | null;
  /** The file's own change record, written back while it is the only record and pending. */
  readonly fileRecord?: OoxmlElement;
  /** Recorded even when it writes the same properties back (a table op touches every cell). */
  readonly recordUnchanged?: boolean;
  /** Shown while pending with no change record (a control's value: Word records none). */
  readonly untracked?: boolean;
  /** Applied only once accepted. */
  readonly acceptedOnly?: boolean;
  /** A table record: the spaces of the table's cells that showed when it was taken. */
  readonly tableCells?: readonly string[];
  /** The suggestion a later change of the same properties took this record from (policy 2). */
  readonly replaced?: string;
  /**
   * A direct change made over pending records of the same properties: their suggestions. Word
   * keeps one record, the properties before the first pending change, so a reject of one of
   * them takes this change away too, as the engine's reject does.
   */
  readonly over?: readonly string[];
  /**
   * The children the record sets, as `<namespace URI> <local name>` (a control's `w:sdtPr`):
   * it replaces those and keeps the others; a key with no child in `element` removes it. Absent:
   * the record gives the whole element.
   */
  readonly merge?: readonly string[];
}

/**
 * A paragraph property press on a paragraph mark: the engine's op, run again by compose.
 *
 * @public
 */
export interface OrderedParagraphOpRecord {
  readonly kind: 'paragraphOp';
  readonly suggestion: string;
  readonly seq: number;
  readonly target: 'pPr';
  readonly op: OrderedParagraphOp;
  /** Shown while pending with no change record (a press on its author's own new paragraph). */
  readonly untracked?: boolean;
  /** The suggestion a later change of the same properties took this record from (policy 2). */
  readonly replaced?: string;
  /** A direct press made over pending records of the paragraph's properties: as for an element. */
  readonly over?: readonly string[];
}

/**
 * The note properties one write gives: the fields it names replace those the document or the
 * section had, and the others stay (`w:pos`, `w:numFmt`, `w:numStart`, `w:numRestart`).
 *
 * @public
 */
export interface OrderedNoteProperties {
  readonly position?: string;
  readonly numFmt?: string;
  readonly numStart?: number;
  readonly numRestart?: string;
}

/**
 * A write of the notes' properties, as the engine's `setNoteProperties` without its scope: the
 * footnotes', the endnotes', or both.
 *
 * @public
 */
export interface OrderedNotePropertiesOp {
  readonly op: 'setNoteProperties';
  readonly footnote?: OrderedNoteProperties;
  readonly endnote?: OrderedNoteProperties;
}

/** One node a whole field covers, kept as the file had it. @public */
export interface OrderedFieldPiece {
  readonly node: OoxmlNode;
  /** The run the piece came from: pieces of one run stay one run. */
  readonly sourceRun?: string;
  /** Under a file's format records: the run's own face under each record, by record id. */
  readonly recordFaces?: ReadonlyMap<string, OoxmlElement | null>;
  /** The run's own change record (`w:rPrChange`) under each file record, by record id. */
  readonly recordFiles?: ReadonlyMap<string, OoxmlElement>;
}

/**
 * The labels of an item: what the suggestions did to it. Only `relabel` writes them, inside an
 * undo record.
 *
 * @public
 */
export interface OrderedItemLabels {
  /** The suggestion that inserted it; null for the file's own content. */
  by: string | null;
  /** The suggestions that struck it, in order. */
  dels: readonly string[];
  /** Format records, applied in `seq` order. */
  fmts: readonly OrderedFormatRecord[];
  /**
   * Pending format records a later press gave the unit's earlier face back over, so the engine
   * dropped the run's one change record (`withPressFormat`). They draw nothing and hold no card;
   * a take-back of that press gives them back.
   */
  cancelledFmts: readonly OrderedFormatRecord[];
  /** Property records (marks and container open items), applied in `seq` order. */
  props: readonly OrderedPropertyRecord[];
  /** A move's name: it links the struck source and the inserted destination. */
  move: string | null;
  /** Taken back by its own author: it never existed for anyone else. */
  retracted: boolean;
  /**
   * The change that took the item back for its author (a deletion over the author's own pending
   * item): the undo of that change shows it again. Absent or null otherwise, also once its
   * suggestion is taken back whole.
   */
  taken?: string | null;
  /** Its container went (a rejected row, an accepted struck row): it shows nowhere. */
  gone: boolean;
  /** Inside a new note or row: the container's own marker covers it. */
  quiet: boolean;
  /**
   * The suggestions that inserted or struck the rows and cells around it. It goes when they
   * do (a rejected row insertion, an accepted row deletion), and their marker is the
   * container's, never its own.
   */
  within: OrderedWithin;
}

/** The insertions and deletions of an item's containers. @public */
export interface OrderedWithin {
  readonly by: readonly string[];
  readonly dels: readonly string[];
}

/** What the tree sums for each item. @public */
export interface OrderedItemSummary {
  /** Units, shown or not. */
  units: number;
  /** Units the view shows, as the engine counts them: characters and atoms. */
  shown: number;
  /** Paragraph marks the view shows. */
  marks: number;
}

/**
 * One item. Fields other than the labels never change once the item is in a story. Nothing is
 * removed or moved: struck text that is accepted, and inserted text that is rejected, stay in
 * the list, hidden.
 *
 * @public
 */
export interface OrderedItem
  extends OrderedItemLabels, SumTreeItem<OrderedItem, OrderedItemSummary> {
  readonly space: string;
  /** Offset of the first unit in `space`. */
  readonly offset: number;
  /** Number of units: characters for text, 1 for every other kind. */
  readonly length: number;
  readonly kind: OrderedItemKind;
  /** The model text: characters for text, one character for an atom, '' otherwise. */
  readonly text: string;
  /**
   * The payload: an atom's or a marker's element; a mark's paragraph shell (its attributes and
   * its `w:pPr`); an open item's container shell (attributes and properties children).
   */
  readonly node: OoxmlNode | null;
  /** Text only: the source text element without its text (`w:t`, `w:instrText`), if any. */
  readonly textShell?: OoxmlElement;
  /** A whole field (complex or simple): the nodes it covers, in order. */
  readonly pieces?: readonly OrderedFieldPiece[];
  /** Run properties of text and atoms (null: none). A mark's are its paragraph's mark face. */
  readonly face: OoxmlElement | null;
  /** Units of one source run, or of one typed insertion, draw as one run. */
  readonly run: string;
  /** A marker inside a run (a positional tab, an empty `w:t`), not between runs. */
  readonly inRun?: boolean;
  /** A unit in the paragraph itself, not in a run (an equation, a simple field). */
  readonly bare?: boolean;
  /** A closed item between blocks (a section's properties, a block bookmark). */
  readonly block?: boolean;
  /** Text inside a field instruction. */
  readonly instr?: boolean;
  /**
   * The file labelled it at load (a revision record its loaded paragraph or block shows), so
   * they are drawn again from the items, also once a take-back took the labels off.
   */
  readonly rebuilt?: boolean;
  /**
   * A control's prompt unit, or (on an open item) a control that shows one or sits in another
   * control's prompt, or the mark of a paragraph of a block control's prompt of two or more
   * paragraphs.
   */
  readonly prompt?: boolean;
  /** Open and close items: the container kind. */
  readonly token?: OrderedContainerToken;
  /** Open and close items: the open item's space, shared by its close item. */
  readonly pair?: string;
  /** The paragraph the item was loaded in (its node id), for reuse of untouched paragraphs. */
  readonly home?: string;
  /** A mark: the units its paragraph held at load, so a whole untouched paragraph is reused. */
  readonly unitCount?: number;
  /**
   * A marker that sits between the pieces of a whole field (a form field's bookmark): the field
   * item's address (`space|offset`) and the piece it comes before. It draws there while the
   * field shows.
   */
  readonly fieldSlot?: { readonly field: string; readonly before: number };
  /** A file's move range marker: drawn bare, never wrapped, while its move is pending. */
  readonly moveRange?: boolean;
}

/**
 * A suggestion. Only `setStatus` writes `status`, and only the model's writes change `version`,
 * inside an undo record.
 *
 * @public
 */
export interface OrderedSuggestion {
  readonly id: string;
  /** Who made it, and when. Only `setAttribution` writes it after it is made. */
  attribution: OrderedAttribution;
  status: OrderedSuggestionStatus;
  /** Its order of creation: format and property records apply in this order. */
  readonly seq: number;
  /**
   * Goes up with every write that adds, removes or changes an item it labels, or changes the
   * labels of one: a decision names the version its organiser saw, and a change in between
   * makes the decision stale.
   */
  version: number;
  /**
   * A direct edit's suggestion, accepted when it was made: its author takes it back and makes
   * it again by turning its status (`revert`). A decided suggestion is not one.
   */
  direct?: boolean;
}

/** One story part (the main document, a header, a footer, the notes) as items. @public */
export interface OrderedStory {
  /** The part name, such as `/word/document.xml`. */
  readonly name: string;
  /**
   * The part the story was loaded from; compose keeps its root and replaces its blocks. A paste
   * declares on its root the prefixes its content uses that the root lacks (`declarePrefixes`).
   */
  part: OoxmlPart;
  readonly tree: SumTree<OrderedItem, OrderedItemSummary>;
  /** Space to its items, by offset. */
  readonly spaces: Map<string, OrderedItem[]>;
}

/**
 * A file's tracked change outside the stories: the change records of one element with one
 * `w:id`, author and date in a part the list does not hold as items (a style's `w:pPrChange` or
 * `w:rPrChange` in the styles part, a table style's `w:tblPrChange`, `w:trPrChange` or
 * `w:tcPrChange`, the document defaults'). The editor alone shows it as one review card. It is a
 * suggestion with no items: a decision turns its status, and compose writes the part from the
 * file's with the decided records resolved by the engine.
 *
 * @public
 */
export interface OrderedPartRevision {
  /** The part name, such as `/word/styles.xml`. */
  readonly part: string;
  /** The change record's local name, such as `pPrChange`. */
  readonly element: string;
  /** The records' `w:id`, author and date, as the file wrote them. */
  readonly revision: { readonly id: string; readonly author: string; readonly date?: string };
  /** The style that holds the records (`w:styleId`); absent for the document defaults'. */
  readonly styleId?: string;
}

/** Counters the model mints ids and orders from. @public */
export interface OrderedCounters {
  suggestion: number;
  /** Not read by the model: kept so the public shape stays the same. */
  insertion: number;
  seq: number;
}

/**
 * Word revision ids (`w:id`): the file's own, by suggestion and role, and every `w:id` the file
 * holds, which a new mark's id avoids. Load writes it; an id the file did not give is a function
 * of the suggestion's id and the mark's role (`revisionIdFor`), so every copy writes the same.
 *
 * @public
 */
export interface OrderedRevisionIds {
  /** The file's own `w:id` by suggestion and role. */
  readonly byRole: Map<string, string>;
  /** The file's own attributes past `w:id` (and a move range's `w:name`) by suggestion and role. */
  readonly attributesByRole: Map<string, readonly OoxmlAttribute[]>;
  /** Every revision `w:id` the loaded parts hold. */
  readonly fileIds: Set<string>;
  /**
   * The `w:id`s a room checkpoint lists for each suggestion it makes, those of a part the room
   * does not load among them (a prompt's records in the glossary), in the checkpoint's order.
   */
  readonly listed: Map<string, readonly string[]>;
}

/** One step of an undo record: how to take back one write. @public */
export type OrderedUndoStep = () => void;

/** The units `[from, to)` of one space of a story. @public */
export interface OrderedUnitRange {
  /** The story part, such as `/word/document.xml`. */
  readonly story: string;
  /** The space the units were inserted in. */
  readonly space: string;
  /** The offset of the first unit. */
  readonly from: number;
  /** The offset right after the last unit. */
  readonly to: number;
}

/** The whole document: its stories, its suggestions, and the indexes over them. @public */
export interface OrderedDocument {
  /**
   * The room's Word file number the document was loaded from. Every space a load makes holds
   * it, so an address from another Word file of the room never resolves here (design D8).
   */
  readonly file: number;
  readonly stories: Map<string, OrderedStory>;
  readonly suggestions: Map<string, OrderedSuggestion>;
  /** Suggestion to the items that carry one of its labels, for status changes. */
  readonly labelled: Map<string, Set<OrderedItem>>;
  /**
   * Change id to the items an edit of its author's own pending words changed: took back
   * (`taken`) or formatted (a format record's `change`). The undo of the change reads it.
   */
  readonly ownEdited: Map<string, Set<OrderedItem>>;
  /**
   * Change id to the units it added to its author's earlier pending suggestions (typing beside
   * or over one's own words, a tab, a break, Enter or a note inside them): the undo of the
   * change takes them back.
   */
  readonly added: Map<string, readonly OrderedUnitRange[]>;
  readonly counters: OrderedCounters;
  /** Run id to the run's shell (its attributes, no children), for compose. Load writes it. */
  readonly runShells: Map<string, OoxmlElement>;
  /** Node id to the loaded node, for untouched paragraphs and blocks. Load writes it. */
  readonly originals: Map<string, OoxmlNode>;
  readonly revisionIds: OrderedRevisionIds;
  /**
   * The glossary's placeholder blocks by `w:docPartPr/w:name`: the prompts that controls name
   * through `w:placeholder`. Load writes it.
   */
  readonly placeholders: Map<string, readonly OoxmlNode[]>;
  /** Moves whose range markers the file holds: compose writes no other range for them. */
  readonly movesWithRanges: Set<string>;
  /** The next number for a node id compose mints, so ids stay unique across composes. */
  readonly composeIds: { next: number };
  /**
   * The `w14:paraId` values new paragraphs took this session, upper-case: a new paragraph's id
   * avoids them and its story's own. Edits write it inside an undo record.
   */
  readonly paraIds: Set<string>;
  /**
   * Comment ids (`w:id`) the loaded parts use, in their range markers, references and comments:
   * a new comment's id avoids them. Load writes it.
   */
  readonly commentIds: Set<string>;
  /**
   * The comments by `w:id`, with their threads and resolved state: the main document's comments
   * part's, in the part's order, then those the changes added, in the order they were added.
   * Load and the comment changes write it, the changes inside an undo record.
   */
  readonly comments: Map<string, OrderedComment>;
  /** Where each comment's range sits, by `w:id`. Load and the comment changes write it. */
  readonly commentRanges: Map<string, OrderedCommentRange>;
  /**
   * The list numbers (`w:num/@w:numId`) the package's numbering defines, or null when the
   * document was loaded without its package. A new list definition is not covered.
   */
  numbering: ReadonlySet<string> | null;
  /**
   * What edits added outside the story text (media, story parts, relationships, list
   * definitions, styles, settings), and the loaded file's names they must not take (D14).
   */
  readonly resources: OrderedResources;
  /**
   * The file's tracked changes outside the stories, by suggestion id (the styles part's change
   * records). Load writes it; only a decision changes their suggestions.
   */
  readonly partRevisions: Map<string, OrderedPartRevision>;
  /**
   * The person of each namespace whose changes the document applied, which `orderedUnitTyper`
   * names: the apply writes it inside an undo record, a checkpoint's load for each namespace that
   * typed.
   */
  readonly namespaces: Map<string, OrderedPerson>;
  /** While an edit runs inside an undo record: how to take back each write, in order. */
  undo: OrderedUndoStep[] | null;
}

// Spaces a change mints (`i:`, `k:`, `c:`, `x:`) hold the change id. Spaces a load makes hold
// the Word file number and the node id (`LOADED_SPACE`).

/** A space name: units of one paragraph's loaded text and atoms. */
export const paragraphSpace = (file: number, paragraphId: string): string =>
  `b:${file}:${paragraphId}`;

/** A space name: one loaded paragraph's mark. */
export const markSpace = (file: number, paragraphId: string): string => `m:${file}:${paragraphId}`;

/** A space name: one loaded container's open (offset 0) and close (offset 1) items. */
export const tokenSpace = (file: number, nodeId: string): string => `t:${file}:${nodeId}`;

/** A space name: one loaded marker. */
export const zeroSpace = (file: number, nodeId: string): string => `z:${file}:${nodeId}`;

/**
 * A loaded space: `<kind>:<file>:<node id>`, kind `b`, `m`, `t` or `z`. Node ids are positions in
 * the file, so the same position in the room's next Word file has the same id; the file number
 * keeps an address from one file from naming a unit of another (design D8).
 */
const LOADED_SPACE = /^([bmtz]):(\d{1,15}):/;

/**
 * A space a load made: its kind letter, the Word file number and the node id. Null for a space
 * a change minted.
 */
export function loadedSpace(
  space: string
): { readonly kind: 'b' | 'm' | 't' | 'z'; readonly file: number; readonly node: string } | null {
  const match = LOADED_SPACE.exec(space);
  if (match === null) {
    return null;
  }

  // The pattern admits only these letters
  const kind = match[1] as 'b' | 'm' | 't' | 'z';
  return { kind, file: Number(match[2]), node: space.slice(match[0].length) };
}
