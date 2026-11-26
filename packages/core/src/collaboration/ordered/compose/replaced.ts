// The suggestions another took records from, in the review queue's terms (design D5, policy 2).

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import {
  changedFormatting,
  FORMATTING_PROPERTIES,
  type ReviewFormattingChange,
} from '../../../store/store/review-formatting.ts';
import type { ReviewRevisionItem } from '../../../store/store/review-items.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import { isOrderedItemHidden } from '../card-of.ts';
import { isPending } from '../model/document.ts';
import type { OrderedPerson } from '../model/person.ts';
import type { OrderedDocument, OrderedFormatRecord, OrderedItem } from '../model/types.ts';
import { childWhere, isParagraphProperties } from '../ooxml-nodes.ts';
import { ComposeContext, type HeldTakeovers } from './context.ts';
import { containerPropertiesAfter, faceAfter, paragraphPropertiesAfter } from './properties.ts';
import { sectionAfter, sectionPropertiesOf, sectionRecordsOf } from './sections.ts';

/**
 * A suggestion whose records another holds now by takeover (policy 2): a later change of the same
 * properties took them, as Word keeps one change record for each properties element.
 *
 * @public
 */
export interface OrderedReplacedSuggestion {
  /** The suggestion the records were taken from. It is pending: a take-back gives them back. */
  readonly suggestion: string;
  /** Its Word name (`w:author`). */
  readonly author: string;
  /** The room's person who made it; null for a suggestion of an uploaded file. */
  readonly person: OrderedPerson | null;
  /** When it was made (`w:date`); absent when its file wrote none. */
  readonly date?: string;
  /**
   * The local name of the change record that holds its records: `rPrChange`, `pPrChange`,
   * `tblPrChange`, `trPrChange`, `tcPrChange`, `sectPrChange`, and so on. When change records of
   * several kinds hold them, the first one the document draws names it, and only its changes count.
   */
  readonly formattingKind: string;
  /** What it had set, as the review queue's items name formatting changes. */
  readonly formattingChanges: NonNullable<ReviewRevisionItem['formattingChanges']>;
}

/**
 * The suggestions whose records `suggestion` holds now by takeover, each once, in the order the
 * document first draws their records. A copy of the room and a copy opened from its checkpoint
 * give the same answer. Empty when it took over nothing.
 *
 * @public
 */
export function orderedReplacedSuggestions(
  doc: OrderedDocument,
  suggestion: string
): OrderedReplacedSuggestion[] {
  const out: OrderedReplacedSuggestion[] = [];
  for (const [id, { kind, changes }] of takenChanges(doc, suggestion)) {
    const replaced = doc.suggestions.get(id);
    if (replaced !== undefined) {
      const { author, person, date } = replaced.attribution;
      out.push({
        suggestion: id,
        author,
        person,
        ...(date === undefined ? {} : { date }),
        formattingKind: kind,
        formattingChanges: changes,
      });
    }
  }

  return out;
}

/** What a replaced suggestion had set: the kind of its change record, and the changes. */
interface TakenChanges {
  readonly kind: string;
  readonly changes: ReviewFormattingChange[];
}

/**
 * What each suggestion `suggestion` took records from had set, by its id, in the order the
 * document first draws their records.
 */
function takenChanges(doc: OrderedDocument, suggestion: string): Map<string, TakenChanges> {
  const found = new Map<string, TakenChanges>();

  // Cheap first: the items with a record `suggestion` took from one still pending, before any
  // drawing. A record taken from a decided one takes over nothing, and one on an item that shows
  // nowhere (a direct deletion took it, or content took the place of its prompt) holds no change
  // record in the story, as cards read units (`isOrderedItemHidden`)
  const holders = new Map<OrderedItem, ReadonlySet<string>>();
  const trees = new Set<object>();
  for (const item of doc.labelled.get(suggestion) ?? []) {
    if (isOrderedItemHidden(doc, item)) {
      continue;
    }

    const taken = new Set(
      [...item.fmts, ...item.props].flatMap((r) => {
        const replaced = r.replaced ?? null;
        return r.suggestion === suggestion && isPending(doc, replaced) ? [replaced] : [];
      })
    );
    if (taken.size > 0 && item.leaf !== undefined) {
      holders.set(item, taken);
      trees.add(item.leaf.tree);
    }
  }

  // In document order, the same on every copy: a load labels items in another order than edits
  let scratch = 0;
  const mint = () => `replaced#${(scratch += 1)}`;
  for (const story of doc.stories.values()) {
    if (!trees.has(story.tree)) {
      continue;
    }

    const ctx = new ComposeContext(doc, story, {}, mint);
    for (const item of story.tree.items()) {
      const taken = holders.get(item);
      if (taken === undefined || addsNothing(item, taken, found)) {
        continue;
      }

      for (const { record, taken: drawn } of drawnTakeovers(ctx, item)) {
        for (const { replaced, before, after } of drawn) {
          if (!taken.has(replaced)) {
            continue;
          }

          let entry = found.get(replaced);
          if (entry === undefined) {
            entry = { kind: record, changes: [] };
            found.set(replaced, entry);
          }

          if (record === entry.kind) {
            addChanges(entry.changes, formattingChanges(record, before, after));
          }
        }
      }
    }
  }

  return found;
}

/**
 * Whether drawing an item can add nothing to `found`: each record of it taken from a suggestion
 * in `taken` can change only properties that suggestion's entry names already. A record the
 * review table cannot name is drawn.
 */
function addsNothing(
  item: OrderedItem,
  taken: ReadonlySet<string>,
  found: ReadonlyMap<string, TakenChanges>
): boolean {
  return [...item.fmts, ...item.props].every((record) => {
    if (record.replaced === undefined || !taken.has(record.replaced)) {
      return true;
    }

    // The property is enough: the answer keeps one entry per property, the first (`addChanges`)
    const named = found.get(record.replaced)?.changes;
    const properties = record.kind === 'properties' ? namedProperties(record) : null;
    return (
      named !== undefined &&
      properties !== null &&
      properties.every((property) => named.some((change) => change.property === property))
    );
  });
}

/**
 * The review properties a run format record can change, or null when the table cannot name all
 * it writes (a style, or a property the table does not list). It writes only the elements it
 * names (`faceWithProperties`), so only their properties can change.
 */
function namedProperties(
  record: Extract<OrderedFormatRecord, { kind: 'properties' }>
): ReviewFormattingChange['property'][] | null {
  const names = [...record.properties.map(({ localName }) => localName), ...(record.remove ?? [])];
  const rows = FORMATTING_PROPERTIES.filter(({ element }) => names.includes(element));
  if (record.style !== undefined || names.some((name) => !rows.some((r) => r.element === name))) {
    return null;
  }

  return rows.map(({ property }) => property);
}

/** The takeovers an item's records hold, drawn as compose draws them. */
function drawnTakeovers(ctx: ComposeContext, item: OrderedItem): Iterable<HeldTakeovers> {
  // The context keeps the takeovers of every record it drew: only this item's count
  ctx.takeovers.clear();
  const shell = item.node && item.node.kind !== 'textValue' ? item.node : null;
  faceAfter(ctx, item.fmts, item.face);
  if (item.kind === 'mark') {
    const pPr = shell === null ? null : (childWhere(shell, isParagraphProperties) ?? null);
    paragraphPropertiesAfter(ctx, item, pPr);
  }

  if (item.kind === 'open' && shell !== null) {
    containerPropertiesAfter(ctx, item, shell.children);
  }

  if (sectionRecordsOf(item).length > 0) {
    sectionAfter(ctx, item, sectionPropertiesOf(item));
  }

  return ctx.takeovers.values();
}

/**
 * A record's properties as an element of their own: `properties` under the element the change
 * record names (`w:rPr` for `w:rPrChange`). It is never placed in a part.
 */
const takenProperties = (record: string, properties: readonly OoxmlNode[]): OoxmlElement =>
  build('taken-properties', 'generic', record.slice(0, -'Change'.length), [], properties);

/** What properties a record gave over the ones before it, as the review queue names them. */
function formattingChanges(
  record: string,
  before: readonly OoxmlNode[],
  after: readonly OoxmlNode[]
): ReviewFormattingChange[] {
  return changedFormatting({
    node: build('replaced-record', 'generic', record, [], [takenProperties(record, before)]),
    parent: takenProperties(record, after),
    refused: false,
    paragraphMark: false,
    propertyChange: true,
    nesting: 0,
  });
}

/** Add the changes of a property `changes` does not name yet: its first record's value stays. */
function addChanges(
  changes: ReviewFormattingChange[],
  more: readonly ReviewFormattingChange[]
): void {
  for (const change of more) {
    if (!changes.some((other) => other.property === change.property)) {
      changes.push(change);
    }
  }
}
