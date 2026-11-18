// What a session's saved file must give back. The view is the same, always. Decisions are the
// same for every suggestion but those a documented limit of Word's form covers (the matrix
// lists them in `known.ts`, codes with `L`):
//
//  - Word keeps one change record for each cell's or table's properties, so every other
//    pending change of the same properties shares that record; a column a decision removed
//    gives its width to its neighbour's live properties only, so a pending change of the
//    neighbour's properties (or of the grid) gives back the width from before the removal;
//  - Word keeps one paragraph (or table, row, cell) change record, so a direct change of those
//    properties made after a pending one goes with that one's rejection; and one run change
//    record (`w:rPrChange`, the face before the first pending format), so a direct format made
//    after a pending one goes with that one's rejection too;
//  - a paragraph whose mark went for good right before a table runs into the table's first
//    cell, as the engine's own reject merges it there; the saved file holds its words in that
//    cell, so a decision on the first row then takes them with it, where the live list keeps
//    them (the engine's result depends on which came first);
//  - a control's prompt that holds a discussion marker: while text is typed or pasted into the
//    control, the file writes no `w:showingPlcHdr`, and once that text goes the control still
//    holds the marker, so a reopened file shows the control empty where the live list shows
//    the prompt around the marker again.
//
// A reopened file's suggestions are new ones (one for each revision triple): each is matched to
// the live suggestion that wrote its `w:id`.

import { REVISION_ROLE, revisionIdFor } from '../../model/revision-ids.ts';
import { isHidden } from '../../model/document.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from '../../model/types.ts';

const FIXED_ROLES = Object.values(REVISION_ROLE).filter((role) => typeof role === 'string');

/** Live pending suggestions a documented limit of Word's form covers. */
export function wordFormLimited(doc: OrderedDocument): Set<string> {
  const pending = (id: string) => doc.suggestions.get(id)?.status === 'pending';
  const accepted = (id: string) => doc.suggestions.get(id)?.status === 'accepted';
  const out = new Set<string>();
  for (const story of doc.stories.values()) {
    for (const id of markedPromptInsertions(doc, story)) out.add(id);
    for (const item of story.tree.items()) {
      // Two pending changes of one cell's or table's properties share Word's one record
      const owners = new Set(
        item.props
          .filter((r) => pending(r.suggestion) && (r.target === 'tblGrid' || r.target === 'tcPr'))
          .map((r) => r.suggestion)
      );
      for (const record of item.props) {
        if (!pending(record.suggestion)) continue;
        if (owners.size > 1) out.add(record.suggestion);
        // A pending paragraph, table, row or cell change with a direct one after it on the
        // same properties: Word's one record holds the properties before the pending one
        for (const later of item.props)
          if (
            later.target === record.target &&
            later.seq > record.seq &&
            accepted(later.suggestion)
          )
            out.add(record.suggestion);
      }
      // A pending format with a direct one after it on the same unit
      for (const record of item.fmts)
        if (
          pending(record.suggestion) &&
          item.fmts.some((later) => later.seq > record.seq && accepted(later.suggestion))
        )
          out.add(record.suggestion);
    }
  }
  for (const story of doc.stories.values()) {
    carriedIntoTable(doc, story, out);
    besideRemovedColumns(doc, story, out);
  }
  return out;
}

/** Pending insertions inside a control whose prompt holds a marker (a discussion's range). */
function markedPromptInsertions(doc: OrderedDocument, story: OrderedStory): string[] {
  const out: string[] = [];
  const items = [...story.tree.items()];
  items.forEach((open, k) => {
    if (open.kind !== 'open' || open.prompt !== true) return;
    let marker = false;
    const inserted = new Set<string>();
    for (let j = k + 1; j < items.length; j += 1) {
      const item = items[j];
      if (!item || (item.kind === 'close' && item.pair === open.space)) break;
      if (item.kind === 'zero') marker = true;
      if (
        (item.kind === 'text' || item.kind === 'atom') &&
        item.prompt !== true &&
        item.by !== null
      )
        if (doc.suggestions.get(item.by)?.status === 'pending') inserted.add(item.by);
    }
    if (marker) out.push(...inserted);
  });
  return out;
}

/** Pending property changes in a row, or of a grid, beside cells a column decision removed. */
function besideRemovedColumns(doc: OrderedDocument, story: OrderedStory, out: Set<string>): void {
  const status = (id: string) => doc.suggestions.get(id)?.status;
  // A column a decision removed, or one a pending suggestion may remove
  const column = (cell: OrderedItem): string[] =>
    cell.within.dels.length > 0 || cell.within.by.length > 0
      ? []
      : [
          ...(cell.by !== null && status(cell.by) !== 'accepted' ? [cell.by] : []),
          ...cell.dels.filter((id) => status(id) !== 'rejected'),
        ];
  const tables: OrderedItem[] = [];
  let cells: OrderedItem[] = [];
  let depth = 0;
  for (const item of story.tree.items()) {
    if (item.kind !== 'open' && item.kind !== 'close') continue;
    if (item.token === 'tbl') {
      if (item.kind === 'open') tables.push(item);
      else tables.pop();
      continue;
    }
    if (item.token === 'tr' && item.kind === 'open') {
      depth = tables.length;
      cells = [];
    } else if (item.token === 'tc' && item.kind === 'open' && tables.length === depth)
      cells.push(item);
    else if (item.token === 'tr' && item.kind === 'close' && tables.length === depth) {
      const columns = cells.flatMap(column);
      if (isHidden(doc, item) || columns.length === 0) continue;
      const table = tables[tables.length - 1];
      const pending = [...cells, ...(table ? [table] : [])].flatMap((owner) =>
        owner.props.filter((r) => status(r.suggestion) === 'pending').map((r) => r.suggestion)
      );
      if (pending.length === 0) continue;
      for (const id of [...pending, ...columns]) if (status(id) === 'pending') out.add(id);
    }
  }
}

/**
 * The pending suggestions on the first row of a table that words run into: a shown unit, then
 * a paragraph mark that went for good, then the table.
 */
function carriedIntoTable(doc: OrderedDocument, story: OrderedStory, out: Set<string>): void {
  let carried = false;
  let row: OrderedItem | null = null;
  for (const item of story.tree.items()) {
    if (row !== null) {
      // Every pending suggestion that labels the first row
      if (item.kind === 'close' && item.space === row.space) {
        row = null;
        continue;
      }
      for (const id of [item.by, ...item.dels, ...item.within.by, ...item.within.dels])
        if (id !== null && doc.suggestions.get(id)?.status === 'pending') out.add(id);
      continue;
    }
    // A mark that went leaves its words carried on to the next mark
    if (isHidden(doc, item)) continue;
    if (item.kind === 'open' && item.token === 'tbl') continue;
    if (item.kind === 'open' && item.token === 'tr' && carried) {
      row = item;
      for (const id of [item.by, ...item.dels])
        if (id !== null && doc.suggestions.get(id)?.status === 'pending') out.add(id);
      carried = false;
      continue;
    }
    if (item.kind === 'text' || item.kind === 'atom') carried = true;
    else if (item.kind === 'mark') carried = false;
  }
}

/** Each reopened suggestion's live suggestion: the one whose `w:id` it was read from. */
export function liveOf(live: OrderedDocument, again: OrderedDocument): Map<string, string> {
  const roles: string[] = [...FIXED_ROLES];
  for (const story of live.stories.values())
    for (const item of story.tree.items())
      if (item.kind === 'mark')
        roles.push(REVISION_ROLE.deleteIn(item.space), REVISION_ROLE.moveFromIn(item.space));
  const writers = new Map<string, string[]>();
  for (const suggestion of live.suggestions.values()) {
    if (suggestion.status !== 'pending') continue;
    for (const role of roles) {
      const id = revisionIdFor(live, suggestion.id, role);
      writers.set(id, [...(writers.get(id) ?? []), suggestion.id]);
    }
  }
  const out = new Map<string, string>();
  for (const [key, id] of again.revisionIds.byRole) {
    const reopened = key.slice(0, key.indexOf('/'));
    if (out.has(reopened)) continue;
    const attribution = again.suggestions.get(reopened)?.attribution;
    const match = (writers.get(id) ?? []).find((candidate) => {
      const theirs = live.suggestions.get(candidate)?.attribution;
      return theirs?.author === attribution?.author && theirs?.date === attribution?.date;
    });
    if (match !== undefined) out.set(reopened, match);
  }
  return out;
}
