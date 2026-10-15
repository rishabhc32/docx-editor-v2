// Proposed columns in the views that resolve tracked changes.
//
// Suggesting mode proposes deleting a column as `w:tcPr/w:cellDel` on every cell of it, and
// inserting one as `w:cellIns` (17.13.5.1 / 17.13.5.2). A tracked ROW already leaves the views
// that show the document as if every change were accepted or rejected; its cells did not, so
// No Markup still drew the column a reviewer had deleted, and Original the one they had added.
//
// The geometry is not re-derived here. Accepting a deleted cell gives its width to the cell
// before it (or the first one after it), compacts the grid, and repairs vertical merges, and a
// second copy of those rules would drift from the first. The view runs the store's own resolver
// over a detached copy of the table instead, so the table it lays out is the one Accept or
// Reject leaves. The canonical tree is never touched: the copy lives in a throwaway part, and
// journal capture is suspended while it resolves. The resolver keeps every subtree it does not
// change, so the copy's paragraphs are the authored nodes and their cached layout still hits.

import type { OoxmlElement, OoxmlPart } from '@docx-editor.dev/core/store';
import { runWithoutJournalCapture } from '../store/package/canonical-primitive-capture.ts';
import { isTableWrapper, tableChildren } from '../store/store/revision-table-children.ts';
import { isWmlElement } from '../store/store/tree-op-table-shared.ts';
import { collectRevisionSites, resolveRevisions } from '../store/store/tree-op-revisions.ts';
import { cacheProjection } from './bounded-projection-cache.ts';
import { wmlRevisionAttribute } from './revision-cell-shading.ts';
import {
  revisionFilterSelects,
  revisionNodeProjectionMode,
  type RevisionAuthorFilter,
  type RevisionDisplayMode,
} from './revision-projection.ts';

const projections = new WeakMap<OoxmlElement, Map<string, OoxmlElement>>();

/** The cell decisions a view takes: deletions it accepts and insertions it rejects. */
interface ResolvedCellMarkers {
  readonly accept: readonly string[];
  readonly reject: readonly string[];
}

/** Which of the table's own cell markers the view resolves; nested tables answer for theirs. */
function resolvedCellMarkers(
  table: OoxmlElement,
  mode: RevisionDisplayMode,
  filter: RevisionAuthorFilter | undefined
): ResolvedCellMarkers {
  const accept: string[] = [];
  const reject: string[] = [];
  for (const row of tableChildren(table, 'tableRow')) {
    for (const cell of tableChildren(row, 'tableCell')) {
      for (const properties of cell.children) {
        if (!isWmlElement(properties, 'tcPr')) continue;
        for (const marker of properties.children) {
          if (!isWmlElement(marker, 'cellDel') && !isWmlElement(marker, 'cellIns')) continue;
          const deletion = marker.localName === 'cellDel';
          const author = wmlRevisionAttribute(marker, 'author') ?? '';
          // The same per-decision answer a tracked row gets in `readTableStructure`.
          const projected = filter
            ? revisionNodeProjectionMode(filter, marker.id, author, mode)
            : mode;
          if (deletion && projected === 'proposed') accept.push(marker.id);
          else if (!deletion && projected === 'original') reject.push(marker.id);
        }
      }
    }
  }
  return { accept, reject };
}

/**
 * The table as resolving the given cell markers leaves it, or null when the resolver refuses.
 *
 * A marker the resolver cannot answer stays visible, as it does on the page in All Markup.
 */
function resolvedCopy(table: OoxmlElement, markers: ResolvedCellMarkers): OoxmlElement | null {
  const root: OoxmlElement = {
    kind: 'generic',
    id: `${table.id}/resolved-cells`,
    namespaceUri: table.namespaceUri,
    localName: 'resolvedCells',
    namespaceBindings: [],
    attributes: [],
    children: [table],
  };
  // The part name prefixes every id the resolver mints, so none can match a real node.
  let part: OoxmlPart = { id: root.id, name: root.id, contentType: '', root };
  const steps = [
    ['accept', markers.accept],
    ['reject', markers.reject],
  ] as const;
  return runWithoutJournalCapture(() => {
    for (const [action, ids] of steps) {
      const answerable = new Set(
        collectRevisionSites(part)
          .filter((site) => !site.refused)
          .map((site) => site.node.id)
      );
      const siteNodeIds = ids.filter((id) => answerable.has(id));
      if (siteNodeIds.length === 0) continue;
      const resolved = resolveRevisions(part, action, undefined, {
        siteNodeIds,
        deferValidation: true,
      });
      if (!resolved.ok || !resolved.part) return null;
      part = resolved.part;
    }
    const survivor = part.root.children.find((child) => child.kind === 'table');
    if (survivor && survivor.kind !== 'textValue') return survivor;
    // Every row went, so the resolver removed the table. Lay it out with no rows, as a table
    // whose every row is a hidden tracked deletion is laid out.
    return {
      ...table,
      children: table.children.filter(
        (child) => child.kind !== 'tableRow' && !isTableWrapper(child)
      ),
    } as OoxmlElement;
  });
}

/**
 * The table a resolved view lays out: the cells it shows as deleted (`proposed`) or never
 * inserted (`original`) removed, with their space where Accept or Reject puts it.
 *
 * Answers the table itself when the view resolves none of its cell markers, and always in All
 * Markup unless a reviewer filter changes which revisions it keeps. A filter's hidden author
 * resolves as `excludedNodeMode` says, exactly as a tracked row does.
 */
export function tableWithResolvedCells<T extends OoxmlElement>(
  table: T,
  mode: RevisionDisplayMode,
  filter?: RevisionAuthorFilter
): T {
  if (mode === 'all-markup' && !revisionFilterSelects(filter)) return table;
  const key = `${mode}|${filter?.cacheKey ?? ''}`;
  const known = projections.get(table)?.get(key);
  if (known) return known as T;
  const markers = resolvedCellMarkers(table, mode, filter);
  if (markers.accept.length === 0 && markers.reject.length === 0) return table;
  const projected = resolvedCopy(table, markers) ?? table;
  let entries = projections.get(table);
  if (!entries) projections.set(table, (entries = new Map()));
  cacheProjection(entries, key, projected);
  return projected as T;
}
