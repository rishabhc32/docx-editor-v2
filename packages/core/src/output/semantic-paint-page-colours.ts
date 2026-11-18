// The colours a kept page shows its authors in.

import type { SemanticLayout } from '../layout/index.ts';
import {
  pageAuthorsOf,
  revisionAuthorStyleKey,
  revisionStyleContextKey,
  type RevisionStyleContext,
} from './revision-presentation.ts';

type PageRecord = SemanticLayout['pages'][number];

/** The colour contexts one paint used: tracked text, and table cells. */
export interface PaintColours {
  readonly text: RevisionStyleContext | null;
  readonly cells: RevisionStyleContext | null;
}

const scheme = (context: RevisionStyleContext | null): string =>
  context === null ? 'kind' : `author:${context.others}`;

/**
 * The part of the colours every page shares, for the paint parameters. It holds only the
 * scheme: the colour context holds the whole document's author map, which one author coming
 * or going changes, so each page is checked against its own authors (`pageColoursKept`).
 */
export function paintColoursKey(colours: PaintColours): string {
  return `rev:${scheme(colours.text)}|cells:${scheme(colours.cells)}`;
}

/**
 * Whether what a sheet painted for `previous` with `before` keeps for `next` shows its authors
 * as `after` paints them: every author both records name keeps its slot and style. A block or
 * a layer kept from the old sheet is in both records, so its authors are among them; anything
 * that names an author only one record names is painted again. All pages keep their colours
 * when the contexts agree for everyone; none does with no earlier paint.
 */
export function pageColoursKept(
  before: PaintColours | undefined,
  after: PaintColours
): (previous: PageRecord, next: PageRecord) => boolean {
  if (!before) return () => false;
  const same = (left: RevisionStyleContext | null, right: RevisionStyleContext | null) =>
    revisionStyleContextKey(left) === revisionStyleContextKey(right);
  if (same(before.text, after.text) && same(before.cells, after.cells)) return () => true;
  const sameFor =
    (left: RevisionStyleContext | null, right: RevisionStyleContext | null) => (name: string) =>
      left === right || revisionAuthorStyleKey(left, name) === revisionAuthorStyleKey(right, name);
  const text = sameFor(before.text, after.text);
  const cells = sameFor(before.cells, after.cells);
  return (previous, next) => {
    const shared = previous === next ? null : new Set(pageAuthorsOf(previous));
    return pageAuthorsOf(next).every(
      (name) => (shared !== null && !shared.has(name)) || (text(name) && cells(name))
    );
  };
}
