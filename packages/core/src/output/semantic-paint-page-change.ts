// How a page record differs from the one its sheet was painted from, when the sheet can stay.
//
// Paint keeps a sheet and repaints the parts that moved: the body blocks, and the header and
// footer layer. Anything else that differs rebuilds the sheet.

import type { PageRecord } from '../layout/semantic-records.ts';
import { sameDataWithin } from './semantic-paint-record-equality.ts';

/** Objects one page-field comparison may visit before adoption gives up on it. */
const PAGE_RECORD_COMPARE_BUDGET = 20_000;

/** Every own key equal by identity. */
function samePlainValues(left: object, right: object): boolean {
  const keys = Object.keys(left);
  if (keys.length !== Object.keys(right).length) return false;
  for (const key of keys) {
    const a = (left as Record<string, unknown>)[key];
    const b = (right as Record<string, unknown>)[key];
    if (a !== b) return false;
  }

  return true;
}

function sameItems(left: readonly unknown[], right: readonly unknown[]): boolean {
  if (left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) return false;
  }

  return true;
}

function samePageValue(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (Array.isArray(a) && Array.isArray(b) && sameItems(a, b)) return true;
  if (
    typeof a === 'object' &&
    typeof b === 'object' &&
    a !== null &&
    b !== null &&
    !Array.isArray(a) &&
    !Array.isArray(b) &&
    samePlainValues(a, b)
  ) {
    return true;
  }

  // A record layout rebuilt with the same data paints the same: a footer re-projected for an
  // unchanged page number, after a table edit replaced the page. Bounded, so a large or
  // changed story refuses and repaints.
  return sameDataWithin(a, b, PAGE_RECORD_COMPARE_BUDGET);
}

/** What a kept sheet must repaint besides its changed body blocks. */
export interface AdoptablePageChange {
  /** The header or footer changed: the sheet's header and footer layer is painted again. */
  readonly furniture: boolean;
}

/**
 * The change from `previous` to `next` when it is only in the body blocks or the header and
 * footer stories, else `null`.
 *
 * The test is written against the record's own keys rather than a hand-listed set, so a
 * field added to `PageRecord` later disables adoption (a slower repaint) instead of silently
 * painting the page without it. Everything else must be the same object, the same flat values
 * (a box), or the same array members: layout hands untouched sub-records back by identity, so
 * anything that really changed fails this and takes the full path.
 *
 * A header or footer edit gives every page that shows the story a new record. Only a story
 * present on both sides counts as a furniture change: a page that gains or loses one also
 * changes its blank-band invitation, so it is painted again in full.
 */
export function adoptablePageChange(
  previous: PageRecord,
  next: PageRecord
): AdoptablePageChange | null {
  if (previous === next) return null;
  let furniture = false;
  const keys = new Set([...Object.keys(previous), ...Object.keys(next)]);
  for (const key of keys) {
    // Paint never reads the page-field source: the header, footer and body field records that
    // show a page number carry its effect, and they are compared like everything else.
    if (key === 'fragments' || key === 'pageFieldSource') continue;
    const a = (previous as unknown as Record<string, unknown>)[key];
    const b = (next as unknown as Record<string, unknown>)[key];
    if (samePageValue(a, b)) continue;
    if ((key === 'header' || key === 'footer') && a !== undefined && b !== undefined) {
      furniture = true;
      continue;
    }

    return null;
  }

  return { furniture };
}
