// The `section-property` change kind (task 9.8): page size, orientation and margins
// (`setSectionProperties`), and a different first page and the header and footer distances
// (`setSectionFurnitureOptions`; its odd and even pages is a `setting` change, group 9.10).
//
// Page setup goes to the sections the document has when the change applies: every one (Word's
// "Whole document"), or the one that governs a paragraph then ("This section"), so a section
// break another person made first takes it too, as it would in Word one after the other. The
// furniture options go to the sections the change names. Both are a suggestion when suggested:
// Word records them as `w:sectPrChange` (ECMA-376 Part 1, §17.13.5.32).

import {
  othersFurnitureRecord,
  othersSectionRecord,
  setFurnitureOptions,
  setPageSetup,
} from '../../../edits/sections.ts';
import { suggestionOrNull } from '../../../model/document.ts';
import type { OrderedSectionOp } from '../../../model/sections.ts';
import { address } from '../../values.ts';
import { runKindEdit } from '../run.ts';
import { at, Invalid, list, type Fields } from '../../read.ts';
import {
  ORDERED_CHANGE_LIMITS,
  type OrderedChangeAddress,
  type OrderedEditBase,
} from '../../types.ts';
import { orderedKindSpec } from '../spec.ts';
import { readFurniture, readPage, touchSection } from './read.ts';

/** Page size, orientation and margins, in twips, as `setSectionProperties` takes them. @public */
export interface OrderedChangePageSetup {
  readonly pageWidthTwips?: number;
  readonly pageHeightTwips?: number;
  readonly orientation?: 'portrait' | 'landscape';
  readonly marginTopTwips?: number;
  readonly marginRightTwips?: number;
  readonly marginBottomTwips?: number;
  readonly marginLeftTwips?: number;
}

/** A different first page and the header and footer distances, in twips. @public */
export interface OrderedChangeFurnitureOptions {
  readonly titlePage?: boolean;
  readonly headerDistanceTwips?: number;
  readonly footerDistanceTwips?: number;
}

/**
 * Change section properties. `page` goes to every section the document has when the change
 * applies, or with `pageAt` (a paragraph mark) to the section that governs that paragraph then.
 * `furniture` goes to each of `sections`, named by the item that holds its properties: a
 * paragraph mark, or the body's own section item.
 *
 * @public
 */
export interface OrderedSectionPropertyChange extends OrderedEditBase {
  readonly kind: 'section-property';
  readonly sections: readonly OrderedChangeAddress[];
  readonly page?: OrderedChangePageSetup;
  readonly pageAt?: OrderedChangeAddress;
  readonly furniture?: OrderedChangeFurnitureOptions;
}

function read(fields: Fields): Record<string, unknown> {
  const path = fields.path;

  // No field means no sections; a field of the wrong type, `null` too, is refused
  const given = fields.has('sections') ? fields.get('sections') : [];
  const sections = list(given, at(path, 'sections'), 0, ORDERED_CHANGE_LIMITS.targets).map(
    (value, k) => address(value, at(at(path, 'sections'), k))
  );
  const out: Record<string, unknown> = { sections };
  if (fields.has('page')) out.page = readPage(fields.get('page'), at(path, 'page'));
  if (fields.has('pageAt')) {
    if (out.page === undefined) throw new Invalid('unknown-field', at(path, 'pageAt'));
    out.pageAt = address(fields.get('pageAt'), at(path, 'pageAt'));
  }

  if (fields.has('furniture'))
    out.furniture = readFurniture(fields.get('furniture'), at(path, 'furniture'));
  if (out.page === undefined && out.furniture === undefined)
    throw new Invalid('empty', path, 'no section property');

  // The named sections are the furniture options' alone
  if ((out.furniture !== undefined) !== sections.length > 0)
    throw new Invalid(sections.length > 0 ? 'unknown-field' : 'empty', at(path, 'sections'));
  return out;
}

export const SECTION_PROPERTY_SPEC = orderedKindSpec<OrderedSectionPropertyChange>({
  kind: 'section-property',
  fields: ['sections', 'page', 'pageAt', 'furniture'],
  read,
  touches: (change, reader) => {
    if (change.pageAt !== undefined) reader.item(change.pageAt, 'mark');
    for (const holder of change.sections) touchSection(reader, holder);
  },
  apply: (change, input) =>
    runKindEdit(
      input,
      (ctx) => {
        let pageSuggestion: string | null = null;
        if (change.page !== undefined) {
          const op: Extract<OrderedSectionOp, { op: 'page' }> = { op: 'page', ...change.page };
          pageSuggestion = setPageSetup(ctx, change.pageAt ?? null, op);
        }

        if (change.furniture === undefined) {
          return pageSuggestion;
        }

        // One change is one suggestion: the furniture options join the page setup's
        return setFurnitureOptions(
          ctx,
          change.sections,
          { op: 'furniture', ...change.furniture },
          suggestionOrNull(ctx.doc, pageSuggestion)
        );
      },
      { settleDirect: true }
    ),
  policy: {
    text: false,
    properties: true,
    // A section holds one `w:sectPrChange`, so a party's write on a section that holds another
    // person's pending record is refused, as the host's check refuses it: that person's change
    // would lose its author (`others_suggestions`)
    party(change, doc, maker) {
      if (change.kind !== 'section-property') {
        return null;
      }

      const story = doc.stories.get(change.part);
      if (story === undefined) {
        return null;
      }

      let pageRecord: string | null = null;
      if (change.page !== undefined) {
        pageRecord = othersSectionRecord(doc, story, change.pageAt ?? null, maker);
      }

      const other = pageRecord ?? othersFurnitureRecord(doc, story, change.sections, maker);
      if (other === null) {
        return null;
      }

      return {
        ok: false,
        code: 'policy-not-own',
        reason: 'others_suggestions',
        detail: `${other} holds the section's change record`,
      };
    },
  },
  publish: 'story',
});
