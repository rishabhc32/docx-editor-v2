// Sections in the ordered model. A section's properties (`w:sectPr`) end the section: a
// paragraph mark holds one in its `w:pPr`, and the body holds the last section's as a block item.

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { zeroSpace, type OrderedNotePropertiesOp } from './types.ts';

/** Which story a section's reference names. @public */
export type OrderedFurnitureKind = 'header' | 'footer';

/** Which pages a header or footer is for (`ST_HdrFtr`, §17.18.36). @public */
export type OrderedFurnitureVariant = 'default' | 'first' | 'even';

/**
 * One write to one section, as the engine's op makes it on that section alone.
 *
 * - `page`: page size, orientation and margins (`setSectionProperties`).
 * - `furniture`: a different first page (`w:titlePg`) and the header and footer distances
 *   (`setSectionFurnitureOptions`).
 * - `reference`: the section's header or footer of a variant names the relationship `id`, or
 *   none when null (create, delete, link and unlink).
 * - `type`: how the section starts (`w:type`, written on the section a break starts).
 * - `mark`: the section properties a paragraph mark takes when a section break ends a section
 *   there (`setSectionMark`): a copy of the section it was cut from.
 * - `setNoteProperties`: the section's footnote and endnote properties (`w:footnotePr`,
 *   `w:endnotePr`), a direct edit only.
 *
 * @public
 */
export type OrderedSectionOp =
  | {
      readonly op: 'page';
      readonly pageWidthTwips?: number;
      readonly pageHeightTwips?: number;
      readonly orientation?: 'portrait' | 'landscape';
      readonly marginTopTwips?: number;
      readonly marginRightTwips?: number;
      readonly marginBottomTwips?: number;
      readonly marginLeftTwips?: number;
    }
  | {
      readonly op: 'furniture';
      readonly titlePage?: boolean;
      readonly headerDistanceTwips?: number;
      readonly footerDistanceTwips?: number;
    }
  | {
      readonly op: 'reference';
      readonly kind: OrderedFurnitureKind;
      readonly variant: OrderedFurnitureVariant;
      readonly id: string | null;
    }
  | { readonly op: 'type'; readonly type: 'nextPage' | 'continuous' }
  | { readonly op: 'mark'; readonly sectPr: OoxmlElement }
  | OrderedNotePropertiesOp;

/**
 * A write to a section, as a record on the item that holds it (`props`, target `sectPr`): the
 * engine's own write for that one section, which compose runs again in the order changes made
 * them (accepted first, then pending). So a take-back or a reject takes the write away exactly,
 * as a paragraph press record does (design D6). A pending record that is not `untracked` draws
 * Word's `w:sectPrChange` (the section's properties before it).
 *
 * @public
 */
export interface OrderedSectionRecord {
  readonly kind: 'sectionOp';
  readonly suggestion: string;
  readonly seq: number;
  readonly target: 'sectPr';
  readonly op: OrderedSectionOp;
  /** Shown while pending with no change record: a section break on its author's own new mark. */
  readonly untracked?: boolean;
  /**
   * Applied when its suggestion is rejected too: Word writes it untracked, so a reject leaves it
   * (the type a suggested section break gives the section after it). A take-back removes it.
   */
  readonly keptOnReject?: boolean;
  /** The suggestion a later change of the same section took this record from (policy 2). */
  readonly replaced?: string;
  /** The suggestions it was taken from before `replaced`, oldest first. */
  readonly replacedEarlier?: readonly string[];
  /**
   * A direct write made over pending records of the section: their suggestions. Word keeps one
   * record, the section before the first pending change, so a reject of one of them takes this
   * write away too, as the engine's reject does.
   */
  readonly over?: readonly string[];
}

/**
 * The node id of the body's own section in a document that wrote none. Such a document still has
 * one (ECMA-376 Part 1, §17.6.17: Word's defaults when `w:body` has no `w:sectPr`). Load gives it
 * an item with no node, at a fixed address, so every copy names it alike; compose writes it only
 * once a record gives it properties.
 */
export const IMPLICIT_SECTION = 'body-section';

/** The space of the item that holds the body's section when the file wrote none. */
export const implicitSectionSpace = (file: number): string => zeroSpace(file, IMPLICIT_SECTION);
