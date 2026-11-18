// Change kinds of group 9.8 (headers, footers and sections: create, delete, link and unlink,
// page furniture options, section properties and breaks, watermark). The group's own file: its
// change types and their specs, and nothing else edits it. Each spec lives in `sections/`.

import { SECTION_BREAK_SPEC, type OrderedSectionBreakChange } from './sections/break.ts';
import { HEADER_FOOTER_SPEC, type OrderedHeaderFooterChange } from './sections/furniture.ts';
import { SECTION_PROPERTY_SPEC, type OrderedSectionPropertyChange } from './sections/property.ts';
import { WATERMARK_SPEC, type OrderedWatermarkChange } from './sections/watermark.ts';
import type { OrderedKindSpec } from './spec.ts';

export type { OrderedSectionBreakChange } from './sections/break.ts';
export type { OrderedHeaderFooterChange } from './sections/furniture.ts';
export type {
  OrderedChangeFurnitureOptions,
  OrderedChangePageSetup,
  OrderedSectionPropertyChange,
} from './sections/property.ts';
export type {
  OrderedChangePictureWatermark,
  OrderedChangeTextWatermark,
  OrderedChangeWatermark,
  OrderedChangeWatermarkBox,
  OrderedChangeWatermarkSize,
  OrderedWatermarkChange,
} from './sections/watermark.ts';

/**
 * The change kinds of group 9.8: a union of the group's change interfaces, each extending
 * `OrderedEditBase` with its own `kind`.
 *
 * @public
 */
export type OrderedSectionChange =
  | OrderedSectionPropertyChange
  | OrderedHeaderFooterChange
  | OrderedSectionBreakChange
  | OrderedWatermarkChange;

/**
 * The group's specs: one for each kind of {@link OrderedSectionChange}, made with
 * `orderedKindSpec`.
 */
export const SECTION_KINDS: { readonly [K in OrderedSectionChange['kind']]: OrderedKindSpec } = {
  'section-property': SECTION_PROPERTY_SPEC,
  'header-footer': HEADER_FOOTER_SPEC,
  'section-break': SECTION_BREAK_SPEC,
  watermark: WATERMARK_SPEC,
};
