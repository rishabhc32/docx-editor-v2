// The full types of the relationships a change adds, as a package names them.

import {
  FRAGMENT_OBJECT_RELATIONSHIP_TYPES,
  FRAGMENT_PART_RELATIONSHIP_TYPES,
} from '../change/fragment-names-objects.ts';
import type { OrderedRelationshipResource } from '../model/resources.ts';

export const OFFICE_RELATIONSHIPS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The full type of each kind of relationship a change adds. */
export const RELATIONSHIP_TYPES: Readonly<Record<OrderedRelationshipResource['type'], string>> = {
  image: `${OFFICE_RELATIONSHIPS}/image`,
  hyperlink: `${OFFICE_RELATIONSHIPS}/hyperlink`,
  header: `${OFFICE_RELATIONSHIPS}/header`,
  footer: `${OFFICE_RELATIONSHIPS}/footer`,
  footnotes: `${OFFICE_RELATIONSHIPS}/footnotes`,
  endnotes: `${OFFICE_RELATIONSHIPS}/endnotes`,
  oleObject: FRAGMENT_OBJECT_RELATIONSHIP_TYPES.oleObject,
  ink: FRAGMENT_OBJECT_RELATIONSHIP_TYPES.ink,
  ...FRAGMENT_PART_RELATIONSHIP_TYPES,
};
