// Change kinds of group 9.7 (tables: insert and delete columns, column widths, the right edge,
// table design). The group's own file: its change types and their specs, and nothing else edits
// it. Empty until the group builds them (design, "How to add a group").

import type { OrderedKindSpec } from './spec.ts';

/**
 * The change kinds of group 9.7: a union of the group's change interfaces, each extending
 * `OrderedEditBase` with its own `kind`. `never` until the group adds one.
 *
 * @public
 */
export type OrderedTableChange = never;

/**
 * The group's specs: one for each kind of {@link OrderedTableChange}, made with
 * `orderedKindSpec`.
 */
export const TABLE_KINDS: { readonly [K in OrderedTableChange['kind']]: OrderedKindSpec } = {};
