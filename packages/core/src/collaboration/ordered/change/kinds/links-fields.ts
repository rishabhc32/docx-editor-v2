// Change kinds of group 9.5 (links and fields: insert and remove links, page fields, bookmarks,
// field codes and results, a table of contents and its refresh). The group's own file: its
// change types and their specs, and nothing else edits it (design, "How to add a group").

import { insertLink, removeLink } from '../../edits/links.ts';
import { shownItemAt } from '../../edits/places.ts';
import { fieldPath, required } from '../read.ts';
import { changeStyle, styleId, type OrderedChangeStyle } from '../resources.ts';
import {
  type OrderedChangeAddress,
  type OrderedChangePlace,
  type OrderedEditBase,
} from '../types.ts';
import { address, linkTarget, optionalRange, place } from '../values.ts';
import { FIELD_KINDS, type OrderedFieldChange } from './fields.ts';
import { paragraphRange, runKindEdit } from './run.ts';
import { orderedKindSpec, type OrderedKindSpec } from './spec.ts';

/**
 * Wrap the units from `from` up to `to` (one paragraph) in a new link to an address (`href`) or
 * a bookmark (`anchor`), exactly one: the engine's `insertHyperlink`. The words take the
 * character style `style`; `newStyle` declares it when the document has none (Word's Hyperlink
 * style, as Word adds it on a first link). Suggesting, the units are the suggestion's own copy of
 * the words it struck (Word has no revision for a link around words).
 *
 * @public
 */
export interface OrderedInsertLinkChange extends OrderedEditBase {
  readonly kind: 'insert-link';
  readonly from: OrderedChangePlace;
  readonly to: OrderedChangePlace;
  readonly href?: string;
  readonly anchor?: string;
  readonly tooltip?: string;
  readonly style?: string;
  readonly newStyle?: OrderedChangeStyle;
}

/**
 * Take a link off its words, named by its open item, or off the part from `from` to `to`: the
 * engine's `removeHyperlink`. `style` is the character style the words lose with it.
 * Suggesting, the part is the suggestion's own copy at the link's end (the editor's proposal).
 *
 * @public
 */
export interface OrderedRemoveLinkChange extends OrderedEditBase {
  readonly kind: 'remove-link';
  readonly link: OrderedChangeAddress;
  readonly from?: OrderedChangePlace;
  readonly to?: OrderedChangePlace;
  readonly style?: string;
}

/**
 * The change kinds of group 9.5.
 *
 * @public
 */
export type OrderedLinkFieldChange =
  | OrderedInsertLinkChange
  | OrderedRemoveLinkChange
  | OrderedFieldChange;

const INSERT_LINK = orderedKindSpec<OrderedInsertLinkChange>({
  kind: 'insert-link',
  fields: ['from', 'to', 'href', 'anchor', 'tooltip', 'style', 'newStyle'],
  read: (f) => ({
    from: place(required(f, 'from'), fieldPath(f, 'from')),
    to: place(required(f, 'to'), fieldPath(f, 'to')),
    ...linkTarget(f),
    ...(f.has('style') ? { style: styleId(f.get('style'), fieldPath(f, 'style')) } : {}),
    ...(f.has('newStyle')
      ? { newStyle: changeStyle(f.get('newStyle'), fieldPath(f, 'newStyle')) }
      : {}),
  }),
  touches: (change, reader) => reader.range(change.from, change.to),
  apply: (change, input) =>
    runKindEdit(input, (ctx) => {
      const { p, start, end } = paragraphRange(ctx, change.from, change.to);
      const target =
        change.href !== undefined ? { url: change.href } : { anchor: change.anchor ?? '' };
      return insertLink(ctx, p, start, end, target, {
        ...(change.tooltip === undefined ? {} : { tooltip: change.tooltip }),
        ...(change.style === undefined ? {} : { style: change.style }),
        ...(change.newStyle === undefined ? {} : { newStyle: change.newStyle }),
      });
    }),
  // A link changes no text a reader sees: a party's untracked link is a package change
  policy: { text: false, properties: false },
  publish: 'paragraphs',
  caret: (change) => change.to,
});

const REMOVE_LINK = orderedKindSpec<OrderedRemoveLinkChange>({
  kind: 'remove-link',
  fields: ['link', 'from', 'to', 'style'],
  read: (f) => ({
    link: address(required(f, 'link'), fieldPath(f, 'link')),
    ...optionalRange(f),
    ...(f.has('style') ? { style: styleId(f.get('style'), fieldPath(f, 'style')) } : {}),
  }),
  touches: (change, reader) => {
    reader.item(change.link, 'open', ['hyperlink']);
    if (change.from !== undefined && change.to !== undefined) reader.range(change.from, change.to);
  },
  apply: (change, input) =>
    runKindEdit(input, (ctx) => {
      const open = shownItemAt(ctx.doc, ctx.story, change.link, 'open', ['hyperlink']);
      const part =
        change.from !== undefined && change.to !== undefined
          ? paragraphRange(ctx, change.from, change.to)
          : null;
      return removeLink(ctx, open, part, change.style);
    }),
  policy: { text: false, properties: false },
  publish: 'paragraphs',
});

/** The group's specs: one for each kind of {@link OrderedLinkFieldChange}. */
export const LINK_FIELD_KINDS: { readonly [K in OrderedLinkFieldChange['kind']]: OrderedKindSpec } =
  {
    'insert-link': INSERT_LINK,
    'remove-link': REMOVE_LINK,
    ...FIELD_KINDS,
  };
