// Change kinds of group 9.5 for fields, bookmarks and a table of contents (`links-fields.ts`
// joins them into the group's table). Text that depends on layout comes in the change as text,
// from the sender's layout, so every copy writes the same text.

import { bookmarkNameProblem } from '../../../../store/package/bookmark-names.ts';
import {
  parseTocInstruction,
  TOC_MAX_ENTRIES,
  TOC_MAX_INSTRUCTION_CHARS,
} from '../../../../store/package/toc-instruction.ts';
import {
  MAX_FIELD_RESULT_TEXT_CHARS,
  MAX_FIELD_RESULT_UPDATES,
} from '../../../../store/store/tree-op-field-results.ts';
import type { Ctx } from '../../edits/context.ts';
import {
  editBookmark,
  insertPageField,
  setFieldCode,
  setFieldResult,
  setRunText,
  type OrderedPageField,
} from '../../edits/fields.ts';
import { resolveOrderedPlace, shownItemAt, type OrderedResolvedPlace } from '../../edits/places.ts';
import {
  insertToc,
  replaceTocResult,
  type OrderedTocBookmark,
  type OrderedTocContent,
  type OrderedTocRow,
} from '../../edits/toc.ts';
import { OrderedRefusal } from '../../model/refusal.ts';
import type { OrderedItem } from '../../model/types.ts';
import {
  at,
  bool,
  fieldPath,
  type Fields,
  fieldsOf,
  int,
  Invalid,
  isPlainText,
  list,
  oneOf,
  plainText,
  required,
  str,
} from '../read.ts';
import {
  ORDERED_CHANGE_LIMITS as LIMITS,
  type OrderedChangeAddress,
  type OrderedChangePlace,
  type OrderedEditBase,
} from '../types.ts';
import { resolveChangePlace } from '../change-place.ts';
import { address, place } from '../values.ts';
import { orderedKindSpec, type OrderedKindSpec, type OrderedTargetReader } from './spec.ts';
import { paragraphRange, runKindEdit } from './run.ts';

/**
 * Type a page field at a place: `PAGE`, `NUMPAGES`, `SECTIONPAGES`, or `PAGE_X_OF_Y` (both
 * fields with " of " between), as the engine's `insertPageField` builds it.
 *
 * @public
 */
export interface OrderedInsertFieldChange extends OrderedEditBase {
  readonly kind: 'insert-field';
  readonly place: OrderedChangePlace;
  readonly field: OrderedPageField;
}

/**
 * A whole field's new instruction, `PAGE` or `NUMPAGES` (the engine's `setFieldCode`). The field
 * is named by its item; its result is kept.
 *
 * @public
 */
export interface OrderedFieldCodeChange extends OrderedEditBase {
  readonly kind: 'field-code';
  readonly field: OrderedChangeAddress;
  readonly code: 'PAGE' | 'NUMPAGES';
}

/**
 * A whole field's new result text: the engine's `refreshFieldResults` (a tab is `\t`, a line
 * break `\n`), or with `last` only the field's last text element (a contents row's page number
 * in a `PAGEREF` field, as `rewriteTocPageNumbers` writes it).
 *
 * @public
 */
export interface OrderedFieldResultChange extends OrderedEditBase {
  readonly kind: 'field-result';
  readonly fields: readonly {
    readonly field: OrderedChangeAddress;
    readonly text: string;
    readonly last?: true;
  }[];
}

/**
 * New text for contents rows' page numbers that are plain text: each `from` to `to` is the text
 * of one text element, which keeps its run and face (`rewriteTocPageNumbers`).
 *
 * @public
 */
export interface OrderedTocPagesChange extends OrderedEditBase {
  readonly kind: 'toc-pages';
  readonly pages: readonly {
    readonly from: OrderedChangePlace;
    readonly to: OrderedChangePlace;
    readonly text: string;
  }[];
}

/**
 * Add a bookmark `name` from `from` to `to` (equal places: a point), and remove the markers
 * `remove` names: the same name elsewhere, which a new bookmark moves, or a bookmark deleted
 * (no `from` and `to`). A direct edit only: Word has no revision for a bookmark.
 *
 * @public
 */
export interface OrderedBookmarkChange extends OrderedEditBase {
  readonly kind: 'bookmark';
  readonly name: string;
  readonly from?: OrderedChangePlace;
  readonly to?: OrderedChangePlace;
  readonly remove: readonly OrderedChangeAddress[];
}

/**
 * A table of contents in a block control before a body paragraph (named by its mark), with its
 * rows and heading bookmarks, as the engine's `insertToc` builds it. A direct edit only.
 *
 * @public
 */
export interface OrderedInsertTocChange extends OrderedEditBase {
  readonly kind: 'insert-toc';
  readonly before: OrderedChangeAddress;
  readonly alias: string;
  readonly instruction: string;
  readonly rows: readonly OrderedTocRow[];
  readonly bookmarks: readonly OrderedTocBookmark[];
}

/**
 * A table of contents' result made again: everything between the field's separator and its end
 * goes, and the rows go in (the engine's `replaceTocResult`). A direct edit only.
 *
 * @public
 */
export interface OrderedTocResultChange extends OrderedEditBase {
  readonly kind: 'toc-result';
  readonly separate: OrderedChangeAddress;
  readonly end: OrderedChangeAddress;
  readonly instruction: string;
  readonly rows: readonly OrderedTocRow[];
  readonly bookmarks: readonly OrderedTocBookmark[];
}

/**
 * The field, bookmark and table of contents kinds of group 9.5.
 *
 * @public
 */
export type OrderedFieldChange =
  | OrderedInsertFieldChange
  | OrderedFieldCodeChange
  | OrderedFieldResultChange
  | OrderedTocPagesChange
  | OrderedBookmarkChange
  | OrderedInsertTocChange
  | OrderedTocResultChange;

/** Text a field result writes: a tab and a line break are its only control characters. */
function resultText(value: unknown, path: string, max: number): string {
  if (typeof value !== 'string') throw new Invalid('invalid-type', path, 'not a string');
  if (value.length > max) throw new Invalid('too-long', path, `over ${max}`);
  if (!isPlainText(value.replace(/[\t\n]/g, ' '))) throw new Invalid('invalid-text', path);
  return value;
}

/** An item of a list, read as an object with these fields. */
const entries = <T>(
  f: Fields,
  key: string,
  min: number,
  max: number,
  known: readonly string[],
  read: (e: Fields) => T
): T[] =>
  list(required(f, key), fieldPath(f, key), min, max).map((value, k) =>
    read(fieldsOf(value, at(fieldPath(f, key), k), new Set(known)))
  );

const BOOKMARK_NAME_MAX = 40;

const rowsOf = (f: Fields): OrderedTocRow[] =>
  entries(f, 'rows', 0, TOC_MAX_ENTRIES, ['level', 'text', 'heading', 'bookmark', 'page'], (e) => ({
    level: int(required(e, 'level'), fieldPath(e, 'level'), 0, 8),
    text: resultText(required(e, 'text'), fieldPath(e, 'text'), 200),
    heading: address(required(e, 'heading'), fieldPath(e, 'heading')),
    bookmark: plainText(required(e, 'bookmark'), fieldPath(e, 'bookmark'), 1, BOOKMARK_NAME_MAX),
    page: resultText(required(e, 'page'), fieldPath(e, 'page'), 32),
  }));

const bookmarksOf = (f: Fields): OrderedTocBookmark[] =>
  entries(f, 'bookmarks', 0, TOC_MAX_ENTRIES, ['heading', 'name'], (e) => ({
    heading: address(required(e, 'heading'), fieldPath(e, 'heading')),
    name: plainText(required(e, 'name'), fieldPath(e, 'name'), 1, BOOKMARK_NAME_MAX),
  }));

function instructionOf(f: Fields): string {
  const raw = str(
    required(f, 'instruction'),
    fieldPath(f, 'instruction'),
    TOC_MAX_INSTRUCTION_CHARS
  );
  if (!isPlainText(raw) || !parseTocInstruction(raw))
    throw new Invalid('invalid-value', fieldPath(f, 'instruction'), 'not a TOC field instruction');
  return raw;
}

/** The headings a table of contents change reads. */
function touchesHeadings(change: OrderedTocContent, reader: OrderedTargetReader): void {
  for (const row of change.rows) reader.item(row.heading, 'mark');
  for (const bookmark of change.bookmarks) reader.item(bookmark.heading, 'mark');
}

/** The shown item at an address, of one kind, and its place in the view. */
function addressedItem(
  ctx: Ctx,
  where: OrderedChangeAddress,
  kind: OrderedItem['kind']
): { item: OrderedItem; at: OrderedResolvedPlace } {
  const item = shownItemAt(ctx.doc, ctx.story, where, kind);
  return { item, at: resolveOrderedPlace(ctx.doc, ctx.story.name, where) };
}

const INSERT_FIELD = orderedKindSpec<OrderedInsertFieldChange>({
  kind: 'insert-field',
  fields: ['place', 'field'],
  read: (f) => ({
    place: place(required(f, 'place'), fieldPath(f, 'place')),
    field: oneOf(required(f, 'field'), fieldPath(f, 'field'), [
      'PAGE',
      'NUMPAGES',
      'SECTIONPAGES',
      'PAGE_X_OF_Y',
    ]),
  }),
  touches: (change, reader) => reader.caret(change.place),
  apply: (change, input) =>
    runKindEdit(input, (ctx) => {
      const at = resolveChangePlace(ctx.doc, ctx.story.name, change.place);
      return insertPageField(ctx, at, change.field);
    }),
  policy: { text: true, properties: false },
  publish: 'paragraphs',
  caret: (change) => change.place,
  types: true,
});

const FIELD_CODE = orderedKindSpec<OrderedFieldCodeChange>({
  kind: 'field-code',
  fields: ['field', 'code'],
  read: (f) => ({
    field: address(required(f, 'field'), fieldPath(f, 'field')),
    code: oneOf(required(f, 'code'), fieldPath(f, 'code'), ['PAGE', 'NUMPAGES']),
  }),
  touches: (change, reader) => void reader.item(change.field, 'atom'),
  apply: (change, input) =>
    runKindEdit(input, (ctx) => {
      const { item, at: place } = addressedItem(ctx, change.field, 'atom');
      return setFieldCode(ctx, item, place, change.code);
    }),
  // An instruction is no text a reader sees: a party's untracked one is a package change
  policy: { text: false, properties: false },
  publish: 'paragraphs',
});

const FIELD_RESULT = orderedKindSpec<OrderedFieldResultChange>({
  kind: 'field-result',
  fields: ['fields'],
  read: (f) => ({
    fields: entries(f, 'fields', 1, MAX_FIELD_RESULT_UPDATES, ['field', 'text', 'last'], (e) => ({
      field: address(required(e, 'field'), fieldPath(e, 'field')),
      text: resultText(required(e, 'text'), fieldPath(e, 'text'), MAX_FIELD_RESULT_TEXT_CHARS),
      ...(e.has('last') && bool(e.get('last'), fieldPath(e, 'last'))
        ? { last: true as const }
        : {}),
    })),
  }),
  touches: (change, reader) => {
    for (const entry of change.fields) reader.item(entry.field, 'atom');
  },
  apply: (change, input) =>
    runKindEdit(input, (ctx) => {
      let made: string | null = null;
      for (const entry of change.fields) {
        const { item, at: place } = addressedItem(ctx, entry.field, 'atom');
        made = setFieldResult(ctx, item, place, entry.text, entry.last === true) ?? made;
      }

      return made;
    }),
  policy: { text: true, properties: false },
  publish: 'paragraphs',
});

const TOC_PAGES = orderedKindSpec<OrderedTocPagesChange>({
  kind: 'toc-pages',
  fields: ['pages'],
  read: (f) => ({
    pages: entries(f, 'pages', 1, TOC_MAX_ENTRIES, ['from', 'to', 'text'], (e) => ({
      from: place(required(e, 'from'), fieldPath(e, 'from')),
      to: place(required(e, 'to'), fieldPath(e, 'to')),
      text: resultText(required(e, 'text'), fieldPath(e, 'text'), 32),
    })),
  }),
  touches: (change, reader) => {
    for (const page of change.pages) reader.range(page.from, page.to);
  },
  apply: (change, input) =>
    runKindEdit(input, (ctx) => {
      let made: string | null = null;
      for (const page of change.pages) {
        const { p, start, end } = paragraphRange(ctx, page.from, page.to);
        const units = [...new Set(p.chars.slice(start, end).map((c) => c.item))];
        made = setRunText(ctx, units, page.text) ?? made;
      }

      return made;
    }),
  policy: { text: true, properties: false },
  publish: 'paragraphs',
});

const BOOKMARK = orderedKindSpec<OrderedBookmarkChange>({
  kind: 'bookmark',
  fields: ['name', 'from', 'to', 'remove'],
  read: (f) => {
    if (f.has('from') !== f.has('to'))
      throw new Invalid('missing-field', f.path, 'both of from and to, or neither');
    const name = plainText(required(f, 'name'), fieldPath(f, 'name'), 1, 256);

    // A new bookmark's name follows Word's rules; a removed one is named as the file has it
    if (f.has('from') && bookmarkNameProblem(name) !== null)
      throw new Invalid('invalid-value', fieldPath(f, 'name'), 'not a bookmark name');
    const remove = list(required(f, 'remove'), fieldPath(f, 'remove'), 0, LIMITS.targets).map(
      (value, k) => address(value, at(fieldPath(f, 'remove'), k))
    );
    if (!f.has('from') && remove.length === 0)
      throw new Invalid('empty', fieldPath(f, 'remove'), 'a bookmark to add or markers to remove');
    return {
      name,
      ...(f.has('from')
        ? {
            from: place(f.get('from'), fieldPath(f, 'from')),
            to: place(f.get('to'), fieldPath(f, 'to')),
          }
        : {}),
      remove,
    };
  },
  touches: (change, reader) => {
    if (change.from !== undefined) reader.caret(change.from);
    if (change.to !== undefined) reader.caret(change.to);
    for (const marker of change.remove) reader.item(marker, 'zero');
  },
  apply: (change, input) =>
    runKindEdit(input, (ctx) => {
      const remove = change.remove.map((marker) => addressedItem(ctx, marker, 'zero').item);
      const range =
        change.from !== undefined && change.to !== undefined
          ? {
              from: resolveChangePlace(ctx.doc, ctx.story.name, change.from),
              to: resolveChangePlace(ctx.doc, ctx.story.name, change.to),
            }
          : null;
      if (
        range !== null &&
        (range.to.index < range.from.index ||
          (range.to.index === range.from.index && range.to.offset < range.from.offset))
      )
        throw new OrderedRefusal('invalid-edit', 'a bookmark ends after it starts');
      return editBookmark(ctx, change.name, range, remove);
    }),
  // A bookmark is no text a reader sees: a party's is a package change
  policy: { text: false, properties: false },
  // A removed marker may sit between blocks, where no paragraph's publish reaches it
  publish: 'story',
});

const INSERT_TOC = orderedKindSpec<OrderedInsertTocChange>({
  kind: 'insert-toc',
  fields: ['before', 'alias', 'instruction', 'rows', 'bookmarks'],
  read: (f) => ({
    before: address(required(f, 'before'), fieldPath(f, 'before')),
    alias: plainText(required(f, 'alias'), fieldPath(f, 'alias'), 1, 128),
    instruction: instructionOf(f),
    rows: rowsOf(f),
    bookmarks: bookmarksOf(f),
  }),
  touches: (change, reader) => {
    reader.item(change.before, 'mark');
    touchesHeadings(change, reader);
  },
  apply: (change, input) =>
    runKindEdit(input, (ctx) => {
      const before = resolveOrderedPlace(ctx.doc, ctx.story.name, change.before);
      return insertToc(ctx, before, change.alias, change);
    }),
  policy: { text: true, properties: false },
  publish: 'story',
});

const TOC_RESULT = orderedKindSpec<OrderedTocResultChange>({
  kind: 'toc-result',
  fields: ['separate', 'end', 'instruction', 'rows', 'bookmarks'],
  read: (f) => ({
    separate: address(required(f, 'separate'), fieldPath(f, 'separate')),
    end: address(required(f, 'end'), fieldPath(f, 'end')),
    instruction: instructionOf(f),
    rows: rowsOf(f),
    bookmarks: bookmarksOf(f),
  }),
  touches: (change, reader) => {
    reader.item(change.separate, 'zero');
    reader.item(change.end, 'zero');
    touchesHeadings(change, reader);
  },
  apply: (change, input) =>
    runKindEdit(input, (ctx) => {
      const separate = addressedItem(ctx, change.separate, 'zero').item;
      const end = addressedItem(ctx, change.end, 'zero').item;
      return replaceTocResult(ctx, separate, end, change);
    }),
  policy: { text: true, properties: false },
  publish: 'story',
});

/** The specs of {@link OrderedFieldChange}. */
export const FIELD_KINDS: { readonly [K in OrderedFieldChange['kind']]: OrderedKindSpec } = {
  'insert-field': INSERT_FIELD,
  'field-code': FIELD_CODE,
  'field-result': FIELD_RESULT,
  'toc-pages': TOC_PAGES,
  bookmark: BOOKMARK,
  'insert-toc': INSERT_TOC,
  'toc-result': TOC_RESULT,
};
