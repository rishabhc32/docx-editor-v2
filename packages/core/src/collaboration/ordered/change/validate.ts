// The change validator, shared by the client and the room (design D6). It checks a change's
// shape and caps only; whether its addresses exist is `targets.ts`, against a document. The
// result is a new object with the known fields in a fixed order, never the input object, so
// nothing a sender adds survives and the room logs one canonical form.

import { ORDERED_NAMESPACE_PATTERN, parseOrderedChangeId } from './ids.ts';
import {
  at,
  attempt,
  fieldPath,
  fieldsOf,
  int,
  Invalid,
  isPlainText,
  list,
  oneOf,
  plainText,
  required,
  str,
  type Fields,
  type OrderedChangeResult,
} from './read.ts';
import { ORDERED_CHANGE_KINDS, orderedKindSpecOf } from './kinds/registry.ts';
import {
  ORDERED_CHANGE_LIMITS as LIMITS,
  isOrderedCoreKind,
  type OrderedChange,
  type OrderedChangeKind,
  type OrderedCoreChangeKind,
  type OrderedStampedChange,
} from './types.ts';
import { readPerson } from '../model/person.ts';
import { tableOp } from './table-op.ts';
import {
  address,
  addresses,
  commentId,
  commentText,
  controlValue,
  linkTarget,
  optionalRange,
  place,
  properties,
  seenCards,
  suggestionId,
  suggestionIds,
} from './values.ts';

/** What the validator checks a change against. @public */
export interface OrderedValidateOptions {
  /** The sender's namespace: every change id must be in it. Omit only for changes the room sent. */
  readonly namespace?: string;
  /** The room's story parts: a change that names another part is refused. */
  readonly parts?: { has(name: string): boolean };
}

const SERVER_FIELDS: ReadonlySet<string> = new Set(['n', 'by', 'at']);
const BASE = ['v', 'id', 'kind'] as const;
const EDIT = [...BASE, 'part', 'mode', 'into'] as const;
const PART = /^\/[^\s]{1,511}$/;
const STAMP_DATE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/;

// ---- kinds ----

type Payload = Record<string, unknown>;
interface KindSpec {
  readonly fields: readonly string[];
  readonly read: (f: Fields) => Payload;
}

const get = (f: Fields, key: string) => required(f, key);
const range = (f: Fields) => ({
  from: place(get(f, 'from'), fieldPath(f, 'from')),
  to: place(get(f, 'to'), fieldPath(f, 'to')),
});

/** A take-back's `undo`: earlier change ids of the sender's own namespace, none twice. */
function ownEarlierChanges(f: Fields): string[] {
  const path = fieldPath(f, 'undo');
  const namespace = parseOrderedChangeId(str(get(f, 'id'), fieldPath(f, 'id'), 96))?.namespace;
  const ids = list(get(f, 'undo'), path, 1, LIMITS.targets).map((value, k) => {
    const id = str(value, at(path, k), 96);
    const parsed = parseOrderedChangeId(id);
    if (parsed === null) throw new Invalid('invalid-id', at(path, k));
    if (parsed.namespace !== namespace) throw new Invalid('foreign-id', at(path, k));
    return id;
  });
  if (new Set(ids).size !== ids.length) throw new Invalid('repeated-id', path);
  return ids;
}

const KINDS: Readonly<Record<OrderedCoreChangeKind, KindSpec>> = {
  insert: {
    fields: [...EDIT, 'place', 'text'],
    read: (f) => ({
      place: place(get(f, 'place'), fieldPath(f, 'place')),
      text: plainText(get(f, 'text'), fieldPath(f, 'text'), 1, LIMITS.text),
    }),
  },
  'insert-tab': {
    fields: [...EDIT, 'place'],
    read: (f) => ({ place: place(get(f, 'place'), fieldPath(f, 'place')) }),
  },
  'insert-break': {
    fields: [...EDIT, 'place', 'break', 'clear'],
    read: (f) => {
      const kind = oneOf(get(f, 'break'), fieldPath(f, 'break'), [
        'line',
        'page',
        'column',
      ] as const);
      const out: Payload = { place: place(get(f, 'place'), fieldPath(f, 'place')), break: kind };
      if (f.has('clear')) {
        // Only a line break restarts below floating objects (ECMA-376 17.3.3.1)
        if (kind !== 'line') throw new Invalid('unknown-field', fieldPath(f, 'clear'));
        out.clear = oneOf(f.get('clear'), fieldPath(f, 'clear'), ['left', 'right', 'all'] as const);
      }

      return out;
    },
  },
  delete: { fields: [...EDIT, 'from', 'to'], read: range },
  format: {
    fields: [...EDIT, 'from', 'to', 'properties'],
    read: (f) => ({
      ...range(f),
      properties: properties(get(f, 'properties'), fieldPath(f, 'properties'), 'run', true),
    }),
  },
  'mark-format': {
    fields: [...EDIT, 'marks', 'properties'],
    read: (f) => ({
      marks: addresses(get(f, 'marks'), fieldPath(f, 'marks')),
      properties: properties(get(f, 'properties'), fieldPath(f, 'properties'), 'run', true),
    }),
  },
  split: {
    fields: [...EDIT, 'place'],
    read: (f) => ({ place: place(get(f, 'place'), fieldPath(f, 'place')) }),
  },
  join: {
    fields: [...EDIT, 'mark'],
    read: (f) => ({ mark: address(get(f, 'mark'), fieldPath(f, 'mark')) }),
  },
  'paragraph-property': {
    fields: [...EDIT, 'marks', 'properties'],
    read: (f) => ({
      marks: addresses(get(f, 'marks'), fieldPath(f, 'marks')),
      properties: properties(get(f, 'properties'), fieldPath(f, 'properties'), 'paragraph', true),
    }),
  },
  'table-property': {
    fields: [...EDIT, 'on', 'op'],
    read: (f) => ({
      on: addresses(get(f, 'on'), fieldPath(f, 'on')),
      op: tableOp(get(f, 'op'), fieldPath(f, 'op')),
    }),
  },
  'insert-row': {
    fields: [...EDIT, 'row', 'where'],
    read: (f) => ({
      row: address(get(f, 'row'), fieldPath(f, 'row')),
      where: oneOf(get(f, 'where'), fieldPath(f, 'where'), ['above', 'below'] as const),
    }),
  },
  'delete-row': {
    fields: [...EDIT, 'row'],
    read: (f) => ({ row: address(get(f, 'row'), fieldPath(f, 'row')) }),
  },
  'insert-column': {
    fields: [...EDIT, 'cell', 'where'],
    read: (f) => ({
      cell: address(get(f, 'cell'), fieldPath(f, 'cell')),
      where: oneOf(get(f, 'where'), fieldPath(f, 'where'), ['left', 'right'] as const),
    }),
  },
  'delete-column': {
    fields: [...EDIT, 'cell'],
    read: (f) => ({ cell: address(get(f, 'cell'), fieldPath(f, 'cell')) }),
  },
  'insert-table': {
    fields: [...EDIT, 'place', 'rows', 'columns', 'columnWidth'],
    read: (f) => ({
      place: place(get(f, 'place'), fieldPath(f, 'place')),
      rows: int(get(f, 'rows'), fieldPath(f, 'rows'), 1, LIMITS.tableRows),
      columns: int(get(f, 'columns'), fieldPath(f, 'columns'), 1, LIMITS.tableColumns),
      columnWidth: int(get(f, 'columnWidth'), fieldPath(f, 'columnWidth'), 1, LIMITS.columnWidth),
    }),
  },
  'control-value': {
    fields: [...EDIT, 'control', 'value'],
    read: (f) => ({
      control: address(get(f, 'control'), fieldPath(f, 'control')),
      value: controlValue(get(f, 'value'), fieldPath(f, 'value')),
    }),
  },
  'link-target': {
    fields: [...EDIT, 'link', 'href', 'anchor', 'tooltip', 'from', 'to'],
    read: (f) => ({
      link: address(get(f, 'link'), fieldPath(f, 'link')),
      ...linkTarget(f),
      ...optionalRange(f),
    }),
  },
  move: {
    fields: [...EDIT, 'from', 'to', 'place'],
    read: (f) => ({ ...range(f), place: place(get(f, 'place'), fieldPath(f, 'place')) }),
  },
  'insert-note': {
    fields: [...EDIT, 'place', 'note', 'text'],
    read: (f) => ({
      place: place(get(f, 'place'), fieldPath(f, 'place')),
      note: oneOf(get(f, 'note'), fieldPath(f, 'note'), ['footnote', 'endnote'] as const),
      text: plainText(get(f, 'text'), fieldPath(f, 'text'), 0, LIMITS.text),
    }),
  },
  decide: {
    fields: [...BASE, 'cards', 'status'],
    read: (f) => ({
      cards: seenCards(get(f, 'cards'), fieldPath(f, 'cards')),
      status: oneOf(get(f, 'status'), fieldPath(f, 'status'), ['accepted', 'rejected'] as const),
    }),
  },
  retract: {
    fields: [...BASE, 'suggestions', 'undo'],
    read: (f) => {
      // A take-back that only undoes edits of one's own words names no suggestion
      const undo = f.has('undo') ? ownEarlierChanges(f) : undefined;
      const min = undo === undefined ? 1 : 0;
      return {
        suggestions: suggestionIds(get(f, 'suggestions'), fieldPath(f, 'suggestions'), min),
        ...(undo === undefined ? {} : { undo }),
      };
    },
  },
  revert: {
    fields: [...BASE, 'suggestions', 'status'],
    read: (f) => ({
      suggestions: suggestionIds(get(f, 'suggestions'), fieldPath(f, 'suggestions')),
      status: oneOf(get(f, 'status'), fieldPath(f, 'status'), ['accepted', 'rejected'] as const),
    }),
  },
  comment: {
    fields: [...BASE, 'part', 'from', 'to', 'text'],
    read: (f) => ({ ...range(f), text: commentText(get(f, 'text'), fieldPath(f, 'text')) }),
  },
  'comment-reply': {
    fields: [...BASE, 'comment', 'text'],
    read: (f) => ({
      comment: commentId(get(f, 'comment'), fieldPath(f, 'comment')),
      text: commentText(get(f, 'text'), fieldPath(f, 'text')),
    }),
  },
  'comment-resolve': {
    fields: [...BASE, 'comment', 'note'],
    read: (f) => {
      const comment = commentId(get(f, 'comment'), fieldPath(f, 'comment'));
      if (!f.has('note')) {
        return { comment };
      }

      return { comment, note: commentText(f.get('note'), fieldPath(f, 'note')) };
    },
  },
};

/** A kind's spec: the first kinds' own, or the group spec's, past the edit's own fields. */
function specOf(kind: OrderedChangeKind): KindSpec {
  if (isOrderedCoreKind(kind)) {
    return KINDS[kind];
  }

  const group = orderedKindSpecOf(kind);
  if (group === undefined) {
    throw new Error(`No spec for the change kind ${kind}.`);
  }

  return { fields: [...EDIT, ...group.fields], read: (f) => group.read(f) };
}

const KNOWN_FIELDS = new Map<OrderedChangeKind, ReadonlySet<string>>(
  ORDERED_CHANGE_KINDS.map((kind) => [kind, new Set(specOf(kind).fields)])
);

/** Every field any kind has. */
const ANY_FIELDS: ReadonlySet<string> = new Set(
  ORDERED_CHANGE_KINDS.flatMap((kind) => specOf(kind).fields)
);

function partOf(f: Fields, options: OrderedValidateOptions): string {
  const part = str(get(f, 'part'), fieldPath(f, 'part'), 512);
  if (!PART.test(part) || !isPlainText(part))
    throw new Invalid('invalid-part', fieldPath(f, 'part'));
  if (options.parts && !options.parts.has(part))
    throw new Invalid('unknown-part', fieldPath(f, 'part'));
  return part;
}

function readChange(
  value: unknown,
  path: string,
  options: OrderedValidateOptions,
  stamped: boolean
) {
  const kindValue =
    typeof value === 'object' && value !== null && !Array.isArray(value)
      ? Object.getOwnPropertyDescriptor(value, 'kind')?.value
      : undefined;
  const kind =
    typeof kindValue === 'string' ? ORDERED_CHANGE_KINDS.find((k) => k === kindValue) : undefined;

  // An unknown kind is reported as such, not as an unknown field
  const known = kind === undefined ? ANY_FIELDS : KNOWN_FIELDS.get(kind)!;
  const fields = stamped
    ? fieldsOf(value, path, new Set([...known, ...SERVER_FIELDS]))
    : fieldsOf(value, path, known, SERVER_FIELDS);
  if (get(fields, 'v') !== 1) throw new Invalid('invalid-version', fieldPath(fields, 'v'));
  const id = str(get(fields, 'id'), fieldPath(fields, 'id'), 96);
  const parsed = parseOrderedChangeId(id);
  if (parsed === null) throw new Invalid('invalid-id', fieldPath(fields, 'id'));
  if (options.namespace !== undefined && parsed.namespace !== options.namespace)
    throw new Invalid('foreign-id', fieldPath(fields, 'id'));
  if (kind === undefined) {
    if (typeof get(fields, 'kind') !== 'string')
      throw new Invalid('invalid-type', fieldPath(fields, 'kind'));
    throw new Invalid('unknown-kind', fieldPath(fields, 'kind'));
  }

  const spec = specOf(kind);
  const head: Payload = { v: 1, id, kind };
  if (spec.fields.includes('part')) head.part = partOf(fields, options);
  if (spec.fields.includes('mode'))
    head.mode = oneOf(get(fields, 'mode'), fieldPath(fields, 'mode'), ['suggest', 'edit'] as const);
  if (fields.has('into')) head.into = suggestionId(fields.get('into'), fieldPath(fields, 'into'));
  const change = { ...head, ...spec.read(fields) };
  if (!stamped) return change;
  return {
    ...change,
    n: int(get(fields, 'n'), fieldPath(fields, 'n'), 1),
    by: readPerson(get(fields, 'by'), fieldPath(fields, 'by')),
    at: stampDate(get(fields, 'at'), fieldPath(fields, 'at')),
  };
}

function stampDate(value: unknown, path: string): string {
  const date = str(value, path, 20);
  if (!STAMP_DATE.test(date)) throw new Invalid('invalid-date', path);
  return date;
}

// ---- entry points ----

/**
 * Check one change from a client. `n`, `by` and `at` are refused (`server-field`).
 *
 * @public
 */
export function validateOrderedChange(
  value: unknown,
  options: OrderedValidateOptions = {}
): OrderedChangeResult<OrderedChange> {
  checkNamespaceOption(options);

  // Every kind's reader builds exactly that kind's fields, so the object is that change
  return attempt(() => readChange(value, '', options, false) as unknown as OrderedChange);
}

/**
 * Check a batch from a client: 1 to {@link ORDERED_CHANGE_LIMITS}.batch changes, each valid,
 * no id twice. One bad change refuses the whole batch.
 *
 * @public
 */
export function validateOrderedBatch(
  value: unknown,
  options: OrderedValidateOptions = {}
): OrderedChangeResult<readonly OrderedChange[]> {
  checkNamespaceOption(options);
  return attempt(() => readClientChanges(value, '', options));
}

/**
 * Check a change the room ordered: the client fields and `n`, `by` and `at`.
 *
 * @public
 */
export function validateOrderedStampedChange(
  value: unknown,
  options: OrderedValidateOptions = {}
): OrderedChangeResult<OrderedStampedChange> {
  checkNamespaceOption(options);
  return attempt(() => readChange(value, '', options, true) as unknown as OrderedStampedChange);
}

/** Check a list of ordered changes at a path (a message's `changes`). */
export function readStampedChanges(
  value: unknown,
  path: string,
  max: number
): OrderedStampedChange[] {
  return list(value, path, 0, max).map(
    (item, k) => readChange(item, at(path, k), {}, true) as unknown as OrderedStampedChange
  );
}

/** Check a list of client changes at a path (a message's `changes`). */
export function readClientChanges(
  value: unknown,
  path: string,
  options: OrderedValidateOptions
): OrderedChange[] {
  const seen = new Set<string>();
  return list(value, path, 1, LIMITS.batch).map((item, k) => {
    const change = readChange(item, at(path, k), options, false) as unknown as OrderedChange;
    if (seen.has(change.id)) throw new Invalid('repeated-id', at(at(path, k), 'id'));
    seen.add(change.id);
    return change;
  });
}

/** Check one client change at a path. */
export function readClientChange(
  value: unknown,
  path: string,
  options: OrderedValidateOptions
): OrderedChange {
  return readChange(value, path, options, false) as unknown as OrderedChange;
}

function checkNamespaceOption(options: OrderedValidateOptions): void {
  if (options.namespace !== undefined && !ORDERED_NAMESPACE_PATTERN.test(options.namespace))
    throw new RangeError('The namespace option is not a valid namespace.');
}
