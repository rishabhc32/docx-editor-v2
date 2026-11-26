// Who made a room item: the person the login gives the room, stamped on every change. The ID
// decides whose an item is; the name and the company only make the Word name (`w:author`).

import { at, fieldsOf, Invalid, isPlainText, plainText, required } from '../change/read.ts';
import { ORDERED_CHANGE_LIMITS } from '../change/types.ts';
import type { OrderedAttribution, OrderedDocument } from './types.ts';

/**
 * A person as the host's login gives them: a stable ID, a display name, and a company (empty
 * for none). The room stamps it on every change; exported Word files carry only the Word name
 * ({@link orderedAuthorName}).
 *
 * @public
 */
export interface OrderedPerson {
  /** The login's stable ID: it decides whose an item is. */
  readonly id: string;
  /** The display name. */
  readonly name: string;
  /** The company; empty for none. */
  readonly org: string;
}

/** UTF-16 units of a person's ID, name, or company. */
const PERSON_FIELD = 256;

/**
 * The Word name of a person, written as `w:author`: `Name (Org)`, or `Name` with no company.
 *
 * @public
 */
export function orderedAuthorName(person: OrderedPerson): string {
  if (person.org === '') {
    return person.name;
  }

  return `${person.name} (${person.org})`;
}

/** Whether a value is plain text of `min` to the field cap, with no space at either end. */
const isPersonField = (value: string, min: number): boolean =>
  isPlainText(value) &&
  value.length >= min &&
  value.length <= PERSON_FIELD &&
  value.trim() === value;

/** What a person lacks, or undefined when the person is valid. */
function personFault(person: OrderedPerson): string | undefined {
  if (!isPersonField(person.id, 1)) {
    return `an id of 1 to ${PERSON_FIELD} characters`;
  }

  if (!isPersonField(person.name, 1)) {
    return `a name of 1 to ${PERSON_FIELD} characters`;
  }

  if (!isPersonField(person.org, 0)) {
    return `a company of 0 to ${PERSON_FIELD} characters`;
  }

  if (orderedAuthorName(person).length > ORDERED_CHANGE_LIMITS.author) {
    return `a name and company of at most ${ORDERED_CHANGE_LIMITS.author} characters together`;
  }

  return undefined;
}

/**
 * The person, trimmed and checked, where the host gives one (the login, the session options).
 *
 * @throws RangeError when the person is not valid.
 */
export function checkedPerson(person: OrderedPerson): OrderedPerson {
  const trimmed = { id: person.id.trim(), name: person.name.trim(), org: person.org.trim() };
  const fault = personFault(trimmed);
  if (fault !== undefined) {
    throw new RangeError(`A person needs ${fault}.`);
  }

  return Object.freeze(trimmed);
}

/** A person read from untrusted JSON (a log, a frame, a checkpoint), as the room stamped it. */
export function readPerson(value: unknown, path: string): OrderedPerson {
  const f = fieldsOf(value, path, new Set(['id', 'name', 'org']));
  const person = {
    id: plainText(required(f, 'id'), at(path, 'id'), 1, PERSON_FIELD),
    name: plainText(required(f, 'name'), at(path, 'name'), 1, PERSON_FIELD),
    org: plainText(required(f, 'org'), at(path, 'org'), 0, PERSON_FIELD),
  };
  const fault = personFault(person);
  if (fault !== undefined) {
    throw new Invalid('invalid-value', path, `not ${fault}`);
  }

  return person;
}

/** Who made an item, as the rules for own and other items compare it. */
export type OrderedMaker = Pick<OrderedAttribution, 'author' | 'person'>;

/** A person as a maker: their Word name and the person. */
export const makerOf = (person: OrderedPerson): OrderedMaker => ({
  author: orderedAuthorName(person),
  person,
});

/** The attribution of what a person makes at a date: their Word name, the person, the date. */
export const personAttribution = (person: OrderedPerson, date: string): OrderedAttribution => ({
  ...makerOf(person),
  date,
});

/**
 * Whether two makers are one. A room item's maker is its person, by ID. A file's item has no
 * person and belongs to nobody in the room: only an edit made as that item's own author
 * (`asAuthor`) is the same maker, by the name Word saved.
 */
export function sameMaker(a: OrderedMaker, b: OrderedMaker): boolean {
  if (a.person === null || b.person === null) {
    return a.person === b.person && a.author === b.author;
  }

  return a.person.id === b.person.id;
}

/**
 * Whether every suggestion and comment that names a person carries that person's Word name, and
 * so does each attribution a suggestion had before a change gave it another. A room's file is
 * untrusted: a card must never show one name for another person.
 */
export function personsFit(doc: OrderedDocument): boolean {
  const fits = ({ author, person }: OrderedMaker) =>
    person === null || author === orderedAuthorName(person);
  const makers = [
    ...[...doc.suggestions.values()].map(({ attribution }) => attribution),
    ...doc.revisionIds.reattributed.values(),
    ...doc.comments.values(),
  ];
  return makers.every(fits);
}
