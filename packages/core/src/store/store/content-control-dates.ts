// Date values of a date-picker content control: reading a `w:fullDate`, and formatting one with
// the control's own `w:dateFormat` and reading it back. Split from `tree-op-content-controls.ts`.

import type { ContentControlDateFormat } from '../package/content-control-nodes.ts';
import { dateNamesFor, type DateNames } from './date-names.ts';

/** What a date's text depends on: the control's `w:dateFormat` and its language (`w:lid`). */
export type DateForm = Pick<ContentControlDateFormat, 'dateFormat' | 'lid'>;

/** A date control's date and time, as its `w:fullDate` writes them, with no time zone shift. */
export interface ControlDate {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly hour: number;
  readonly minute: number;
  readonly second: number;
}

const FULL_DATE =
  /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2})(?:\.(\d{1,9}))?(Z|[+-]\d{2}:\d{2})?)?$/;

/** Days in `month` (1 to 12) for a Gregorian `year`, or 0 when the month is out of range. */
function daysInGregorianMonth(year: number, month: number): number {
  if (month < 1 || month > 12) {
    return 0;
  }

  if (month === 2) {
    const leap = (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
    return leap ? 29 : 28;
  }

  return month === 4 || month === 6 || month === 9 || month === 11 ? 30 : 31;
}

/**
 * Strict calendar date / date-time parse for content-control values destined for
 * `@w:fullDate`. Accepts `YYYY-MM-DD` or an ISO-8601 local/Zulu/offset date-time; rejects
 * impossible calendar dates (e.g. Feb 31), non-leap Feb 29, and malformed time suffixes.
 * Returns a normalized `xsd:dateTime` string (`…T…Z` or with a numeric offset).
 */
export function normalizeSdtFullDate(value: string): string | null {
  return readFullDate(value)?.fullDate ?? null;
}

/** A `w:fullDate` value checked and normalized (`normalizeSdtFullDate`), with its fields. */
function readFullDate(
  value: string
): { readonly fullDate: string; readonly date: ControlDate } | null {
  const match = FULL_DATE.exec(value.trim());
  if (!match) {
    return null;
  }

  const [year = 0, month = 0, day = 0, hour = 0, minute = 0, second = 0] = match
    .slice(1, 7)
    .map((field) => Number(field ?? 0));
  const maxDay = daysInGregorianMonth(year, month);
  if (maxDay === 0 || day < 1 || day > maxDay) {
    return null;
  }

  const date: ControlDate = { year, month, day, hour, minute, second };
  const ymd = `${match[1]}-${match[2]}-${match[3]}`;
  if (match[4] === undefined) {
    return { fullDate: `${ymd}T00:00:00Z`, date };
  }

  if (date.hour > 23 || date.minute > 59 || date.second > 59) {
    return null;
  }

  const fraction = match[7] !== undefined ? `.${match[7]}` : '';
  const zone = match[8] ?? 'Z';
  if (!isValidXsdDateTimeZone(zone)) {
    return null;
  }

  return { fullDate: `${ymd}T${match[4]}:${match[5]}:${match[6]}${fraction}${zone}`, date };
}

/** Whether `zone` is a legal `xsd:dateTime` timezone: `Z` or a numeric offset within ±14:00. */
function isValidXsdDateTimeZone(zone: string): boolean {
  if (zone === 'Z') {
    return true;
  }

  if (zone.length !== 6 || (zone[0] !== '+' && zone[0] !== '-')) {
    return false;
  }

  const sign = zone[0] === '+' ? 1 : -1;
  const zoneHour = Number(zone.slice(1, 3));
  const zoneMinute = Number(zone.slice(4, 6));
  if (!Number.isInteger(zoneHour) || !Number.isInteger(zoneMinute)) {
    return false;
  }

  if (zoneHour < 0 || zoneMinute < 0 || zoneMinute > 59) {
    return false;
  }

  const totalMinutes = sign * (zoneHour * 60 + zoneMinute);
  if (totalMinutes > 14 * 60 || totalMinutes < -14 * 60) {
    return false;
  }

  // The bound is exactly ±14:00; a fourteen-hour offset must carry `:00` minutes.
  if (Math.abs(totalMinutes) === 14 * 60 && zoneMinute !== 0) {
    return false;
  }

  return true;
}

/** The date and time a `w:fullDate` names, as it writes them, or null for one that names none. */
export function controlDateOf(fullDate: string): ControlDate | null {
  return readFullDate(fullDate)?.date ?? null;
}

/** The `w:fullDate` of a date and time, at UTC, as Word writes it. */
export function fullDateOf(date: ControlDate): string {
  const pad = (value: number, width = 2) => String(value).padStart(width, '0');
  const day = `${pad(date.year, 4)}-${pad(date.month)}-${pad(date.day)}`;
  return `${day}T${pad(date.hour)}:${pad(date.minute)}:${pad(date.second)}Z`;
}

/**
 * A value set in a date control: the `w:fullDate` it writes and the text the control's format
 * shows, or null for a value that is no date. The editor's string value and the editing API's
 * typed value both take it.
 */
export function dateControlValue(
  value: string,
  form: DateForm | undefined
): { readonly fullDate: string; readonly text: string } | null {
  const read = readFullDate(value);
  return read === null
    ? null
    : { fullDate: read.fullDate, text: formatContentControlDate(read.date, form) };
}

/**
 * The form of a date control with no `w:dateFormat`: the application's default (ECMA-376,
 * `w:dateFormat`), which an en-US Word writes as `M/d/yyyy`.
 */
const DEFAULT_DATE_FORMAT = 'M/d/yyyy';

/**
 * The letters of a date and time field (ECMA-376 Part 1, date and time formatting), each with
 * the longest run one field takes: Word reads `MMMMM` as `MMMM` and `M` (`November11`).
 */
const FIELD_RUNS: ReadonlyMap<string, number> = new Map([
  ['y', 4],
  ['M', 4],
  ['d', 4],
  ['h', 2],
  ['H', 2],
  ['m', 2],
  ['s', 2],
]);

type FieldLetter = 'y' | 'M' | 'd' | 'h' | 'H' | 'm' | 's';

/**
 * The forms of the morning and afternoon mark. Word shows `AM` or `PM` for each, in capitals
 * also for `am/pm` and `a/p` (Word for Mac, en-US, 2026-11-25); `A/P` is no mark in Word.
 */
const DAY_HALVES = ['AM/PM', 'am/pm', 'a/p'];

/** One part of a date format as read: words shown as they are, a field, or a day-half mark. */
type FormatPart =
  | { readonly kind: 'literal'; readonly text: string }
  | { readonly kind: 'field'; readonly letter: FieldLetter; readonly run: number }
  | { readonly kind: 'half' };

/**
 * One part of a date format in the control's language: a field shown as digits (with the text
 * of day 1 where the language writes one), or a month or weekday shown by name.
 */
type DateToken =
  | Exclude<FormatPart, { kind: 'field' }>
  | (Extract<FormatPart, { kind: 'field' }> & { readonly firstDay?: string })
  | { readonly kind: 'name'; readonly letter: 'M' | 'd'; readonly names: readonly string[] };

/**
 * The parts of a date control's format in its language (the default form when it names none),
 * or null for a format too long to be one Word writes.
 */
function dateTokens(form: DateForm | undefined): DateToken[] | null {
  const parts = formatParts(form?.dateFormat || DEFAULT_DATE_FORMAT);
  return parts && withNames(parts, dateNamesFor(form?.lid));
}

/**
 * The parts of a date format, or null for one too long to be one Word writes. A BOUNDED walk:
 * the format comes out of an untrusted file, so it is read once, left to right, with no
 * backtracking and no repetition driven by a file-supplied count.
 */
function formatParts(pattern: string): FormatPart[] | null {
  if (pattern.length > 64) {
    return null;
  }

  const parts: FormatPart[] = [];
  let index = 0;
  while (index < pattern.length) {
    const char = pattern[index]!;
    const half = DAY_HALVES.find((form) => pattern.startsWith(form, index));
    if (half) {
      parts.push({ kind: 'half' });
      index += half.length;
    } else if (char === "'") {
      // Quoted words show as they are, without their quotes
      const end = pattern.indexOf("'", index + 1);
      const close = end < 0 ? pattern.length : end;
      parts.push({ kind: 'literal', text: pattern.slice(index + 1, close) });
      index = close + 1;
    } else if (FIELD_RUNS.has(char)) {
      const longest = FIELD_RUNS.get(char)!;
      let run = 0;
      while (run < longest && pattern[index + run] === char) {
        run += 1;
      }

      parts.push({ kind: 'field', letter: char as FieldLetter, run });
      index += run;
    } else {
      parts.push({ kind: 'literal', text: char });
      index += 1;
    }
  }

  return parts;
}

/**
 * The names each part shows in a language, as Word for Mac shows them (2026-11-25): a month
 * after a day number takes its after-day form (`d MMMM`, also `MMMM d` where the language says
 * so), and day 1 right before a month name takes the language's own text (French `1er`).
 */
function withNames(parts: readonly FormatPart[], names: DateNames): DateToken[] {
  const firstDayNumber = parts.findIndex(
    (part) => part.kind === 'field' && part.letter === 'd' && part.run <= 2
  );
  return parts.map((part, index) => {
    if (part.kind !== 'field' || (part.letter !== 'M' && part.letter !== 'd')) {
      return part;
    }

    const { letter, run } = part;
    if (letter === 'M' && run === 4) {
      const afterDay = firstDayNumber >= 0 && (firstDayNumber < index || names.dayEitherSide);
      return { kind: 'name', letter, names: afterDay ? names.monthsAfterDay : names.months };
    }

    if (run >= 3) {
      const full = letter === 'M' ? names.months : names.weekdays;
      const short = letter === 'M' ? names.shortMonths : names.shortWeekdays;
      return { kind: 'name', letter, names: run === 3 ? short : full };
    }

    if (
      letter === 'd' &&
      run === 1 &&
      names.firstDay !== undefined &&
      beforeMonthName(parts, index)
    ) {
      return { ...part, firstDay: names.firstDay };
    }

    return part;
  });
}

/** Whether the part at `index` stands right before a month name: one space, then `MMMM`. */
function beforeMonthName(parts: readonly FormatPart[], index: number): boolean {
  const space = parts[index + 1];
  const month = parts[index + 2];
  return (
    space?.kind === 'literal' &&
    space.text === ' ' &&
    month?.kind === 'field' &&
    month.letter === 'M' &&
    month.run === 4
  );
}

/** The day of the week of a date, 0 for Sunday. */
function weekdayOf(date: ControlDate): number {
  // `Date.UTC` reads years 0 to 99 as 1900 to 1999; `setUTCFullYear` takes the year as it is
  const day = new Date(0);
  day.setUTCFullYear(date.year, date.month - 1, date.day);
  return day.getUTCDay();
}

/** A field's text in digits, padded to the run, or the language's text of day 1. */
function fieldText(date: ControlDate, token: DateToken & { kind: 'field' }): string {
  const digits = (value: number) => String(value).padStart(Math.min(token.run, 2), '0');
  switch (token.letter) {
    case 'y':
      return token.run <= 2 ? String(date.year % 100).padStart(2, '0') : String(date.year);
    case 'M':
      return digits(date.month);
    case 'd':
      return date.day === 1 && token.firstDay !== undefined ? token.firstDay : digits(date.day);
    case 'h':
      return digits(date.hour % 12 || 12);
    case 'H':
      return digits(date.hour);
    case 'm':
      return digits(date.minute);
    case 's':
      return digits(date.second);
  }
}

/**
 * Format a date the way the control's own `w:dateFormat` asks, or in the default form when it
 * names none: the fields of Word's date and time pictures, with the names of the control's
 * language (`w:lid`). A format too long to be one Word writes gives the ISO date.
 */
export function formatContentControlDate(date: ControlDate, form: DateForm | undefined): string {
  const tokens = dateTokens(form);
  if (!tokens) {
    return fullDateOf(date).slice(0, 10);
  }

  return tokens
    .map((token) => {
      if (token.kind === 'literal') {
        return token.text;
      }

      if (token.kind === 'half') {
        return date.hour < 12 ? 'AM' : 'PM';
      }

      if (token.kind === 'name') {
        return token.names[token.letter === 'M' ? date.month - 1 : weekdayOf(date)]!;
      }

      return fieldText(date, token);
    })
    .join('');
}

/** The fields read from a date's text so far. */
interface ReadFields {
  year?: number;
  month?: number;
  day?: number;
  hour?: number;
  minute?: number;
  second?: number;
  /** A 12-hour clock's hour, and whether a mark said afternoon. */
  clock?: number;
  afternoon?: boolean;
}

/**
 * The date and time a control's text shows, read back through its own `w:dateFormat`: the
 * inverse of {@link formatContentControlDate}, or null for text that formatter does not write.
 * A field the format does not show reads as its first value (the first day, the first month,
 * midnight); a year it does not show is the year of `previous`, and a two-digit year falls in
 * the century of `previous` (2000 without one).
 *
 * Bounded like the formatter: one walk of the format and the text together, left to right.
 * The result is formatted again and compared, so a format whose fields run together cannot
 * read a date that would not print as this text.
 */
export function parseContentControlDate(
  text: string,
  form: DateForm | undefined,
  previous: ControlDate | null
): ControlDate | null {
  const tokens = dateTokens(form);
  if (!tokens) {
    return controlDateOf(text);
  }

  const century = previous ? Math.floor(previous.year / 100) * 100 : 2000;
  const fields: ReadFields = {};
  let at: number | null = 0;
  for (const token of tokens) {
    at = readToken(text, at, token, century, fields);
    if (at === null) {
      return null;
    }
  }

  const year = fields.year ?? previous?.year;
  const clock =
    fields.clock === undefined ? undefined : (fields.clock % 12) + (fields.afternoon ? 12 : 0);
  const date: ControlDate = {
    year: year ?? 0,
    month: fields.month ?? 1,
    day: fields.day ?? 1,
    hour: clock ?? fields.hour ?? 0,
    minute: fields.minute ?? 0,
    second: fields.second ?? 0,
  };
  const valid =
    at === text.length &&
    year !== undefined &&
    date.day <= daysInGregorianMonth(date.year, date.month) &&
    date.hour < 24 &&
    date.minute < 60 &&
    date.second < 60;
  return valid && formatContentControlDate(date, form) === text ? date : null;
}

/**
 * Read one part of a format from `text` at `at` into `fields`: the position after it, or null
 * when the text does not hold it there. A weekday sets no field: formatting the date again
 * checks it.
 */
function readToken(
  text: string,
  at: number,
  token: DateToken,
  century: number,
  fields: ReadFields
): number | null {
  if (token.kind === 'literal') {
    return text.startsWith(token.text, at) ? at + token.text.length : null;
  }

  if (token.kind === 'half') {
    const mark = text.slice(at, at + 2);
    if (mark !== 'AM' && mark !== 'PM') {
      return null;
    }

    fields.afternoon = mark === 'PM';
    return at + 2;
  }

  if (token.kind === 'name') {
    const index = longestNameAt(text, at, token.names);
    if (index < 0) {
      return null;
    }

    if (token.letter === 'M') {
      fields.month = index + 1;
    }

    return at + token.names[index]!.length;
  }

  const { letter, run, firstDay } = token;
  if (firstDay !== undefined && text.startsWith(firstDay, at)) {
    fields.day = 1;
    return at + firstDay.length;
  }

  const fullYear = letter === 'y' && run > 2;
  const read = digitsAt(text, at, fullYear ? 1 : Math.min(run, 2), fullYear ? 4 : 2);
  if (!read) {
    return null;
  }

  storeField(fields, letter, letter === 'y' && !fullYear ? century + read.value : read.value);
  return read.at;
}

/** The name in `names` that `text` holds at `at`, the longest one (`Tháng 10`, not `Tháng 1`). */
function longestNameAt(text: string, at: number, names: readonly string[]): number {
  let found = -1;
  names.forEach((name, index) => {
    if (text.startsWith(name, at) && (found < 0 || name.length > names[found]!.length)) {
      found = index;
    }
  });
  return found;
}

/** Keep one field's value. */
function storeField(fields: ReadFields, letter: FieldLetter, value: number): void {
  switch (letter) {
    case 'y':
      fields.year = value;
      break;
    case 'M':
      fields.month = value;
      break;
    case 'd':
      fields.day = value;
      break;
    case 'h':
      fields.clock = value;
      break;
    case 'H':
      fields.hour = value;
      break;
    case 'm':
      fields.minute = value;
      break;
    case 's':
      fields.second = value;
      break;
  }
}

/** The number of `min` to `max` digits in `text` at `at`, or null when fewer digits stand there. */
function digitsAt(
  text: string,
  at: number,
  min: number,
  max: number
): { readonly value: number; readonly at: number } | null {
  let end = at;
  while (end < text.length && end - at < max && text[end]! >= '0' && text[end]! <= '9') {
    end += 1;
  }

  if (end - at < min) {
    return null;
  }

  return { value: Number(text.slice(at, end)), at: end };
}
