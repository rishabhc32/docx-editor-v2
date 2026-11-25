// Date values of a date-picker content control: reading a `w:fullDate`, and formatting one with
// the control's own `w:dateFormat` and reading it back. Split from `tree-op-content-controls.ts`.

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

/** Days in `month` (1–12) for a Gregorian `year`, or 0 when the month is out of range. */
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
  const match = FULL_DATE.exec(value.trim());
  if (!match) {
    return null;
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const maxDay = daysInGregorianMonth(year, month);
  if (maxDay === 0 || day < 1 || day > maxDay) {
    return null;
  }

  const date = `${match[1]}-${match[2]}-${match[3]}`;
  if (match[4] === undefined) {
    return `${date}T00:00:00Z`;
  }

  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = Number(match[6]);
  if (hour > 23 || minute > 59 || second > 59) {
    return null;
  }

  const fraction = match[7] !== undefined ? `.${match[7]}` : '';
  const zone = match[8] ?? 'Z';
  if (!isValidXsdDateTimeZone(zone)) {
    return null;
  }

  return `${date}T${match[4]}:${match[5]}:${match[6]}${fraction}${zone}`;
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
  const normalized = normalizeSdtFullDate(fullDate);
  const match = normalized === null ? null : FULL_DATE.exec(normalized);
  if (!match) {
    return null;
  }

  const [year, month, day, hour, minute, second] = match.slice(1, 7).map(Number);
  return {
    year: year!,
    month: month!,
    day: day!,
    hour: hour!,
    minute: minute!,
    second: second!,
  };
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
  dateFormat: string | undefined
): { readonly fullDate: string; readonly text: string } | null {
  const fullDate = normalizeSdtFullDate(value);
  const date = fullDate === null ? null : controlDateOf(fullDate);
  if (fullDate === null || date === null) {
    return null;
  }

  return { fullDate, text: formatContentControlDate(date, dateFormat) };
}

const MONTH_NAMES = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];

const WEEKDAY_NAMES = [
  'Sunday',
  'Monday',
  'Tuesday',
  'Wednesday',
  'Thursday',
  'Friday',
  'Saturday',
];

/**
 * The form of a date control with no `w:dateFormat`: the application's default (ECMA-376,
 * `w:dateFormat`), which an en-US Word writes as `M/d/yyyy`.
 */
const DEFAULT_DATE_FORMAT = 'M/d/yyyy';

/** The letters of a date and time field (ECMA-376 Part 1, date and time formatting). */
type FieldLetter = 'y' | 'M' | 'd' | 'h' | 'H' | 'm' | 's';

const FIELD_LETTERS: ReadonlySet<string> = new Set(['y', 'M', 'd', 'h', 'H', 'm', 's']);

/** The four forms of the morning and afternoon mark, each with its two marks. */
const DAY_HALVES: readonly (readonly [string, string, string])[] = [
  ['AM/PM', 'AM', 'PM'],
  ['am/pm', 'am', 'pm'],
  ['A/P', 'A', 'P'],
  ['a/p', 'a', 'p'],
];

/** One part of a date format: words shown as they are, a field, or a morning or afternoon mark. */
type DateToken =
  | { readonly kind: 'literal'; readonly text: string }
  | { readonly kind: 'field'; readonly letter: FieldLetter; readonly run: number }
  | { readonly kind: 'half'; readonly marks: readonly [string, string] };

/**
 * The parts of a date format (the default form when it names none), or null for one too long to
 * be one Word writes. A BOUNDED walk: the format comes out of an untrusted file, so it is read
 * once, left to right, with no backtracking and no repetition driven by a file-supplied count.
 */
function dateTokens(dateFormat: string | undefined): DateToken[] | null {
  const pattern = dateFormat || DEFAULT_DATE_FORMAT;
  if (pattern.length > 64) {
    return null;
  }

  const tokens: DateToken[] = [];
  let index = 0;
  while (index < pattern.length) {
    const char = pattern[index]!;
    const half = DAY_HALVES.find(([form]) => pattern.startsWith(form, index));
    if (half) {
      tokens.push({ kind: 'half', marks: [half[1], half[2]] });
      index += half[0].length;
    } else if (char === "'") {
      // Quoted words show as they are, without their quotes
      const end = pattern.indexOf("'", index + 1);
      const close = end < 0 ? pattern.length : end;
      tokens.push({ kind: 'literal', text: pattern.slice(index + 1, close) });
      index = close + 1;
    } else if (FIELD_LETTERS.has(char)) {
      let run = 0;
      while (pattern[index + run] === char) {
        run += 1;
      }

      tokens.push({ kind: 'field', letter: char as FieldLetter, run });
      index += run;
    } else {
      tokens.push({ kind: 'literal', text: char });
      index += 1;
    }
  }

  return tokens;
}

/** A month's or a weekday's name: in full, or its first three letters for a run of three. */
function nameFor(names: readonly string[], index: number, run: number): string {
  const name = names[index]!;
  return run === 3 ? name.slice(0, 3) : name;
}

/** A field's text: names for long months and weekdays, else digits padded to the run. */
function fieldText(date: ControlDate, letter: FieldLetter, run: number): string {
  const digits = (value: number) => String(value).padStart(Math.min(run, 2), '0');
  switch (letter) {
    case 'y':
      return run <= 2 ? String(date.year % 100).padStart(2, '0') : String(date.year);
    case 'M':
      return run >= 3 ? nameFor(MONTH_NAMES, date.month - 1, run) : digits(date.month);
    case 'd': {
      const weekday = new Date(Date.UTC(date.year, date.month - 1, date.day)).getUTCDay();
      return run >= 3 ? nameFor(WEEKDAY_NAMES, weekday, run) : digits(date.day);
    }
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
 * names none: the fields of Word's date and time pictures, with English names. A format too
 * long to be one Word writes gives the ISO date.
 */
export function formatContentControlDate(
  date: ControlDate,
  dateFormat: string | undefined
): string {
  const tokens = dateTokens(dateFormat);
  if (!tokens) {
    return fullDateOf(date).slice(0, 10);
  }

  return tokens
    .map((token) => {
      if (token.kind === 'literal') {
        return token.text;
      }

      if (token.kind === 'half') {
        return token.marks[date.hour < 12 ? 0 : 1];
      }

      return fieldText(date, token.letter, token.run);
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
  dateFormat: string | undefined,
  previous: ControlDate | null
): ControlDate | null {
  const tokens = dateTokens(dateFormat);
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
  return valid && formatContentControlDate(date, dateFormat) === text ? date : null;
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
    const index = token.marks.findIndex((mark) => text.startsWith(mark, at));
    if (index < 0) {
      return null;
    }

    fields.afternoon = index === 1;
    return at + token.marks[index]!.length;
  }

  const { letter, run } = token;
  if ((letter === 'M' || letter === 'd') && run >= 3) {
    const names = letter === 'M' ? MONTH_NAMES : WEEKDAY_NAMES;
    const index = names.findIndex((_, k) => text.startsWith(nameFor(names, k, run), at));
    if (index < 0) {
      return null;
    }

    if (letter === 'M') {
      fields.month = index + 1;
    }

    return at + nameFor(names, index, run).length;
  }

  const fullYear = letter === 'y' && run > 2;
  const read = digitsAt(text, at, fullYear ? 1 : Math.min(run, 2), fullYear ? 6 : 2);
  if (!read) {
    return null;
  }

  storeField(fields, letter, letter === 'y' && !fullYear ? century + read.value : read.value);
  return read.at;
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
