// Date values of a date-picker content control: reading an ISO date, and formatting one with
// the control's own `w:dateFormat`. Split from `tree-op-content-controls.ts`.

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:T[\d:.]{1,15}Z?)?$/;

/** ISO input, validated as a real calendar date rather than a well-shaped string. */
export function parseIsoDate(raw: string): { year: number; month: number; day: number } | null {
  const match = ISO_DATE.exec(raw);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1) return null;
  if (date.getUTCDate() !== day) return null;
  return { year, month, day };
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

/**
 * Format a date the way the control's own `w:dateFormat` asks.
 *
 * A BOUNDED token substitution over the patterns Word writes, not a locale engine: the format
 * comes out of an untrusted file, so it is walked once, left to right, with no backtracking and
 * no repetition driven by a file-supplied count.
 */
export function formatContentControlDate(
  date: { year: number; month: number; day: number },
  pattern: string | undefined
): string {
  const iso = `${String(date.year).padStart(4, '0')}-${String(date.month).padStart(2, '0')}-${String(date.day).padStart(2, '0')}`;
  if (!pattern || pattern.length === 0 || pattern.length > 64) return iso;
  let out = '';
  let index = 0;
  while (index < pattern.length) {
    const char = pattern[index]!;
    if (char !== 'y' && char !== 'M' && char !== 'd') {
      out += char;
      index += 1;
      continue;
    }
    let run = 0;
    while (index + run < pattern.length && pattern[index + run] === char) run += 1;
    if (char === 'y')
      out += run <= 2 ? String(date.year % 100).padStart(2, '0') : String(date.year);
    else if (char === 'M') {
      out +=
        run >= 4
          ? MONTH_NAMES[date.month - 1]!
          : run === 3
            ? MONTH_NAMES[date.month - 1]!.slice(0, 3)
            : String(date.month).padStart(Math.min(run, 2), '0');
    } else {
      out += String(date.day).padStart(Math.min(run, 2), '0');
    }
    index += run;
  }
  return out;
}

/**
 * The date a control's display text shows, read back through its own `w:dateFormat`: the
 * inverse of {@link formatContentControlDate}, or null for text that formatter does not write.
 * A two-digit year falls in `century`.
 *
 * Bounded like the formatter: one walk of the pattern and the text together, left to right.
 * The result is formatted again and compared, so a pattern whose fields run together cannot
 * read a date that would not print as this text.
 */
export function parseContentControlDate(
  text: string,
  pattern: string | undefined,
  century = 2000
): { year: number; month: number; day: number } | null {
  if (!pattern || pattern.length === 0 || pattern.length > 64) return parseIsoDate(text);
  const fields: Record<DateField, number | null> = { y: null, M: null, d: null };
  let at = 0;
  let index = 0;
  while (index < pattern.length) {
    const char = pattern[index]!;
    if (!isDateField(char)) {
      if (text[at] !== char) return null;
      at += 1;
      index += 1;
      continue;
    }
    let run = 0;
    while (index + run < pattern.length && pattern[index + run] === char) run += 1;
    index += run;
    const read = readDateField(text, at, char, run, century);
    if (!read) return null;
    fields[char] = read.value;
    at = read.at;
  }
  const { y: year, M: month, d: day } = fields;
  if (at !== text.length || year === null || month === null || day === null) return null;
  const parsed = parseIsoDate(
    `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`
  );
  return parsed && formatContentControlDate(parsed, pattern) === text ? parsed : null;
}
/** A field letter of a `w:dateFormat`: year, month or day. */
type DateField = 'y' | 'M' | 'd';
function isDateField(char: string): char is DateField {
  return char === 'y' || char === 'M' || char === 'd';
}
/** A field's value read from the text, and the text position after it. */
interface DateFieldRead {
  readonly value: number;
  readonly at: number;
}
/**
 * The field of `run` pattern letters `char` in `text` at `at`, or null when the text does not
 * hold it there. A two-digit year falls in `century`.
 */
function readDateField(
  text: string,
  at: number,
  char: DateField,
  run: number,
  century: number
): DateFieldRead | null {
  if (char === 'M' && run >= 3) {
    const found = MONTH_NAMES.findIndex((name) =>
      text.startsWith(run >= 4 ? name : name.slice(0, 3), at)
    );
    if (found < 0) return null;
    return { value: found + 1, at: at + (run >= 4 ? MONTH_NAMES[found]!.length : 3) };
  }
  if (char !== 'y') return digitsAt(text, at, Math.min(run, 2), 2);
  if (run > 2) return digitsAt(text, at, 1, 6);
  const twoDigits = digitsAt(text, at, 2, 2);
  return twoDigits === null ? null : { value: century + twoDigits.value, at: twoDigits.at };
}
/** The number of `min` to `max` digits in `text` at `at`, or null when fewer digits stand there. */
function digitsAt(text: string, at: number, min: number, max: number): DateFieldRead | null {
  let end = at;
  while (end < text.length && end - at < max && text[end]! >= '0' && text[end]! <= '9') {
    end += 1;
  }
  if (end - at < min) return null;
  return { value: Number(text.slice(at, end)), at: end };
}
