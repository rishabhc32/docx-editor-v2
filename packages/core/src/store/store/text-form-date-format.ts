import { dateNamesFor } from './date-names.ts';
import { normalizeDateDigits, textFormDateConvention } from './text-form-date-locale.ts';

/**
 * The month a name in `language` or in English names (January is 1), or 0. Lower case on both
 * sides; a day number before the month takes its genitive where the language has one.
 */
function monthNamed(name: string, language: string | undefined): number {
  const wanted = name.toLowerCase();
  for (const names of [dateNamesFor(language), dateNamesFor(undefined)]) {
    for (const list of [names.months, names.monthsAfterDay]) {
      const index = list.findIndex((month) => month.toLowerCase() === wanted);
      if (index >= 0) {
        return index + 1;
      }
    }
  }

  return 0;
}

export function parseTextFormDate(
  text: string,
  locale = 'en-US',
  language?: string
): { year: number; month: number; day: number } | null {
  const convention = textFormDateConvention(locale);
  text = normalizeDateDigits(text.trim(), convention);
  let year: number, month: number, day: number;
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
  const regional = convention.pattern.exec(text);
  // Accept common numeric separators as well as the exact regional pattern.
  const numeric = /^(\d{1,4})\s*([./-])\s*(\d{1,2})\s*\2\s*(\d{1,4})\.?$/.exec(text);
  const named = /^(?:(\d{1,2}) (\p{L}+)|(\p{L}+) (\d{1,2}),) (\d{4})$/u.exec(text);
  if (iso) [year, month, day] = [Number(iso[1]), Number(iso[2]), Number(iso[3])];
  else if (regional || numeric) {
    const values = regional ? regional.slice(1) : [numeric![1]!, numeric![3]!, numeric![4]!];
    const components = { year: 0, month: 0, day: 0 };
    for (const [index, component] of convention.order.entries()) {
      const value = values[index]!;
      if (component === 'year' ? !/^(\d{2}|\d{4})$/.test(value) : !/^\d{1,2}$/.test(value))
        return null;
      components[component] = Number(value);
      if (component === 'year' && value.length === 2)
        components.year += components.year <= 29 ? 2000 : 1900;
    }
    ({ year, month, day } = components);
    // Preserve support for unambiguous day/month and month/day input. Never move the year.
    if (convention.order[0] !== 'year' && !validDate(year, month, day)) [month, day] = [day, month];
  } else if (named) {
    year = Number(named[5]);
    day = Number(named[1] ?? named[4]);
    month = monthNamed((named[2] ?? named[3])!, language);
  } else return null;
  return validDate(year, month, day) ? { year, month, day } : null;
}
function validDate(year: number, month: number, day: number): boolean {
  if (year < 100 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return false;
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}
