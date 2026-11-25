// A date control's text read back through its own `w:dateFormat`, the inverse of the formatter.

import { describe, expect, test } from 'bun:test';
import {
  formatContentControlDate,
  parseContentControlDate,
} from '../store/content-control-dates.ts';

describe('reading a date back from its text', () => {
  test('gives back every date the formatter writes, in the patterns Word writes', () => {
    const patterns = [
      'M/d/yyyy',
      'dd-MM-yy',
      'MMMM d, yyyy',
      'd MMM yyyy',
      'yyyy-MM-dd',
      undefined,
    ];
    const dates = [
      { year: 2026, month: 1, day: 2, hour: 0, minute: 0, second: 0 },
      { year: 2026, month: 12, day: 31, hour: 0, minute: 0, second: 0 },
      { year: 2031, month: 5, day: 9, hour: 0, minute: 0, second: 0 },
    ];
    for (const pattern of patterns) {
      for (const date of dates) {
        const text = formatContentControlDate(date, pattern);
        expect(parseContentControlDate(text, pattern)).toEqual(date);
      }
    }
  });

  test('puts a two-digit year in the century it is given', () => {
    const previous = { year: 1950, month: 1, day: 1, hour: 0, minute: 0, second: 0 };
    expect(parseContentControlDate('02-01-99', 'dd-MM-yy', previous)).toEqual({
      year: 1999,
      month: 1,
      day: 2,
      hour: 0,
      minute: 0,
      second: 0,
    });
  });

  test('refuses text the format does not write, and dates the calendar does not have', () => {
    for (const text of ['13/2/2026', '2/30/2026', 'Jan 2, 2026', '1/2/2026 noon', '']) {
      expect(parseContentControlDate(text, 'M/d/yyyy')).toBeNull();
    }
  });
});

describe('the fields of a date format Word writes', () => {
  // Tuesday, 15 September 2026, 2:05:09 in the afternoon
  const date = { year: 2026, month: 9, day: 15, hour: 14, minute: 5, second: 9 };

  test('show weekdays, hours, minutes, seconds, AM or PM and quoted text', () => {
    const shown: [string, string][] = [
      ['dddd, MMMM d, yyyy', 'Tuesday, September 15, 2026'],
      ['ddd d MMM yy', 'Tue 15 Sep 26'],
      ['M/d/yyyy h:mm am/pm', '9/15/2026 2:05 pm'],
      ['h:mm:ss AM/PM', '2:05:09 PM'],
      ['hh:mm A/P', '02:05 P'],
      ['H:m:s', '14:5:9'],
      ['HH:mm a/p', '14:05 p'],
      ["d 'of' MMMM", '15 of September'],
    ];
    for (const [pattern, text] of shown) {
      expect({ pattern, text: formatContentControlDate(date, pattern) }).toEqual({ pattern, text });
    }
  });

  test('read back what they show, a field they do not show at its first value', () => {
    const midnight = { ...date, hour: 0, minute: 0, second: 0 };
    const read: [string, string, typeof date][] = [
      ['dddd, MMMM d, yyyy', 'Tuesday, September 15, 2026', midnight],
      ['M/d/yyyy h:mm:ss AM/PM', '9/15/2026 2:05:09 PM', date],
      ['yyyy-MM-dd HH:mm', '2026-09-15 14:05', { ...date, second: 0 }],
    ];
    for (const [pattern, text, expected] of read) {
      expect({ pattern, read: parseContentControlDate(text, pattern, null) }).toEqual({
        pattern,
        read: expected,
      });
    }

    // A weekday that is not the date's own does not read
    expect(parseContentControlDate('Monday, September 15, 2026', 'dddd, MMMM d, yyyy', null)).toBe(
      null
    );
  });

  test('read a format that shows no day or no year with the first day, in the year before', () => {
    const previous = { year: 2026, month: 6, day: 30, hour: 0, minute: 0, second: 0 };
    expect(parseContentControlDate('September 2026', 'MMMM yyyy', null)).toEqual({
      ...previous,
      month: 9,
      day: 1,
    });
    expect(parseContentControlDate('September 15', 'MMMM d', previous)).toEqual({
      ...previous,
      month: 9,
      day: 15,
    });
    expect(parseContentControlDate('September 15', 'MMMM d', null)).toBeNull();
  });
});
