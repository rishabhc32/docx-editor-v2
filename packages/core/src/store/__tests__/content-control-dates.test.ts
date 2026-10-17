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
      { year: 2026, month: 1, day: 2 },
      { year: 2026, month: 12, day: 31 },
      { year: 2031, month: 5, day: 9 },
    ];
    for (const pattern of patterns) {
      for (const date of dates) {
        const text = formatContentControlDate(date, pattern);
        expect(parseContentControlDate(text, pattern)).toEqual(date);
      }
    }
  });

  test('puts a two-digit year in the century it is given', () => {
    expect(parseContentControlDate('02-01-99', 'dd-MM-yy', 1900)).toEqual({
      year: 1999,
      month: 1,
      day: 2,
    });
  });

  test('refuses text the format does not write, and dates the calendar does not have', () => {
    for (const text of ['13/2/2026', '2/30/2026', 'Jan 2, 2026', '1/2/2026 noon', '']) {
      expect(parseContentControlDate(text, 'M/d/yyyy')).toBeNull();
    }
  });
});
