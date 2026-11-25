// A date control's text read back through its own `w:dateFormat`, the inverse of the formatter.

import { describe, expect, test } from 'bun:test';
import {
  contentControlPropertiesOf,
  contentControlsIn,
  contentControlTextOf,
  readOoxmlPart,
} from '../index.ts';
import {
  formatContentControlDate,
  parseContentControlDate,
} from '../store/content-control-dates.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** The text and the `w:fullDate` of a date control in `dateFormat` (none: no format) after `value`. */
function dateSet(dateFormat: string | undefined, value: string | { kind: 'date'; iso: string }) {
  const format = dateFormat === undefined ? '' : `<w:dateFormat w:val="${dateFormat}"/>`;
  const read = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:p><w:sdt><w:sdtPr><w:date>${format}` +
      `<w:lid w:val="en-US"/></w:date></w:sdtPr><w:sdtContent><w:r><w:t>Soon</w:t></w:r>` +
      `</w:sdtContent></w:sdt></w:p></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!read.ok) {
    throw new Error(read.reason);
  }

  const controlId = contentControlsIn(read.part.root)[0]!.node.id;
  const result = applyTreeOp(read.part, { op: 'setContentControlValue', controlId, value });
  if (!result.ok) {
    throw new Error(`refused: ${result.reason}`);
  }

  const control = contentControlsIn(result.part.root)[0]!.node;
  const fullDate = contentControlPropertiesOf(control).date?.fullDate;
  return { text: contentControlTextOf(control), fullDate };
}

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

  // As Word for Mac (en-US) shows its date pictures (DATE fields, 2026-11-25): `AM` or `PM` in
  // capitals for every mark, and no field longer than four letters
  test('show weekdays, hours, minutes, seconds, AM or PM and quoted text', () => {
    const shown: [string, string][] = [
      ['dddd, MMMM d, yyyy', 'Tuesday, September 15, 2026'],
      ['ddd d MMM yy', 'Tue 15 Sep 26'],
      ['M/d/yyyy h:mm am/pm', '9/15/2026 2:05 PM'],
      ['h:mm:ss AM/PM', '2:05:09 PM'],
      ['hh:mm:ss a/p', '02:05:09 PM'],
      ['H:m:s', '14:5:9'],
      ["d 'of' MMMM", '15 of September'],
      ['MMMMM', 'September9'],
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

describe('a date value, whichever form it takes: the editor string or the typed value', () => {
  const both = (iso: string) => [iso, { kind: 'date' as const, iso }];

  test("shows Word's en-US default with no format", () => {
    for (const value of both('2026-09-15')) {
      expect(dateSet(undefined, value)).toEqual({
        text: '9/15/2026',
        fullDate: '2026-09-15T00:00:00Z',
      });
    }
  });

  test('takes a full ISO-8601 instant, with its time and offset', () => {
    for (const value of both('2026-09-15T10:30:00+02:00')) {
      expect(dateSet(undefined, value)).toEqual({
        text: '9/15/2026',
        fullDate: '2026-09-15T10:30:00+02:00',
      });
    }
  });

  test('shows a weekday and a time as Word shows them', () => {
    for (const value of both('2026-09-15T14:05:00Z')) {
      expect(dateSet('dddd, MMMM d, yyyy h:mm AM/PM', value).text).toBe(
        'Tuesday, September 15, 2026 2:05 PM'
      );
    }
  });
});
