import { expect, test } from 'bun:test';
import { createT, en, localizeDisabledReason } from './index';

// Repeat Header Rows is refused on a header row that the table style makes one, because the row
// has no flag of its own to clear. The engine says so in English (`HEADER_ROW_FROM_STYLE_REASON`
// in core's table lane); the table here makes it reach every locale.
test('the refusal for a header row the table style makes is translatable', () => {
  const reason = 'the table style makes this row a header row';
  expect(localizeDisabledReason(reason, createT(en))).toBe(en.disabledReason.headerRowFromStyle);
});

// Aligning a table that text wraps around is refused, because its position places it and
// `w:jc` would not move it (`FLOATING_TABLE_ALIGNMENT_REASON`).
test('the refusal for aligning a table that text wraps around is translatable', () => {
  const reason = 'text wraps around this table, so alignment does not move it';
  expect(localizeDisabledReason(reason, createT(en))).toBe(
    en.disabledReason.floatingTableAlignment
  );
});
