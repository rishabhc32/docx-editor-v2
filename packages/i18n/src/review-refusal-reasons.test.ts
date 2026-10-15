import { expect, test } from 'bun:test';
import { createT, en, localizeDisabledReason } from './index';

// Accept or Reject on a change that removes a row, cell or table is refused while another
// reviewer's change inside it is pending. The engine says so in English (`HELD_CHANGE_REASON`
// in core's review lane); the table here makes it reach every locale.
test('the refusal for a change holding another pending change is translatable', () => {
  const reason = 'another pending change sits inside this one; accept or reject that change first';
  expect(localizeDisabledReason(reason, createT(en))).toBe(
    en.disabledReason.revisionHoldsPendingChange
  );
});

// A suggested table change is refused on a row that holds another author's pending formatting
// change. The engine says so in English (`HELD_ROW_REASON` in core's tracked table lane).
test("the refusal for a row holding another author's change is translatable", () => {
  const reason =
    "another author's formatting change is pending on a row this change affects; accept or reject that change first";
  expect(localizeDisabledReason(reason, createT(en))).toBe(
    en.disabledReason.tableRowHoldsPendingChange
  );
});
