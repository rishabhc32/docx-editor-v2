// A suggested row deletion over a cell that holds a change of its own: another person's new
// column puts a `w:cellIns` in it, a pending column deletion a `w:cellDel`. A cell holds at most
// one of `w:cellIns`, `w:cellDel` and `w:cellMerge` (`CT_TcPr`), and the engine's tracked row
// deletion refuses such a row (`unsupported-revision`). The room struck the row and wrote both
// markers in one cell (acceptance room s2-short, `ana-1:347`).

import { describe, expect, test } from 'bun:test';
import {
  addressOf,
  apply,
  DIRECT,
  EDIT,
  load,
  opens,
  partsOf,
  refusal,
} from './change-apply-fixture.ts';

/** The fixture's table after Ben's suggested change to the column of its first cell. */
function withColumnChange(kind: 'insert-column' | 'delete-column') {
  const doc = load();
  const [cell] = opens(doc, 'tc');
  if (!cell) {
    throw new Error('fixture');
  }

  const at = { ...EDIT, cell: addressOf(cell) };
  if (kind === 'insert-column') {
    apply(doc, 'ben:1', { kind, ...at, where: 'right' }, 'Ben');
  } else {
    apply(doc, 'ben:1', { kind, ...at }, 'Ben');
  }

  return doc;
}

describe('a suggested row deletion over a cell with a change of its own', () => {
  for (const kind of ['insert-column', 'delete-column'] as const) {
    test(`is refused as the engine refuses it: ${kind}`, () => {
      const doc = withColumnChange(kind);
      const row = () => {
        const [first] = opens(doc, 'tr');
        if (!first) {
          throw new Error('fixture');
        }

        return { kind: 'delete-row' as const, ...EDIT, row: addressOf(first) };
      };
      const before = partsOf(doc);
      expect(refusal(doc, 'ana:1', row())).toBe('engine-refused');
      expect(refusal(doc, 'olga:1', row(), 'Olga')).toBe('engine-refused');
      expect(partsOf(doc)).toEqual(before);

      // A direct deletion removes the row with the cell, as the engine's untracked one does
      apply(doc, 'olga:2', { ...row(), ...DIRECT }, 'Olga');
      expect(partsOf(doc).join()).not.toContain('A1');
    });
  }

  test('is struck as before when no cell holds a change', () => {
    const doc = load();
    const [first] = opens(doc, 'tr');
    if (!first) {
      throw new Error('fixture');
    }

    expect(refusal(doc, 'ana:1', { kind: 'delete-row', ...EDIT, row: addressOf(first) })).toBe(
      'applied'
    );
    expect(partsOf(doc).join()).toMatch(/<w:trPr><w:del [^>]*w:author="Ana"/);
  });
});
