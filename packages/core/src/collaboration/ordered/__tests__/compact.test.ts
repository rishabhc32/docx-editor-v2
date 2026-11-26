// The cleanup at a close (`compactOrdered`, plan D8): what it removes, what it keeps, and how a
// removed address reads. `fold-check.ts` compares a cleaned room with the live one in every fold
// session (`withCleanup`).

import { describe, expect, test } from 'bun:test';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { compactOrdered } from '../change/compact.ts';
import { itemAt, storyNamed } from '../model/document.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import {
  apply,
  at,
  BODY,
  DIRECT,
  EDIT,
  load,
  partsOf,
  refusal,
  textOf,
} from './change-apply-fixture.ts';

/** The organiser's decision on the card of `id`, as the organiser sees it now. */
function decide(
  doc: OrderedDocument,
  change: string,
  id: string,
  status: 'accepted' | 'rejected'
): void {
  const seen = orderedCardSeen(doc, id);
  if (seen === null) {
    throw new Error(`no card for ${id}`);
  }

  apply(doc, change, { kind: 'decide', cards: [{ suggestions: seen }], status }, 'Olive');
}

/** The items a suggestion labels now. */
const labelledBy = (doc: OrderedDocument, id: string): OrderedItem[] => [
  ...(doc.labelled.get(id) ?? []),
];

/** Whether an item is in its story's tree. */
const inTree = (item: OrderedItem): boolean => item.leaf !== undefined;

describe('a cleanup at a close', () => {
  test('removes a rejected insertion and an accepted deletion, and keeps their suggestions', () => {
    const doc = load();
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 4), text: 'new ' });
    apply(doc, 'ana:2', { kind: 'delete', ...EDIT, from: at(doc, 8, 0), to: at(doc, 8, 4) });
    decide(doc, 'olive:1', 'S:ana:1.1', 'rejected');
    decide(doc, 'olive:2', 'S:ana:2.1', 'accepted');
    const went = [...labelledBy(doc, 'S:ana:1.1'), ...labelledBy(doc, 'S:ana:2.1')];
    const parts = partsOf(doc);
    const text = textOf(doc, 8);

    expect(compactOrdered(doc)).toEqual({ items: went.length });
    expect(went.filter(inTree)).toEqual([]);
    expect([doc.suggestions.has('S:ana:1.1'), doc.suggestions.has('S:ana:2.1')]).toEqual([
      true,
      true,
    ]);
    expect(partsOf(doc)).toEqual(parts);
    expect(textOf(doc, 8)).toBe(text);
    expect(compactOrdered(doc)).toEqual({ items: 0 });
  });

  test('keeps a pending suggestion; removes what a direct edit hid, as the room refuses its revert', () => {
    const doc = load();
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 4), text: 'new ' });
    const pending = labelledBy(doc, 'S:ana:1.1');
    const deletion = { kind: 'delete', ...DIRECT, from: at(doc, 8, 0), to: at(doc, 8, 4) } as const;
    apply(doc, 'olive:1', deletion, 'Olive');
    const insertion = { kind: 'insert', ...DIRECT, place: at(doc, 8, 8), text: 'more ' } as const;
    apply(doc, 'olive:2', insertion, 'Olive');
    const revert = { kind: 'revert', suggestions: ['S:olive:2.1'], status: 'rejected' } as const;
    apply(doc, 'olive:3', revert, 'Olive');
    const went = [...labelledBy(doc, 'S:olive:1.1'), ...labelledBy(doc, 'S:olive:2.1')];
    const parts = partsOf(doc);
    const text = textOf(doc, 8);

    expect(compactOrdered(doc)).toEqual({ items: went.length });
    expect(went.filter(inTree)).toEqual([]);
    expect(pending.every(inTree)).toBe(true);
    expect(partsOf(doc)).toEqual(parts);
    expect(textOf(doc, 8)).toBe(text);
  });

  test('a removed address names no item: a change there is refused by name', () => {
    const doc = load();
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 4), text: 'new ' });
    apply(doc, 'ana:2', { kind: 'delete', ...EDIT, from: at(doc, 8, 0), to: at(doc, 8, 4) });
    decide(doc, 'olive:1', 'S:ana:1.1', 'rejected');
    decide(doc, 'olive:2', 'S:ana:2.1', 'accepted');

    // A typed unit, and a unit of a loaded paragraph that keeps its other units
    const places = ['S:ana:1.1', 'S:ana:2.1'].map((id) => {
      const [removed] = labelledBy(doc, id);
      if (removed === undefined) {
        throw new Error(`no item of ${id}`);
      }

      return { space: removed.space, offset: removed.offset + 1 };
    });
    compactOrdered(doc);

    for (const place of places) {
      expect(itemAt(storyNamed(doc, BODY), place)).toBeNull();
      expect(refusal(doc, 'ana:3', { kind: 'insert', ...EDIT, place, text: 'x' })).toBe(
        'unknown-address'
      );
    }
  });

  test("ends the undo of one's own edits: no client holds one after a close", () => {
    const doc = load();
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 0), text: 'abcd ' });
    apply(doc, 'ana:2', { kind: 'delete', ...EDIT, from: at(doc, 0, 1), to: at(doc, 0, 3) });
    expect(doc.ownEdited.has('ana:2')).toBe(true);
    const taken = labelledBy(doc, 'S:ana:1.1').filter((item) => item.retracted);
    const parts = partsOf(doc);

    expect(compactOrdered(doc)).toEqual({ items: taken.length });
    expect(doc.ownEdited.size).toBe(0);
    expect(labelledBy(doc, 'S:ana:1.1').filter((item) => item.taken)).toEqual([]);
    expect(partsOf(doc)).toEqual(parts);
  });

  test('keeps the last unit of a space: typing on never takes a removed offset again', () => {
    const doc = load();
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 0), text: 'hello' });
    apply(doc, 'ana:2', { kind: 'delete', ...EDIT, from: at(doc, 0, 3), to: at(doc, 0, 5) });
    const [first] = labelledBy(doc, 'S:ana:1.1');
    const space = first?.space ?? '';

    expect(compactOrdered(doc)).toEqual({ items: 0 });
    apply(doc, 'ana:3', { kind: 'insert', ...EDIT, place: at(doc, 0, 3), text: 'p' });
    const typed = labelledBy(doc, 'S:ana:1.1').filter((item) => item.text === 'p');
    expect(typed.map((item) => item.space === space)).toEqual([false]);
    expect(textOf(doc, 0).startsWith('help')).toBe(true);
  });
});
