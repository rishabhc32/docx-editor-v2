// A change record with a record made over its pending ones lists in the manifest's made-over
// records, or the checkpoint refuses: no section of the manifest keeping it would lose the
// records it holds with no refusal (design D14).

import { expect, test } from 'bun:test';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { checkpointFacts } from '../compose/checkpoint-facts.ts';
import type { HeldTakeovers } from '../compose/context.ts';
import { composedWithoutMedia, type OrderedPartTakeovers } from '../compose/package.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';
import { apply, at, DIRECT, EDIT } from './change-apply-fixture.ts';
import { docx, p, t } from './word-fixtures.ts';

/** A format of words of the first paragraph. */
const format = (doc: OrderedDocument, mode: typeof EDIT | typeof DIRECT, localName: string) =>
  ({
    kind: 'format',
    ...mode,
    from: at(doc, 0, 2),
    to: at(doc, 0, 8),
    properties: [{ localName }],
  }) as const;

/** A room with Olive's direct italic over Ana's pending bold, composed with its takeovers. */
function composed() {
  const { document: doc, base } = openOrderedWordFile(docx(p(t('Cloud and serverless.'))), {
    file: 1,
  });
  apply(doc, 'ana:1', format(doc, EDIT, 'b'));
  apply(doc, 'olive:1', format(doc, DIRECT, 'i'), 'Olive');
  return { doc, ...composedWithoutMedia(doc, base, {}) };
}

/** The takeovers with each made-over record changed by `edit`. */
function edited(
  held: OrderedPartTakeovers,
  edit: (found: HeldTakeovers) => HeldTakeovers
): OrderedPartTakeovers {
  return new Map(
    [...held].map(([part, records]) => [
      part,
      new Map([...records].map(([id, found]) => [id, found.drawn ? edit(found) : found])),
    ])
  );
}

/** The refusal's reason, or `none` with the count of manifest entries that keep the record. */
function outcome(doc: OrderedDocument, pkg: OoxmlPackage, held: OrderedPartTakeovers): string {
  try {
    const { facts } = checkpointFacts(doc, pkg, held, new Map(), new Map());
    return `none: ${facts.madeOver.length + facts.takeovers.length}`;
  } catch (error) {
    return error instanceof OrderedRefusal ? error.reason : String(error);
  }
}

test('a room with a record made over a pending one lists it', () => {
  const { doc, pkg, takeovers } = composed();
  const { facts } = checkpointFacts(doc, pkg, takeovers, new Map(), new Map());
  expect(facts.madeOver.length).toBe(1);
});

test('a made-over record that names a suggestion the room does not hold refuses', () => {
  const { doc, pkg, takeovers } = composed();
  const held = edited(takeovers, (found) => ({
    ...found,
    drawn: found.drawn!.map((record) => ({ ...record, suggestion: 'S:nobody:1.1' })),
  }));
  expect(outcome(doc, pkg, held)).toBe('engine-refused');
});

test('a made-over record on a change record of no made-over target refuses', () => {
  const { doc, pkg, takeovers } = composed();
  const held = edited(takeovers, (found) => ({ ...found, record: 'tcPrChange' }));
  expect(outcome(doc, pkg, held)).toBe('engine-refused');
});
