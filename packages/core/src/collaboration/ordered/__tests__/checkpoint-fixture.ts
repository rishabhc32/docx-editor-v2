// Room changes that leave a fact in every section of a checkpoint's manifest (design D14).

import type { OrderedDocument } from '../model/types.ts';
import {
  addressOf,
  apply,
  at,
  BODY,
  DIRECT,
  EDIT,
  markOf,
  opens,
  textOf,
} from './change-apply-fixture.ts';

/** The index of the shown body paragraph that starts with `text`. */
export function paragraph(doc: OrderedDocument, text: string): number {
  for (let k = 0; textOf(doc, k) !== '<none>'; k += 1) {
    if (textOf(doc, k).startsWith(text)) {
      return k;
    }
  }

  throw new Error(`no paragraph ${text}`);
}

/**
 * Changes to `APPLY_FILE` that leave a fact of every section past its first paragraph and first
 * cell: a takeover (Pia's press over Quin's), a row height in Ana's new table (an untracked record
 * with its table's cells), Ben's control value (an untracked record of `w:sdtPr`), Olive's words
 * typed inside Ana's insertion (typers), and Ben's comment. Each is found by what it shows.
 */
export function withRoomFacts(doc: OrderedDocument): OrderedDocument {
  for (const [id, who, property] of [
    ['quin:1', 'Quin', { localName: 'jc', attributes: { val: 'center' } }],
    ['pia:1', 'Pia', { localName: 'spacing', attributes: { after: '240' } }],
  ] as const) {
    const marks = [markOf(doc, paragraph(doc, 'Signed'))];
    apply(doc, id, { kind: 'paragraph-property', ...EDIT, marks, properties: [property] }, who);
  }

  const due = () => paragraph(doc, 'Due');
  const table = { rows: 1, columns: 2, columnWidth: 3000 };
  apply(doc, 'ana:1', { kind: 'insert-table', ...EDIT, place: at(doc, due(), 0), ...table });
  const row = addressOf(opens(doc, 'tr').at(-1)!);
  const height = { op: 'setTableRowHeight', heightTwips: 480 } as const;
  apply(doc, 'ana:2', { kind: 'table-property', ...EDIT, on: [row], op: height });
  apply(doc, 'ana:3', { kind: 'insert', ...EDIT, place: at(doc, due(), 4), text: 'new words' });
  apply(
    doc,
    'olive:1',
    { kind: 'insert', ...DIRECT, place: at(doc, due(), 7), text: 'XY' },
    'Olive'
  );
  const control = addressOf(opens(doc, 'sdt').at(-1)!);
  const value = { kind: 'text', text: 'Beta Ltd' } as const;
  apply(doc, 'ben:1', { kind: 'control-value', ...EDIT, control, value }, 'Ben');
  const [from, to] = [at(doc, due(), 0), at(doc, due(), 3)];
  apply(doc, 'ben:2', { kind: 'comment', part: BODY, from, to, text: 'Which date?' }, 'Ben');
  return doc;
}
