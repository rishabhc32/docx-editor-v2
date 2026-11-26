// What a checkpoint keeps of who typed and of a control's last value (design D14): live, after one
// fold and after two.

import { describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import type { OrderedDocument } from '../model/types.ts';
import { orderedUnitTyper } from '../model/typers.ts';
import {
  addressOf,
  apply,
  at,
  BODY,
  DIRECT,
  EDIT,
  FOOTNOTES,
  opens,
} from './change-apply-fixture.ts';
import { roomBlobs, TYPED_BARE, TYPED_RUNS } from './checkpoint-fixture.ts';
import { checkFold, folded, mediaStore } from './fold-check.ts';
import { docx, p, t } from './word-fixtures.ts';

describe("a combo box's last value", () => {
  // Found by the fold audit (run-2, quick-short-1): an edit after compose (a list control's last
  // value following its text) once refused a checkpoint's part
  test('follows its text, live and after a fold', async () => {
    const combo =
      '<w:sdt><w:sdtPr><w:id w:val="7"/><w:comboBox w:lastValue="Acme">' +
      '<w:listItem w:displayText="Acme" w:value="Acme"/></w:comboBox></w:sdtPr>' +
      `<w:sdtContent>${t('Acme')}</w:sdtContent></w:sdt>`;
    const { document: doc, base } = openOrderedWordFile(docx(p(t('Party: ') + combo)), { file: 1 });

    // Olive's value, a record the last value follows, and Ana's letter typed in it, pending
    const control = addressOf(opens(doc, 'sdt')[0]!);
    const value = { kind: 'text', text: 'Beta' } as const;
    apply(doc, 'olive:1', { kind: 'control-value', ...DIRECT, control, value }, 'Olive');
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 9), text: 'X' }, 'Ana');
    const blobs = mediaStore(base);
    const lastValue = (bytes: Uint8Array) =>
      /w:lastValue="([^"]*)"/.exec(strFromU8(unzipSync(bytes)['word/document.xml']!))?.[1];
    expect(lastValue(orderedPackageBytes(doc, base, { blobs }))).toBe('BeXta');
    await checkFold(doc, base, blobs, 'a combo box');
  });
});

/** Each unit Olive typed, by its story and what it shows. */
function typedByOlive(doc: OrderedDocument): string[] {
  const out: string[] = [];
  for (const story of doc.stories.values()) {
    const where = isTextBoxStoryName(story.name) ? 'text box' : story.name;
    for (const item of story.tree.items()) {
      const node = item.node?.kind === 'textValue' ? '' : (item.node?.localName ?? '');
      for (let k = 0; k < item.length; k += 1) {
        if (orderedUnitTyper(doc, item, item.offset + k)?.namespace === 'olive') {
          // A complex field's unit has no node of its own
          out.push(`${where} ${node || (item.pieces === undefined ? item.text[k] : 'field')}`);
        }
      }
    }
  }

  return out.sort();
}

describe('who typed each unit kind', () => {
  for (const { name, source, edit, typed } of [
    {
      ...TYPED_RUNS,
      typed: [
        ...[...'XXYY', 'br', 'drawing', 'field', 'footnoteReference', 'tab'].map(
          (u) => `${BODY} ${u}`
        ),
        `${FOOTNOTES}  `,
        `${FOOTNOTES} N`,
        `${FOOTNOTES} footnoteRef`,
        'text box X',
        'text box Y',
      ],
    },
    { ...TYPED_BARE, typed: ['X', 'Y', 'fldSimple', 'oMath'].map((u) => `${BODY} ${u}`) },
  ]) {
    test(`${name}: the same after one fold and two`, async () => {
      const { document: doc, base } = openOrderedWordFile(source, { file: 1 });
      edit(doc);
      expect(typedByOlive(doc)).toEqual(typed);
      const blobs = roomBlobs(base);
      await checkFold(doc, base, blobs, name);
      const once = await folded(doc, base, 2, blobs);
      const twice = await folded(once.document, once.base, 3, blobs);
      expect([once.document, twice.document].map(typedByOlive)).toEqual([typed, typed]);
    });
  }
});
