// A control's prompt that shows again after a fold, card by card as live (design D14): its value
// record read back from the prompt, its format records, and its place among a comment's markers.

import { describe, test } from 'bun:test';
import { openOrderedWordFile } from '../client/open.ts';
import { addressOf, apply, at, BODY, EDIT, opens } from './change-apply-fixture.ts';
import { checkEveryCard, folded, mediaStore } from './fold-check.ts';
import { docx, p, t } from './word-fixtures.ts';

describe("a control's prompt shown again, decided card by card after a fold", () => {
  // After a fold, Olive takes back her value over Ana's pick of the value the prompt stores: a
  // reject of the pick reads the record back from the prompt, as the engine does, and her words
  // taken back are no words that went
  test("a pick of a stored value, once another person's value over it is taken back", async () => {
    const prompt = t('Choose an item.', '<w:rStyle w:val="PlaceholderText"/>');
    const combo =
      '<w:sdt><w:sdtPr><w:id w:val="2"/><w:showingPlcHdr/><w:comboBox w:lastValue="EUR">' +
      `<w:listItem w:displayText="EUR" w:value="EUR"/></w:comboBox></w:sdtPr><w:sdtContent>${prompt}` +
      '</w:sdtContent></w:sdt>';
    const file = docx(p(t('Paid in ') + combo + t(' only.')));
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    const blobs = mediaStore(base);
    const control = addressOf(opens(doc, 'sdt')[0]!);
    const value = { kind: 'control-value', ...EDIT, control } as const;
    apply(doc, 'ana:1', { ...value, value: { kind: 'listItem', value: 'EUR' } });
    apply(doc, 'olive:1', { ...value, value: { kind: 'text', text: 'Acme' } }, 'Olive');
    const copy = await folded(doc, base, 2, blobs);
    apply(copy.document, 'olive:2', { kind: 'retract', suggestions: ['S:olive:1.1'] }, 'Olive');
    await checkEveryCard(copy.document, copy.base, blobs, 'a value taken back after a fold');
  });

  // Ben's pick taken back never filled the prompt: Ana's typing over the prompt she formatted is
  // its first fill, so a reject of the typing shows her format again, after a fold too
  test('a format on a prompt whose first fill was taken back, then typing over it', async () => {
    const items =
      '<w:listItem w:displayText="Net 30" w:value="30"/>' +
      '<w:listItem w:displayText="Net 60" w:value="60"/>';
    const prompt = t('Choose an item.', '<w:rStyle w:val="PlaceholderText"/>');
    const list =
      '<w:sdt><w:sdtPr><w:id w:val="3"/><w:showingPlcHdr/><w:dropDownList w:lastValue="30">' +
      `${items}</w:dropDownList></w:sdtPr><w:sdtContent>${prompt}</w:sdtContent></w:sdt>`;
    const file = docx(p(t('Terms are ') + list + t(' from today.')));
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    const control = addressOf(opens(doc, 'sdt')[0]!);
    const pick = { kind: 'control-value', ...EDIT, control } as const;
    apply(doc, 'ben:1', { ...pick, value: { kind: 'listItem', value: '30' } }, 'Ben');
    apply(doc, 'ben:2', { kind: 'retract', suggestions: ['S:ben:1.1'] }, 'Ben');
    const bold = { kind: 'format', ...EDIT, properties: [{ localName: 'b' }] } as const;
    apply(doc, 'ana:1', { ...bold, from: at(doc, 0, 10), to: at(doc, 0, 24) });
    apply(doc, 'ana:2', { kind: 'insert', ...EDIT, place: at(doc, 0, 10), text: 'net ' });
    await checkEveryCard(doc, base, mediaStore(base), 'a fill taken back');
  });

  // Ana's format on the prompt stays hers when the prompt shows again inside Ben's comment on
  // Olive's value over it: the room reads the record it wrote for her, as a fold does
  test("a format on a prompt shown again inside a comment's range", async () => {
    const prompt = t('Click here to enter text.', '<w:rStyle w:val="PlaceholderText"/>');
    const control =
      '<w:sdt><w:sdtPr><w:id w:val="4"/><w:showingPlcHdr/><w:text/></w:sdtPr>' +
      `<w:sdtContent>${prompt}</w:sdtContent></w:sdt>`;
    const file = docx(p(t('Signed by ') + control + t(' for the buyer.')));
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    const bold = { kind: 'format', ...EDIT, properties: [{ localName: 'b' }] } as const;
    apply(doc, 'ana:1', { ...bold, from: at(doc, 0, 10), to: at(doc, 0, 34) });
    const value = { kind: 'text', text: 'Acme due' } as const;
    const box = addressOf(opens(doc, 'sdt')[0]!);
    apply(doc, 'olive:1', { kind: 'control-value', ...EDIT, control: box, value }, 'Olive');
    const [from, to] = [at(doc, 0, 12), at(doc, 0, 14)];
    apply(doc, 'ben:1', { kind: 'comment', part: BODY, from, to, text: 'Why?' }, 'Ben');
    await checkEveryCard(doc, base, mediaStore(base), "a prompt inside a comment's range");
  });
});
