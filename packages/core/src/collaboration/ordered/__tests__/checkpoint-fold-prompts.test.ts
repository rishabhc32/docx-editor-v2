// A control's prompt that shows again after a fold, card by card as live (design D14): its value
// record read back from the prompt, its format records, and its place among a comment's markers.

import { describe, expect, test } from 'bun:test';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { openOrderedWordFile } from '../client/open.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  addressOf,
  apply,
  at,
  BODY,
  DIRECT,
  EDIT,
  opens,
  partsOf,
} from './change-apply-fixture.ts';
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

  // Ben's italic on the prompt, taken back while his value fills the control, never was: a reject
  // of the value shows the prompt plain, after a fold too
  test('a format on a prompt taken back while a value fills the control', async () => {
    const prompt = t('Click here to enter text.', '<w:rStyle w:val="PlaceholderText"/>');
    const control =
      '<w:sdt><w:sdtPr><w:id w:val="4"/><w:showingPlcHdr/><w:text/></w:sdtPr>' +
      `<w:sdtContent>${prompt}</w:sdtContent></w:sdt>`;
    const file = docx(p(t('Signed by ') + control + t(' for the buyer.')));
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    const italic = { kind: 'format', ...EDIT, properties: [{ localName: 'i' }] } as const;
    apply(doc, 'ben:1', { ...italic, from: at(doc, 0, 10), to: at(doc, 0, 34) }, 'Ben');
    const value = { kind: 'text', text: 'Acme' } as const;
    const box = addressOf(opens(doc, 'sdt')[0]!);
    apply(doc, 'ben:2', { kind: 'control-value', ...EDIT, control: box, value }, 'Ben');
    apply(doc, 'ben:3', { kind: 'retract', suggestions: ['S:ben:1.1'] }, 'Ben');
    await checkEveryCard(doc, base, mediaStore(base), 'a format taken back');
  });

  // Olive's move into a combo box keeps the file's range markers after a fold. Her direct
  // deletion takes the source and her value takes back the copy, so the move holds no words: a
  // reject of the value shows the prompt with no move range around it, after a fold too
  test("a prompt shown again where a move's words all went", async () => {
    const prompt = t('Choose an item.', '<w:rStyle w:val="PlaceholderText"/>');
    const combo =
      '<w:sdt><w:sdtPr><w:id w:val="2"/><w:showingPlcHdr/><w:comboBox w:lastValue="EUR">' +
      `<w:listItem w:displayText="EUR" w:value="EUR"/></w:comboBox></w:sdtPr><w:sdtContent>${prompt}` +
      '</w:sdtContent></w:sdt>';
    const file = docx(p(t('Paid in ') + combo + t(' only.')) + p(t('Fees net of tax.')));
    const first = openOrderedWordFile(file, { file: 1 });
    const blobs = mediaStore(first.base);
    const moved = { from: at(first.document, 1, 5), to: at(first.document, 1, 8) };
    const place = at(first.document, 0, 8);
    apply(first.document, 'olive:1', { kind: 'move', ...EDIT, ...moved, place }, 'Olive');
    const { document: doc, base } = await folded(first.document, first.base, 2, blobs);
    const source = { from: at(doc, 1, 5), to: at(doc, 1, 8) };
    apply(doc, 'olive:2', { kind: 'delete', ...DIRECT, ...source }, 'Olive');
    const control = addressOf(opens(doc, 'sdt')[0]!);
    const value = { kind: 'text', text: 'Acme' } as const;
    apply(doc, 'olive:3', { kind: 'control-value', ...EDIT, control, value }, 'Olive');
    await checkEveryCard(doc, base, blobs, "a move's words all went");
  });

  // The prompt a party typed over holds a combo box: after a fold the glossary gives it back, and a
  // reject of the typing shows the combo box in it again, not its words alone, and each of its
  // paragraphs. They are the prompt's own: typing in its first drops the second, as live
  const face = '<w:rStyle w:val="PlaceholderText"/>';
  const city =
    '<w:sdt><w:sdtPr><w:id w:val="9204"/><w:showingPlcHdr/><w:comboBox>' +
    '<w:listItem w:displayText="Mumbai" w:value="Mumbai"/></w:comboBox></w:sdtPr>' +
    `<w:sdtContent>${t('Choose a city.', face)}</w:sdtContent></w:sdt>`;
  const courts = p(t('Courts at ', face) + city + t(' have jurisdiction.', face));
  for (const [name, prompt] of [
    ['of one paragraph', courts],
    ['of two paragraphs', p(t('Terms apply.', face)) + courts],
  ] as const) {
    test(`a block prompt ${name} that holds a control, typed over and rejected`, async () => {
      const venue =
        '<w:sdt><w:sdtPr><w:id w:val="9203"/><w:showingPlcHdr/></w:sdtPr>' +
        `<w:sdtContent>${prompt}</w:sdtContent></w:sdt>`;
      const file = docx(p(t('Venue:')) + venue + p(t('Signed.')));
      const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
      const blobs = mediaStore(base);
      apply(doc, 'ben:1', { kind: 'insert', ...EDIT, place: at(doc, 1, 0), text: 'QR' }, 'Ben');
      await checkEveryCard(doc, base, blobs, 'a prompt that holds a control');
      const copy = await folded(doc, base, 2, blobs);
      for (const room of [doc, copy.document]) {
        const seen = orderedCardSeen(room, 'S:ben:1.1') ?? [];
        const reject = {
          kind: 'decide',
          cards: [{ suggestions: seen }],
          status: 'rejected',
        } as const;
        apply(room, 'olive:1', reject, 'Olive');
        apply(room, 'ana:1', { kind: 'insert', ...EDIT, place: at(room, 1, 0), text: 'S' });
      }

      // A paragraph the prompt gives back after a fold is a new one: its ids are new
      const shown = (room: OrderedDocument) =>
        partsOf(room).map((xml) => xml.replace(/ w14:(?:paraId|textId)="[^"]*"/g, ''));
      expect(shown(copy.document)).toEqual(shown(doc));
    });
  }
});
