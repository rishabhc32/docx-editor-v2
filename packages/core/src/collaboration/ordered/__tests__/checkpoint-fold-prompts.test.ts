// A control's prompt that shows again after a fold, card by card as live (design D14): its value
// record read back from the prompt, its format records, and its place among a comment's markers.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { readOoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedIdMinter } from '../change/ids.ts';
import { openOrderedCheckpoint, openOrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import type { OrderedDocument } from '../model/types.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import {
  addressOf,
  apply,
  at,
  BODY,
  DATE,
  DIRECT,
  EDIT,
  opens,
  partsOf,
  textOf,
} from './change-apply-fixture.ts';
import { checkEveryCard, checkFold, checkpointBytes, folded, mediaStore } from './fold-check.ts';
import { testPerson } from './test-persons.ts';
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

  // Typing over a prompt, taken back after a fold, never was: the prompt shows again, as live,
  // where the fold kept it in the glossary
  const face = '<w:rStyle w:val="PlaceholderText"/>';
  const terms =
    '<w:sdt><w:sdtPr><w:id w:val="3"/><w:showingPlcHdr/><w:dropDownList w:lastValue="30">' +
    '<w:listItem w:displayText="Net 30" w:value="30"/></w:dropDownList></w:sdtPr>' +
    `<w:sdtContent>${t('Choose an item.', face)}</w:sdtContent></w:sdt>`;
  const notes =
    '<w:sdt><w:sdtPr><w:id w:val="5"/><w:showingPlcHdr/></w:sdtPr>' +
    `<w:sdtContent>${p(t('Add notes.', face))}</w:sdtContent></w:sdt>`;
  for (const [name, body, place, shown] of [
    ['an inline', p(t('Terms are ') + terms + t(' from today.')), [0, 24], 'Choose an item.'],
    ['a block', p(t('Notes:')) + notes + p(t('Signed.')), [1, 3], 'Add notes.'],
  ] as const) {
    test(`typing over ${name} control's prompt taken back after a fold`, async () => {
      const { document: doc, base } = openOrderedWordFile(docx(body), { file: 1 });
      const blobs = mediaStore(base);
      const [index, n] = place;
      const typing = { kind: 'insert', ...EDIT, place: at(doc, index, n), text: ' due' } as const;
      apply(doc, 'olive:1', typing, 'Olive');
      const copy = await folded(doc, base, 2, blobs);
      for (const room of [doc, copy.document]) {
        apply(room, 'olive:2', { kind: 'retract', suggestions: ['S:olive:1.1'] }, 'Olive');
        expect(textOf(room, index)).toContain(shown);
      }

      await checkFold(copy.document, copy.base, blobs, `${name} prompt after a fold`);
    });
  }

  // After a fold the control names the glossary block Ben's value kept: a reject of Ana's value
  // over the prompt shows that block, as the engine does. A link on the prompt goes; a comment's
  // range made before the fold is the prompt's and goes with it; one made after the fold keeps its
  // range round the block
  const signed = () => {
    const prompt = t('Click here to enter text.', '<w:rStyle w:val="PlaceholderText"/>');
    const control =
      '<w:sdt><w:sdtPr><w:id w:val="4"/><w:showingPlcHdr/><w:text/></w:sdtPr>' +
      `<w:sdtContent>${prompt}</w:sdtContent></w:sdt>`;
    return docx(p(t('Signed by ') + control + t(' for the buyer.')));
  };
  const value = (room: OrderedDocument, text: string) =>
    ({
      kind: 'control-value',
      ...EDIT,
      control: addressOf(opens(room, 'sdt')[0]!),
      value: { kind: 'text', text },
    }) as const;
  const reject = (room: OrderedDocument, id: string) =>
    ({
      kind: 'decide',
      cards: [{ suggestions: orderedCardSeen(room, id) ?? [] }],
      status: 'rejected',
    }) as const;
  const link = { kind: 'insert-link', ...DIRECT, href: 'https://example.com/fees' } as const;
  const comment = { kind: 'comment', part: BODY, text: 'Why?' } as const;
  const block =
    '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Click here to enter text.</w:t></w:r>';
  const around = new RegExp(`^<w:commentRangeStart w:id="\\d+"/>${block}<w:commentRangeEnd `);
  for (const [name, mark, early, content] of [
    ['a link', link, false, new RegExp(`^${block}$`)],
    ["a comment's range made before the fold", comment, true, new RegExp(`^${block}$`)],
    ["a comment's range made after the fold", comment, false, around],
  ] as const) {
    test(`${name} on a prompt that names a block, under a value rejected`, async () => {
      const { document: doc, base } = openOrderedWordFile(signed(), { file: 1 });
      const blobs = mediaStore(base);
      const marked = (room: OrderedDocument) =>
        apply(room, 'olive:2', { ...mark, from: at(room, 0, 31), to: at(room, 0, 33) }, 'Olive');
      apply(doc, 'ben:1', value(doc, 'Acme'), 'Ben');
      apply(doc, 'olive:1', reject(doc, 'S:ben:1.1'), 'Olive');
      if (early) {
        marked(doc);
      }

      const { document: room, base: named } = await folded(doc, base, 2, blobs);
      if (!early) {
        marked(room);
      }

      apply(room, 'ana:1', value(room, 'Acme due'));
      await checkFold(room, named, blobs, `${name} on a named prompt`);
      apply(room, 'olive:3', reject(room, 'S:ana:1.1'), 'Olive');
      expect(textOf(room, 0)).toBe('Signed by Click here to enter text. for the buyer.');
      const [, shown] = partsOf(room)[0]!.match(/<w:sdtContent>(.*?)<\/w:sdtContent>/) ?? [];
      expect(shown).toMatch(content);
    });
  }

  // Seed 91 of the fold shapes: after a fold, Olive's direct deletion takes words of Ben's date
  // and of her own over it; a reject of the two shows the prompt beside the date the file stores,
  // as after a second fold, where those words are no more: a direct edit took them, not the
  // decision. With no fold before, the prompt's own units show again, as they do
  for (const fold of [false, true]) {
    test(`a date the file stores, under picks a direct deletion cut, fold ${fold}`, async () => {
      const date =
        '<w:sdt><w:sdtPr><w:id w:val="1"/><w:showingPlcHdr/>' +
        '<w:date w:fullDate="2026-06-30T00:00:00Z"><w:dateFormat w:val="M/d/yyyy"/>' +
        '<w:lid w:val="en-US"/></w:date></w:sdtPr>' +
        `<w:sdtContent>${t('Click to enter a date.', '<w:rStyle w:val="PlaceholderText"/>')}` +
        '</w:sdtContent></w:sdt>';
      const file = docx(p(t('Due on ') + date + t(' at the latest.')));
      const first = openOrderedWordFile(file, { file: 1 });
      const blobs = mediaStore(first.base);
      const pick = (room: OrderedDocument, iso: string) =>
        ({
          kind: 'control-value',
          ...EDIT,
          control: addressOf(opens(room, 'sdt')[0]!),
          value: { kind: 'date', iso },
        }) as const;
      apply(first.document, 'ben:1', pick(first.document, '2026-09-15'), 'Ben');
      const { document: doc, base } = fold
        ? await folded(first.document, first.base, 2, blobs)
        : first;
      apply(doc, 'olive:1', pick(doc, '2026-09-18'), 'Olive');
      const cut = { kind: 'delete', ...DIRECT, from: at(doc, 0, 11), to: at(doc, 0, 22) } as const;
      apply(doc, 'olive:2', cut, 'Olive');
      await checkEveryCard(doc, base, blobs, 'picks a direct deletion cut');
    });
  }

  // Olive's words filled the prompt first and were rejected; Ben then makes the prompt italic and
  // types in it. The control kept its prompt at the first fill, as the engine keeps it in the
  // glossary: a reject of Ben's typing shows that prompt, not his italic, live and folded, as the
  // editor alone does
  test('a format on a prompt shown again, then its maker typing in it', async () => {
    const items = '<w:listItem w:displayText="Net 30" w:value="30"/>';
    const prompt = t('Choose an item.', '<w:rStyle w:val="PlaceholderText"/>');
    const list =
      '<w:sdt><w:sdtPr><w:id w:val="3"/><w:showingPlcHdr/><w:dropDownList>' +
      `${items}</w:dropDownList></w:sdtPr><w:sdtContent>${prompt}</w:sdtContent></w:sdt>`;
    const file = docx(p(t('Terms are ') + list + t(' from today.')));
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    const blobs = mediaStore(base);
    const typing = (id: string, text: string, who: string) =>
      apply(doc, id, { kind: 'insert', ...EDIT, place: at(doc, 0, 24), text }, who);
    const reject = (id: string) => {
      const suggestions = orderedCardSeen(doc, id) ?? [];
      const decide = { kind: 'decide', cards: [{ suggestions }], status: 'rejected' } as const;
      apply(doc, 'olive:3', decide, 'Olive');
    };
    typing('olive:1', 'x', 'Olive');
    reject('S:olive:1.1');
    const italic = { kind: 'format', ...EDIT, properties: [{ localName: 'i' }] } as const;
    apply(doc, 'ben:1', { ...italic, from: at(doc, 0, 10), to: at(doc, 0, 24) }, 'Ben');
    typing('ben:2', 'a', 'Ben');
    await checkFold(doc, base, blobs, 'typing in a prompt formatted after a fill');
    reject('S:ben:2.1');
    expect(partsOf(doc)[0]).not.toContain('<w:i/>');
  });

  // Ben wraps words in a new drop-down list and deletes his copy of them: the control he inserted
  // is empty, so it shows its prompt inside his insertion, in the list as in the file, live and
  // after a fold, alone and in a row Ana struck
  for (const struck of [false, true]) {
    test(`a new control its maker emptied${struck ? " in another person's struck row" : ''}`, async () => {
      const cell =
        '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>' +
        p(t('Two per cent')) +
        '</w:tc>';
      const table =
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
        `<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell}</w:tr></w:tbl>`;
      const file = docx(p(t('Before.')) + table + p(t('After.')));
      const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
      if (struck) {
        apply(doc, 'ana:1', { kind: 'delete-row', ...EDIT, row: addressOf(opens(doc, 'tr')[0]!) });
      }

      const wrap = {
        from: at(doc, 1, 9),
        to: at(doc, 1, 11),
        control: { type: 'dropDownList' },
      } as const;
      apply(doc, 'ben:1', { kind: 'wrap-control', ...EDIT, ...wrap }, 'Ben');
      const copy = { kind: 'delete', ...EDIT, from: at(doc, 1, 11), to: at(doc, 1, 13) } as const;
      apply(doc, 'ben:2', copy, 'Ben');
      expect(textOf(doc, 1)).toBe('Two per cenChoose an item.t');
      expect(partsOf(doc)[0]).toMatch(/<w:ins w:author="Ben"[^>]*><w:sdt>/);
      await checkFold(doc, base, mediaStore(base), 'a new control its maker emptied');
    });
  }

  // The prompt a party typed over holds a combo box: after a fold the glossary gives it back, and a
  // reject of the typing shows the combo box in it again, not its words alone, and each of its
  // paragraphs. They are the prompt's own: typing in its first drops the second, as live
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

  // Olive's link round two letters of the prompt, kept by a fold, her bold on the prompt, then
  // her value: the prompt goes whole, the link and its words too. Before, a load marked the
  // link's edges as the prompt's, so the link read as a control in the prompt and its words with
  // her bold stayed beside the value
  test('a link round prompt words, formatted, then a value over the prompt', async () => {
    const prompt = t('Click here to enter text.', '<w:rStyle w:val="PlaceholderText"/>');
    const control =
      '<w:sdt><w:sdtPr><w:id w:val="4"/><w:showingPlcHdr/><w:text/></w:sdtPr>' +
      `<w:sdtContent>${prompt}</w:sdtContent></w:sdt>`;
    const file = docx(p(t('Signed by ') + control + t(' for the buyer.')));
    const opened = openOrderedWordFile(file, { file: 1 });
    const blobs = mediaStore(opened.base);
    const [from, to] = [at(opened.document, 0, 11), at(opened.document, 0, 13)];
    const href = 'https://example.com/fees';
    const link = { kind: 'insert-link', ...DIRECT, from, to, href } as const;
    apply(opened.document, 'olive:1', link, 'Olive');
    const { document: doc, base } = await folded(opened.document, opened.base, 2, blobs);
    const bold = { kind: 'format', ...EDIT, properties: [{ localName: 'b' }] } as const;
    apply(doc, 'olive:2', { ...bold, from: at(doc, 0, 10), to: at(doc, 0, 35) }, 'Olive');
    const box = addressOf(opens(doc, 'sdt')[0]!);
    const value = { kind: 'text', text: 'Party a' } as const;
    apply(doc, 'olive:3', { kind: 'control-value', ...EDIT, control: box, value }, 'Olive');
    expect(partsOf(doc).join('')).not.toContain('li</w:t>');
    await checkEveryCard(doc, base, blobs, 'a link round prompt words');
  });

  // Seed 83 of the fold shapes: Ben's comment ends inside the prompt, which a fold keeps in the
  // glossary with the comment's end and reference. A reject of Ana's typing over the prompt shows
  // the prompt with them again, as the engine's restore does, live as after a fold. A marker whose
  // comment the file does not hold never comes back: the block is the file's content
  describe("a comment's end inside a prompt kept in the glossary", () => {
    const terms = () => {
      const list =
        '<w:sdt><w:sdtPr><w:id w:val="3"/><w:showingPlcHdr/><w:dropDownList w:lastValue="30">' +
        '<w:listItem w:displayText="Net 30" w:value="30"/></w:dropDownList></w:sdtPr>' +
        `<w:sdtContent>${t('Choose an item.', '<w:rStyle w:val="PlaceholderText"/>')}` +
        '</w:sdtContent></w:sdt>';
      return docx(p(t('Terms are ') + list + t(' from today.')));
    };
    const reject = (room: OrderedDocument) =>
      ({
        kind: 'decide',
        cards: [{ suggestions: orderedCardSeen(room, 'S:ana:1.1') ?? [] }],
        status: 'rejected',
      }) as const;
    const typed = async () => {
      const { document: doc, base } = openOrderedWordFile(terms(), { file: 1 });
      const blobs = mediaStore(base);
      const comment = { kind: 'comment', part: BODY, text: 'Why?' } as const;
      apply(doc, 'ben:1', { ...comment, from: at(doc, 0, 1), to: at(doc, 0, 11) }, 'Ben');
      const copy = await folded(doc, base, 2, blobs);
      apply(copy.document, 'ana:1', {
        kind: 'insert',
        ...EDIT,
        place: at(copy.document, 0, 25),
        text: 'a',
      });
      return { ...copy, blobs };
    };

    test('shows again with the prompt', async () => {
      const { document: room, base, blobs } = await typed();
      await checkFold(room, base, blobs, "a comment's end in a kept prompt");

      // Folded again, the markers are in the glossary only. A reject taken back, as a refused
      // batch is, and made again under another id shows them, as the first does
      const { document: again } = await folded(room, base, 3, blobs);
      const first = { v: 1, id: 'olive:1', ...reject(again) } as const;
      const by = { person: testPerson('Olive'), date: DATE, ids: orderedIdMinter('olive:1') };
      takeBack(again, withUndoRecord(again, () => applyOrderedChange(again, first, by)).record);
      apply(again, 'olive:2', reject(again), 'Olive');
      expect(partsOf(again)[0]).toMatch(/<w:t>C<\/w:t><\/w:r><w:commentRangeEnd /);
    });

    test('of a comment the file does not hold never comes back', async () => {
      const { document: room, base, blobs } = await typed();
      const files = unzipSync(await checkpointBytes(room, base, blobs));
      const glossary = 'word/glossary/document.xml';
      files[glossary] = strToU8(
        strFromU8(files[glossary]!).replace(/(comment\w+ w:id=")\d+/g, '$1999')
      );
      const forged = openOrderedCheckpoint(zipSync(files), { file: 3 });
      apply(forged.document, 'olive:1', reject(forged.document), 'Olive');
      const bytes = orderedPackageBytes(forged.document, forged.base, { blobs });
      expect(readOoxmlPackage(bytes, {}).ok).toBe(true);
      expect(strFromU8(unzipSync(bytes)['word/document.xml']!)).not.toContain('w:id="999"');
    });
  });
});
