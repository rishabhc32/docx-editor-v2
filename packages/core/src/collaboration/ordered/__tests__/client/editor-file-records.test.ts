// A file's own row exceptions record (`w:tblPrExChange`), first-edition numbering record
// (`w:numberingChange`) and tracked list membership (`w:numPr/w:ins`) in an ordered room (task
// 9.8). Word for Mac shows, accepts and rejects the first two (2026-10-29), and the engine decides
// all three, so the organiser decides them through the room as the editor alone does, also after
// a save and reopen, and a party's decision is refused by the party rule (`policy-decide`).

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import type { EditorCommand } from '../../../../contracts/editor.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  join,
  localRoom,
  ORGANISER,
  settle,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { converged, saved } from './sections-fixtures.ts';

afterEach(closeAll);

const border = (value: string) =>
  `<w:tblBorders><w:top w:val="${value}" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>`;
const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(t(text))}</w:tc>`;

/** Row two's borders changed from single to double, as tblprex2.docx records it. */
const ROW_RECORD = docx(
  p(t('Before the table.')) +
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${border('single')}</w:tblPr>` +
    '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    `<w:tr>${cell('Row one A')}${cell('Row one B')}</w:tr>` +
    `<w:tr><w:tblPrEx>${border('double')}` +
    '<w:tblPrExChange w:id="5" w:author="Cara" w:date="2026-01-01T00:00:00Z">' +
    `<w:tblPrEx>${border('single')}</w:tblPrEx></w:tblPrExChange></w:tblPrEx>` +
    `${cell('Row two A')}${cell('Row two B')}</w:tr></w:tbl>` +
    p(t('After the table.'))
);

const NUMBERING =
  '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>' +
  '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
const LIST = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';

/** A list whose item one holds `record` in its `w:numPr`. */
const listWith = (record: string) =>
  docx(
    p(t('Before the list.')) +
      p(
        t('Numbered item one.'),
        `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/>${record}</w:numPr>`
      ) +
      p(t('Numbered item two.'), LIST) +
      p(t('After the list.')),
    {
      parts: { 'word/numbering.xml': NUMBERING },
      contentTypes:
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
      relationships:
        '<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
    }
  );

/** Item one's numbering record, as numchange.docx holds it. */
const NUMBERING_RECORD = listWith(
  '<w:numberingChange w:id="7" w:author="Cara" w:date="2026-01-01T00:00:00Z" w:original="%1)"/>'
);

/** Item one's tracked list membership. */
const LIST_INSERTED = listWith('<w:ins w:id="8" w:author="Cara" w:date="2026-01-01T00:00:00Z"/>');

const RECORDS: readonly (readonly [string, Uint8Array, RegExp])[] = [
  ['a row exceptions record', ROW_RECORD, /<w:tblPrExChange w:author="Cara"[^>]*w:id="5"/],
  ['a numbering record', NUMBERING_RECORD, /<w:numberingChange w:author="Cara"[^>]*w:id="7"/],
  ['a list membership', LIST_INSERTED, /<w:numId w:val="1"\/><w:ins w:author="Cara"[^>]*w:id="8"/],
];

async function roomOn(file: Uint8Array) {
  const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
  const review = reviewTestModule();
  const ana = await join(room, 'Ana', { wordFile: file, review });
  const olive = await join(room, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    wordFile: file,
    review,
  });
  return { room, ana, olive };
}

function onlyCard(person: Person): string {
  const cards = person.editor.getReviewItems().filter((item) => item.kind === 'revision');
  expect(cards).toHaveLength(1);
  return cards[0]!.key;
}

const body = async (person: Person) => (await saved(person))['word/document.xml'] ?? '';

/** Item one's paragraph properties in a saved main part. */
const itemOne = (xml: string) =>
  /<w:pPr>(?:(?!<w:pPr>)[\s\S])*?<\/w:pPr>(?=(?:(?!<w:p[ >])[\s\S])*Numbered item one\.)/.exec(
    xml
  )?.[0] ?? '';

/** The editor alone on `file`, the organiser's command on item one: its save and cards. */
async function alone(file: Uint8Array, command: EditorCommand) {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: file, author: 'Olive', mode: 'edit' });
  try {
    const id = editor.surface!.session.paragraphIds()[1]!;
    editor.surface!.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 0 },
    });
    expect(editor.exec(command).ok).toBe(true);
    const files = unzipSync(new Uint8Array(await editor.save()));
    return strFromU8(files['word/document.xml'] ?? new Uint8Array());
  } finally {
    editor.destroy();
    container.remove();
  }
}

/** Row two's part of a saved main part. */
const rowTwo = (xml: string) => xml.slice(xml.indexOf('Row one B'), xml.indexOf('Row two B'));

describe("a file's own records that Word shows and decides", () => {
  test('a row exceptions record: the organiser accepts or rejects it, as the editor alone does', async () => {
    for (const decision of ['accept', 'reject'] as const) {
      const { room, ana, olive } = await roomOn(ROW_RECORD);
      expect(await body(olive)).toContain('<w:tblPrExChange ');
      const card = onlyCard(olive);
      const done =
        decision === 'accept'
          ? olive.editor.acceptReviewItem(card)
          : olive.editor.rejectReviewItem(card);
      expect(done.ok).toBe(true);
      await settle();
      converged(room, ana, olive);
      const row = rowTwo(await body(ana));
      expect(row).not.toContain('tblPrExChange');
      expect(row).toContain(decision === 'accept' ? 'w:val="double"' : 'w:val="single"');
      // The engine's reject gives the row's cells with no record of their own default
      // formatting (`restoreUnrecordedCells`), as Word does
      expect(row).toContain(decision === 'accept' ? 'w:type="dxa"' : 'w:type="auto"');
      closeAll();
    }
  });

  test('a numbering record or a list membership: accepted or rejected, it goes and the list stays', async () => {
    for (const [file, decision] of [NUMBERING_RECORD, LIST_INSERTED].flatMap((f) =>
      (['accept', 'reject'] as const).map((d) => [f, d] as const)
    )) {
      const { room, ana, olive } = await roomOn(file);
      expect(await body(olive)).toMatch(/<w:(numberingChange|ins) /);
      const card = onlyCard(olive);
      const done =
        decision === 'accept'
          ? olive.editor.acceptReviewItem(card)
          : olive.editor.rejectReviewItem(card);
      expect(done.ok).toBe(true);
      await settle();
      converged(room, ana, olive);
      const xml = await body(ana);
      expect(xml).not.toMatch(/numberingChange|<w:ins /);
      expect(xml.split(LIST)).toHaveLength(3);
      closeAll();
    }
  });

  test("a party's decision is refused by the party rule, and the record stays", async () => {
    for (const [, file, record] of RECORDS) {
      const { room, ana, olive } = await roomOn(file);
      ana.editor.acceptReviewItem(onlyCard(ana));
      await settle();
      expect(ana.refusals).toEqual([
        expect.objectContaining({ code: 'ordered-refused', reason: 'policy-decide' }),
      ]);
      ana.refusals.length = 0;
      converged(room, ana, olive);
      expect(onlyCard(olive)).toBeString();
      expect(await body(olive)).toMatch(record);
      closeAll();
    }
  });

  test('saved while pending, the record reopens as the file had it, and the organiser decides it', async () => {
    for (const [name, file, record] of RECORDS) {
      const first = await roomOn(file);
      const bytes = new Uint8Array(await first.olive.editor.save());
      expect([name, (await body(first.olive)).match(record) !== null]).toEqual([name, true]);
      closeAll();
      const { room, ana, olive } = await roomOn(bytes);
      expect([name, (await body(ana)).match(record) !== null]).toEqual([name, true]);
      expect(olive.editor.rejectReviewItem(onlyCard(olive)).ok).toBe(true);
      await settle();
      converged(room, ana, olive);
      expect([name, /tblPrExChange|numberingChange|<w:ins /.test(await body(ana))]).toEqual([
        name,
        false,
      ]);
      closeAll();
    }
  });

  test('a list edit over a pending numbering record keeps, drops or copies it as the editor alone', async () => {
    const commands: readonly EditorCommand[] = [
      { type: 'adjustIndent', direction: 'increase' },
      { type: 'toggleList', kind: 'ordered' },
    ];
    for (const [file, command] of [NUMBERING_RECORD, LIST_INSERTED].flatMap((f) =>
      commands.map((c) => [f, c] as const)
    )) {
      const expected = itemOne(await alone(file, command));
      // A new level keeps the record; leaving the list takes `w:numPr` and the record with it
      if (command.type === 'adjustIndent')
        expect(expected).toMatch(/<w:ilvl w:val="1"\/>.*<w:(ins|numberingChange) /);
      else expect(expected).not.toContain('numPr');
      const { room, ana, olive } = await roomOn(file);
      caretAt(olive, 1, 0);
      expect(olive.editor.exec(command).ok).toBe(true);
      await settle();
      converged(room, ana, olive);
      expect([command.type, itemOne(await body(ana))]).toEqual([command.type, expected]);
      closeAll();
    }
  });

  test("a party's list change over another person's pending numbering record is refused", async () => {
    const commands: readonly EditorCommand[] = [
      { type: 'adjustIndent', direction: 'increase' },
      { type: 'toggleList', kind: 'ordered' },
    ];
    for (const file of [NUMBERING_RECORD, LIST_INSERTED]) {
      const { room, ana, olive } = await roomOn(file);
      // On item one it would drop Cara's record, or copy it under Ana's: her suggestion would go
      for (const command of commands) {
        caretAt(ana, 1, 0);
        ana.editor.exec(command);
        await settle();
        expect([command.type, ana.refusals.splice(0).map((r) => `${r.code}:${r.reason}`)]).toEqual([
          command.type,
          ['ordered-refused:others_suggestions'],
        ]);
        converged(room, ana, olive);
      }
      expect(await body(olive)).toMatch(/<w:(numberingChange|ins) w:author="Cara"/);
      // Item two holds no record: Ana's level change there is a suggestion of her own
      caretAt(ana, 2, 0);
      expect(ana.editor.exec(commands[0]!).ok).toBe(true);
      await settle();
      converged(room, ana, olive);
      expect(await body(olive)).toMatch(/<w:pPrChange w:author="Ana"/);
      // The organiser's direct change over the record is not affected
      caretAt(olive, 1, 0);
      expect(olive.editor.exec(commands[0]!).ok).toBe(true);
      await settle();
      converged(room, ana, olive);
      expect(itemOne(await body(olive))).toMatch(/<w:ilvl w:val="1"\/>.*<w:(ins|numberingChange) /);
      closeAll();
    }
  });
});
