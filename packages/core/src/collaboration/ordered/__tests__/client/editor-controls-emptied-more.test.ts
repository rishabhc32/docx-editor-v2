// Typing into a content control that an edit or a decision emptied, in the cases the first
// settle left as they were, with real editors and one ordered room. Each step
// is compared with the editor alone, on every copy, and the last one with the saved file:
//  - a part of a control that Enter split: each part is a control of its own;
//  - a block control of several paragraphs: the engine keeps the first one, with the prompt,
//    and typing in an empty paragraph after it goes after it;
//  - Enter in a prompt that a decision showed, and its undo;
//  - Enter in a loaded prompt that showed again after a reject of a typing over it;
//  - a party's new control round words whose copy she deleted, alone and in a struck row;
//  - a party's pick in a drop-down, then its row struck and accepted: the prompt's glossary block.
// The room once refused the typing (`text-differs`), drew a prompt over the paragraphs Enter
// made, or left a control empty where the engine shows its prompt again: in a row another person
// struck, and with the party's own insertion taken off her control. Its file once lost the
// glossary block that keeps a filled prompt once the control went with its row.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { buildingBlocksOf } from '../../../../store/package/building-blocks.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import { docx, p, t, W } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { tableCommand } from './tables-harness.ts';

afterEach(closeAll);

const FACE = '<w:rStyle w:val="PlaceholderText"/>';
const REV = (id: number) => `w:id="${id}" w:author="Cara" w:date="2026-10-28T08:00:00Z"`;
const ins = (id: number, text: string) => `<w:ins ${REV(id)}>${t(text)}</w:ins>`;
const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const GLOSSARY = {
  parts: {
    'word/glossary/document.xml':
      `<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr>` +
      '<w:name w:val="PartyName"/><w:category><w:name w:val="General"/>' +
      '<w:gallery w:val="placeholder"/></w:category></w:docPartPr>' +
      `<w:docPartBody><w:p>${t('Enter the party.', FACE)}</w:p></w:docPartBody></w:docPart>` +
      '</w:docParts></w:glossaryDocument>',
  },
  contentTypes:
    '<Override PartName="/word/glossary/document.xml" ContentType="application/' +
    'vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>',
  relationships:
    '<Relationship Id="rIdGlossary" Type="http://schemas.openxmlformats.org/officeDocument/' +
    '2006/relationships/glossaryDocument" Target="glossary/document.xml"/>',
};

const MULTI = '<w:text w:multiLine="1"/>';
const split = (content: string) =>
  docx(p(t('Name: ') + sdt(MULTI, content) + t(' end.')) + p(t('After.')));
const block = (content: string, after = p(t('After.'))) =>
  docx(p(t('Before.')) + sdt('<w:id w:val="602"/>', content) + after);
const loaded = (named: boolean) =>
  docx(
    p(t('Name:')) +
      sdt(
        `${named ? '<w:placeholder><w:docPart w:val="PartyName"/></w:placeholder>' : ''}` +
          '<w:showingPlcHdr/>',
        p(t('Click or tap here.', FACE))
      ) +
      p(t('After.')),
    named ? GLOSSARY : {}
  );

/** One step: who takes it, in which mode, and what. */
interface Step {
  readonly who: 'olive' | 'ana';
  readonly mode: 'edit' | 'suggesting';
  readonly act:
    | 'type'
    | 'pick'
    | 'delete'
    | 'enter'
    | 'wrap'
    | 'deleteRow'
    | 'accept'
    | 'reject'
    | 'undo'
    | 'redo';
  readonly line?: number;
  readonly at?: number;
  readonly to?: number;
}
const olive = (act: Step['act'], line?: number, at?: number, to?: number): Step => ({
  who: 'olive',
  mode: 'edit',
  act,
  ...(line === undefined ? {} : { line }),
  ...(at === undefined ? {} : { at }),
  ...(to === undefined ? {} : { to }),
});
const suggested = (step: Step): Step => ({ ...step, mode: 'suggesting' });
const ana = (act: Step['act'], line?: number, at?: number, to?: number): Step =>
  suggested({ ...olive(act, line, at, to), who: 'ana' });

function run(person: Person, { mode, act, line = 0, at = 0, to }: Step): void {
  person.editor.setMode(mode);
  const s = surface(person);
  if (act === 'type') {
    caretAt(person, line, at);
    s.insertPlainText('Beta');
  } else if (act === 'delete') {
    caretAt(person, line, at, to);
    s.deleteBackward();
  } else if (act === 'enter') {
    caretAt(person, line, at);
    s.splitParagraph();
  } else if (act === 'pick') {
    caretAt(person, line, at);
    const pick = { type: 'setContentControlValue', value: '60' };
    expect(person.editor.exec(pick as never).ok).toBe(true);
  } else if (act === 'wrap') {
    caretAt(person, line, at, to);
    expect(person.editor.exec({ type: 'insertContentControl', subtype: 'plainText' }).ok).toBe(
      true
    );
  } else if (act === 'deleteRow') {
    caretAt(person, line, at);
    expect(tableCommand(person, { type: 'deleteRow' }).ok).toBe(true);
  } else if (act === 'accept' || act === 'reject') {
    const command = { type: 'resolveAllReviewChanges', action: act, scope: 'document' };
    expect(person.editor.exec(command as never).ok).toBe(true);
  } else if (act === 'undo') s.undo();
  else s.redo();
}

/** The body as a reader compares it: no ids, dates or paragraph ids; a new control's id too. */
const bodyOf = (xml: string): string =>
  xml
    .slice(xml.indexOf('<w:body>'), xml.indexOf('<w:sectPr'))
    .replace(/ w14:(paraId|textId)="[^"]*"/g, '')
    .replace(/ w:(id|date)="[^"]*"/g, '')
    .replace(/<w:id w:val="\d+"\/>/g, '');

/** The body after each step on the editor alone, the organiser's and the party's steps in turn. */
function aloneSteps(file: Uint8Array, steps: readonly Step[]): string[] {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: file,
    author: 'Olive',
    mode: 'edit',
    modules: [reviewTestModule()],
  });
  try {
    const person = { name: 'alone', editor, container, refusals: [] } as unknown as Person;
    return steps.map((step) => {
      editor.setAuthor(step.who === 'ana' ? 'Ana' : 'Olive');
      run(person, step);
      return bodyOf(editorXml(person));
    });
  } finally {
    editor.destroy();
    container.remove();
  }
}

/** The names of the glossary's blocks in a Word file. */
function blocksOf(bytes: Uint8Array): string[] {
  const read = readOoxmlPackage(bytes, {});
  if (!read.ok) throw new Error(read.reason);
  return buildingBlocksOf(read.package).map((block) => block.name);
}

async function savedBody(person: Person): Promise<string> {
  const saved = new Uint8Array(await person.editor.save());
  const again = openOrderedWordFile(saved, { file: 2 }).document;
  return bodyOf(serializeOoxmlPart(composeOrderedPart(again, BODY).part));
}

async function inRoom(file: Uint8Array, steps: readonly Step[]): Promise<void> {
  const expected = aloneSteps(file, steps);
  const review = reviewTestModule();
  const room = localRoom(file);
  const people = {
    olive: await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: file, review }),
    ana: await join(room, 'Ana', { wordFile: file, review }),
  };
  const shown: string[] = [];
  for (const [k, step] of steps.entries()) {
    run(people[step.who], step);
    await settle();
    for (const person of [people.olive, people.ana]) {
      const state = {
        step: k,
        person: person.name,
        refusals: person.refusals,
        room: sameAsRoom(person, room),
        screen: screenFollowsList(person),
      };
      expect(state).toEqual({ ...state, refusals: [], room: true, screen: true });
    }
    shown.push(bodyOf(editorXml(people.ana)));
  }
  expect(shown).toEqual(expected);
  expect(await savedBody(people.ana)).toBe(expected.at(-1) ?? '');
}

describe('a part of a control that Enter split, emptied, then typed into', () => {
  test("the organiser's direct Enter, the second part emptied directly, typed, undone", () =>
    inRoom(split(t('Acme Corp')), [
      olive('enter', 0, 10),
      olive('delete', 1, 0, 5),
      olive('type', 1, 0),
      olive('undo'),
      olive('undo'),
      olive('undo'),
    ]));

  test('the first part emptied directly, typed, and every step undone', () =>
    inRoom(split(t('Acme Corp')), [
      olive('enter', 0, 10),
      olive('delete', 0, 6, 10),
      olive('type', 0, 6),
      olive('undo'),
      olive('undo'),
      olive('undo'),
    ]));

  test("a decision empties the second part, and a party's typing there is a suggestion", () =>
    inRoom(split(t('Acme') + ins(6, ' Corp')), [
      olive('enter', 0, 10),
      olive('reject'),
      ana('type', 1, 0),
    ]));
});

describe('the Enter that split a control taken back while a part shows its prompt', () => {
  test("a party's Enter, the organiser empties the second part, the party takes the Enter back", async () => {
    // Two people, so the editor alone has no such sequence: the parts are one control again,
    // and it holds the first part's words, with no prompt, on every copy and in the save
    const file = split(t('Acme Corp'));
    const room = localRoom(file);
    const review = reviewTestModule();
    const organiser = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      wordFile: file,
      review,
    });
    const party = await join(room, 'Ana', { wordFile: file, review });
    run(party, ana('enter', 0, 10));
    await settle();
    run(organiser, olive('delete', 1, 0, 5));
    await settle();
    expect(editorTexts(party)).toEqual([
      'Name: Acme',
      'Click or tap here to enter text. end.',
      'After.',
    ]);
    run(party, ana('undo'));
    await settle();
    for (const person of [organiser, party]) {
      expect(person.refusals).toEqual([]);
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
      expect(editorTexts(person)).toEqual(['Name: Acme end.', 'After.']);
    }
    const saved = await savedBody(party);
    expect(saved.match(/<w:sdt>/g)).toHaveLength(1);
    expect(saved).not.toContain('Click or tap');
  });
});

describe('a block control of several paragraphs, emptied, then typed into', () => {
  test("a reject of all its words, then a party's typing", () =>
    inRoom(block(p(ins(6, 'Acme')) + p(ins(7, 'Corp'))), [
      olive('reject'),
      ana('type', 1, 0),
      ana('undo'),
    ]));

  test('direct deletions of each paragraph, a typing, and every step undone', () =>
    inRoom(
      block(
        p(t('Acme'), '<w:pPr><w:jc w:val="center"/></w:pPr>') +
          p(t('Corp')) +
          p(t('Inc'), '<w:pPr><w:ind w:left="720"/></w:pPr>')
      ),
      [
        olive('delete', 1, 0, 4),
        olive('delete', 2, 0, 4),
        olive('delete', 3, 0, 3),
        ana('type', 1, 0),
        ana('undo'),
        olive('undo'),
        olive('undo'),
        olive('undo'),
      ]
    ));
});

// Before: the paragraph after the control took the items of the paragraphs it dropped as
// its own, and typing in it went into the control, which then showed no prompt
describe('typing in an empty paragraph after a block control that dropped paragraphs', () => {
  const EMPTY_AFTER = '<w:p/>' + p(t('End.'));

  test("a reject of all its words, then a party's typing and its undo", () =>
    inRoom(block(p(ins(6, 'Acme')) + p(ins(7, 'Corp')), EMPTY_AFTER), [
      olive('reject'),
      ana('type', 2, 0),
      ana('undo'),
    ]));

  test("direct deletions of each paragraph, a party's typing, and every step undone", () =>
    inRoom(block(p(t('Acme')) + p(t('Corp')), EMPTY_AFTER), [
      olive('delete', 1, 0, 4),
      olive('delete', 2, 0, 4),
      ana('type', 2, 0),
      ana('undo'),
      olive('undo'),
      olive('undo'),
    ]));
});

describe('Enter in a prompt a decision showed', () => {
  test('a block control: Enter, typing, and both undone show the prompt again', () =>
    inRoom(block(p(ins(6, 'Acme'))), [
      olive('reject'),
      olive('enter', 1, 2),
      olive('type', 1, 0),
      olive('undo'),
      olive('undo'),
    ]));

  for (const named of [true, false])
    test(`a loaded prompt ${named ? 'named in the glossary' : 'of its own'}: a reject of the typing over it, then Enter and its undo`, () =>
      inRoom(loaded(named), [
        suggested(olive('type', 1, 5)),
        olive('reject'),
        olive('enter', 1, 5),
        olive('undo'),
      ]));
});

describe("a control in another person's pending insertion, emptied directly", () => {
  // The editor alone drops the person's insertion around the control whose words the organiser
  // deleted directly: the control is no longer a suggestion, shows its prompt and takes typing
  test("the party's new control, its words deleted by the organiser, typed into, undone", () =>
    inRoom(docx(p(t('Name: Acme Corp end.')) + p(t('After.'))), [
      ana('wrap', 0, 6, 15),
      olive('delete', 0, 15, 24),
      olive('type', 0, 15),
      olive('undo'),
      olive('undo'),
    ]));

  test("the party's new control emptied by the organiser, the party's suggestion accepted, typed into", () =>
    inRoom(docx(p(t('Name: Acme Corp end.')) + p(t('After.'))), [
      ana('wrap', 0, 6, 15),
      olive('delete', 0, 15, 24),
      olive('accept'),
      olive('type', 0, 6),
    ]));
});

describe("a party's new control round words, its copy then deleted by her", () => {
  // The control she inserted holds nothing then: it shows its prompt inside her insertion, and
  // her typing replaces the prompt, as in the editor alone
  test('in a paragraph', () =>
    inRoom(docx(p(t('Name: Acme Corp')) + p(t('After.'))), [
      ana('wrap', 0, 6, 10),
      ana('delete', 0, 10, 14),
      ana('type', 0, 10),
    ]));

  test('in a row the organiser struck', () => {
    const cell = `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t('Acme Corp'))}</w:tc>`;
    const table =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
      `<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell}</w:tr></w:tbl>`;
    return inRoom(docx(p(t('Name:')) + table + p(t('After.'))), [
      suggested(olive('deleteRow', 1, 0)),
      ana('wrap', 1, 0, 4),
      ana('delete', 1, 4, 8),
      ana('type', 1, 4),
    ]);
  });
});

describe("a prompt's glossary block, once its control goes with its row", () => {
  // Ana's pick hides the drop-down's prompt, so the engine keeps the prompt as a glossary block,
  // which stays when the organiser's accept takes the row away with the control
  test("a party's pick, then her deletion of the row, accepted", async () => {
    const list =
      '<w:sdt><w:sdtPr><w:id w:val="604"/><w:showingPlcHdr/><w:dropDownList>' +
      '<w:listItem w:displayText="Net 60" w:value="60"/></w:dropDownList></w:sdtPr>' +
      `<w:sdtContent>${t('Choose an item.', FACE)}</w:sdtContent></w:sdt>`;
    const cell = (inner: string) =>
      `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(inner)}</w:tc>`;
    const table =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>' +
      '<w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr>${cell(t('Terms'))}${cell(t('Net: ') + list)}</w:tr></w:tbl>`;
    const file = docx(p(t('Before.')) + table + p(t('After.')));
    const steps = [ana('pick', 2, 7), ana('deleteRow', 2, 0), olive('accept')];
    await inRoom(file, steps);

    const container = document.createElement('div');
    document.body.append(container);
    const review = reviewTestModule();
    const editor = createDocxEditor({ container, document: file, mode: 'edit', modules: [review] });
    const alone = { name: 'alone', editor, container, refusals: [] } as unknown as Person;
    const room = localRoom(file);
    const people = {
      olive: await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: file, review }),
      ana: await join(room, 'Ana', { wordFile: file, review }),
    };
    try {
      for (const step of steps) {
        editor.setAuthor(step.who === 'ana' ? 'Ana' : 'Olive');
        run(alone, step);
        run(people[step.who], step);
        await settle();
      }

      const kept = blocksOf(new Uint8Array(await editor.save()));
      expect(kept).toHaveLength(1);
      const { base } = openOrderedWordFile(file, { file: 1 });
      expect(blocksOf(orderedPackageBytes(room.room.document, base))).toEqual(kept);
    } finally {
      editor.destroy();
      container.remove();
    }
  });
});
