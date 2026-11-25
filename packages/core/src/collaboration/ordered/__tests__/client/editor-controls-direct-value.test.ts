// The organiser's direct value in a control that a pending suggestion put in, with real editors
// and one ordered room, against the editor alone: in their own suggested wrap (undone and redone
// too) and in another person's pending control. Each is accepted and rejected live, after a save
// and reopen, and after one and two folds.
//
// A direct edit inside a pending insertion is typed into it as its author's (design D9), as Word
// extends a tracked insertion with untracked typing: the value is part of the control's insertion,
// so a reject takes the control with its value, and an accept keeps both. The room once kept the
// value as settled text in the control and the wrapped words hidden in it: a reject live left the
// value as plain text, where the file, the fold and the editor alone left the words as they were.
//
// A party's suggested value in its own suggested wrap is part of the wrap's insertion too, as the
// editor alone writes one insertion: the room once made it a second suggestion, a `w:ins` of the
// same author inside the wrap's. Its undo takes back the value alone, and the next one the wrap,
// as the editor alone undoes them: the room once took back the wrap and the value in one step.

import { afterEach, describe, expect, test } from 'bun:test';
import { docx, p, t } from '../word-fixtures.ts';
import { decidedCheck, type Act } from './decided-harness.ts';
import { copiedBetween } from './paste-harness.ts';
import { caretAt, closeAll, surface } from './editor-harness.ts';

afterEach(closeAll);

const LINE = docx(p(t('Line: Revenue')) + p(t('Signed.')));

const ITEMS =
  '<w:listItem w:displayText="Net 30" w:value="30"/>' +
  '<w:listItem w:displayText="Net 60" w:value="60"/>';
const DATE_FORMAT = '<w:dateFormat w:val="M/d/yyyy"/><w:lid w:val="en-US"/>';
const CHECK_BOX =
  '<w14:checkbox><w14:checked w14:val="0"/>' +
  '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>' +
  '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>';

/** "Term: <control> end." with the control inserted by Ana's pending suggestion. */
const anasControl = (props: string, content: string) =>
  docx(
    p(
      t('Term: ') +
        '<w:ins w:id="1" w:author="Ana" w:date="2026-10-28T07:00:00Z">' +
        `<w:sdt><w:sdtPr><w:id w:val="5"/>${props}</w:sdtPr><w:sdtContent>${content}` +
        '</w:sdtContent></w:sdt></w:ins>' +
        t(' end.')
    ) + p(t('Signed.'))
  );

/**
 * Each wrap: its label, the control type, and the value set in it. A date picker the wrap makes
 * has no `w:dateFormat`: the room once wrote its date as `2026-09-15` where the editor alone wrote
 * `9/15/2026`, Word's en-US default.
 */
const WRAPS = [
  ['a date picker', 'date', '2026-09-15'],
  ['a combo box', 'comboBox', 'Net 30'],
  ['a plain text control', 'plainText', 'Acme'],
  ['a rich text control', 'richText', 'Acme'],
] as const;

/** Each control of Ana's: its label, the Word file, and the value set in it. */
const ANAS = [
  ['a date picker', anasControl(`<w:date>${DATE_FORMAT}</w:date>`, t('Soon')), '2026-09-15'],
  ['a drop-down list', anasControl(`<w:dropDownList>${ITEMS}</w:dropDownList>`, t('Net 30')), '60'],
  ['a combo box', anasControl(`<w:comboBox>${ITEMS}</w:comboBox>`, t('Net 30')), 'Net 90'],
  ['a check box', anasControl(CHECK_BOX, t('☐')), 'true'],
  ['a plain text control', anasControl('<w:text/>', t('Acme')), 'Beta'],
  ['a rich text control', anasControl('', t('Acme')), 'Beta'],
] as const;

/**
 * Each control a party pastes as its own suggestion, as a wrap makes none of these kinds with a
 * value to choose: its label, the Word file it is copied from, and the value set in it.
 */
const PASTED = [
  ['a pasted check box', CHECK_BOX, '☐', 'true'],
  ['a pasted drop-down list', `<w:dropDownList>${ITEMS}</w:dropDownList>`, 'Net 30', '60'],
] as const;

/** A Word file of one line, "[<control>].". */
const copySource = (props: string, content: string) =>
  docx(
    p(
      t('[') +
        `<w:sdt><w:sdtPr><w:id w:val="7"/>${props}</w:sdtPr><w:sdtContent>${t(content)}` +
        '</w:sdtContent></w:sdt>' +
        t('].')
    )
  );

/** Wrap "Revenu" in a new control of `type`, as a suggestion. */
const wrap =
  (type: string): Act =>
  (person) => {
    person.editor.setMode('suggesting');
    caretAt(person, 0, 6, 12);
    const command = { type: 'insertContentControl', subtype: type, tag: 'clause.86' };
    expect(person.editor.exec(command as never).ok).toBe(true);
  };

/** Paste "[<control>]", copied from `copySource`, before "Revenue", as a suggestion. */
const paste =
  (props: string, content: string): Act =>
  (person) => {
    person.editor.setMode('suggesting');
    caretAt(person, 0, 6);
    const copy = copiedBetween(0, 0, content.length + 2, copySource(props, content));
    expect(surface(person).pasteRich(copy.text, copy.html)).toBe(true);
  };

/** Set the first control's value directly. */
const value =
  (text: string): Act =>
  (person) => {
    person.editor.setMode('edit');
    const id = surface(person).layout().contentControls?.[0]?.id ?? '';
    expect(surface(person).contentControls.setValue(id, text)).toBe(true);
  };

/** Set the first control's value as a suggestion. */
const suggestedValue =
  (text: string): Act =>
  (person) => {
    const id = surface(person).layout().contentControls?.[0]?.id ?? '';
    expect(surface(person).contentControls.setValue(id, text)).toBe(true);
  };

const undo: Act = (person) => surface(person).undo();
const redo: Act = (person) => surface(person).redo();

/** The first paragraph as a reader of the file sees it: no ids, dates or revision sessions. */
const lineOf = (xml: string): string =>
  (/<w:p\b.*?<\/w:p>/s.exec(/<w:body>(.*)<\/w:body>/s.exec(xml)?.[1] ?? '')?.[0] ?? 'no line')
    .replace(/ w:(id|date|rsidR|rsidRPr|rsidDel)="[^"]*"| w14:(paraId|textId)="[^"]*"/g, '')
    .replace(/<w:id w:val="\d+"\/>/g, '')
    .replace(/<w:placeholder>.*?<\/w:placeholder>/g, '')
    .replace(/ xml:space="preserve"|<w:rPr>.*?<\/w:rPr>/g, '')
    .replace(/<\/w:r><w:r>/g, '')
    .replace(/<\/w:t><w:t>/g, '');

const check = decidedCheck(lineOf);

describe("the organiser's direct value in their own suggested wrap, as the editor alone", () => {
  for (const [label, type, text] of WRAPS) {
    for (const decision of ['accept', 'reject'] as const) {
      test(`${label}, ${decision}ed`, async () => {
        await check(label, LINE, [wrap(type), value(text)], decision);
      });
    }
  }
});

describe("a party's suggested value in its own suggested wrap, as the editor alone", () => {
  for (const [label, type, text] of WRAPS) {
    for (const decision of ['accept', 'reject'] as const) {
      test(`${label}, ${decision}ed`, async () => {
        await check(label, LINE, [wrap(type), suggestedValue(text)], decision, 'ben');
      });
    }
  }
});

describe("the organiser's direct value in their own suggested wrap, undone, as alone", () => {
  const steps = [wrap('plainText'), value('Acme')];
  for (const [label, more] of [
    ['undone', [undo]],
    ['undone and redone', [undo, redo]],
  ] as const) {
    for (const decision of ['accept', 'reject'] as const) {
      test(`${label}, ${decision}ed`, async () => {
        await check(label, LINE, [...steps, ...more], decision);
      });
    }
  }
});

describe("a party's suggested value in its own pending control, undone and redone, as alone", () => {
  const made = [
    ...WRAPS.map(([label, type, text]) => [label, [wrap(type), suggestedValue(text)]] as const),
    ...PASTED.map(
      ([label, props, content, text]) =>
        [label, [paste(props, content), suggestedValue(text)]] as const
    ),
  ];
  for (const [label, steps] of made) {
    for (const [then, more] of [
      ['undone', [undo]],
      ['undone twice', [undo, undo]],
      ['undone and redone', [undo, redo]],
      ['undone twice and redone', [undo, undo, redo]],
      ['undone twice and redone twice', [undo, undo, redo, redo]],
    ] as const) {
      for (const decision of ['accept', 'reject'] as const) {
        test(`${label}, ${then}, ${decision}ed`, async () => {
          await check(`${label}, ${then}`, LINE, [...steps, ...more], decision, 'ben');
        });
      }
    }
  }
});

describe("the organiser's direct value in another person's pending control, as alone", () => {
  for (const [label, file, text] of ANAS) {
    for (const decision of ['accept', 'reject'] as const) {
      test(`${label}, ${decision}ed`, async () => {
        await check(label, file, [value(text)], decision);
      });
    }
  }
});
