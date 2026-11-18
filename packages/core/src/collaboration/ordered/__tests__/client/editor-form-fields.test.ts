// Legacy form fields with real editors and one ordered room (task 9.6), in a file protected for
// filling in forms, as Word fills them: a check box, a drop-down, a text field's default and its
// result formatted on leaving it, and typing in its result. Each equals the editor alone, every
// copy follows, and the organiser's undo and redo are exact. Forms protection turns tracking off
// in Word, so a party's edit is a direct edit, refused by the host's party rule.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { stubReviewModule } from '../../../../editor/__tests__/review-test-module.ts';
import { findNode } from '../../../../store/package/ooxml-edit.ts';
import { legacyCheckboxFieldsOf } from '../../../../store/store/legacy-checkbox-fields.ts';
import { legacyDropdownFieldsOf } from '../../../../store/store/legacy-dropdown-fields.ts';
import { textFormFieldsOf } from '../../../../store/store/text-form-fields.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t, W, WML_CT } from '../word-fixtures.ts';
import {
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

afterEach(closeAll);

const formField = (instr: string, data: string, result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="F${instr.length}"/><w:enabled/>` +
  `<w:calcOnExit w:val="0"/>${data}</w:ffData></w:fldChar></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

const BODY_XML =
  p(t('Name: ') + formField('FORMTEXT', '<w:textInput/>', t('Bob')) + t(' end.')) +
  p(
    t('Box: ') +
      formField(
        'FORMCHECKBOX',
        '<w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox>',
        ''
      ) +
      t(' end.')
  ) +
  p(
    t('Pick: ') +
      formField(
        'FORMDROPDOWN',
        '<w:ddList><w:listEntry w:val="One"/><w:listEntry w:val="Two"/></w:ddList>',
        t('One')
      ) +
      t(' end.')
  );

/** The fields, unprotected: a field's definition may change. */
const PLAIN = docx(BODY_XML);

/** The fields, protected for filling in forms, as Word fills them. */
const FILE = docx(BODY_XML, {
  parts: {
    'word/settings.xml': `<w:settings xmlns:w="${W}"><w:documentProtection w:edit="forms" w:enforcement="1"/></w:settings>`,
  },
  contentTypes: `<Override PartName="/word/settings.xml" ContentType="${WML_CT}.settings+xml"/>`,
  relationships:
    '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>',
});

const bodyOf = (xml: string): string =>
  xml
    .slice(xml.indexOf('<w:body>'), xml.indexOf('</w:body>'))
    .replace(/ w14:(paraId|textId)="[^"]*"/g, '');

function apply(person: Person, op: TreeDocOp): void {
  const s = surface(person) as unknown as {
    applyGatedTreeOps(ops: readonly TreeDocOp[]): unknown;
  };
  s.applyGatedTreeOps([op]);
}

/** The editor's paragraph and the field of a kind in it. */
function field(person: Person, index: number, kind: 'text' | 'checkbox' | 'dropdown') {
  const session = surface(person).session;
  const paragraphId = session.paragraphIds()[index] ?? '';
  const node = findNode(session.part(), paragraphId);
  if (!node || node.kind !== 'paragraph') throw new Error('no paragraph');
  const found =
    kind === 'text'
      ? textFormFieldsOf(node)[0]
      : kind === 'checkbox'
        ? legacyCheckboxFieldsOf(node)[0]
        : legacyDropdownFieldsOf(node)[0];
  if (!found) throw new Error(`no ${kind} field`);
  return { paragraphId, fieldNodeId: found.fieldNodeId, start: found.start };
}

type Edit = (person: Person) => void;

/** Edits, and the file each runs on (a definition changes only without forms protection). */
const FILES: Readonly<Record<string, Uint8Array>> = {
  "a text field's default and options": PLAIN,
};

const EDITS: Readonly<Record<string, Edit>> = {
  'a check box on': (person) =>
    apply(person, { op: 'setLegacyCheckbox', ...field(person, 1, 'checkbox'), checked: true }),
  'a drop-down entry': (person) =>
    apply(person, { op: 'setLegacyDropdown', ...field(person, 2, 'dropdown'), selectedIndex: 1 }),
  "a text field's default and options": (person) => {
    const { paragraphId, fieldNodeId } = field(person, 0, 'text');
    apply(person, {
      op: 'setTextFormFieldDefault',
      paragraphId,
      fieldNodeId,
      text: 'Name',
      options: { type: 'regular', maxLength: 10, format: 'Uppercase', enabled: true },
    });
  },
  'typing at the start and the end of a text field': (person) => {
    const at = field(person, 0, 'text');
    caretAt(person, 0, at.start);
    surface(person).insertPlainText('X');
    caretAt(person, 0, at.start + 4);
    surface(person).insertPlainText('Y');
  },
};

function alone(edit: Edit, file: Uint8Array): string {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: file,
    author: 'Olive',
    mode: 'edit',
    modules: [stubReviewModule()],
  });
  try {
    const person = { name: 'alone', editor, container } as unknown as Person;
    edit(person);
    return bodyOf(editorXml(person));
  } finally {
    editor.destroy();
    container.remove();
  }
}

function converged(room: ReturnType<typeof localRoom>, ...people: Person[]): void {
  for (const person of people) {
    expect({ person: person.name, refusals: person.refusals }).toEqual({
      person: person.name,
      refusals: [],
    });
    expect({ person: person.name, room: sameAsRoom(person, room) }).toEqual({
      person: person.name,
      room: true,
    });
    expect({ person: person.name, screen: screenFollowsList(person) }).toEqual({
      person: person.name,
      screen: true,
    });
  }
}

describe('legacy form fields under forms protection', () => {
  for (const [name, edit] of Object.entries(EDITS))
    test(`the organiser's ${name}, then undo and redo`, async () => {
      const file = FILES[name] ?? FILE;
      const room = localRoom(file);
      const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: file });
      const ben = await join(room, 'Ben', { mode: 'edit', login: ORGANISER, wordFile: file });
      const before = bodyOf(editorXml(olive));
      edit(olive);
      await settle();
      converged(room, olive, ben);
      const after = bodyOf(editorXml(olive));
      expect(after).not.toBe(before);
      expect(after).toBe(alone(edit, file));
      expect(bodyOf(editorXml(ben))).toBe(after);
      surface(olive).undo();
      await settle();
      converged(room, olive, ben);
      if (name.startsWith('typing')) {
        // Two keystrokes in two places are two steps: one undo takes back the last
        surface(olive).undo();
        await settle();
        converged(room, olive, ben);
      }
      expect(bodyOf(editorXml(olive))).toBe(before);
      surface(olive).redo();
      await settle();
      converged(room, olive, ben);
      if (name.startsWith('typing')) {
        surface(olive).redo();
        await settle();
      }
      converged(room, olive, ben);
      expect(bodyOf(editorXml(olive))).toBe(after);
    });

  test("a party's check box is a direct edit: refused by the host's party rule", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: FILE });
    EDITS['a check box on']?.(ana);
    await settle();
    expect(ana.refusals.map((r) => r.reason)).toEqual(['package']);
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(screenFollowsList(ana)).toBe(true);
    expect(editorTexts(ana)[1]).toBe('Box: ￼ end.');
  });
});
