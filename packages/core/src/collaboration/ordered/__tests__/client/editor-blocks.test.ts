// Blocks and containers through real editor commands (task 6.1): paragraph properties, a list
// level, a new list, a new table, rows, a note, a control's value, and an organiser's header
// edit. Each is one change the room orders; the screen follows the list and the copy equals the
// room. Columns: `editor-tables.test.ts`.

import { afterEach, describe, expect, test } from 'bun:test';
import { planTableCommand } from '../../../../editor/table-command-plan.ts';
import { SESSION_FILE } from '../sessions/fixture.ts';
import { docx, p, R, t, W } from '../word-fixtures.ts';
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
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import type { OrderedChange } from '../../change/types.ts';

afterEach(closeAll);

/** A document whose section takes a default header. */
const HEADED = docx(
  p(t('The fee is 100 USD.')) +
    `<w:sectPr xmlns:r="${R}"><w:headerReference w:type="default" r:id="rIdHeader1"/></w:sectPr>`,
  {
    parts: {
      'word/header1.xml': `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header text.</w:t></w:r></w:p></w:hdr>`,
    },
    relationships: `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes:
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
  }
);

async function open(): Promise<{ room: OrderedLocalRoom; ana: Person }> {
  const room = localRoom(SESSION_FILE);
  const ana = await join(room, 'Ana', { wordFile: SESSION_FILE });
  return { room, ana };
}

function checked(
  person: Person,
  room: OrderedLocalRoom,
  order: number,
  kind?: OrderedChange['kind']
): void {
  expect(person.refusals).toEqual([]);
  expect(room.room.order).toBe(order);
  if (kind !== undefined) expect(room.room.joinSnapshot(order - 1)[0]?.kind).toBe(kind);
  expect(screenFollowsList(person)).toBe(true);
  expect(sameAsRoom(person, room)).toBe(true);
}

/** Run one of the editor's table commands at the caret. */
function tableCommand(person: Person, command: Readonly<Record<string, unknown>>) {
  const s = surface(person);
  const plan = planTableCommand({
    // The command shapes are the editor's; the test names them by their type
    command: command as never,
    part: s.session.part(),
    layout: s.layout(),
    storeRevision: s.session.packageRevision(),
    selection: s.state().selection,
    cellSelection: s.state().cellSelection,
    themeColors: s.session.documentThemeColors(),
    editable: s.session.editable,
    viewing: false,
  });
  return s.applyTableCommandPlan(plan);
}

const indexOf = (person: Person, text: string): number => {
  const index = editorTexts(person).findIndex((t) => t === text);
  if (index < 0) throw new Error(`no paragraph "${text}"`);
  return index;
};

describe('blocks and containers', () => {
  test("a paragraph's alignment and a list level", async () => {
    const { room, ana } = await open();
    caretAt(ana, 8, 2);
    surface(ana).setParagraphProperty('jc', { val: 'center' });
    await settle();
    checked(ana, room, 1, 'paragraph-property');
    expect(editorXml(ana)).toContain('<w:pPrChange');
    caretAt(ana, 0, 1);
    expect(surface(ana).toggleList('bullet')).toBe(true);
    await settle();
    // A new list definition travels with it (task 9.10)
    checked(ana, room, 2, 'new-list');
  });

  test('a new table, a row, and a row deleted', async () => {
    const { room, ana } = await open();
    caretAt(ana, 1, 0);
    expect(surface(ana).insertTable(2, 2)).toBe(true);
    await settle();
    checked(ana, room, 1, 'insert-table');
    caretAt(ana, indexOf(ana, 'Fee'), 0);
    expect(tableCommand(ana, { type: 'insertRow', where: 'below' }).ok).toBe(true);
    await settle();
    checked(ana, room, 2, 'insert-row');
    caretAt(ana, indexOf(ana, 'Fee'), 0);
    expect(tableCommand(ana, { type: 'deleteRow' }).ok).toBe(true);
    await settle();
    checked(ana, room, 3, 'delete-row');
    expect(editorXml(ana)).toContain('<w:trPr><w:del ');
  });

  test("a footnote and a control's value", async () => {
    const { room, ana } = await open();
    caretAt(ana, 0, 6);
    expect(surface(ana).insertNote('footnote')).toBe(true);
    await settle();
    checked(ana, room, 1, 'insert-note');
    const part = surface(ana).session.part();
    const id = findControl(part.root, 'Click here');
    expect(id).not.toBeNull();
    expect(surface(ana).contentControls.setValue(id ?? '', 'Acme Ltd')).toBe(true);
    await settle();
    checked(ana, room, 2, 'control-value');
    expect(editorTexts(ana)).toContain('Alpha Acme Ltd Omega.');
  });

  test("an organiser's direct edit in the header", async () => {
    const room = localRoom(HEADED);
    const olive = await join(room, 'Olive', { wordFile: HEADED, mode: 'edit', login: ORGANISER });
    const ana = await join(room, 'Ana', { wordFile: HEADED });
    expect(surface(olive).enterHeaderFooter({ rId: 'rIdHeader1', kind: 'header' })).toBe(true);
    const header = surface(olive).session.paragraphIdsIn({
      kind: 'headerFooter',
      rId: 'rIdHeader1',
    });
    const first = header[0];
    if (first === undefined) throw new Error('no header paragraph');
    surface(olive).setSelection({
      anchor: { paragraphId: first, offset: 0 },
      head: { paragraphId: first, offset: 0 },
    });
    surface(olive).insertPlainText('Draft: ');
    await settle();
    checked(olive, room, 1, 'insert');
    expect(room.room.joinSnapshot(0)[0]).toMatchObject({ part: '/word/header1.xml', mode: 'edit' });
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(screenFollowsList(ana)).toBe(true);
  });
});

/** The id of the first content control whose text holds `text`. */
function findControl(
  node: {
    readonly kind: string;
    readonly id: string;
    readonly localName?: string;
    readonly children?: readonly unknown[];
  },
  text: string
): string | null {
  const stack: unknown[] = [node];
  for (let next = stack.pop(); next !== undefined; next = stack.pop()) {
    if (typeof next !== 'object' || next === null || !('kind' in next) || next.kind === 'textValue')
      continue;
    if (
      'localName' in next &&
      next.localName === 'sdt' &&
      'id' in next &&
      typeof next.id === 'string'
    )
      if (JSON.stringify(next).includes(text)) return next.id;
    if ('children' in next && Array.isArray(next.children))
      for (const child of next.children) stack.push(child);
  }
  return null;
}
