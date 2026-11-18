// The editing API's table authoring (`authorTable`, task 9.7) with real editors and one room:
// each kind of change as a party's suggestion and as the organiser's direct edit, through the
// editor's automation path. Each result equals the editor alone with the same op, every copy
// equals the room, and each screen shows its list. A party's suggestion is decided, saved and
// reopened; what the editor alone refuses is refused by the same reason, with nothing applied.

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { paragraphTextOf } from '../../../../store/store/tree-op-apply.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import {
  closeAll,
  editorTexts,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import {
  alone,
  comparedBody,
  converged,
  gridOf,
  TABLE_FILE,
  tableIdsIn,
  type Surface,
  type TableIds,
} from './tables-harness.ts';

afterEach(closeAll);

const review = reviewTestModule();

type Action = Extract<TreeDocOp, { readonly op: 'authorTable' }>['action'];

/** An op built from the ids of the editor it runs in: each editor has ids of its own. */
type Build = (ids: TableIds, paragraph: (text: string) => string) => Action;

/**
 * Run an `authorTable` op through the editor's automation path, as the editing API does. The
 * API attributes a tracked one itself (`automation/plan-tables.ts`): `by` names the author.
 */
function author(s: Surface, build: Build, by?: string) {
  const part = s.session.part();
  const paragraph = (text: string): string => {
    const found = s.session.paragraphIds().find((id) => paragraphTextOf(part, id) === text);
    if (!found) throw new Error(`no paragraph ${text}`);
    return found;
  };
  const action = build(tableIdsIn(part.root), paragraph);
  const revision = by === undefined ? {} : { revision: { author: by, date: DATE } };
  return s.applyAutomationOps(() => [{ op: 'authorTable', action, ...revision }]);
}

const DATE = '2026-10-28T08:00:00Z';

/** What the editor alone makes of the op, as the matrix compares bodies (`alone`). */
const aloneWith = (
  build: Build,
  name: string,
  mode: 'suggesting' | 'edit',
  decide?: 'accept' | 'reject'
) =>
  alone({ author: name, mode, review, ...(decide === undefined ? {} : { decide }) }, (s) => {
    expect(author(s, build, mode === 'suggesting' ? name : undefined).committed).toBe(true);
  });

const kinds = (room: ReturnType<typeof localRoom>): string[] =>
  room.room.joinSnapshot(0).map((change) => change.kind);

async function seats(file = TABLE_FILE) {
  const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { wordFile: file, review });
  const olive = await join(room, 'Olive', {
    wordFile: file,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  return { room, ana, olive };
}

async function decideAll(person: Person, action: 'accept' | 'reject'): Promise<void> {
  expect(
    person.editor.exec({ type: 'resolveAllReviewChanges', action, scope: 'document' }).ok
  ).toBe(true);
  await settle();
}

/** The op in the room by one person: the same result as the editor alone, every copy equal. */
async function made(who: 'ana' | 'olive', build: Build) {
  const seated = await seats();
  const person = seated[who];
  const by = who === 'ana' ? person.name : undefined;
  expect(author(surface(person), build, by).committed).toBe(true);
  await settle();
  converged(seated.room, seated.ana, seated.olive);
  const mode = who === 'ana' ? 'suggesting' : 'edit';
  expect(comparedBody(surface(person).session.part())).toBe(aloneWith(build, person.name, mode));
  return seated;
}

/**
 * A party's op in the room gives what the editor alone gives, pending and decided, and leaves
 * the suggestion's kinds of change.
 */
async function decidedAlike(build: Build, action: 'accept' | 'reject'): Promise<string[]> {
  const { room, ana, olive } = await seats();
  expect(author(surface(ana), build, ana.name).committed).toBe(true);
  await settle();
  converged(room, ana, olive);
  expect(comparedBody(surface(ana).session.part())).toBe(aloneWith(build, ana.name, 'suggesting'));
  const made = kinds(room);
  await decideAll(olive, action);
  converged(room, ana, olive);
  expect(comparedBody(surface(olive).session.part())).toBe(
    aloneWith(build, ana.name, 'suggesting', action)
  );
  closeAll();
  return made;
}

const withTable =
  (mutation: Extract<Action, { kind: 'existing' }>['mutation']): Build =>
  (ids) => ({ kind: 'existing', tableId: ids.table, mutation });

describe("the editing API's table authoring in an ordered room (task 9.7)", () => {
  test("a party's new rows with values: a suggestion, rejected and accepted", async () => {
    const build = withTable({
      kind: 'addRows',
      location: 'end',
      count: 2,
      values: [
        ['Tax', '7', '107'],
        ['Fine', '1', '101'],
      ],
    });
    const { room, ana, olive } = await made('ana', build);
    expect(kinds(room)).toContain('insert-row');
    expect(editorTexts(olive)).toContain('Fine');
    expect(editorXml(olive)).toContain('<w:trPr><w:ins ');
    await decideAll(olive, 'reject');
    converged(room, ana, olive);
    expect(editorTexts(olive)).not.toContain('Tax');
    const again = await made('ana', build);
    await decideAll(again.olive, 'accept');
    converged(again.room, again.ana, again.olive);
    expect(editorTexts(again.olive)).toEqual(expect.arrayContaining(['Tax', '7', '107', 'Fine']));
    expect(editorXml(again.olive)).not.toContain('<w:ins ');
  });

  test("a party's values in the table's cells and rows inserted before a row", async () => {
    const values = withTable({
      kind: 'values',
      values: [
        ['Fee', 'Rate', 'Sum'],
        ['200', '5', '205'],
      ],
    });
    const { room, ana, olive } = await made('ana', values);
    // Each value typed before the old text, which is struck, as the editor alone writes it
    expect(editorTexts(olive)).toEqual(expect.arrayContaining(['SumTotal', '200100']));
    await decideAll(olive, 'reject');
    converged(room, ana, olive);
    expect(editorTexts(olive)).toEqual(expect.arrayContaining(['Total', '100']));
    await made('ana', (ids) => ({
      kind: 'existing',
      tableId: ids.table,
      mutation: {
        kind: 'insertRows',
        rowId: ids.rows[1] ?? '',
        location: 'before',
        count: 1,
        values: [['A', 'B', 'C']],
      },
    }));
  });

  test("a party's deleted rows: struck, and the deletion of every row is refused as the editor refuses it", async () => {
    const { room, ana, olive } = await made(
      'ana',
      withTable({ kind: 'deleteRows', index: 1, count: 1 })
    );
    expect(editorXml(olive)).toContain('<w:trPr><w:del ');
    await decideAll(olive, 'accept');
    converged(room, ana, olive);
    expect(editorTexts(olive)).not.toContain('100');
    const seated = await seats();
    const before = editorXml(seated.ana);
    expect(
      author(surface(seated.ana), withTable({ kind: 'deleteRows', index: 0, count: 2 }), 'Ana')
        .committed
    ).toBe(false);
    await settle();
    expect(editorXml(seated.ana)).toBe(before);
    expect(seated.room.room.order).toBe(0);
  });

  test('a new table with values at a place in a paragraph, as a suggestion and directly', async () => {
    const build: Build = (_ids, paragraph) => ({
      kind: 'insert',
      paragraphId: paragraph('Signed here.'),
      offset: 6,
      rowCount: 2,
      columnCount: 2,
      values: [
        ['Name', 'Date'],
        ['Ana', 'Today'],
      ],
    });
    expect((await decidedAlike(build, 'accept')).slice(0, 2)).toEqual(['split', 'insert-table']);
    await decidedAlike(build, 'reject');
    // The values are the author's inserted text, inside the inserted rows, as Word writes them
    // (a Word file: `e2e/fixtures/list-pagination-break.docx`)
    const alone = aloneWith(build, 'Ana', 'suggesting');
    expect(alone).toMatch(/<w:ins [^>]*><w:r><w:t>Name<\/w:t><\/w:r><\/w:ins>/);
    await made('olive', build);
  });

  test("the organiser's direct changes: style and header rows, columns, a cell's properties", async () => {
    await made('olive', withTable({ kind: 'properties', styleId: 'TableGrid', headerRowCount: 1 }));
    const columns = await made(
      'olive',
      withTable({
        kind: 'addColumns',
        location: 'end',
        count: 2,
        values: [
          ['X', 'Y'],
          ['1', '2'],
        ],
      })
    );
    expect(gridOf(columns.olive)).toHaveLength(5);
    await made('olive', withTable({ kind: 'deleteColumns', index: 0, count: 2 }));
    await made('olive', (ids) => ({
      kind: 'existing',
      tableId: ids.table,
      mutation: {
        kind: 'cell',
        cellId: ids.cells[1]?.[1] ?? '',
        value: 'Six',
        shadingColor: '#FFCC00',
        verticalAlignment: 'Center',
        columnWidth: 100,
      },
    }));
  });

  test("the organiser's direct change is undone and made again", async () => {
    const { room, ana, olive } = await made(
      'olive',
      withTable({ kind: 'addRows', location: 'start', count: 1, values: [['No', 'Rate', 'Sum']] })
    );
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect(editorTexts(olive)).not.toContain('No');
    surface(olive).redo();
    await settle();
    converged(room, ana, olive);
    expect(editorTexts(olive)).toContain('No');
  });

  test("a party's table and cell properties and columns: suggestions decided as the editor alone decides them", async () => {
    const builds: [Build, string][] = [
      [
        withTable({ kind: 'properties', styleId: 'TableGrid', headerRowCount: 1 }),
        'table-property',
      ],
      [
        withTable({ kind: 'addColumns', location: 'end', count: 1, values: [['X'], ['1']] }),
        'insert-column',
      ],
      [withTable({ kind: 'deleteColumns', index: 0, count: 1 }), 'delete-column'],
      [
        (ids) => ({
          kind: 'existing',
          tableId: ids.table,
          mutation: {
            kind: 'cell',
            cellId: ids.cells[1]?.[1] ?? '',
            shadingColor: '#FFCC00',
            verticalAlignment: 'Center',
            columnWidth: 100,
          },
        }),
        'table-property',
      ],
    ];
    for (const [build, kind] of builds)
      for (const action of ['accept', 'reject'] as const)
        expect(await decidedAlike(build, action)).toContain(kind);
  });

  test('what the editor alone refuses is refused by its reason, with nothing applied', async () => {
    const seated = await seats();
    const before = editorXml(seated.ana);
    // A tracked deletion of the whole table: the editor alone refuses it too
    expect(author(surface(seated.ana), withTable({ kind: 'delete' }), 'Ana').committed).toBe(false);
    // The whole table's deletion is a block deletion (group 9.10)
    expect(author(surface(seated.olive), withTable({ kind: 'delete' })).committed).toBe(false);
    await settle();
    expect(editorXml(seated.ana)).toBe(before);
    expect(editorXml(seated.olive)).toBe(before);
    expect(seated.room.room.order).toBe(0);
    expect(seated.olive.refusals).toEqual([
      expect.objectContaining({ code: 'ordered-unsupported-op', reason: 'deleteBlock' }),
    ]);
  });

  test("a party's new rows decided on the saved file give what they give live", async () => {
    for (const action of ['accept', 'reject'] as const) {
      const live = await made(
        'ana',
        withTable({ kind: 'addRows', location: 'end', count: 1, values: [['Tax', '7', '107']] })
      );
      const saved = new Uint8Array(await live.ana.editor.save());
      expect(readOoxmlPackage(saved).ok).toBe(true);
      await decideAll(live.olive, action);
      const expected = editorTexts(live.olive);
      const reopened = await seats(saved);
      await decideAll(reopened.olive, action);
      converged(reopened.room, reopened.ana, reopened.olive);
      expect(editorTexts(reopened.olive)).toEqual(expected);
      closeAll();
    }
  });
});
