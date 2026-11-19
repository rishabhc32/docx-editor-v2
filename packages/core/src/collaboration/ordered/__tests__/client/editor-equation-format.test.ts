// A format over an equation, with real editors and one room. Word and the editor alone format
// the runs and the paragraph marks a selection covers and leave an equation as it is (Word for
// Mac, 2026-11-20): an equation has no run to hold a change record. A host's server agent sends
// its format as one change over a range, so the room keeps no record on an equation either, and
// never refuses another party's later edit there for a record the file does not hold.

import { afterEach, describe, expect, test } from 'bun:test';
import type { OrderedLocalLogin, OrderedLocalRoom } from '../../client/local-room.ts';
import { orderedViewParagraph } from '../../compose/compose.ts';
import { ORDERED_PROTOCOL_VERSION } from '../../protocol/messages.ts';
import { decodeOrderedServerMessage, encodeOrderedClientMessage } from '../../protocol/codec.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { testPerson } from '../test-persons.ts';
import { docx, M, MATRIX_EXTRAS, p, t } from '../word-fixtures.ts';
import {
  BODY,
  closeAll,
  CODE,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  PARTY,
  ROOM,
  settle,
  surface,
} from './editor-harness.ts';
import { alone, exec, revisionAuthors, roomBody, select, type Step } from './format-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { converged } from './tables-harness.ts';

afterEach(closeAll);

const review = reviewTestModule();
const ITALIC = [{ localName: 'i' }];
const math = (text: string) => `<m:oMath xmlns:m="${M}"><m:r><m:t>${text}</m:t></m:r></m:oMath>`;

/** A paragraph that holds only an equation, between two of words. */
const ALONE_FILE = docx(
  p(t('Before the formula.')) +
    p(`<m:oMathPara xmlns:m="${M}">${math('Total=Fee×Units')}</m:oMathPara>`) +
    p(t('After the formula.')),
  MATRIX_EXTRAS
);

/** Words, an equation in the line, words. */
const MIXED_FILE = docx(
  p(t('Price is ') + math('a+b') + t(' today.')) + p(t('Next line.')),
  MATRIX_EXTRAS
);

/** A simple field in the line. */
const FIELD_FILE = docx(
  p(t('Page ') + '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>' + t('.')) +
    p(t('Next line.')),
  MATRIX_EXTRAS
);

/** Italic from the start of paragraph `index` to the start of the next one: its mark too. */
const italicWithMark =
  (index: number): Step =>
  (editor) => {
    const ids = editor.surface?.session.paragraphIds() ?? [];
    const [from, to] = [ids[index], ids[index + 1]];
    if (from === undefined || to === undefined) {
      throw new Error(`no paragraph ${index + 1}`);
    }

    editor.surface?.setSelection({
      anchor: { paragraphId: from, offset: 0 },
      head: { paragraphId: to, offset: 0 },
    });
    exec(editor, { type: 'toggleMark', mark: 'italic' });
  };

/** Italic over `[start, end)` of a paragraph; the editor alone may find nothing to do. */
const italicOver =
  (start: number, end: number, index = 0): Step =>
  (editor) => {
    select(editor, start, end, index);
    editor.exec({ type: 'toggleMark', mark: 'italic' });
  };

/** The room's format records on the units of a paragraph, by character. */
function formatsOf(shared: OrderedLocalRoom, index: number): string[] {
  const view = orderedViewParagraph(shared.room.document, BODY, index);
  return (view?.chars ?? []).map((c) => `${c.ch}:${c.item.fmts.length}`);
}

/** The address of a character of a paragraph of the room, or of the paragraph's end. */
function placeOf(shared: OrderedLocalRoom, index: number, offset: number) {
  const view = orderedViewParagraph(shared.room.document, BODY, index);
  if (!view) {
    throw new Error(`no paragraph ${index}`);
  }

  return (
    view.chars[offset]?.address ?? { mark: { space: view.mark.space, offset: view.mark.offset } }
  );
}

let sent = 0;

/**
 * A host's server agent sends one change through the room, as the host's own runs do (a format
 * as one range, which the editor sends run by run). Returns the room's refusal reasons.
 */
async function hostSends(
  shared: OrderedLocalRoom,
  name: string,
  login: Omit<OrderedLocalLogin, 'person'>,
  body: Readonly<Record<string, unknown>>
): Promise<string[]> {
  const namespace = `${name.toLowerCase()}-host`;
  const refusals: string[] = [];
  const connection = shared.transport({ person: testPerson(name), ...login }).connect({
    open: () => {},
    close: () => {},
    message: (frame) => {
      const read = decodeOrderedServerMessage(frame);
      if (read.ok && read.value.type === 'refuse') {
        refusals.push(...read.value.refusals.map((r) => r.reason ?? r.code));
      }
    },
  });
  const v = ORDERED_PROTOCOL_VERSION;
  const hello = { v, type: 'hello', room: ROOM, file: 1, last: 0, code: CODE, namespace } as const;
  connection.send(encodeOrderedClientMessage(hello));
  sent += 1;
  const change = { v: 1, id: `${namespace}:${sent}`, part: BODY, ...body } as never;
  connection.send(encodeOrderedClientMessage({ v, type: 'changes', changes: [change] }));
  await settle();
  connection.close();
  return refusals;
}

async function roomWith(file: Uint8Array) {
  const shared = localRoom(file, { policy: suggestionsOnlyPolicy() });
  const olive = await join(shared, 'Olive', {
    wordFile: file,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  const ana = await join(shared, 'Ana', { wordFile: file, review });
  return { shared, olive, ana };
}

describe('a format over an equation (Word leaves the equation as it is)', () => {
  test("the host's case: a party's italic over a paragraph that holds only an equation keeps no record, and the other parties' edits there go in", async () => {
    const { shared, olive, ana } = await roomWith(ALONE_FILE);
    const italic = {
      kind: 'format',
      mode: 'suggest',
      from: placeOf(shared, 1, 0),
      to: placeOf(shared, 1, 1),
      properties: ITALIC,
    };
    expect(await hostSends(shared, 'Cy', PARTY, italic)).toEqual([]);
    converged(shared, olive, ana);
    expect(formatsOf(shared, 1)).toEqual(['￼:0']);
    expect(revisionAuthors(olive)).toEqual([]);
    expect(await roomBody(olive)).toBe(
      await alone(ALONE_FILE, review, [
        { author: 'Cy', mode: 'suggesting', step: italicOver(0, 1, 1) },
      ])
    );

    // Another party's format there, and another's deletion of the paragraph's content
    expect(await hostSends(shared, 'Dee', PARTY, italic)).toEqual([]);
    select(ana.editor, 0, 1, 1);
    surface(ana).deleteSelection();
    await settle();
    converged(shared, olive, ana);
    expect(revisionAuthors(olive)).toEqual(['Ana']);
  });

  test("a party's italic over words, an equation and the mark: the words and the mark take it, as in Word", async () => {
    const { shared, olive, ana } = await roomWith(MIXED_FILE);
    const italic = {
      kind: 'format',
      mode: 'suggest',
      from: placeOf(shared, 0, 0),
      to: placeOf(shared, 1, 0),
      properties: ITALIC,
    };
    expect(await hostSends(shared, 'Ana', PARTY, italic)).toEqual([]);
    converged(shared, olive, ana);
    expect(formatsOf(shared, 0).filter((entry) => !entry.endsWith(':1'))).toEqual(['￼:0']);
    expect(await roomBody(olive)).toBe(
      await alone(MIXED_FILE, review, [
        { author: 'Ana', mode: 'suggesting', step: italicWithMark(0) },
      ])
    );

    // Word's form: a change record on each run of words and on the mark; the equation untouched
    const xml = editorXml(olive);
    expect(xml.match(/<w:rPrChange /g)?.length).toBe(3);
    expect(xml).toMatch(/<m:oMath[^>]*><m:r><m:t>a\+b<\/m:t><\/m:r><\/m:oMath>/);
  });

  test("the organiser's direct italic over an equation: no record", async () => {
    const { shared, olive, ana } = await roomWith(MIXED_FILE);
    const italic = {
      kind: 'format',
      mode: 'edit',
      from: placeOf(shared, 0, 9),
      to: placeOf(shared, 0, 10),
      properties: ITALIC,
    };
    expect(await hostSends(shared, 'Olive', ORGANISER, italic)).toEqual([]);
    converged(shared, olive, ana);
    expect(formatsOf(shared, 0)[9]).toBe('￼:0');
    expect(await roomBody(olive)).toBe(
      await alone(MIXED_FILE, review, [{ author: 'Olive', mode: 'edit', step: italicOver(9, 10) }])
    );
  });

  test("a party's italic over their own new equation: no record", async () => {
    const { shared, olive, ana } = await roomWith(MIXED_FILE);
    const rewrite = {
      kind: 'equation',
      mode: 'suggest',
      object: placeOf(shared, 0, 9),
      equation: 0,
      linear: '{a+b}/{2}',
    };
    expect(await hostSends(shared, 'Ana', PARTY, rewrite)).toEqual([]);
    const own = formatsOf(shared, 0).lastIndexOf('￼:0');
    const italic = {
      kind: 'format',
      mode: 'suggest',
      from: placeOf(shared, 0, own),
      to: placeOf(shared, 0, own + 1),
      properties: ITALIC,
    };
    expect(await hostSends(shared, 'Ana', PARTY, italic)).toEqual([]);
    converged(shared, olive, ana);
    expect(formatsOf(shared, 0)[own]).toBe('￼:0');
  });

  test("a party's italic over a simple field reaches the field's runs", async () => {
    const { shared, olive, ana } = await roomWith(FIELD_FILE);
    const italic = {
      kind: 'format',
      mode: 'suggest',
      from: placeOf(shared, 0, 0),
      to: placeOf(shared, 1, 0),
      properties: ITALIC,
    };
    expect(await hostSends(shared, 'Ana', PARTY, italic)).toEqual([]);
    converged(shared, olive, ana);
    expect(formatsOf(shared, 0)).toEqual(['P:1', 'a:1', 'g:1', 'e:1', ' :1', '￼:1', '.:1']);
    expect(editorXml(olive)).toMatch(/<w:fldSimple [^>]*><w:r><w:rPr><w:i\/><w:rPrChange /);
    expect(await roomBody(olive)).toBe(
      await alone(FIELD_FILE, review, [
        { author: 'Ana', mode: 'suggesting', step: italicWithMark(0) },
      ])
    );
  });
});
