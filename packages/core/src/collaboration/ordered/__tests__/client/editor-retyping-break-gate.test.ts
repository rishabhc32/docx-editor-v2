// A party's section break that gives the next section another type, refused by the client before
// it sends. The host's party rule refuses it in the room (`policy-direct`, reason
// `package`): Word writes that type untracked, so it stays when the break is rejected. With the
// client's `suggests` option the editor refuses it at once, with the room's reason, and nothing
// goes to the room. The room still checks every change, and a break that keeps the type is a
// suggestion as before.

import { afterEach, describe, expect, test } from 'bun:test';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import {
  closeAll,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
} from './editor-harness.ts';
import { converged, FILE } from './sections-fixtures.ts';
import type { Person } from './editor-harness.ts';

afterEach(closeAll);

async function people(suggests: boolean | undefined) {
  const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', {
    wordFile: FILE,
    ...(suggests === undefined ? {} : { suggests }),
  });
  const olive = await join(room, 'Olive', { wordFile: FILE, mode: 'edit', login: ORGANISER });
  return { room, ana, olive };
}

/** A section break at the first paragraph's fourth character. */
function breakAt(person: Person, type: 'continuous' | 'nextPage'): boolean {
  const s = surface(person);
  const first = s.session.paragraphIds()[0]!;
  s.setSelection({
    anchor: { paragraphId: first, offset: 3 },
    head: { paragraphId: first, offset: 3 },
  });
  return s.insertSectionBreak(type);
}

const pending = (room: ReturnType<typeof localRoom>) =>
  [...room.room.document.suggestions.values()].filter((s) => s.status === 'pending');

describe("a party's break that retypes the next section", () => {
  test('is refused by the client before it sends, with the room reason', async () => {
    const { room, ana, olive } = await people(true);
    const order = room.room.order;
    const before = editorXml(ana);
    expect(breakAt(ana, 'continuous')).toBe(false);
    // At once: refused by name, nothing waits to be sent, and the editor shows no break
    expect(ana.refusals).toEqual([
      {
        code: 'ordered-refused',
        reason: 'package',
        detail: 'the type of the section after the break is written untracked',
      },
    ]);
    expect(ana.handle.session.unsentChanges()).toEqual([]);
    expect(editorXml(ana)).toBe(before);
    await settle();
    expect(room.room.order).toBe(order);
    ana.refusals.length = 0;
    converged(room, ana, olive);
  });

  test('without the option the room refuses it, as before', async () => {
    const { room, ana, olive } = await people(undefined);
    expect(breakAt(ana, 'continuous')).toBe(true);
    expect(ana.handle.session.unsentChanges().length).toBeGreaterThan(0);
    await settle();
    expect(ana.refusals.map((r) => r.reason)).toEqual(['package']);
    for (const person of [ana, olive])
      expect(editorXml(person)).not.toContain('<w:type w:val="continuous"/>');
    ana.refusals.length = 0;
    converged(room, ana, olive);
  });

  test('a break that keeps the type is still a suggestion', async () => {
    const { room, ana, olive } = await people(true);
    expect(breakAt(ana, 'nextPage')).toBe(true);
    await settle();
    expect(pending(room)).toHaveLength(1);
    converged(room, ana, olive);
  });

  test("the organiser's retyping break is a direct edit", async () => {
    const { room, ana, olive } = await people(true);
    expect(breakAt(olive, 'continuous')).toBe(true);
    await settle();
    expect(editorXml(ana)).toContain('<w:type w:val="continuous"/>');
    converged(room, ana, olive);
  });
});
