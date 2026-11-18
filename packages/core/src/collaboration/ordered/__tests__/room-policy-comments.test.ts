// The suggestions-only policy's comment rules: everyone the login gives comments writes them,
// on any words but not in a header or a footer; only the organiser resolves; a resolved thread
// takes no reply; a document holds at most 5,000 comments, the Word file's counted.

import { describe, expect, test } from 'bun:test';
import { ORDERED_CHANGE_LIMITS } from '../change/types.ts';
import { loadOrderedPackage } from '../load/load.ts';
import type { OrderedDocument } from '../model/types.ts';
import { OrderedRoom } from '../room/room.ts';
import { suggestionsOnlyPolicy } from '../room/suggestions-only.ts';
import { actor, BODY, HEADER, loadRoom, paragraphOf, realApply } from './room-harness.ts';
import { docx, p, packageOf, t, W } from './word-fixtures.ts';

const OLGA = actor('Olga', 'olga-1', 'organiser');
const ANA = actor('Ana', 'ana-1', 'party');
const RITA = actor('Rita', 'rita-1', 'reader');

/** A room on `doc` with the suggestions-only policy, and `send` for one change at a time. */
function setup(doc: OrderedDocument = loadRoom()) {
  let tick = 0;
  const room = new OrderedRoom({
    document: doc,
    applyChange: realApply,
    policy: suggestionsOnlyPolicy(),
    now: () => Date.UTC(2026, 10, 16) + 1000 * tick++,
  });
  const counters = new Map<string, number>();

  /** One change in a batch: `accepted`, or the refusal as `code/reason`. */
  const send = (who: typeof ANA, change: Record<string, unknown>) => {
    const k = (counters.get(who.namespace) ?? 0) + 1;
    counters.set(who.namespace, k);
    const result = room.submit(who, [{ v: 1, id: `${who.namespace}:${k}`, ...change }]);
    const refusal = result.refused[0];
    return refusal ? `${refusal.code}${refusal.reason ? `/${refusal.reason}` : ''}` : 'accepted';
  };
  const p0 = paragraphOf(room.document, 0);
  const comment = (text = 'Why?', part = BODY, from = paragraphOf(room.document, 0, part)) => ({
    kind: 'comment',
    part,
    from: { space: from.text, offset: 0 },
    to: { space: from.text, offset: 3 },
    text,
  });

  /** The id of the comment the room added last. */
  const last = () => [...room.document.comments.keys()].at(-1) ?? '';
  return { room, send, comment, last, p0 };
}

/** A Word file whose comments part holds `count` comments with no range. */
function withComments(count: number): OrderedDocument {
  const comments = Array.from(
    { length: count },
    (_, k) =>
      `<w:comment w:id="${k}" w:author="Cara"><w:p><w:r><w:t>Note ${k}</w:t></w:r></w:p></w:comment>`
  ).join('');
  const file = docx(p(t('The fee is 100 USD.')), {
    parts: { 'word/comments.xml': `<w:comments xmlns:w="${W}">${comments}</w:comments>` },
    relationships:
      '<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>',
    contentTypes:
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>',
  });
  return loadOrderedPackage(packageOf(file));
}

describe('the comment rules', () => {
  test('a party and the organiser comment, also at a caret and over a pending suggestion', () => {
    const { send, comment, p0 } = setup();
    expect(
      send(ANA, {
        kind: 'insert',
        part: BODY,
        mode: 'suggest',
        place: { space: p0.text, offset: 0 },
        text: 'Now ',
      })
    ).toBe('accepted');
    expect(send(OLGA, comment())).toBe('accepted');
    expect(send(ANA, comment())).toBe('accepted');
    const caret = { space: p0.text, offset: 2 };
    expect(send(ANA, { ...comment(), from: caret, to: caret })).toBe('accepted');
  });

  test('without comments in the login, every comment kind is refused', () => {
    const { send, comment, last } = setup();
    expect(send(OLGA, comment())).toBe('accepted');
    const id = last();
    for (const who of [
      { ...ANA, comments: false },
      { ...OLGA, comments: false },
    ]) {
      expect(send(who, comment())).toBe('policy-comment/comments');
      expect(send(who, { kind: 'comment-reply', comment: id, text: 'Yes.' })).toBe(
        'policy-comment/comments'
      );
      expect(send(who, { kind: 'comment-resolve', comment: id })).toBe('policy-comment/comments');
    }

    expect(send(RITA, comment())).toBe('policy-read-only');
  });

  test('no comment in a header or a footer, for anyone', () => {
    const { send, comment } = setup();
    expect(send(OLGA, comment('Why?', HEADER))).toBe('policy-furniture/furniture');
    expect(send(ANA, comment('Why?', HEADER))).toBe('policy-furniture/furniture');
  });

  test('only the organiser resolves; a resolved thread takes no reply and no resolve', () => {
    const { send, comment, last } = setup();
    expect(send(ANA, comment())).toBe('accepted');
    const id = last();
    expect(send(ANA, { kind: 'comment-resolve', comment: id })).toBe('policy-comment/resolve');
    expect(send(OLGA, { kind: 'comment-resolve', comment: id, note: 'Done.' })).toBe('accepted');
    expect(send(ANA, { kind: 'comment-reply', comment: last(), text: 'Late.' })).toBe(
      'policy-comment/resolved'
    );
    expect(send(OLGA, { kind: 'comment-resolve', comment: id })).toBe('policy-comment/resolved');
    expect(send(OLGA, { kind: 'comment-reply', comment: 'none', text: 'Who?' })).toBe(
      'unknown-comment'
    );
  });

  test('a text past 10,000 UTF-16 units is refused', () => {
    const { send, comment } = setup();
    expect(send(ANA, comment('a'.repeat(ORDERED_CHANGE_LIMITS.comment)))).toBe('accepted');
    expect(send(ANA, comment('a'.repeat(ORDERED_CHANGE_LIMITS.comment + 1)))).toBe('too-long');
  });

  test("a document holds 5,000 comments, the Word file's counted", () => {
    const { room, send, comment } = setup(withComments(4999));
    expect(send(ANA, comment())).toBe('accepted');
    expect(room.document.comments.size).toBe(5000);
    expect(send(ANA, comment())).toBe('policy-comment/limit');
    expect(send(ANA, { kind: 'comment-reply', comment: '7', text: 'Yes.' })).toBe(
      'policy-comment/limit'
    );
    expect(send(OLGA, { kind: 'comment-resolve', comment: '7', note: 'Done.' })).toBe(
      'policy-comment/limit'
    );

    // A resolve with no note adds no comment
    expect(send(OLGA, { kind: 'comment-resolve', comment: '7' })).toBe('accepted');
  });
});
