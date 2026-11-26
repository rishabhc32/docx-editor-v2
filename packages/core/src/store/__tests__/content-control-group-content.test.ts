// A group control's own content (`w:group`), and the controls nested in it.
//
// Word refuses every edit of what a group holds outside its nested controls: typing, deleting,
// a deletion of the author's own pending words, and a decision on a revision there (Accept and
// Reject; Accept All leaves it). The controls nested in a group stay editable. The engine once
// let an author delete their own pending words in an inline group, and the group was gone.
// Word evidence: desktop Word for Mac, 2026-11-26.

import { describe, expect, test } from 'bun:test';
import {
  bodyStoryRoot,
  readOoxmlPart,
  serializeOoxmlPart,
  storyParagraphs,
  type OoxmlPart,
} from '../index.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function parseDoc(bodyInner: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body>${bodyInner}</w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!result.ok) {
    throw new Error(result.reason);
  }

  return result.part;
}

const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (inner: string) => `<w:p>${inner}</w:p>`;
const sdt = (props: string, inner: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const GROUP = '<w:tag w:val="group"/><w:group/>';
const ANA = { id: '5', author: 'Ana', date: '2026-11-26T00:00:00Z' };
const ANAS = `<w:ins w:id="5" w:author="Ana" w:date="${ANA.date}">${r('ab')}</w:ins>`;
const BEAS = `<w:del w:id="6" w:author="Bea" w:date="${ANA.date}"><w:r><w:delText>ab</w:delText></w:r></w:del>`;
const BEA = { id: '6', author: 'Bea', date: ANA.date };

/** "A: " + an outer control > an inline group > `inner` + " end.": the group holds 3 to 5. */
const inOuter = (inner: string) =>
  parseDoc(p(r('A: ') + sdt('<w:tag w:val="outer"/>', sdt(GROUP, inner)) + r(' end.')));

/** "A: " + a group > "G " + a nested control "ab" + " end.": own 3 to 5, nested 5 to 7. */
const withField = (lock = '') =>
  parseDoc(
    p(r('A: ') + sdt(GROUP + lock, r('G ') + sdt('<w:tag w:val="n1"/>', r('ab'))) + r(' end.'))
  );

/** A block group: "Own words.", a nested block control "Field.", then "After.". */
const blockGroup = () =>
  parseDoc(
    sdt(GROUP, p(r('Own words.')) + sdt('<w:tag w:val="n2"/>', p(r('Field.')))) + p(r('After.'))
  );

function paragraphIds(part: OoxmlPart): string[] {
  const body = bodyStoryRoot(part);
  return body ? storyParagraphs(body).map((paragraph) => paragraph.id) : [];
}

/** The op's refusal, and that the part did not change. */
function refused(part: OoxmlPart, op: TreeDocOp): string {
  const before = serializeOoxmlPart(part);
  const result = applyTreeOp(part, op);
  expect(serializeOoxmlPart(part)).toBe(before);
  return result.ok ? 'applied' : result.reason;
}

describe("a group control's own content refuses every edit (Word)", () => {
  test("an author's deletion of their own pending words keeps the words, the revision and the group", () => {
    const part = inOuter(ANAS);
    const [id] = paragraphIds(part);
    const op: TreeDocOp = { op: 'deleteText', paragraphId: id!, start: 3, end: 5, revision: ANA };
    expect(refused(part, op)).toBe('locked');
  });

  test('typing, deleting and Enter in the own content are refused', () => {
    const part = withField();
    const [id] = paragraphIds(part);
    const ops: TreeDocOp[] = [
      { op: 'insertText', paragraphId: id!, offset: 4, text: 'Q' },
      { op: 'deleteText', paragraphId: id!, start: 3, end: 4 },
      { op: 'deleteText', paragraphId: id!, start: 4, end: 6 },
      { op: 'deleteText', paragraphId: id!, start: 0, end: 12 },
      { op: 'splitParagraph', paragraphId: id!, offset: 4 },
    ];
    expect(ops.map((op) => refused(part, op))).toEqual(ops.map(() => 'locked'));
  });

  test('a block group refuses an edit of its own paragraph', () => {
    const part = blockGroup();
    const [own] = paragraphIds(part);
    const ops: TreeDocOp[] = [
      { op: 'insertText', paragraphId: own!, offset: 3, text: 'Q' },
      { op: 'deleteText', paragraphId: own!, start: 0, end: 3 },
      { op: 'splitParagraph', paragraphId: own!, offset: 3 },
    ];
    expect(ops.map((op) => refused(part, op))).toEqual(ops.map(() => 'locked'));
  });

  test('a decision on a revision in the own content is refused (Word keeps it)', () => {
    const accept: TreeDocOp = { op: 'acceptRevision', revision: BEA };
    const reject: TreeDocOp = { op: 'rejectRevision', revision: ANA };
    expect(refused(inOuter(BEAS), accept)).toBe('locked');
    expect(refused(inOuter(ANAS), reject)).toBe('locked');
  });
});

describe('the controls nested in a group stay editable', () => {
  test('typing and deleting in a nested inline control', () => {
    const part = withField();
    const [id] = paragraphIds(part);
    const typed = applyTreeOp(part, { op: 'insertText', paragraphId: id!, offset: 6, text: 'Q' });
    expect(typed.ok).toBe(true);
    const deleted = applyTreeOp(part, { op: 'deleteText', paragraphId: id!, start: 5, end: 7 });
    expect(deleted.ok).toBe(true);
  });

  test('typing and deleting in a nested block control', () => {
    const part = blockGroup();
    const nested = paragraphIds(part)[1]!;
    expect(
      applyTreeOp(part, { op: 'insertText', paragraphId: nested, offset: 3, text: 'Q' }).ok
    ).toBe(true);
    expect(applyTreeOp(part, { op: 'deleteText', paragraphId: nested, start: 0, end: 3 }).ok).toBe(
      true
    );
  });

  test('a decision in a nested control', () => {
    const part = parseDoc(
      p(r('A: ') + sdt(GROUP, r('G ') + sdt('<w:tag w:val="n1"/>', BEAS)) + r(' end.'))
    );
    expect(applyTreeOp(part, { op: 'acceptRevision', revision: BEA }).ok).toBe(true);
  });

  test('text outside the group stays editable', () => {
    const part = withField();
    const [id] = paragraphIds(part);
    expect(applyTreeOp(part, { op: 'deleteText', paragraphId: id!, start: 0, end: 2 }).ok).toBe(
      true
    );
  });
});
