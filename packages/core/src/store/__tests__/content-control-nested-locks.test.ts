// What a content lock binds: the control's own content, not the content of a control nested in it.
//
// Word lets a person edit a control nested in a `contentLocked` control or a group, and keeps the
// outer control's own words (desktop Word for Mac, 2026-11-26: a group and a rich text control,
// each content-locked, each with own words and a nested control; the nested words were deleted,
// the own words stayed). Word saves every group with `contentLocked`, so the fields in a
// Word-saved group are filled this way. The engine once passed the lock down and refused them.

import { describe, expect, test } from 'bun:test';
import {
  bodyStoryRoot,
  contentControlsIn,
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
const lock = (value: string) => `<w:lock w:val="${value}"/>`;
const WORD_GROUP = `<w:tag w:val="group"/>${lock('contentLocked')}<w:group/>`;
const DATE = '2026-11-26T00:00:00Z';
const BEAS = `<w:del w:id="6" w:author="Bea" w:date="${DATE}"><w:r><w:delText>ab</w:delText></w:r></w:del>`;

/** "A: " + an outer control (`outer`) > "G " + a nested control "ab" + " end.": own 3 to 5, nested 5 to 7. */
const nested = (outer: string, inner = r('ab')) =>
  parseDoc(p(r('A: ') + sdt(outer, r('G ') + sdt('<w:tag w:val="n1"/>', inner)) + r(' end.')));

/** A Word-saved block group: "Own words.", a nested plain text field "Field.", then "After.". */
const blockGroup = () =>
  parseDoc(
    sdt(WORD_GROUP, p(r('Own words.')) + sdt('<w:tag w:val="n2"/><w:text/>', p(r('Field.')))) +
      p(r('After.'))
  );

function paragraphIds(part: OoxmlPart): string[] {
  const body = bodyStoryRoot(part);
  return body ? storyParagraphs(body).map((paragraph) => paragraph.id) : [];
}

/** The op's refusal, and that the part did not change; `applied` when it applied. */
function outcome(part: OoxmlPart, op: TreeDocOp): string {
  const before = serializeOoxmlPart(part);
  const result = applyTreeOp(part, op);
  if (result.ok) {
    return 'applied';
  }

  expect(serializeOoxmlPart(part)).toBe(before);
  return result.reason;
}

describe('a content lock keeps its own content, not a control nested in it (Word)', () => {
  for (const [name, outer] of [
    ['a Word-saved group', WORD_GROUP],
    ['a content-locked rich text control', `<w:tag w:val="rich"/>${lock('contentLocked')}`],
  ] as const) {
    test(`${name}: the nested control takes typing and deletion`, () => {
      const part = nested(outer);
      const [id] = paragraphIds(part);
      const ops: TreeDocOp[] = [
        { op: 'insertText', paragraphId: id!, offset: 6, text: 'Q' },
        { op: 'deleteText', paragraphId: id!, start: 5, end: 7 },
        { op: 'deleteText', paragraphId: id!, start: 6, end: 7 },
      ];
      expect(ops.map((op) => outcome(part, op))).toEqual(ops.map(() => 'applied'));
    });

    test(`${name}: its own words, and a deletion that crosses into the nested control, are refused`, () => {
      const part = nested(outer);
      const [id] = paragraphIds(part);
      const ops: TreeDocOp[] = [
        { op: 'insertText', paragraphId: id!, offset: 4, text: 'Q' },
        { op: 'deleteText', paragraphId: id!, start: 3, end: 5 },
        { op: 'deleteText', paragraphId: id!, start: 4, end: 6 },
      ];
      expect(ops.map((op) => outcome(part, op))).toEqual(ops.map(() => 'locked'));
    });

    test(`${name}: a decision in the nested control applies`, () => {
      const part = nested(outer, BEAS);
      const op: TreeDocOp = {
        op: 'acceptRevision',
        revision: { id: '6', author: 'Bea', date: DATE },
      };
      expect(outcome(part, op)).toBe('applied');
    });
  }

  test("a Word-saved block group's field is filled; its own paragraph stays refused", () => {
    const part = blockGroup();
    const [own, field] = paragraphIds(part);
    const fieldControl = contentControlsIn(part.root)[1]!.node;
    expect(outcome(part, { op: 'insertText', paragraphId: field!, offset: 6, text: 'Q' })).toBe(
      'applied'
    );
    expect(
      outcome(part, { op: 'setContentControlValue', controlId: fieldControl.id, value: 'Ana' })
    ).toBe('applied');
    expect(outcome(part, { op: 'insertText', paragraphId: own!, offset: 3, text: 'Q' })).toBe(
      'locked'
    );
  });

  test('each control keeps its own lock, at every depth', () => {
    // outer (contentLocked) > "G " + middle (unlocked) > "M " + inner (contentLocked) > "ab"
    const part = parseDoc(
      p(
        sdt(
          lock('contentLocked'),
          r('G ') + sdt('<w:tag w:val="m"/>', r('M ') + sdt(lock('contentLocked'), r('ab')))
        )
      )
    );
    const [id] = paragraphIds(part);
    expect(outcome(part, { op: 'insertText', paragraphId: id!, offset: 3, text: 'Q' })).toBe(
      'applied'
    );
    expect(outcome(part, { op: 'insertText', paragraphId: id!, offset: 5, text: 'Q' })).toBe(
      'locked'
    );
    expect(outcome(part, { op: 'insertText', paragraphId: id!, offset: 1, text: 'Q' })).toBe(
      'locked'
    );
  });
});

describe("typing at a nested control's leading edge in a locked control", () => {
  // The place beside the nested control is the locked control's own content, so the nested
  // control is the one place there that takes the words, tracked or not, as the room puts them
  for (const revision of [undefined, { author: 'Ana', date: DATE }]) {
    test(`${revision ? 'a tracked' : 'an'} insertion goes into the nested control`, () => {
      const part = nested(WORD_GROUP);
      const [id] = paragraphIds(part);
      const op: TreeDocOp = {
        op: 'insertText',
        paragraphId: id!,
        offset: 5,
        text: 'Q',
        ...(revision ? { revision } : {}),
      };
      const result = applyTreeOp(part, op);
      if (!result.ok) {
        throw new Error(result.reason);
      }

      const field = contentControlsIn(result.part.root)[1]!.node;
      expect(serializeOoxmlPart(result.part)).toContain('G </w:t>');
      // Untracked, the words join the field's run; tracked, they are a run of their own
      expect(JSON.stringify(field)).toMatch(/"Q(ab)?"/);
    });
  }
});

describe('what still binds a nested control (Word not checked)', () => {
  test('an sdtContentLocked control keeps the content of a control nested in it', () => {
    const part = nested(lock('sdtContentLocked'));
    const [id] = paragraphIds(part);
    expect(outcome(part, { op: 'insertText', paragraphId: id!, offset: 6, text: 'Q' })).toBe(
      'locked'
    );
  });
});
