// Accept All and Reject All next to a content control's lock.
//
// Word's Accept All and Reject All leave the changes a lock keeps and decide the others (desktop
// Word for Mac, 2026-11-26: a deletion in a group and one outside it; Accept All accepted the one
// outside and left the one in the group). A decision on that change alone stays refused, as Word
// refuses it. The engine once refused the whole Accept All.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart, type OoxmlPart } from '../index.ts';
import { planRevisionBatch } from '../store/revision-batch.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';

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
const del = (id: number, text: string) =>
  `<w:del w:id="${id}" w:author="Bea" w:date="2026-11-26T00:00:00Z"><w:r>` +
  `<w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
const GROUP = '<w:tag w:val="group"/><w:lock w:val="contentLocked"/><w:group/>';

/** Bea's deletion of "ab" in a Word-saved group, and of "." after "Signed". */
const FILE = () =>
  parseDoc(
    p(
      r('A: ') +
        `<w:sdt><w:sdtPr>${GROUP}</w:sdtPr><w:sdtContent>${del(6, 'ab')}</w:sdtContent></w:sdt>`
    ) + p(r('Signed') + del(7, '.'))
  );

/** Apply the plan's ops in order; the part after them, or the first refusal. */
function decided(part: OoxmlPart, action: 'accept' | 'reject', skipLocked: boolean) {
  const plan = planRevisionBatch(part, action, undefined, undefined, { skipLocked });
  let current = part;
  for (const op of plan.ops) {
    const result = applyTreeOp(current, op);
    if (!result.ok) {
      return { plan, refused: result.reason, xml: serializeOoxmlPart(part) };
    }

    current = result.part;
  }

  return { plan, refused: null, xml: serializeOoxmlPart(current) };
}

describe('Accept All and Reject All leave the changes a lock keeps (Word)', () => {
  for (const action of ['accept', 'reject'] as const) {
    test(`${action}: the change outside is decided; the one in the group stays, reported`, () => {
      const { plan, refused, xml } = decided(FILE(), action, true);
      expect(refused).toBeNull();
      expect(plan.result.skipped.map((skip) => skip.reason)).toEqual(['locked']);
      expect(plan.result.resolved).toHaveLength(1);
      expect(plan.result.remaining).toBe(1);
      expect(xml).toContain('w:id="6"');
      expect(xml).not.toContain('w:id="7"');
      expect(xml.includes('.</w:t>')).toBe(action === 'reject');
    });
  }

  test('a decision without the skip is refused whole, as a single decision there is', () => {
    const { refused, xml } = decided(FILE(), 'accept', false);
    expect(refused).toBe('locked');
    expect(xml).toBe(serializeOoxmlPart(FILE()));
  });
});
