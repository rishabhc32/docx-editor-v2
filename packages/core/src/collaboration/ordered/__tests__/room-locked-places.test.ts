// A content control's lock in a room, around and inside the control (security): the acceptance
// run's case as a modified browser sends it, typing at a locked control's edges and at its prompt,
// a lock the organiser sets in the room, a nested control, and objects, notes, links and fields
// inside a locked control. The editor alone refuses the same edits (the last test). The change
// kinds one by one are `room-locked-controls.test.ts`.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { applyTreeOp } from '../../../store/store/tree-op-apply.ts';
import type { TreeDocOp } from '../../../store/store/tree-op-types.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import type { OrderedDocument } from '../model/types.ts';
import { addressOf, at, BODY, FOOTNOTES, textOf } from './change-apply-fixture.ts';
import { bodyFingerprint } from './room-harness.ts';
import {
  BOB,
  BOTH,
  content,
  control,
  edit,
  EXTRAS,
  fileOf,
  firstOf,
  LINK,
  LOCKS,
  lockOf,
  MEDIA,
  noteOf,
  OLGA,
  pictureOf,
  PROMPT_FACE,
  REFERENCE,
  roomOn,
  sdt,
  type Lock,
} from './room-locked-fixture.ts';
import { docx, M, p, pictureRun, t, W } from './word-fixtures.ts';

describe("the acceptance run's case (s100-short seed 1007): typing in a content-locked combo box", () => {
  const COMBO =
    '<w:alias w:val="Currency"/><w:lock w:val="contentLocked"/><w:comboBox>' +
    '<w:listItem w:displayText="USD" w:value="USD"/><w:listItem w:displayText="EUR" w:value="EUR"/>' +
    '</w:comboBox>';
  const FILE = docx(p(t('Pay in ') + sdt(COMBO, t('USD'))) + p(t('after')));

  test('a forged raw change from a modified browser is refused by name, organiser and party', () => {
    for (const who of BOTH) {
      const { room, doc } = roomOn(FILE);
      const words = doc.stories
        .get(BODY)
        ?.tree.toArray()
        .find((item) => item.kind === 'text' && item.text === 'USD');
      if (!words) throw new Error('fixture');
      const before = bodyFingerprint(doc);
      const sender = who === 'organiser' ? OLGA : BOB;
      // The bytes a modified browser sends: a well-formed change at offset 2 of "USD"
      const raw = JSON.parse(
        JSON.stringify({
          v: 1,
          id: `${sender.namespace}:386`,
          kind: 'insert',
          part: BODY,
          mode: who === 'organiser' ? 'edit' : 'suggest',
          place: { space: words.space, offset: words.offset + 2 },
          text: 'émigré ',
        })
      );
      const out = room.submit(sender, [raw]);
      expect(out.accepted).toEqual([]);
      expect(out.refused).toEqual([
        expect.objectContaining({ id: raw.id, code: 'apply-refused', reason: 'locked' }),
      ]);
      expect(bodyFingerprint(doc)).toBe(before);
      expect(textOf(doc, 0)).toBe('Pay in USD');
    }
  });

  test('a value from its list is refused too; typing beside it is not', () => {
    const { doc, send } = roomOn(FILE);
    const combo = control(doc);
    expect(
      send('organiser', {
        kind: 'control-value',
        part: BODY,
        mode: 'edit',
        control: addressOf(combo),
        value: { kind: 'listItem', value: 'EUR' },
      })
    ).toBe('apply-refused/locked');
    expect(
      send('party', { kind: 'insert', ...edit('party'), place: at(doc, 0, 7), text: 'Q' })
    ).toBe('accepted');
    expect(textOf(doc, 0)).toBe('Pay in QUSD');
  });
});

/** The words the view shows inside a control, in order. */
function insideControls(doc: OrderedDocument, index: number): string {
  const view = orderedViewParagraph(doc, BODY, index);
  return (view?.chars ?? [])
    .filter((c) => c.containers.includes('c:'))
    .map((c) => c.ch)
    .join('');
}

describe('at the edges of a locked control, words go beside it, as in the editor alone', () => {
  test('typing at its front and at its end lands outside it', () => {
    for (const lock of LOCKS) {
      const { doc, send } = roomOn(fileOf('inline', lock));
      expect([
        lock,
        send('party', { kind: 'insert', ...edit('party'), place: at(doc, 0, 7), text: 'A' }),
      ]).toEqual([lock, 'accepted']);
      expect([
        lock,
        send('organiser', {
          kind: 'insert',
          ...edit('organiser'),
          place: at(doc, 0, 11),
          text: 'Z',
        }),
      ]).toEqual([lock, 'accepted']);
      expect(textOf(doc, 0)).toBe('Pay in AUSDZ now.');
      expect([lock, insideControls(doc, 0)]).toEqual([lock, 'USD']);
    }
  });

  test('typing at the edge of a locked control that shows its prompt does not replace the prompt', () => {
    const prompt = t('Choose.', PROMPT_FACE);
    for (const lock of LOCKS) {
      const file = docx(
        p(t('Pay in ') + sdt(`<w:showingPlcHdr/>${lockOf(lock)}`, prompt) + t(' now.')),
        EXTRAS
      );
      const { doc, send } = roomOn(file);
      expect([
        lock,
        send('party', { kind: 'insert', ...edit('party'), place: at(doc, 0, 7), text: 'A' }),
      ]).toEqual([lock, 'accepted']);
      // Unlocked, the words replace the prompt (placement rule 7); locked, they go before it
      expect([lock, insideControls(doc, 0)]).toEqual([lock, content(lock) ? 'Choose.' : 'A']);
      // Inside the prompt the lock refuses as anywhere in the content
      const inPrompt = send('party', {
        kind: 'insert',
        ...edit('party'),
        place: at(doc, 0, content(lock) ? 10 : 8),
        text: 'B',
      });
      expect([lock, inPrompt]).toEqual([lock, content(lock) ? 'apply-refused/locked' : 'accepted']);
    }
  });
});

describe('a lock the organiser set in the room binds at once, and its undo frees the content', () => {
  test('a control locked by a control-property change refuses the next typing', () => {
    const { doc, send } = roomOn(fileOf('inline', 'unlocked'));
    const lock = {
      kind: 'control-property',
      ...edit('organiser'),
      control: addressOf(control(doc)),
    };
    expect(send('organiser', { ...lock, lock: 'contentLocked' })).toBe('accepted');
    const typing = { kind: 'insert', ...edit('party'), place: at(doc, 0, 9), text: 'Q' };
    expect(send('party', typing)).toBe('apply-refused/locked');
    expect(
      send('organiser', { kind: 'revert', suggestions: ['S:olga-1:1.1'], status: 'rejected' })
    ).toBe('accepted');
    expect(send('party', typing)).toBe('accepted');
  });
});

describe('a nested control inherits the lock of the control around it', () => {
  test('an unlocked control inside a content-locked one refuses typing; a sibling does not', () => {
    const inner = sdt('<w:tag w:val="inner"/>', t('in'));
    const file = docx(
      p(
        t('A ') +
          sdt('<w:lock w:val="contentLocked"/>', t('x') + inner + t('y')) +
          t(' B ') +
          sdt('<w:tag w:val="free"/>', t('ok'))
      ),
      EXTRAS
    );
    const { doc, send } = roomOn(file);
    // "A x[in]y B ok": inside the inner control at 4, inside the free control at 10
    expect(
      send('party', { kind: 'insert', ...edit('party'), place: at(doc, 0, 4), text: 'Q' })
    ).toBe('apply-refused/locked');
    expect(
      send('party', { kind: 'insert', ...edit('party'), place: at(doc, 0, 10), text: 'Q' })
    ).toBe('accepted');
  });
});

describe('objects, notes and links already inside a locked control', () => {
  const file = (lock: Lock, inner: string) => fileOf('inline', lock, inner);

  test("a picture's size, its image and a note's conversion are refused there", () => {
    for (const lock of LOCKS) {
      const picture = roomOn(file(lock, pictureRun(5)));
      const atom = addressOf(pictureOf(picture.doc));
      const resize = picture.send('organiser', {
        kind: 'object-property',
        ...edit('organiser'),
        object: atom,
        ops: [{ op: 'resize', cx: 228_600, cy: 228_600 }],
      });
      const image = picture.send('organiser', {
        kind: 'object-resource',
        ...edit('organiser'),
        object: atom,
        media: MEDIA,
      });
      const note = roomOn(file(lock, t('USD') + REFERENCE));
      const convert = note.send('organiser', {
        kind: 'convert-note',
        part: FOOTNOTES,
        mode: 'edit',
        notes: [addressOf(noteOf(note.doc))],
      });
      const want = content(lock) ? 'apply-refused/locked' : 'accepted';
      expect([lock, resize, image, convert]).toEqual([lock, want, want, want]);
    }
  });

  test("a field's code, a form field's text and an equation are refused there", () => {
    const formText =
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Name"/><w:enabled/>' +
      '<w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>' +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${t('Bob')}` +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const page =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${t('1')}` +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const math = `<m:oMath xmlns:m="${M}"><m:r><m:t>x=1</m:t></m:r></m:oMath>`;
    for (const lock of LOCKS) {
      const field = roomOn(file(lock, t('USD ') + page));
      const code = field.send('organiser', {
        kind: 'field-code',
        ...edit('organiser'),
        field: at(field.doc, 0, 11),
        code: 'NUMPAGES',
      });
      const form = roomOn(file(lock, formText));
      const typed = form.send('organiser', {
        kind: 'form-text',
        ...edit('organiser'),
        field: at(form.doc, 0, 7),
        op: { op: 'insertText', offset: 1, text: 'x' },
      });
      const equation = roomOn(file(lock, t('USD ') + math));
      const replaced = equation.send('organiser', {
        kind: 'equation',
        ...edit('organiser'),
        object: at(equation.doc, 0, 11),
        equation: 0,
        linear: 'x^2',
      });
      const want = content(lock) ? 'apply-refused/locked' : 'accepted';
      expect([lock, code, typed, replaced]).toEqual([lock, want, want, want]);
    }
  });

  test("a link's target and the link's removal are refused there", () => {
    for (const lock of LOCKS) {
      const retargeted = roomOn(file(lock, LINK));
      const target = retargeted.send('organiser', {
        kind: 'link-target',
        ...edit('organiser'),
        link: addressOf(firstOf(retargeted.doc, 'hyperlink')),
        href: 'https://example.com/other',
      });
      const unlinked = roomOn(file(lock, LINK));
      const removed = unlinked.send('organiser', {
        kind: 'remove-link',
        ...edit('organiser'),
        link: addressOf(firstOf(unlinked.doc, 'hyperlink')),
      });
      const want = content(lock) ? 'apply-refused/locked' : 'accepted';
      expect([lock, target, removed]).toEqual([lock, want, want]);
    }
  });
});

describe('the editor alone refuses the same edits (the reference)', () => {
  const W_DOC = (body: string) =>
    `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`;
  const engine = (lock: Lock, op: (paragraphId: string) => TreeDocOp): string => {
    const read = readOoxmlPart(W_DOC(p(t('Pay in ') + sdt(lockOf(lock), t('USD')) + t(' now.'))), {
      name: '/word/document.xml',
      contentType: 'application/xml',
    });
    if (!read.ok) throw new Error(read.reason);
    const body = read.part.root.children.find((c) => c.kind !== 'textValue');
    const paragraph = body && body.kind !== 'textValue' ? body.children[0] : undefined;
    if (!paragraph) throw new Error('fixture');
    const result = applyTreeOp(read.part, op(paragraph.id));
    return result.ok ? 'applied' : result.reason;
  };

  test('typing, deleting, formatting and Enter inside: locked exactly when the room refuses', () => {
    for (const lock of LOCKS) {
      const rows = [
        engine(lock, (id) => ({ op: 'insertText', paragraphId: id, offset: 9, text: 'Q' })),
        engine(lock, (id) => ({ op: 'deleteText', paragraphId: id, start: 8, end: 10 })),
        engine(lock, (id) => ({
          op: 'setRunProperties',
          paragraphId: id,
          start: 8,
          end: 10,
          properties: [{ localName: 'b' }],
        })),
      ];
      const want = content(lock) ? 'locked' : 'applied';
      expect([lock, ...rows]).toEqual([lock, want, want, want]);
    }
  });
});
