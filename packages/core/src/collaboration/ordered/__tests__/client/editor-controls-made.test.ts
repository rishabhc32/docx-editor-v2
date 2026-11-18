// Controls the room made, and controls inside controls, with two real editors and one ordered
// room (task 9.6, found by the random sessions with controls):
//
// - a control the room made (a new one, or a half of one Enter split) is found by the id the
//   editor shows, so it can be changed and removed;
// - each half of a split control holds its own records, and either half can be removed;
// - a control inside a control keeps its own prompt when the outer one's prompt hides, goes or
//   becomes words;
// - a control removed while it shows its prompt leaves the prompt as words that take Backspace,
//   typing and Enter as the editor alone does, and taking the removal back shows the prompt.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { stubReviewModule } from '../../../../editor/__tests__/review-test-module.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import type { OoxmlNode } from '../../../../store/package/ooxml-tree.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedParagraphCount, orderedViewParagraph } from '../../compose/compose.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
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

const PROMPT_FACE = '<w:rStyle w:val="PlaceholderText"/>';
const NEW_PROMPT = 'Click or tap here to enter text.';
const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

/** A rich text control holding words. */
const FEE = docx(
  p(t('The fee is ') + sdt('<w:tag w:val="fee"/><w:richText/>', t('100 USD')) + t(' a month.')) +
    p(t('Signed here.'))
);

/**
 * A multi-line plain text control holding words: Enter inside it splits it. Word ignores Enter
 * inside an inline rich text control, so the editor and the room refuse it there.
 */
const FEE_LINES = docx(
  p(
    t('The fee is ') +
      sdt('<w:tag w:val="fee"/><w:text w:multiLine="1"/>', t('100 USD')) +
      t(' a month.')
  ) + p(t('Signed here.'))
);

/** A block control showing its prompt, and a block control holding words. */
const TERMS = docx(
  p(t('Plan.')) +
    sdt('<w:tag w:val="terms"/><w:showingPlcHdr/>', p(t('Enter the terms.', PROMPT_FACE))) +
    sdt('<w:tag w:val="notes"/>', p(t('Notes apply.'))) +
    p(t('Signed here.'))
);

/** An inline plain text control showing its prompt. */
const NAME = docx(
  p(
    t('Name: ') +
      sdt('<w:tag w:val="name"/><w:showingPlcHdr/><w:text/>', t('Click here.', PROMPT_FACE)) +
      t(' end.')
  ) + p(t('Signed here.'))
);

function converged(room: ReturnType<typeof localRoom>, ...people: Person[]): void {
  for (const person of people)
    expect({
      person: person.name,
      refusals: person.refusals,
      status: person.handle.session.status(),
      room: sameAsRoom(person, room),
      screen: screenFollowsList(person),
    }).toEqual({ person: person.name, refusals: [], status: 'ready', room: true, screen: true });
}

/** The ids of the controls an editor shows, in document order. */
function controlIds(person: Pick<Person, 'editor'>): string[] {
  const ids: string[] = [];
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'contentControl') ids.push(node.id);
    for (const child of node.children) visit(child);
  };
  const found = person.editor.surface;
  if (!found) throw new Error('no surface');
  visit(found.session.part().root);
  return ids;
}

/** An op through the surface's gated write, as the editing API sends it. */
function apply(person: Pick<Person, 'editor'>, op: TreeDocOp): boolean {
  const gated = person.editor.surface as unknown as {
    applyGatedTreeOps(ops: readonly TreeDocOp[]): { committed: boolean; rejected?: boolean };
  };
  const result = gated.applyGatedTreeOps([op]);
  return result.committed && result.rejected !== true;
}

function textsOf(doc: OrderedDocument): string[] {
  const out: string[] = [];
  for (let k = 0; k < orderedParagraphCount(doc, BODY); k += 1)
    out.push(
      orderedViewParagraph(doc, BODY, k)
        ?.chars.map((c) => c.ch)
        .join('') ?? ''
    );
  return out;
}

/** Every editor shows `expected`, and each one's save reopens as it. */
async function everywhere(people: readonly Person[], expected: readonly string[]): Promise<void> {
  for (const person of people) {
    expect(editorTexts(person)).toEqual([...expected]);
    const saved = new Uint8Array(await person.editor.save());
    expect(readOoxmlPackage(saved).ok).toBe(true);
    expect(textsOf(openOrderedWordFile(saved, { file: 2 }).document)).toEqual([...expected]);
  }
}

/** The same steps on the editor alone (no room), in edit mode: the texts it shows. */
function alone(file: Uint8Array, steps: (person: Pick<Person, 'editor'>) => void): string[] {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: file,
    author: 'Olga',
    mode: 'edit',
    modules: [stubReviewModule()],
  });
  try {
    const person = { editor } as unknown as Person;
    steps(person);
    return editorTexts(person);
  } finally {
    editor.destroy();
    container.remove();
  }
}

async function twoPeople(file: Uint8Array) {
  const room = localRoom(file);
  const olga = await join(room, 'Olga', { wordFile: file, mode: 'edit', login: ORGANISER });
  const ana = await join(room, 'Ana', { wordFile: file });
  return { room, olga, ana };
}

describe('controls the room made (task 9.6)', () => {
  test('a new control takes a tag and goes, by the id the editor shows; undo takes both back', async () => {
    const { room, olga, ana } = await twoPeople(FEE);
    caretAt(olga, 1, 6);
    olga.editor.exec({ type: 'insertContentControl', subtype: 'plainText', tag: 'new' });
    await settle();
    const made = controlIds(olga)[1];
    if (made === undefined) throw new Error('no new control');
    expect(apply(olga, { op: 'setContentControlProperties', controlId: made, tag: 'party' })).toBe(
      true
    );
    await settle();
    converged(room, olga, ana);
    expect(editorXml(ana)).toContain('<w:tag w:val="party"/>');
    expect(apply(olga, { op: 'removeContentControl', controlId: made, keepContent: false })).toBe(
      true
    );
    await settle();
    converged(room, olga, ana);
    await everywhere([olga, ana], ['The fee is 100 USD a month.', 'Signed here.']);
    for (let k = 0; k < 3; k += 1) surface(olga).undo();
    await settle();
    converged(room, olga, ana);
    await everywhere([olga, ana], ['The fee is 100 USD a month.', 'Signed here.']);
    expect(controlIds(olga)).toHaveLength(1);
  });

  for (const half of [0, 1])
    for (const keepContent of [true, false])
      test(`half ${half} of a control Enter split goes${keepContent ? ', its words kept' : ''}`, async () => {
        const { room, olga, ana } = await twoPeople(FEE_LINES);
        // A party's Enter inside the control splits it in two (a suggestion)
        caretAt(ana, 0, 14);
        surface(ana).splitParagraph();
        await settle();
        const id = controlIds(olga)[half];
        if (id === undefined) throw new Error('no half');
        expect(apply(olga, { op: 'removeContentControl', controlId: id, keepContent })).toBe(true);
        await settle();
        converged(room, olga, ana);
        const words = keepContent
          ? ['The fee is 100', ' USD a month.']
          : half === 0
            ? ['The fee is ', ' USD a month.']
            : ['The fee is 100', ' a month.'];
        await everywhere([olga, ana], [...words, 'Signed here.']);
        expect(controlIds(olga)).toHaveLength(1);
      });

  test('a control changed, then split by Enter: each half holds its own records', async () => {
    const { room, olga, ana } = await twoPeople(FEE_LINES);
    expect(
      apply(olga, {
        op: 'setContentControlProperties',
        controlId: controlIds(olga)[0] ?? '',
        alias: 'Fee',
      })
    ).toBe(true);
    await settle();
    caretAt(olga, 0, 14);
    surface(olga).splitParagraph();
    await settle();
    caretAt(ana, 0, 12);
    surface(ana).insertPlainText('9');
    await settle();
    converged(room, olga, ana);
    expect(editorXml(olga).match(/<w:alias w:val="Fee"\/>/g)).toHaveLength(2);
    await everywhere([olga, ana], ['The fee is 1900', ' USD a month.', 'Signed here.']);
  });
});

describe('controls inside controls (task 9.6)', () => {
  test("a new control after a party's words over a block control's prompt", async () => {
    const { room, olga, ana } = await twoPeople(TERMS);
    caretAt(ana, 1, 3);
    surface(ana).insertPlainText('Bobby');
    await settle();
    caretAt(olga, 1, 5);
    olga.editor.exec({ type: 'insertContentControl', subtype: 'plainText', tag: 'new' });
    await settle();
    converged(room, olga, ana);
    await everywhere([olga, ana], ['Plan.', `Bobby${NEW_PROMPT}`, 'Notes apply.', 'Signed here.']);
  });

  test("Enter in a block control's prompt leaves a control inside it its own prompt", async () => {
    const { room, olga, ana } = await twoPeople(TERMS);
    caretAt(olga, 1, 16);
    olga.editor.exec({ type: 'insertContentControl', subtype: 'plainText', tag: 'new' });
    await settle();
    expect(editorTexts(ana)[1]).toBe(`Enter the terms.${NEW_PROMPT}`);
    // Policy 4: Enter in the block's prompt removes that prompt, then breaks
    caretAt(ana, 1, 5);
    surface(ana).splitParagraph();
    await settle();
    converged(room, olga, ana);
    await everywhere([olga, ana], ['Plan.', '', NEW_PROMPT, 'Notes apply.', 'Signed here.']);
  });

  test('a block control removed with a new control inside it: the inner one stays', async () => {
    const { room, olga, ana } = await twoPeople(TERMS);
    caretAt(olga, 2, 5);
    olga.editor.exec({ type: 'insertContentControl', subtype: 'plainText', tag: 'new' });
    await settle();
    const notes = controlIds(olga)[1];
    if (notes === undefined) throw new Error('no block control');
    expect(apply(olga, { op: 'removeContentControl', controlId: notes, keepContent: true })).toBe(
      true
    );
    await settle();
    converged(room, olga, ana);
    await everywhere(
      [olga, ana],
      ['Plan.', 'Enter the terms.', `Notes${NEW_PROMPT} apply.`, 'Signed here.']
    );
    expect(controlIds(olga)).toHaveLength(2);
  });
});

describe('a control removed while it shows its prompt (task 9.6)', () => {
  const EDITS = {
    Backspace: (person: Pick<Person, 'editor'>, index: number) => {
      caretAt(person as Person, index, 9);
      person.editor.surface?.deleteBackward();
    },
    typing: (person: Pick<Person, 'editor'>, index: number) => {
      caretAt(person as Person, index, 9);
      person.editor.surface?.insertPlainText('X');
    },
    Enter: (person: Pick<Person, 'editor'>, index: number) => {
      caretAt(person as Person, index, 9);
      person.editor.surface?.splitParagraph();
    },
  } as const;
  const CASES = [
    ['inline', NAME, 0],
    ['block', TERMS, 1],
  ] as const;

  for (const [kind, file, index] of CASES)
    for (const [name, edit] of Object.entries(EDITS))
      test(`${kind}: its prompt is words that take ${name} as the editor alone does`, async () => {
        const { room, olga, ana } = await twoPeople(file);
        const remove = (person: Pick<Person, 'editor'>) =>
          apply(person, {
            op: 'removeContentControl',
            controlId: controlIds(person)[0] ?? '',
            keepContent: true,
          });
        expect(remove(olga)).toBe(true);
        await settle();
        edit(olga, index);
        await settle();
        converged(room, olga, ana);
        const expected = alone(file, (person) => {
          remove(person);
          edit(person, index);
        });
        await everywhere([olga, ana], expected);
      });

  for (const [kind, file] of CASES)
    test(`${kind}: taking the removal back shows the prompt again`, async () => {
      const { room, olga, ana } = await twoPeople(file);
      const before = editorXml(olga);
      expect(
        apply(olga, {
          op: 'removeContentControl',
          controlId: controlIds(olga)[0] ?? '',
          keepContent: true,
        })
      ).toBe(true);
      await settle();
      expect(controlIds(ana)).toHaveLength(kind === 'block' ? 1 : 0);
      surface(olga).undo();
      await settle();
      converged(room, olga, ana);
      expect(editorXml(ana)).toBe(before);
    });
});
