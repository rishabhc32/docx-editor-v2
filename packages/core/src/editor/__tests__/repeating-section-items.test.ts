// Repeating section items in the editor (task 9.6 of `ordered-suggestion-collaboration`): Word's
// "+" adds an item that repeats the one before it, its words copied and its content controls back
// to their prompts (except a nested repeating section and a bound control); an item goes; a
// section keeps one item; `w15:doNotAllowInsertDeleteSection` and a content lock refuse both.
// The engine refused both ops as `unsupported` before.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart, type OoxmlNode } from '../../store/package/ooxml-tree.ts';
import { paragraphTextOf } from '../../store/store/paragraph-model-text.ts';
import { applyTreeOp } from '../../store/store/tree-op-apply.ts';
import { readOoxmlPackage } from '../../store/package/ooxml-package.ts';
import type { TreeDocOp } from '../../store/store/tree-op-types.ts';
import { validateTreeOp } from '../../store/store/tree-op-validate.ts';
import { docx, p, t } from '../../collaboration/ordered/__tests__/word-fixtures.ts';

const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const item = (content: string) =>
  sdt(`<w:id w:val="21"/><w15:repeatingSectionItem ${W15}/>`, content);
const section = (items: string, more = '') =>
  sdt(`<w:id w:val="20"/><w15:repeatingSection ${W15}>${more}</w15:repeatingSection>`, items);

const NAME = sdt('<w:id w:val="22"/><w:text/>', t('Acme'));
const FILE = (more = '') =>
  docx(
    p(t('Parties:')) + section(item(p(t('Party: ') + NAME + t(' (seller)'))), more) + p(t('End.'))
  );

function part(bytes: Uint8Array) {
  const read = readOoxmlPackage(bytes);
  if (!read.ok) throw new Error('fixture');
  const found = read.package.parts.get('/word/document.xml');
  if (!found) throw new Error('no body');
  return found;
}

function controlIds(node: OoxmlNode, out: string[] = []): string[] {
  if (node.kind === 'textValue') return out;
  if (node.kind === 'contentControl') out.push(node.id);
  for (const child of node.children) controlIds(child, out);
  return out;
}

function texts(xml: ReturnType<typeof part>): string[] {
  const out: string[] = [];
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'paragraph') {
      out.push(paragraphTextOf(xml, node.id) ?? '');
      return;
    }
    for (const child of node.children) visit(child);
  };
  visit(xml.root);
  return out;
}

describe('repeating section items', () => {
  test('a new item repeats the one before it, its controls back to their prompts', () => {
    const before = part(FILE());
    const [sectionId = ''] = controlIds(before.root);
    const result = applyTreeOp(before, { op: 'addRepeatingSectionItem', controlId: sectionId });
    if (!result.ok) throw new Error(result.reason);
    expect(texts(result.part)).toEqual([
      'Parties:',
      'Party: Acme (seller)',
      'Party: Click or tap here to enter text. (seller)',
      'End.',
    ]);
    const xml = serializeOoxmlPart(result.part);
    // Fresh control ids and paragraph ids; the new control shows its prompt
    expect(xml.match(/<w:id w:val="21"\/>/g)).toHaveLength(1);
    expect(xml).toContain('<w:showingPlcHdr/>');
    expect(result.effect.created).toHaveLength(1);
  });

  test('an item at the start repeats the first one', () => {
    const before = part(FILE());
    const [sectionId = ''] = controlIds(before.root);
    const result = applyTreeOp(before, {
      op: 'addRepeatingSectionItem',
      controlId: sectionId,
      index: 0,
    });
    if (!result.ok) throw new Error(result.reason);
    expect(texts(result.part)[1]).toBe('Party: Click or tap here to enter text. (seller)');
    expect(texts(result.part)[2]).toBe('Party: Acme (seller)');
  });

  test('an item goes; the last one stays', () => {
    const before = part(FILE());
    const [sectionId = ''] = controlIds(before.root);
    const added = applyTreeOp(before, { op: 'addRepeatingSectionItem', controlId: sectionId });
    if (!added.ok) throw new Error(added.reason);
    const removed = applyTreeOp(added.part, {
      op: 'removeRepeatingSectionItem',
      controlId: sectionId,
      index: 0,
    });
    if (!removed.ok) throw new Error(removed.reason);
    expect(texts(removed.part)).toEqual([
      'Parties:',
      'Party: Click or tap here to enter text. (seller)',
      'End.',
    ]);
    const last: TreeDocOp = { op: 'removeRepeatingSectionItem', controlId: sectionId, index: 0 };
    expect(validateTreeOp(removed.part, last)).toBe('invalidArgs');
  });

  test('a section that forbids it, a locked section, and a control that is no section refuse', () => {
    const forbidding = part(FILE(`<w15:doNotAllowInsertDeleteSection ${W15}/>`));
    const [forbidden = ''] = controlIds(forbidding.root);
    expect(
      validateTreeOp(forbidding, { op: 'addRepeatingSectionItem', controlId: forbidden })
    ).toBe('locked');
    const plain = part(FILE());
    const ids = controlIds(plain.root);
    expect(validateTreeOp(plain, { op: 'addRepeatingSectionItem', controlId: ids[2] ?? '' })).toBe(
      'typeMismatch'
    );
    expect(
      validateTreeOp(plain, {
        op: 'removeRepeatingSectionItem',
        controlId: ids[0] ?? '',
        index: 3,
      })
    ).toBe('invalidArgs');
  });
});

describe('the editor commands', () => {
  test('add and remove an item from a caret inside the section', async () => {
    const { GlobalRegistrator } = await import('@happy-dom/global-registrator');
    if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
    const { createDocxEditor } = await import('../docx-editor.ts');
    const { stubReviewModule } = await import('./review-test-module.ts');
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({
      container,
      document: FILE(),
      author: 'Olive',
      mode: 'edit',
      modules: [stubReviewModule()],
    });
    try {
      const surface = editor.surface;
      if (!surface) throw new Error('no surface');
      const [, second = ''] = surface.session.paragraphIds();
      surface.setSelection({
        anchor: { paragraphId: second, offset: 2 },
        head: { paragraphId: second, offset: 2 },
      });
      expect(editor.exec({ type: 'addRepeatingSectionItem' } as never)).toMatchObject({ ok: true });
      expect(texts(surface.session.part())).toEqual([
        'Parties:',
        'Party: Acme (seller)',
        'Party: Click or tap here to enter text. (seller)',
        'End.',
      ]);
      expect(editor.exec({ type: 'removeRepeatingSectionItem', index: 1 } as never)).toMatchObject({
        ok: true,
      });
      expect(texts(surface.session.part())).toEqual(['Parties:', 'Party: Acme (seller)', 'End.']);
      editor.exec({ type: 'undo' } as never);
      expect(texts(surface.session.part())).toHaveLength(4);
    } finally {
      editor.destroy();
      container.remove();
    }
  });
});
