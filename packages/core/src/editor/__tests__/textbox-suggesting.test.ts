// Suggesting inside a text box. Word lets a person type in a text box with Track Changes on and
// records the typing as tracked changes inside `w:txbxContent`; the editor did not let a
// suggesting person enter a text box at all. It now does, and every edit there is tracked.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
import { expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { createDocxEditor } from '../docx-editor.ts';
import { stubReviewModule } from './review-test-module.ts';
import { textboxDocx } from './textbox-editing-fixture.ts';

const box = (xml: string) => /<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/.exec(xml)?.[1] ?? '';

test('a suggesting person types, deletes and presses Enter in a text box, all tracked', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: textboxDocx(),
    author: 'Ana',
    mode: 'suggesting',
    modules: [stubReviewModule()],
  });
  try {
    const surface = editor.surface!;
    const match = editor.findMatches('boxed needle')[0]!;
    expect(match.scope?.kind).toBe('frame');
    expect(surface.setActiveScope(match.scope!)).toBe(true);
    expect(editor.exec({ type: 'insertText', text: 'New ' }).ok).toBe(true);
    surface.deleteForward();
    surface.splitParagraph();
    const xml = strFromU8(unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!);
    const content = box(xml);
    expect(content).toMatch(/<w:ins [^>]*w:author="Ana"[^>]*>[\s\S]*New /);
    expect(content).toMatch(
      /<w:del [^>]*w:author="Ana"[^>]*>[\s\S]*<w:delText[^>]*>b<\/w:delText>/
    );
    expect(content.match(/<w:p[ >]/g)?.length).toBe(2);
    // Nothing outside the text box changed
    expect(xml.replace(content, '')).not.toContain('w:author="Ana"');
  } finally {
    editor.destroy();
  }
});

test('viewing still keeps a text box closed', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: textboxDocx(),
    mode: 'view',
  });
  try {
    const match = editor.findMatches('boxed needle')[0]!;
    expect(editor.surface!.setActiveScope(match.scope!)).toBe(false);
  } finally {
    editor.destroy();
  }
});
