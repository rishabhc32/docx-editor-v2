// The name a kept prompt gets in the glossary comes from its content, so equal prompts share one
// block. A save writes a text's `xml:space="preserve"` only where the text needs it, so the same
// prompt read back from a save has no such attribute where it had one: it keeps its name.

import { describe, expect, test } from 'bun:test';
import { contentControlsIn, readOoxmlPart, type OoxmlPart } from '../index.ts';
import { promptToRetain } from '../store/content-control-prompt-glossary.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function parse(text: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:sdt><w:sdtPr><w:showingPlcHdr/></w:sdtPr>` +
      `<w:sdtContent><w:p><w:r>${text}</w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>`,
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

/** The glossary name the prompt of the part's one control gets. */
function nameOf(text: string): string | undefined {
  const part = parse(text);
  const [control] = contentControlsIn(part.root);
  if (control === undefined) {
    throw new Error('no control');
  }

  return promptToRetain(part, control.node, { retainPrompt: () => undefined })?.name;
}

describe("a kept prompt's glossary name", () => {
  test('is the same with and without a space attribute its text does not need', () => {
    const name = nameOf('<w:t>Click or tap here.</w:t>');
    expect(name).toMatch(/^[0-9A-F]{32}$/);
    expect(nameOf('<w:t xml:space="preserve">Click or tap here.</w:t>')).toBe(name);
  });

  test('differs for words with an edge space, which the attribute keeps', () => {
    const name = nameOf('<w:t>Click or tap here.</w:t>');
    expect(nameOf('<w:t xml:space="preserve">Click or tap here. </w:t>')).not.toBe(name);
  });
});
