// Typing over a content control's prompt in suggesting mode, as Word records it.
//
// Word removes the prompt without a revision, clears `w:showingPlcHdr`, and proposes only the
// typed words, inside the control and without the prompt's style. Rejecting them empties the
// control, which shows its prompt again from the glossary block it names. The prompt used to be
// struck like text, and the words kept its style and its flag: Accept left a value that still
// looked and acted like a prompt.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { collectReviewItems, serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const GLOSSARY =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument';

const PROMPT = 'Enter the party name.';
const PROMPT_RUN = `<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>${PROMPT}</w:t></w:r>`;

/** "Name: " and a text control showing the glossary's prompt, then " end.". */
const PROMPTED =
  '<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr>' +
  '<w:placeholder><w:docPart w:val="PartyName"/></w:placeholder><w:showingPlcHdr/><w:text/>' +
  `</w:sdtPr><w:sdtContent>${PROMPT_RUN}</w:sdtContent></w:sdt>` +
  '<w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>';
const PROMPT_START = 'Name: '.length;
const PROMPT_END = `Name: ${PROMPT}`.length;

/** A document whose glossary part holds the `PartyName` placeholder block. */
function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${GLOSSARY}" Target="glossary/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
    'word/glossary/document.xml': strToU8(
      `<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr>` +
        '<w:name w:val="PartyName"/><w:category><w:name w:val="General"/>' +
        '<w:gallery w:val="placeholder"/></w:category></w:docPartPr>' +
        `<w:docPartBody><w:p>${PROMPT_RUN}</w:p></w:docPartBody></w:docPart></w:docParts>` +
        '</w:glossaryDocument>'
    ),
  });
}

const mounted: { surface: PaginatedSurface; container: HTMLElement }[] = [];
afterEach(() => {
  for (const { surface, container } of mounted.splice(0)) {
    surface.destroy();
    container.remove();
  }
});

function suggesting(body = PROMPTED): PaginatedSurface {
  const container = document.createElement('div');
  document.body.append(container);
  const opened = mountPaginatedSurface(container, docx(body), { scale: 1, author: 'Bea' });
  if (!opened.ok) throw new Error(opened.reason);
  mounted.push({ surface: opened.surface, container });
  opened.surface.setEditingMode('suggest');
  return opened.surface;
}

function select(surface: PaginatedSurface, start: number, end = start): void {
  const paragraphId = surface.session.paragraphIds()[0]!;
  surface.setSelection({
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  });
}

/** What the control's `w:sdtContent` holds, ids and dates removed. */
function content(surface: PaginatedSurface): string {
  const xml = serializeOoxmlPart(surface.session.part())
    .replace(/ w:date="[^"]*"/g, '')
    .replace(/ w:id="[^"]*"/g, '');
  return /<w:sdtContent>(.*?)<\/w:sdtContent>/s.exec(xml)?.[1] ?? '';
}

const showsPrompt = (surface: PaginatedSurface) =>
  serializeOoxmlPart(surface.session.part()).includes('<w:showingPlcHdr/>');

describe('typing over a prompt in suggesting mode', () => {
  test('proposes the typed words alone, inside the control, one keystroke after another', () => {
    const surface = suggesting();
    select(surface, PROMPT_START, PROMPT_END);
    for (const character of 'Ana') surface.type(character);
    expect(surface.state().lastRejection).toBeFalsy();
    expect(showsPrompt(surface)).toBe(false);
    expect(content(surface)).toBe('<w:ins w:author="Bea"><w:r><w:t>Ana</w:t></w:r></w:ins>');
    expect(surface.session.bodyText()).toBe('Name: Ana end.');
    const items = collectReviewItems({ storyPart: surface.session.part() });
    expect(items.map((item) => (item.kind === 'revision' ? item.revisionKind : item.kind))).toEqual(
      ['insert']
    );
  });

  test('a caret at either edge of the prompt replaces it whole', () => {
    for (const offset of [PROMPT_START, PROMPT_END]) {
      const surface = suggesting();
      select(surface, offset);
      surface.type('X');
      expect(content(surface)).toBe('<w:ins w:author="Bea"><w:r><w:t>X</w:t></w:r></w:ins>');
    }
  });

  test('deleting the selected prompt changes nothing', () => {
    const surface = suggesting();
    const before = serializeOoxmlPart(surface.session.part());
    select(surface, PROMPT_START, PROMPT_END);
    surface.deleteBackward();
    expect(serializeOoxmlPart(surface.session.part())).toBe(before);
  });

  test('Backspace, Delete or a deletion between a block prompt’s paragraphs changes nothing', () => {
    // The mark between them is the prompt's, as its text is: Word changes a prompt only in
    // Design Mode. Striking it proposed a join that Accept made in the prompt itself.
    const twoLines =
      '<w:sdt><w:sdtPr><w:showingPlcHdr/></w:sdtPr><w:sdtContent>' +
      `<w:p>${PROMPT_RUN}</w:p><w:p>${PROMPT_RUN}</w:p></w:sdtContent></w:sdt>` +
      '<w:p><w:r><w:t>end.</w:t></w:r></w:p>';
    for (const press of ['backspace', 'delete', 'across'] as const) {
      const surface = suggesting(twoLines);
      const before = serializeOoxmlPart(surface.session.part());
      const [first = '', second = ''] = surface.session.paragraphIds();
      const anchor =
        press === 'backspace'
          ? { paragraphId: second, offset: 0 }
          : { paragraphId: first, offset: press === 'delete' ? PROMPT.length : 4 };
      const head = press === 'across' ? { paragraphId: second, offset: 4 } : anchor;
      surface.setSelection({ anchor, head });
      if (press === 'delete') surface.deleteForward();
      else surface.deleteBackward();
      expect({ press, rejection: surface.state().lastRejection ?? null }).toEqual({
        press,
        rejection: null,
      });
      expect(serializeOoxmlPart(surface.session.part())).toBe(before);
    }
  });

  test('rejecting shows the glossary’s prompt again, and accepting keeps a plain value', () => {
    for (const op of ['rejectAllRevisions', 'acceptAllRevisions'] as const) {
      const surface = suggesting();
      select(surface, PROMPT_START, PROMPT_END);
      surface.type('Ana');
      expect(surface.session.applyTreeOps([{ op }]).committed).toBe(true);
      const rejected = op === 'rejectAllRevisions';
      expect(showsPrompt(surface)).toBe(rejected);
      expect(content(surface)).toBe(rejected ? PROMPT_RUN : '<w:r><w:t>Ana</w:t></w:r>');
    }
  });

  // Word for Mac, 2026-11-25 (`w58-after-own-delete.docx`): an author's Backspace over their own
  // new words in a control leaves no revision, and the control shows its prompt again: the one it
  // kept when the words replaced it, else Word's blank prompt (five spaces in no style, with no
  // `w:sdtEndPr`) for a text control that names none, here Bea's own new control
  test("deleting one's own words in a control shows its prompt again, with no revision", () => {
    const own =
      '<w:p><w:r><w:t xml:space="preserve">A: </w:t></w:r>' +
      '<w:ins w:id="7" w:author="Bea" w:date="2026-11-25T00:00:00Z"><w:sdt><w:sdtPr><w:text/>' +
      '</w:sdtPr><w:sdtContent><w:r><w:t>ab</w:t></w:r></w:sdtContent></w:sdt></w:ins>' +
      '<w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>';
    for (const [body, start, typed, prompt] of [
      [PROMPTED, PROMPT_START, true, PROMPT_RUN],
      [own, 'A: '.length, false, '<w:r><w:t xml:space="preserve">     </w:t></w:r>'],
    ] as const) {
      const surface = suggesting(body);
      if (typed) {
        select(surface, start, start + 'Enter'.length);
        surface.type('ab');
      }

      select(surface, start, start + 2);
      surface.deleteBackward();
      expect(surface.state().lastRejection).toBeFalsy();
      expect(showsPrompt(surface)).toBe(true);
      expect(content(surface)).toBe(prompt);
      expect(serializeOoxmlPart(surface.session.part())).not.toContain('sdtEndPr');
      expect(collectReviewItems({ storyPart: surface.session.part() }).map((i) => i.kind)).toEqual(
        typed ? [] : ['revision']
      );
    }
  });

  // The same rule for a block control: it shows its prompt again only when the deletion leaves
  // it with no content, so a second line of other words keeps it as it is.
  test("deleting one's own words in a block control shows its prompt only when it empties", () => {
    const own =
      '<w:p><w:ins w:id="7" w:author="Bea" w:date="2026-11-25T00:00:00Z">' +
      '<w:r><w:t>ab</w:t></w:r></w:ins></w:p>';
    const block = (paragraphs: string) =>
      '<w:sdt><w:sdtPr><w:placeholder><w:docPart w:val="PartyName"/></w:placeholder></w:sdtPr>' +
      `<w:sdtContent>${paragraphs}</w:sdtContent></w:sdt><w:p><w:r><w:t>end.</w:t></w:r></w:p>`;
    for (const [paragraphs, emptied] of [
      [own, true],
      [own + '<w:p><w:r><w:t>kept</w:t></w:r></w:p>', false],
    ] as const) {
      const surface = suggesting(block(paragraphs));
      select(surface, 0, 2);
      surface.deleteBackward();
      expect(surface.state().lastRejection).toBeFalsy();
      expect(showsPrompt(surface)).toBe(emptied);
      expect(content(surface)).toContain(emptied ? PROMPT_RUN : '<w:t>kept</w:t>');
      expect(collectReviewItems({ storyPart: surface.session.part() })).toEqual([]);
    }
  });

  // Innermost first, as a decision empties controls: the inner control shows its prompt, and the
  // control around it holds that prompt, so it keeps the inner control (an inline or a block one).
  test("deleting one's own words in a nested control keeps the inner control", () => {
    const inner =
      '<w:sdt><w:sdtPr><w:tag w:val="inner"/><w:text/></w:sdtPr><w:sdtContent>' +
      '<w:ins w:id="7" w:author="Bea" w:date="2026-11-25T00:00:00Z"><w:r><w:t>ab</w:t></w:r>' +
      '</w:ins></w:sdtContent></w:sdt>';
    const outer = (tagged: string) =>
      `<w:sdt><w:sdtPr><w:tag w:val="outer"/></w:sdtPr><w:sdtContent>${tagged}</w:sdtContent></w:sdt>`;
    for (const [body, start] of [
      [`<w:p><w:r><w:t xml:space="preserve">A: </w:t></w:r>${outer(inner)}</w:p>`, 3],
      [`${outer(`<w:p>${inner}</w:p>`)}<w:p><w:r><w:t>end.</w:t></w:r></w:p>`, 0],
    ] as const) {
      const surface = suggesting(body);
      select(surface, start, start + 2);
      surface.deleteBackward();
      const xml = serializeOoxmlPart(surface.session.part());
      expect(surface.state().lastRejection).toBeFalsy();
      expect(xml).toContain('<w:tag w:val="outer"/></w:sdtPr>');
      expect(xml).toContain('<w:tag w:val="inner"/><w:showingPlcHdr/>');
      expect(xml.match(/<w:showingPlcHdr\/>/g)).toHaveLength(1);
    }
  });

  test('undo gives the prompt back as it was', () => {
    const surface = suggesting();
    const before = serializeOoxmlPart(surface.session.part());
    select(surface, PROMPT_START, PROMPT_END);
    surface.type('Ana');
    surface.session.undo();
    expect(serializeOoxmlPart(surface.session.part())).toBe(before);
  });
});

describe('typing over a prompt in editing mode', () => {
  // Word keeps a filled control's placeholder: emptied again, the control shows the prompt it had.
  // A prompt that lived only in the control's content goes to the glossary, as in suggesting mode.
  const ownPrompt =
    '<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:showingPlcHdr/>' +
    `<w:text/></w:sdtPr><w:sdtContent>${PROMPT_RUN}</w:sdtContent></w:sdt>` +
    '<w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>';

  test('keeps the prompt in the glossary; the emptied control shows it again', () => {
    const surface = suggesting(ownPrompt);
    surface.setEditingMode('edit');
    // A caret at the prompt replaces it whole, as in suggesting mode
    select(surface, PROMPT_START);
    surface.type('Ana');
    expect(showsPrompt(surface)).toBe(false);
    expect(serializeOoxmlPart(surface.session.part())).toMatch(
      /<w:placeholder><w:docPart w:val="[^"]+"\/><\/w:placeholder>/
    );
    select(surface, PROMPT_START, PROMPT_START + 'Ana'.length);
    surface.deleteBackward();
    expect(showsPrompt(surface)).toBe(true);
    expect(surface.session.bodyText()).toBe(`Name: ${PROMPT} end.`);
  });
});

describe('typing at the end of a control’s value in suggesting mode', () => {
  test('stays inside the control, as in Word', () => {
    const surface = suggesting(
      '<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:text/></w:sdtPr>' +
        '<w:sdtContent><w:r><w:t>Acme</w:t></w:r></w:sdtContent></w:sdt>' +
        '<w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>'
    );
    select(surface, 'Name: Acme'.length);
    surface.type('s');
    expect(content(surface)).toBe(
      '<w:r><w:t>Acme</w:t></w:r><w:ins w:author="Bea"><w:r><w:t>s</w:t></w:r></w:ins>'
    );
  });

  test('stays inside when the value ends in a revision this keystroke cannot extend', () => {
    // Another author's words, Bea's own from an earlier session, and struck words all end in a
    // wrapper the new words leave. The control still holds them, as in Word: they used to land
    // after the control, outside the value they continue. Each value reads "Acme" in all markup.
    const OLD = 'w:id="1" w:date="2026-01-01T00:00:00Z"';
    for (const value of [
      `<w:ins ${OLD} w:author="Ann"><w:r><w:t>Acme</w:t></w:r></w:ins>`,
      `<w:ins ${OLD} w:author="Bea"><w:r><w:t>Acme</w:t></w:r></w:ins>`,
      `<w:r><w:t>Ac</w:t></w:r><w:del ${OLD} w:author="Ann"><w:r><w:delText>me</w:delText></w:r></w:del>`,
    ]) {
      const surface = suggesting(
        '<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:text/></w:sdtPr>' +
          `<w:sdtContent>${value}</w:sdtContent></w:sdt>` +
          '<w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>'
      );
      select(surface, 'Name: Acme'.length);
      surface.type('s');
      surface.type('t');
      expect(content(surface)).toBe(
        `${value.replaceAll(` ${OLD}`, '')}<w:ins w:author="Bea"><w:r><w:t>st</w:t></w:r></w:ins>`
      );
      expect(surface.state().selection.head.offset).toBe('Name: Acmest'.length);
    }
  });
});

describe('typing in a block control in suggesting mode', () => {
  // The block control holds the caret's whole paragraph, so it names no place inside it.
  // Asking for one put every keystroke in a `w:ins` of its own, beside the last.
  const block = (properties: string, paragraph: string) =>
    `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent><w:p>${paragraph}</w:p></w:sdtContent></w:sdt>`;
  const paragraphContent = (surface: PaginatedSurface) =>
    content(surface).replace(/<\/?w:p\b[^>]*>/g, '');
  const ins = (text: string) => `<w:ins w:author="Bea"><w:r><w:t>${text}</w:t></w:r></w:ins>`;

  test('extends one suggestion key after key, at the end and in the middle', () => {
    for (const [offset, expected] of [
      [4, `<w:r><w:t>Acme</w:t></w:r>${ins('xyz')}`],
      [2, `<w:r><w:t>Ac</w:t></w:r>${ins('xyz')}<w:r><w:t>me</w:t></w:r>`],
    ] as const) {
      const surface = suggesting(block('<w:richText/>', '<w:r><w:t>Acme</w:t></w:r>'));
      select(surface, offset);
      for (const character of 'xyz') surface.type(character);
      expect(paragraphContent(surface)).toBe(expected);
      expect(surface.state().selection.head.offset).toBe(offset + 3);
    }
  });

  test('typed over its prompt key after key, proposes the words as one suggestion', () => {
    const surface = suggesting(
      block(
        '<w:placeholder><w:docPart w:val="PartyName"/></w:placeholder><w:showingPlcHdr/>',
        PROMPT_RUN
      )
    );
    select(surface, 0, PROMPT.length);
    for (const character of 'Ana') surface.type(character);
    expect(paragraphContent(surface)).toBe(ins('Ana'));
    expect(surface.state().selection.head.offset).toBe(3);
  });
});
