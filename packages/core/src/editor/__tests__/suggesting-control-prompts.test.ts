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

  test('undo gives the prompt back as it was', () => {
    const surface = suggesting();
    const before = serializeOoxmlPart(surface.session.part());
    select(surface, PROMPT_START, PROMPT_END);
    surface.type('Ana');
    surface.session.undo();
    expect(serializeOoxmlPart(surface.session.part())).toBe(before);
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
});
