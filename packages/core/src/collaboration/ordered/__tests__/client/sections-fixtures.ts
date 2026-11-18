// Shared fixtures of the header, footer and section tests (task 9.8): the files, a room with an
// organiser and a party, and readers of an editor's save.

import { expect } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { serializeOoxmlPart, type OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import {
  composeOrderedPart,
  orderedParagraphCount,
  orderedViewParagraph,
} from '../../compose/compose.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, R, t, WML_CT } from '../word-fixtures.ts';
import { join, localRoom, ORGANISER, sameAsRoom, surface, type Person } from './editor-harness.ts';

/** The main story's paragraph texts in a document. */
export function listTextsOf(doc: OrderedDocument): string[] {
  const out: string[] = [];
  for (let k = 0; k < orderedParagraphCount(doc, '/word/document.xml'); k += 1)
    out.push(
      orderedViewParagraph(doc, '/word/document.xml', k)
        ?.chars.map((c) => c.ch)
        .join('') ?? ''
    );
  return out;
}

export const HEADER = '/word/header1.xml';

/** Two paragraphs, a default header and a body section that names it. */
export const FILE = docx(
  p(t('First page text.')) +
    p(t('More text.')) +
    `<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:pgSz w:w="12240" w:h="15840"/>` +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>',
  {
    parts: {
      'word/header1.xml': `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>`,
    },
    relationships: `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes: `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>`,
  }
);

/** Open the default header of section 0 for editing, with the caret at an offset. */
export function inHeader(person: Person, offset: number): void {
  const s = surface(person);
  const rId = s.session.headerFooterResolutionBySection()[0]?.headers.get('default')?.rId;
  if (!rId) throw new Error('no header');
  if (!s.enterHeaderFooter({ rId, kind: 'header', sectionIndex: 0, variant: 'default' }))
    throw new Error('the header did not open');
  const scope = { kind: 'headerFooter' as const, rId };
  const [paragraphId] = s.session.paragraphIdsIn(scope);
  if (paragraphId === undefined) throw new Error('no header paragraph');
  s.setSelection({
    anchor: { paragraphId, offset },
    head: { paragraphId, offset },
  });
}

/** The parts of the editor's save, as text, by name. */
export async function saved(person: Person): Promise<Record<string, string>> {
  const files = unzipSync(new Uint8Array(await person.editor.save()));
  const out: Record<string, string> = {};
  for (const [name, bytes] of Object.entries(files)) out[name] = strFromU8(bytes);
  return out;
}

/** The body's own `w:sectPr` of a saved main part. */
export const bodySection = (xml: string): string =>
  /<w:sectPr\b(?:(?!<w:sectPr\b)[\s\S])*?<\/w:sectPr>(?=<\/w:body>)/.exec(xml)?.[0] ?? '';

/** An organiser and a party in one room on FILE, with the suggestions-only policy. */
export async function room2(file = FILE) {
  const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { wordFile: file });
  const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: file });
  return { room, ana, olive };
}

/** Every copy is the room's, and the editors show the list. */
export function converged(room: ReturnType<typeof localRoom>, ...people: Person[]): void {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    // A publish the editor refused leaves the session in error: the editor no longer shows the list
    expect(person.handle.session.status()).not.toBe('error');
    expect(sameAsRoom(person, room)).toBe(true);
    showsTheList(person);
  }
}

const textOf = (part: OoxmlPart): string => serializeOoxmlPart(part).replace(/<[^>]+>/g, '');

/** Each story the editor shows says what the person's copy composes for it, word for word. */
export function showsTheList(person: Person): void {
  const doc = person.handle.session.orderedDocument();
  const shown = surface(person).session.currentPackage();
  for (const name of doc.stories.keys()) {
    const part = shown.parts.get(name);
    // A story a change added that nothing names is not written
    if (part === undefined) continue;
    expect([name, textOf(part)]).toEqual([name, textOf(composeOrderedPart(doc, name).part)]);
  }
}

/** Two sections: the first ends at its paragraph with the header; the second shows it too. */
export const TWO = docx(
  p(
    t('Section one.'),
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/>' +
      '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
  ) +
    p(t('Section two.')) +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
  {
    parts: {
      'word/header1.xml': `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>`,
    },
    relationships: `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes: `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>`,
  }
);

export const lifecycle = (
  person: Person,
  op: Parameters<NonNullable<ReturnType<typeof surface>['applyHeaderFooterLifecycle']>>[0]
) => surface(person).applyHeaderFooterLifecycle?.(op);

/** The header parts a save holds, by name, with their text. */
export async function headers(person: Person): Promise<Record<string, string>> {
  const out: Record<string, string> = {};
  for (const [name, xml] of Object.entries(await saved(person)))
    if (/^word\/header[^/]*\.xml$/.test(name)) out[name] = xml.replace(/<[^>]+>/g, '');
  return out;
}
