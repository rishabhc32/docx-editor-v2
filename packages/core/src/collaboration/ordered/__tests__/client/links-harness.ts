// Shared steps of the links and fields tests (group 9.5): the editor alone making the same edit,
// and a package compared as Word reads it, past the names that differ by design (revision ids
// and dates, paragraph ids, relationship ids, node ids).

import { expect } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { stubReviewModule } from '../../../../editor/__tests__/review-test-module.ts';
import {
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPackage } from '../../compose/package.ts';
import type { OrderedDocument } from '../../model/types.ts';
import type { EditorModule } from '../../../../contracts/modules.ts';
import {
  join,
  localRoom,
  ORGANISER,
  PARTY,
  sameAsRoom,
  screenFollowsList,
  settle,
  type Editor,
  type Person,
} from './editor-harness.ts';

/** Put an editor's selection in the index-th body paragraph. */
export function select(editor: Editor, index: number, offset: number, end = offset): void {
  const s = editor.surface;
  if (!s) throw new Error('no surface');
  const id = s.session.paragraphIds()[index];
  if (id === undefined) throw new Error(`no paragraph ${index}`);
  s.setSelection({ anchor: { paragraphId: id, offset }, head: { paragraphId: id, offset: end } });
}

/** The editor alone (no room) opens `file`, makes the edits, and gives its package. */
export async function alone(
  file: Uint8Array,
  author: string,
  mode: 'suggesting' | 'edit',
  act: (editor: Editor) => unknown,
  review: EditorModule = stubReviewModule()
): Promise<OoxmlPackage> {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: file, author, mode, modules: [review] });
  try {
    await settle();
    act(editor);
    await settle();
    const s = editor.surface;
    if (!s) throw new Error('no surface');
    return s.session.currentPackage();
  } finally {
    editor.destroy();
    container.remove();
  }
}

/** The room's document saved as a Word file and read again: a refused read fails the test. */
export function roomFile(doc: OrderedDocument, file: Uint8Array): OoxmlPackage {
  const { base } = openOrderedWordFile(file, { file: 1 });
  const bytes = writeOoxmlPackage(composeOrderedPackage(doc, base));
  const read = readOoxmlPackage(bytes, {});
  if (!read.ok) throw new Error(`the saved file does not open: ${read.reason}`);
  return read.package;
}

const MAIN = '/word/document.xml';

/**
 * The main story as Word reads it: relationship ids read as their addresses, and revision ids,
 * dates, paragraph ids, rsids and namespace declarations left out, and neighbour runs with the
 * same properties joined (Word reads them as one).
 */
export function storyAsRead(pkg: OoxmlPackage, part = MAIN): string {
  const story = pkg.parts.get(part);
  if (!story) throw new Error(`no ${part}`);
  const targets = new Map<string, string>();
  for (const target of pkg.externalTargets)
    if (target.ownerPart === part) targets.set(target.id, target.rawTarget);
  return joinedRuns(
    serializeOoxmlPart(story)
      .replace(/ r:id="([^"]*)"/g, (_m, id: string) => ` r:id="${targets.get(id) ?? `?${id}`}"`)
      .replace(/ xmlns:\w+="[^"]*"/g, '')
      .replace(/ w:(id|date|rsid\w*)="[^"]*"/g, '')
      .replace(/ w14:(paraId|textId)="[^"]*"/g, '')
      .replace(/ xml:space="preserve"/g, '')
  );
}

/** Neighbour runs with the same properties read as one run, and neighbour texts as one text. */
function joinedRuns(xml: string): string {
  const neighbours =
    /<w:r>((?:<w:rPr>(?:(?!<\/w:rPr>).)*<\/w:rPr>)?)((?:(?!<\/w:r>).)*)<\/w:r><w:r>\1(?!<w:rPr>)/;
  let joined = xml;
  for (let before = ''; before !== joined; ) {
    before = joined;
    joined = joined.replace(neighbours, '<w:r>$1$2');
  }
  return joined.replace(/<\/w:t><w:t>/g, '');
}

/** A style's element, as Word reads it, or null. */
export function styleAsRead(pkg: OoxmlPackage, styleId: string): string | null {
  for (const part of pkg.parts.values()) {
    if (!part.name.endsWith('/styles.xml')) continue;
    const xml = serializeOoxmlPart(part).replace(/ xmlns:\w+="[^"]*"/g, '');
    const match = new RegExp(`<w:style [^>]*w:styleId="${styleId}"[^>]*>.*?</w:style>`).exec(xml);
    return match?.[0] ?? null;
  }
  return null;
}

/** The addresses of the main story's link relationships. */
export function linkAddresses(pkg: OoxmlPackage): string[] {
  return pkg.externalTargets
    .filter((t) => t.ownerPart === MAIN && t.type.endsWith('/hyperlink'))
    .map((t) => t.rawTarget)
    .sort();
}

export type Act = (editor: Editor) => unknown;

/** Run one action on a person's editor and let the room settle. */
export async function act(person: Person, run: Act): Promise<unknown> {
  const result = run(person.editor);
  await settle();
  return result;
}

/** Both copies are the room's, and each screen shows its list. */
export function converged(room: ReturnType<typeof localRoom>, people: readonly Person[]): void {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
  }
}

/** Steps run in a room by one person, and by the editor alone: the saved stories agree. */
export async function sameAsAlone(
  file: Uint8Array,
  mode: 'suggesting' | 'edit',
  steps: readonly Act[],
  part = MAIN
): Promise<{ room: ReturnType<typeof localRoom>; people: Person[] }> {
  const room = localRoom(file);
  const login = mode === 'edit' ? ORGANISER : PARTY;
  const person = await join(room, 'Ana', { mode, login, wordFile: file });
  const other = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: file });
  for (const step of steps) expect(await act(person, step)).toMatchObject({ ok: true });
  converged(room, [person, other]);
  const editorAlone = await alone(file, 'Ana', mode, (editor) => {
    for (const step of steps) step(editor);
  });
  const saved = roomFile(room.room.document, file);
  expect(storyAsRead(saved, part)).toBe(storyAsRead(editorAlone, part));
  return { room, people: [person, other] };
}
