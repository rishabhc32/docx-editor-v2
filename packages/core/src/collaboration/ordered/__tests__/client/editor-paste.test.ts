// Paste with formatting from the editor in an ordered room (task 9.10): a fragment with a
// style, a numbered list, bold text, a link and a table, copied from another document. The
// organiser's paste is a direct edit that undo and redo turn back and again; a party's is a
// tracked insertion, as Word tracks a paste, that the organiser decides; the list and the style
// the document lacks travel with the paste, and go from the file when it is rejected. A pasted
// picture's bytes travel once, by hash, as a new picture's do. A picture's address never enters
// the room (Word loads it when it opens the file, design D18): a picture with an embedded copy
// keeps the copy without the address, one with no copy is refused by name; never fetched.

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import { orderedBlobHash } from '../../protocol/blobs.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { RED_PNG } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  FILE,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { copied, copiedPicture, copiedShape, LINKED_PICTURE } from './paste-harness.ts';
import { reviewTestModule } from './review-module.ts';

const review = reviewTestModule();

afterEach(closeAll);

const kinds = (room: ReturnType<typeof localRoom>): string[] =>
  room.room.joinSnapshot(0).map((change) => change.kind);

function saved(room: ReturnType<typeof localRoom>): { styles: string; numbering: string } {
  const base = readOoxmlPackage(FILE);
  if (!base.ok) throw new Error('no base');
  const bytes = orderedPackageBytes(room.room.document, base.package);
  const pkg = readOoxmlPackage(bytes);
  if (!pkg.ok) throw new Error('the saved file does not open');
  openOrderedWordFile(bytes, { file: 2 });
  const text = (suffix: string) => {
    const part = [...pkg.package.parts.values()].find((x) => x.name.endsWith(suffix));
    return part ? JSON.stringify(part.root) : '';
  };
  return { styles: text('/styles.xml'), numbering: text('/numbering.xml') };
}

const checked = (room: ReturnType<typeof localRoom>, people: readonly Person[]) => {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
  }
};

const cards = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .filter((placement) => placement.kind === 'revision')
    .map((placement) => placement.key);

describe('paste with formatting in an ordered room', () => {
  test("the organiser's paste is direct; the list and the style travel; undo and redo", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    const { text, html } = copied(6);
    caretAt(olive, 1, 0);
    expect(surface(olive).pasteRich(text, html)).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['insert-fragment']);
    checked(room, [olive, ana]);
    const pasted = editorTexts(ana);
    expect(pasted).toContain('First term');
    expect(pasted).toContain('Fee');
    expect(editorXml(ana)).toContain('ClauseTitle');
    expect(editorXml(ana)).not.toContain('<w:ins ');
    const file = saved(room);
    expect(file.styles).toContain('ClauseTitle');
    expect(file.numbering).toContain('"abstractNum"');

    expect(olive.handle.session.undo()).toBe(true);
    await settle();
    checked(room, [olive, ana]);
    expect(editorTexts(ana)).toEqual(['The fee is 100 USD.', 'Due in 30 days.', 'Signed here.']);
    expect(saved(room).styles).not.toContain('ClauseTitle');
    expect(olive.handle.session.redo()).toBe(true);
    await settle();
    checked(room, [olive, ana]);
    expect(editorTexts(ana)).toEqual(pasted);
  });

  test("a party's paste is a tracked insertion the organiser rejects, and the list goes", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    const { text, html } = copied(3);
    caretAt(ana, 0, 4);
    expect(surface(ana).pasteRich(text, html)).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['insert-fragment']);
    checked(room, [olive, ana]);
    expect(editorXml(olive)).toContain('<w:ins ');
    expect(saved(room).numbering).toContain('"abstractNum"');
    const [card] = cards(olive);
    if (card === undefined) throw new Error('no card');
    expect(olive.editor.rejectReviewItem(card).ok).toBe(true);
    await settle();
    checked(room, [olive, ana]);
    expect(editorTexts(olive)).toEqual(['The fee is 100 USD.', 'Due in 30 days.', 'Signed here.']);
    const file = saved(room);
    expect(file.numbering).not.toContain('"abstractNum"');
    expect(file.styles).not.toContain('ClauseTitle');
  });

  test("a party's paste inside another person's pending insertion is the nesting rule, reported once", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ben = await join(room, 'Ben', { review });
    const ana = await join(room, 'Ana', { review });
    caretAt(ben, 0, 4);
    surface(ben).insertPlainText('really ');
    await settle();
    // Three paragraphs inside a pending insertion: the editor pastes text and cuts it, three
    // changes the room refuses, one cause the person sees
    const { text, html } = copied(3);
    caretAt(ana, 0, 7);
    surface(ana).pasteRich(text, html);
    await settle();
    expect(ana.refusals).toEqual([expect.objectContaining({ reason: 'nested' })]);
    expect(kinds(room)).toEqual(['insert']);
    for (const person of [olive, ben, ana]) expect(sameAsRoom(person, room)).toBe(true);
  });
});

const RED_HASH = orderedBlobHash(RED_PNG);

/** The media parts of a person's editor package, by name. */
function editorMedia(person: Person): Uint8Array[] {
  const out: Uint8Array[] = [];
  for (const [name, bytes] of surface(person).session.currentPackage().partBytes)
    if (name.startsWith('/word/media/')) out.push(bytes);
  return out;
}

/** The media parts of the room's saved file, with the room's blobs. */
function savedMedia(room: ReturnType<typeof localRoom>): string[] {
  const base = readOoxmlPackage(FILE);
  if (!base.ok) throw new Error('no base');
  const bytes = orderedPackageBytes(room.room.document, base.package, {
    blobs: (hash) => room.blob(hash),
  });
  const pkg = readOoxmlPackage(bytes);
  if (!pkg.ok) throw new Error('the saved file does not open');
  openOrderedWordFile(bytes, { file: 2 });
  return [...pkg.package.partBytes.keys()].filter((name) => name.startsWith('/word/media/'));
}

/** The addresses of the saved file's pictures linked by address. */
function savedImageLinks(room: ReturnType<typeof localRoom>): string[] {
  const base = readOoxmlPackage(FILE);
  if (!base.ok) throw new Error('no base');
  const bytes = orderedPackageBytes(room.room.document, base.package, {
    blobs: (hash) => room.blob(hash),
  });
  const pkg = readOoxmlPackage(bytes);
  if (!pkg.ok) throw new Error('the saved file does not open');
  return pkg.package.externalTargets
    .filter((target) => target.type.endsWith('/image'))
    .map((target) => target.rawTarget);
}

const showsPicture = (person: Person): boolean =>
  editorMedia(person).some((bytes) => bytes.length === RED_PNG.length);

describe('a pasted picture in an ordered room', () => {
  test("the organiser's: its bytes go once by hash, every copy shows it; undo and redo", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    const { text, html } = copiedPicture();
    caretAt(olive, 1, 3);
    expect(surface(olive).pasteRich(text, html)).toBe(true);
    await settle(16);
    expect(kinds(room)).toEqual(['insert-fragment']);
    expect(room.blob(RED_HASH)).toEqual(RED_PNG);
    checked(room, [olive, ana]);
    // The other copy read the bytes back from the room, by hash
    expect(showsPicture(ana)).toBe(true);
    expect(editorXml(ana)).toContain('<w:drawing>');
    expect(savedMedia(room)).toHaveLength(1);

    expect(olive.handle.session.undo()).toBe(true);
    await settle(16);
    checked(room, [olive, ana]);
    expect(editorXml(ana)).not.toContain('<w:drawing>');
    expect(savedMedia(room)).toEqual([]);
    expect(olive.handle.session.redo()).toBe(true);
    await settle(16);
    checked(room, [olive, ana]);
    expect(showsPicture(ana)).toBe(true);
  });

  test("a party's is a tracked insertion; a reject takes the picture and its media", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    const { text, html } = copiedPicture();
    caretAt(ana, 0, 4);
    expect(surface(ana).pasteRich(text, html)).toBe(true);
    await settle(16);
    checked(room, [olive, ana]);
    expect(editorXml(olive)).toMatch(
      /<w:ins [^>]*w:author="Ana"[^>]*>(?:(?!<\/w:ins>).)*<w:drawing>/s
    );
    expect(showsPicture(olive)).toBe(true);
    expect(savedMedia(room)).toHaveLength(1);
    const [card] = cards(olive);
    if (card === undefined) throw new Error('no card');
    expect(olive.editor.rejectReviewItem(card).ok).toBe(true);
    await settle(16);
    checked(room, [olive, ana]);
    expect(editorTexts(olive)).toEqual(['The fee is 100 USD.', 'Due in 30 days.', 'Signed here.']);
    expect(savedMedia(room)).toEqual([]);
  });

  // Word loads a linked picture's address when it opens the file, with no click, so a room
  // never writes one (design D18): the picture keeps its embedded copy, never fetched
  test('a picture with an embedded copy loses its address in every copy and the file', async () => {
    const fetched = withoutNetwork();
    try {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
      const ana = await join(room, 'Ana', { review });
      const { text, html } = copiedPicture('both');
      caretAt(olive, 1, 3);
      expect(surface(olive).pasteRich(text, html)).toBe(true);
      await settle(16);
      expect(kinds(room)).toEqual(['insert-fragment']);
      expect(room.blob(RED_HASH)).toEqual(RED_PNG);
      checked(room, [olive, ana]);
      // Word's Smart cut and paste puts a space between "Due" and the pasted word
      expect(editorTexts(ana)[1]).toBe('Due Seal \uFFFC here in 30 days.');
      // A party's paste is a tracked insertion of the picture's embedded copy alone
      caretAt(ana, 2, 0);
      expect(surface(ana).pasteRich(text, html)).toBe(true);
      await settle(16);
      checked(room, [olive, ana]);
      expect(editorXml(olive)).toMatch(
        /<w:ins [^>]*w:author="Ana"[^>]*>(?:(?!<\/w:ins>).)*<w:drawing>/s
      );
      // A VML picture keeps its embedded copy without `o:href`, or without its link `r:href`
      for (const kind of ['vmlBoth', 'vmlHref'] as const) {
        const vml = copiedShape(kind);
        caretAt(olive, 0, 0);
        expect(surface(olive).pasteRich(vml.text, vml.html)).toBe(true);
        await settle(16);
        checked(room, [olive, ana]);
      }
      expect(editorXml(ana)).toContain('<v:imagedata r:id=');
      for (const person of [olive, ana]) {
        expect(showsPicture(person)).toBe(true);
        expect(editorXml(person)).not.toContain('r:link');
        expect(editorXml(person)).not.toContain('o:href');
        expect(editorXml(person)).not.toContain('r:href');
        const rels = surface(person).session.currentPackage().externalTargets;
        expect(rels.filter((r) => r.type.endsWith('/image'))).toEqual([]);
        expect(person.container.innerHTML).not.toContain(LINKED_PICTURE);
      }
      // One embedded copy for each paste, and no address
      expect(savedMedia(room)).toHaveLength(4);
      expect(savedImageLinks(room)).toEqual([]);
      expect(JSON.stringify([...room.room.document.resources.added.values()])).not.toContain(
        LINKED_PICTURE
      );
    } finally {
      fetched.restore();
    }
    expect(fetched.addresses).toEqual([]);
  });

  test('a picture that only an address names is refused by name, for the organiser and a party', async () => {
    const fetched = withoutNetwork();
    try {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
      const ana = await join(room, 'Ana', { review });
      // A DrawingML picture by `r:link`, and a VML picture whose relationship is external
      for (const { text, html } of [copiedPicture('linked'), copiedShape('vmlLinked')])
        for (const person of [olive, ana]) {
          const before = editorXml(person);
          caretAt(person, 1, 3);
          expect(surface(person).pasteRich(text, html)).toBe(false);
          await settle(16);
          expect(person.refusals).toEqual([
            {
              code: 'ordered-refused',
              reason: 'external-picture-link',
              detail: 'a picture by address',
            },
          ]);
          expect(editorXml(person)).toBe(before);
          expect(person.container.innerHTML).not.toContain(LINKED_PICTURE);
          person.refusals.length = 0;
        }
      expect(room.room.order).toBe(0);
      checked(room, [olive, ana]);
      expect(savedImageLinks(room)).toEqual([]);
    } finally {
      fetched.restore();
    }
    expect(fetched.addresses).toEqual([]);
  });
});

describe("a picture's alt text in an ordered room", () => {
  // Word 365 writes the "decorative" flag of a picture's alt text in its drawing's extension
  // list (`adec:decorative`): inert data, which the room keeps as the editor alone keeps it
  test('a picture marked decorative keeps the flag in every copy and the file', async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    const { text, html } = copiedPicture('decorative');
    caretAt(olive, 1, 3);
    expect(surface(olive).pasteRich(text, html)).toBe(true);
    await settle(16);
    caretAt(ana, 2, 0);
    expect(surface(ana).pasteRich(text, html)).toBe(true);
    await settle(16);
    expect(kinds(room)).toEqual(['insert-fragment', 'insert-fragment']);
    checked(room, [olive, ana]);
    const flags = (xml: string) => xml.split('<adec:decorative').length - 1;
    for (const person of [olive, ana]) expect(flags(editorXml(person))).toBe(2);
    const base = readOoxmlPackage(FILE);
    if (!base.ok) throw new Error('no base');
    const saved = unzipSync(
      orderedPackageBytes(room.room.document, base.package, { blobs: (hash) => room.blob(hash) })
    );
    const body = saved['word/document.xml'];
    if (body === undefined) throw new Error('no body');
    expect(flags(strFromU8(body))).toBe(2);
    expect(strFromU8(body)).toMatch(
      /<w:ins [^>]*w:author="Ana"[^>]*>(?:(?!<\/w:ins>).)*adec:decorative/s
    );
  });
});

/** A stand-in for the network that records every address a test asks it for. */
function withoutNetwork(): { readonly addresses: string[]; restore(): void } {
  const addresses: string[] = [];
  const realFetch = globalThis.fetch;
  globalThis.fetch = ((input: Parameters<typeof fetch>[0]) => {
    addresses.push(String(input));
    return Promise.reject(new Error('no network in this test'));
  }) as typeof fetch;
  return {
    addresses,
    restore: () => {
      globalThis.fetch = realFetch;
    },
  };
}
