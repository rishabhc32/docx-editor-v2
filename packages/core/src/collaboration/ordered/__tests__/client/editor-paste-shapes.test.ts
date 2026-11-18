// Paste of shapes, VML pictures and SVG versions in an ordered room (group 9.10): a text box as
// Word writes it (a shape in markup compatibility, its VML copy beside it), a group of a shape
// and a picture, a VML picture of an EMF image, and a picture with an SVG version, each copied
// from a Word file. The organiser's
// paste is a direct edit that undo and redo turn back and again; a party's is a tracked
// insertion that a reject takes out of the file with its media. Every copy equals the room, the
// room's list equals what the editor alone pastes, and the saved file reopens with the same view.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import type { OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { orderedBlobHash } from '../../protocol/blobs.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { grade, normalizedBody } from '../matrix/compare.ts';
import { checkSaveAndLoad } from '../sessions/session.ts';
import { emfBytes, packageOf, RED_PNG, RED_SVG } from '../word-fixtures.ts';
import {
  BODY,
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
import { copiedShape, type Copied, type SHAPES } from './paste-harness.ts';
import { reviewTestModule } from './review-module.ts';

const review = reviewTestModule();

afterEach(closeAll);

type Room = ReturnType<typeof localRoom>;

interface Shape {
  readonly kind: keyof typeof SHAPES;
  readonly name: string;
  /** What the pasted object shows in a copy's file. */
  readonly shows: RegExp;
  /** The media it brings, by bytes. */
  readonly media: readonly Uint8Array[];
  /**
   * The saved file reopens with the same view (`exact`), or with the same body but for the
   * text box's VML copy, which a save writes from the shape's text box as Word does: its
   * paragraph ids and declarations (`body`; a file with a text box does the same with no edit
   * at all, group 9.9).
   */
  readonly reopen: 'exact' | 'body';
  /**
   * What the room holds of the editor alone's paste, when the format leaves out what a
   * namespace it does not hold extends (default: the same).
   */
  readonly roomForm?: (alone: string) => string;
}

const SHAPE_KINDS: readonly Shape[] = [
  {
    kind: 'textbox',
    name: 'a text box shape',
    shows: /<mc:Choice [^>]*Requires="wps"[^>]*>.*<wps:wsp>.*Box text.*<mc:Fallback>.*<v:textbox>/s,
    media: [],
    reopen: 'exact',
  },
  {
    kind: 'group',
    name: 'a group of a shape and a picture',
    shows: /<mc:Choice [^>]*Requires="wpg"[^>]*>.*<wpg:wgp>.*<wps:wsp>.*<pic:pic>.*<v:group /s,
    media: [RED_PNG],
    reopen: 'exact',
  },
  {
    kind: 'canvas',
    name: 'a drawing canvas with a shape and a picture',
    shows:
      /<mc:Choice [^>]*Requires="wpc"[^>]*>.*<wpc:wpc>.*<wpc:bg>.*<wps:wsp>.*<pic:pic>.*<v:group /s,
    media: [RED_PNG],
    reopen: 'exact',
  },
  {
    kind: 'vml',
    name: 'a VML picture of an EMF image',
    shows: /<w:pict>.*<v:imagedata [^>]*r:id="/s,
    media: [emfBytes()],
    reopen: 'exact',
  },
  {
    kind: 'vmlRelid',
    name: 'a VML picture whose image an Office relationship names',
    shows: /<w:pict>.*<v:imagedata [^>]*o:relid="/s,
    media: [emfBytes()],
    reopen: 'exact',
  },
  {
    kind: 'vmlPict',
    name: 'a VML picture with a second image by r:pict',
    shows: /<w:pict>.*<v:imagedata [^>]*r:pict="/s,
    media: [emfBytes(), RED_SVG],
    reopen: 'exact',
  },
  {
    kind: 'sketched',
    name: 'a picture with an extension of a namespace the format does not hold',
    shows: /<a:blip r:embed="[^"]+"><a:extLst><a:ext uri="\{FF2B5EF4/s,
    media: [RED_PNG],
    reopen: 'exact',
    roomForm: (alone) => alone.replace(/<a:ext uri="\{C807C97D[^"]*">.*?<\/a:ext>/s, ''),
  },
  {
    kind: 'model3d',
    name: 'a 3D model, whose namespace the format does not hold, as its fallback picture',
    shows: /<w:r><w:drawing>.*<a:blip r:embed=/s,
    media: [RED_PNG],
    reopen: 'exact',
    roomForm: (alone) =>
      alone.replace(
        /<mc:AlternateContent><mc:Choice [^>]*>.*?<\/mc:Choice><mc:Fallback>(.*?)<\/mc:Fallback><\/mc:AlternateContent>/s,
        '$1'
      ),
  },
  {
    kind: 'svg',
    name: 'a picture with an SVG version',
    shows: /<a:blip [^>]*r:embed="[^"]+">.*<\w+:svgBlip [^>]*r:embed="/s,
    media: [RED_PNG, RED_SVG],
    reopen: 'exact',
  },
];

const BEFORE = ['The fee is 100 USD.', 'Due in 30 days.', 'Signed here.'];

const kinds = (room: Room): string[] => room.room.joinSnapshot(0).map((change) => change.kind);

const checked = (room: Room, people: readonly Person[]) => {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
  }
};

/** The room's saved file, with the room's blobs, opened again. */
function saved(room: Room) {
  const bytes = orderedPackageBytes(room.room.document, packageOf(FILE), {
    blobs: (hash) => room.blob(hash),
  });
  const pkg = readOoxmlPackage(bytes);
  if (!pkg.ok) throw new Error('the saved file does not open');
  return { pkg: pkg.package, doc: openOrderedWordFile(bytes, { file: 2 }).document };
}

/** The media parts' bytes of the room's saved file. */
const savedMedia = (room: Room): Uint8Array[] =>
  [...saved(room).pkg.partBytes]
    .filter(([name]) => name.startsWith('/word/media/'))
    .map(([, bytes]) => bytes);

/** The saved file reopens with the room's view, or (`body`) with the room's body. */
function reopens(room: Room, shape: Shape): void {
  if (shape.reopen === 'exact') {
    checkSaveAndLoad(room.room.document, shape.kind, packageOf(FILE), {
      blobs: (hash) => room.blob(hash),
    });
    return;
  }
  const live = comparable(composeOrderedPart(room.room.document, BODY).part);
  expect(comparable(composeOrderedPart(saved(room).doc, BODY).part)).toBe(live);
}

/** Whether a copy's editor holds each media's bytes: read back from the room, by hash. */
const holds = (person: Person, media: readonly Uint8Array[]): boolean => {
  const held = [...surface(person).session.currentPackage().partBytes.values()];
  return media.every((bytes) =>
    held.some(
      (other) => other.length === bytes.length && orderedBlobHash(other) === orderedBlobHash(bytes)
    )
  );
};

const sameBytes = (a: readonly Uint8Array[], b: readonly Uint8Array[]): boolean =>
  JSON.stringify(a.map(orderedBlobHash).sort()) === JSON.stringify(b.map(orderedBlobHash).sort());

/**
 * A body as both sides write it, without what each names its own way: Word's drawing ids (the
 * room leaves them out) and a VML picture's image relationship id. Images read by order.
 */
const comparable = (part: OoxmlPart): string =>
  normalizedBody(part, () => null)
    .replace(/ (w14|wp14):(anchorId|editId)="[^"]*"/g, '')
    .replace(/ r:id="\?[^"]*"/g, ' r:id="?"');

/** What the editor alone writes for the same paste at the same place. */
function alone(flavours: Copied, party: boolean, paragraph: number, offset: number): string {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: FILE,
    author: party ? 'Ana' : 'Olive',
    mode: party ? 'suggesting' : 'edit',
    modules: [review],
  });
  try {
    const s = editor.surface;
    if (!s) throw new Error('no surface');
    const id = s.session.paragraphIds()[paragraph] ?? '';
    s.setSelection({ anchor: { paragraphId: id, offset }, head: { paragraphId: id, offset } });
    expect(s.pasteRich(flavours.text, flavours.html)).toBe(true);
    return comparable(s.session.part());
  } finally {
    editor.destroy();
    container.remove();
  }
}

const roomForm = (shape: Shape, engine: string): string => shape.roomForm?.(engine) ?? engine;

const agrees = (room: Room, engine: string, party: boolean): string => {
  const g = grade(comparable(composeOrderedPart(room.room.document, BODY).part), engine);
  return g === 'exact' || g === 'same' || (party && g === 'ids') ? 'agrees' : g;
};

const cards = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .filter((placement) => placement.kind === 'revision')
    .map((placement) => placement.key);

for (const shape of SHAPE_KINDS)
  describe(`${shape.name} pasted in an ordered room`, () => {
    test("the organiser's: every copy shows it as the editor alone does; undo and redo", async () => {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
      const ana = await join(room, 'Ana', { review });
      const flavours = copiedShape(shape.kind);
      caretAt(olive, 1, 3);
      expect(surface(olive).pasteRich(flavours.text, flavours.html)).toBe(true);
      await settle(16);
      expect(kinds(room)).toEqual(['insert-fragment']);
      checked(room, [olive, ana]);
      for (const bytes of shape.media) expect(room.blob(orderedBlobHash(bytes))).toEqual(bytes);
      expect(holds(ana, shape.media)).toBe(true);
      expect(editorXml(ana)).toMatch(shape.shows);
      expect(editorXml(ana)).not.toContain('<w:ins ');
      expect(sameBytes(savedMedia(room), shape.media)).toBe(true);
      expect(agrees(room, roomForm(shape, alone(flavours, false, 1, 3)), false)).toBe('agrees');
      reopens(room, shape);

      expect(olive.handle.session.undo()).toBe(true);
      await settle(16);
      checked(room, [olive, ana]);
      expect(editorTexts(ana)).toEqual(BEFORE);
      expect(editorXml(ana)).not.toMatch(shape.shows);
      expect(savedMedia(room)).toEqual([]);
      expect(olive.handle.session.redo()).toBe(true);
      await settle(16);
      checked(room, [olive, ana]);
      expect(editorXml(ana)).toMatch(shape.shows);
      expect(sameBytes(savedMedia(room), shape.media)).toBe(true);
    });

    test("a party's: a tracked insertion; a reject takes it and its media out of the file", async () => {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
      const ana = await join(room, 'Ana', { review });
      const flavours = copiedShape(shape.kind);
      caretAt(ana, 0, 4);
      expect(surface(ana).pasteRich(flavours.text, flavours.html)).toBe(true);
      await settle(16);
      expect(kinds(room)).toEqual(['insert-fragment']);
      checked(room, [olive, ana]);
      expect(editorXml(olive)).toMatch(/<w:ins [^>]*w:author="Ana"/);
      expect(editorXml(olive)).toMatch(shape.shows);
      expect(holds(olive, shape.media)).toBe(true);
      expect(sameBytes(savedMedia(room), shape.media)).toBe(true);
      expect(agrees(room, roomForm(shape, alone(flavours, true, 0, 4)), true)).toBe('agrees');
      reopens(room, shape);

      const [card] = cards(olive);
      if (card === undefined) throw new Error('no card');
      expect(olive.editor.rejectReviewItem(card).ok).toBe(true);
      await settle(16);
      checked(room, [olive, ana]);
      expect(editorTexts(olive)).toEqual(BEFORE);
      expect(editorXml(olive)).not.toMatch(shape.shows);
      expect(savedMedia(room)).toEqual([]);
    });
  });
