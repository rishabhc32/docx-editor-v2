// A list with a picture bullet pasted in an ordered room (group 9.10). Word copies a list's
// picture bullet (`w:numPicBullet`, §17.9.20) and its image with the list; the editor's copy
// does too, and the room carries the bullet beside the list, its image held by hash, as a
// relationship of the numbering part. The organiser's paste is a direct edit that undo and redo
// turn back and again; a party's is a tracked insertion whose reject takes the bullet and its
// image out of the file.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { readOoxmlPackage, type OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { resolveInternalTarget } from '../../../../store/package/opc-names.ts';
import { serializeOoxmlPart, type OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { relationshipsOf } from '../../../../store/package/package-edit.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import { orderedBlobHash } from '../../protocol/blobs.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { grade, normalizedBody } from '../matrix/compare.ts';
import { checkSaveAndLoad } from '../sessions/session.ts';
import { docx, p, packageOf, R, RED_PNG, t, W } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
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
import { copiedBetween, type Copied } from './paste-harness.ts';
import { reviewTestModule } from './review-module.ts';

const review = reviewTestModule();

afterEach(closeAll);

type Room = ReturnType<typeof localRoom>;

const V = 'urn:schemas-microsoft-com:vml';
const O = 'urn:schemas-microsoft-com:office:office';
const WML_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

/** A list item whose level draws a picture, as Word writes a picture bullet, and a line after it. */
const SOURCE = docx(
  p(t('Picture item'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
    p(t('After.')),
  {
    parts: {
      'word/media/bullet.png': RED_PNG,
      'word/numbering.xml':
        `<w:numbering xmlns:w="${W}" xmlns:r="${R}" xmlns:v="${V}" xmlns:o="${O}">` +
        '<w:numPicBullet w:numPicBulletId="0"><w:pict>' +
        '<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:9pt;height:9pt" o:bullet="t">' +
        '<v:imagedata r:id="rIdBullet" o:title="dot"/></v:shape></w:pict></w:numPicBullet>' +
        '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/>' +
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>' +
        '<w:lvlPicBulletId w:val="0"/><w:lvlJc w:val="left"/>' +
        '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
      'word/_rels/numbering.xml.rels':
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rIdBullet" Type="${R}/image" Target="media/bullet.png"/></Relationships>`,
    },
    contentTypes:
      '<Default Extension="png" ContentType="image/png"/>' +
      `<Override PartName="/word/numbering.xml" ContentType="${WML_CT}.numbering+xml"/>`,
    relationships: `<Relationship Id="rIdNumbering" Type="${R}/numbering" Target="numbering.xml"/>`,
  }
);

/** The list item with its paragraph mark, as the editor copies it. */
const copied = (): Copied => copiedBetween(0, 1, 0, SOURCE);

const BEFORE = ['The fee is 100 USD.', 'Due in 30 days.', 'Signed here.'];

const checked = (room: Room, people: readonly Person[]) => {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
  }
};

function saved(room: Room): OoxmlPackage {
  const bytes = orderedPackageBytes(room.room.document, packageOf(FILE), {
    blobs: (hash) => room.blob(hash),
  });
  const pkg = readOoxmlPackage(bytes);
  if (!pkg.ok) throw new Error('the saved file does not open');
  return pkg.package;
}

/** The numbering part of a package, by the main document's relationship. */
function numberingOf(pkg: OoxmlPackage): OoxmlPart | null {
  const record = relationshipsOf(pkg, '/word/document.xml').find((r) =>
    r.type.endsWith('/numbering')
  );
  const target = record ? resolveInternalTarget('/word/document.xml', record.rawTarget) : null;
  return target?.ok ? (pkg.parts.get(target.partName) ?? null) : null;
}

/**
 * The picture bullet the pasted item's level draws: the image bytes of the bullet its level
 * names, through the numbering part's own relationship. Null when no level names one.
 */
function bulletImage(pkg: OoxmlPackage): Uint8Array | null {
  const part = numberingOf(pkg);
  if (!part) return null;
  const xml = serializeOoxmlPart(part);
  const named = /<w:lvlPicBulletId w:val="(\d+)"\/>/.exec(xml)?.[1];
  if (named === undefined) return null;
  const bullet = new RegExp(
    `<w:numPicBullet w:numPicBulletId="${named}">.*?<v:imagedata [^>]*r:id="([^"]+)"`,
    's'
  ).exec(xml);
  const record = relationshipsOf(pkg, part.name).find((r) => r.id === bullet?.[1]);
  const target = record ? resolveInternalTarget(part.name, record.rawTarget) : null;
  return target?.ok ? (pkg.partBytes.get(target.partName) ?? null) : null;
}

/** The picture bullets of a package's numbering part. */
function bullets(pkg: OoxmlPackage): number {
  const part = numberingOf(pkg);
  return part ? (serializeOoxmlPart(part).match(/<w:numPicBullet /g) ?? []).length : 0;
}

/** What the editor alone writes for the same paste at the same place. */
function alone(flavours: Copied, party: boolean, paragraph: number): string {
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
    s.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 0 },
    });
    expect(s.pasteRich(flavours.text, flavours.html)).toBe(true);
    expect(bulletImage(s.session.currentPackage())).toEqual(RED_PNG);
    return comparable(s.session.part());
  } finally {
    editor.destroy();
    container.remove();
  }
}

const comparable = (part: OoxmlPart): string =>
  normalizedBody(part, () => null).replace(/<w:numId w:val="[^"]*"\/>/g, '<w:numId w:val="?"/>');

const agrees = (room: Room, engine: string, party: boolean): string => {
  const g = grade(comparable(composeOrderedPart(room.room.document, BODY).part), engine);
  return g === 'exact' || g === 'same' || (party && g === 'ids') ? 'agrees' : g;
};

const cards = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .filter((placement) => placement.kind === 'revision')
    .map((placement) => placement.key);

describe('a list with a picture bullet pasted in an ordered room', () => {
  test("the organiser's: every copy draws the bullet; undo and redo", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    const flavours = copied();
    caretAt(olive, 1, 0);
    expect(surface(olive).pasteRich(flavours.text, flavours.html)).toBe(true);
    await settle(16);
    checked(room, [olive, ana]);
    expect(room.room.joinSnapshot(0).map((change) => change.kind)).toEqual(['insert-fragment']);
    expect(room.blob(orderedBlobHash(RED_PNG))).toEqual(RED_PNG);
    expect(bulletImage(surface(ana).session.currentPackage())).toEqual(RED_PNG);
    expect(bulletImage(saved(room))).toEqual(RED_PNG);
    expect(agrees(room, alone(flavours, false, 1), false)).toBe('agrees');
    checkSaveAndLoad(room.room.document, 'picture bullet', packageOf(FILE), {
      blobs: (hash) => room.blob(hash),
    });

    expect(olive.handle.session.undo()).toBe(true);
    await settle(16);
    checked(room, [olive, ana]);
    expect(editorTexts(ana)).toEqual(BEFORE);
    expect(bulletImage(saved(room))).toBeNull();
    expect(bullets(saved(room))).toBe(0);
    expect(olive.handle.session.redo()).toBe(true);
    await settle(16);
    checked(room, [olive, ana]);
    expect(bulletImage(saved(room))).toEqual(RED_PNG);
  });

  test("a party's: a tracked insertion; a reject takes the bullet and its image out", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    const flavours = copied();
    caretAt(ana, 1, 0);
    expect(surface(ana).pasteRich(flavours.text, flavours.html)).toBe(true);
    await settle(16);
    checked(room, [olive, ana]);
    expect(bulletImage(surface(olive).session.currentPackage())).toEqual(RED_PNG);
    const file = saved(room);
    expect(serializeOoxmlPart(file.parts.get(BODY)!)).toMatch(/<w:ins [^>]*w:author="Ana"/);
    expect(bulletImage(file)).toEqual(RED_PNG);
    expect(agrees(room, alone(flavours, true, 1), true)).toBe('agrees');
    checkSaveAndLoad(room.room.document, 'picture bullet', packageOf(FILE), {
      blobs: (hash) => room.blob(hash),
    });

    const [card] = cards(olive);
    if (card === undefined) throw new Error('no card');
    expect(olive.editor.rejectReviewItem(card).ok).toBe(true);
    await settle(16);
    checked(room, [olive, ana]);
    expect(editorTexts(olive)).toEqual(BEFORE);
    expect(bullets(saved(room))).toBe(0);
    const media = [...saved(room).partBytes.keys()].filter((name) => name.includes('/media/'));
    expect(media).toEqual([]);
  });
});
