// The watermark in an ordered room (task 9.8), with real editors and one room: the room's save
// writes the shape the editor alone writes, the organiser's undo and redo take it back exactly,
// and a party's is refused by the host's party rule.

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { readPackageWatermark } from '../../../../store/package/watermark-package.ts';
import { orderedBlobHash } from '../../protocol/blobs.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { docx, p, R, RED_PNG, t, WML_CT } from '../word-fixtures.ts';
import { closeAll, listTexts, settle, surface, type Person } from './editor-harness.ts';
import { converged, FILE, HEADER, inHeader, room2, saved } from './sections-fixtures.ts';

afterEach(closeAll);

const DRAFT = { text: 'DRAFT' } as const;
const PICTURE = { imageData: RED_PNG } as const;

/** A watermark the editor's command takes. */
type Watermark = typeof DRAFT | typeof PICTURE;

/**
 * What the pictures of each header part of a save hold, by part name: the shapes, without where
 * the namespaces they use are declared (on the run or on the picture, both valid).
 */
function watermarks(parts: Record<string, string>): Record<string, string[]> {
  const out: Record<string, string[]> = {};
  for (const [name, xml] of Object.entries(parts))
    if (/^word\/header[^/]*\.xml$/.test(name))
      out[name] = [...xml.matchAll(/<w:pict\b[^>]*>([\s\S]*?)<\/w:pict>/g)].map((m) => m[1] ?? '');
  return out;
}

/** The parts of a save, as text, by name. */
function partsOf(bytes: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, data] of Object.entries(unzipSync(bytes))) out[name] = strFromU8(data);
  return out;
}

/** The editor alone, with no room: the same command on the same file, and its save. */
async function alone(file: Uint8Array, watermark: Watermark | null): Promise<Uint8Array> {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: file, author: 'Olive', mode: 'edit' });
  try {
    expect(editor.exec({ type: 'setWatermark', watermark }).ok).toBe(true);
    return new Uint8Array(await editor.save());
  } finally {
    editor.destroy();
    container.remove();
  }
}

const setWatermark = (person: Person, watermark: Watermark | null) =>
  person.editor.exec({ type: 'setWatermark', watermark });

/** A header with a watermark Word wrote, and text after it. */
const MARKED = docx(
  p(t('Body text.')) +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/></w:sectPr>',
  {
    parts: {
      'word/header1.xml':
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
        'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">' +
        '<w:p><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shapetype id="_x0000_t136"/>' +
        '<v:shape id="PowerPlusWaterMarkObject9" o:spid="_x0000_s9" type="#_x0000_t136" ' +
        'style="position:absolute;width:300pt;height:80pt" fillcolor="silver" stroked="f">' +
        '<v:textpath style="font-family:&quot;Calibri&quot;" string="OLD"/></v:shape></w:pict></w:r>' +
        '<w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>',
    },
    relationships: `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes: `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>`,
  }
);

/** One paragraph and a section that names no header. */
const PLAIN = docx(p(t('Body text.')) + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>');

describe("the organiser's watermark", () => {
  test('every copy and the editor alone write the same shape; undo and redo; save and reopen', async () => {
    const { room, ana, olive } = await room2();
    expect(setWatermark(olive, DRAFT).ok).toBe(true);
    await settle();
    converged(room, ana, olive);
    const mine = watermarks(await saved(olive));
    expect(mine['word/header1.xml']?.length).toBe(1);
    expect(mine['word/header1.xml']?.[0]).toContain('string="DRAFT"');
    expect(watermarks(await saved(ana))).toEqual(mine);
    expect(mine).toEqual(watermarks(partsOf(await alone(FILE, DRAFT))));
    // The saved file reopens with the watermark in the header, before its text
    const reopened = openOrderedWordFile(new Uint8Array(await ana.editor.save()), { file: 2 });
    const header = reopened.document.stories.get(HEADER);
    // One unit, as the engine counts a picture it draws: the list's offsets are the editor's
    expect([...(header?.tree.items() ?? [])].map((item) => item.kind)).toEqual([
      'atom',
      'text',
      'mark',
    ]);
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect(watermarks(await saved(ana))).toEqual({ 'word/header1.xml': [] });
    surface(olive).redo();
    await settle();
    converged(room, ana, olive);
    expect(watermarks(await saved(ana))).toEqual(mine);
    // The watermark is one unit in the editor and in the list: typing after it lands right
    inHeader(olive, 7);
    surface(olive).insertPlainText(' new');
    await settle();
    converged(room, ana, olive);
    expect(listTexts(ana, HEADER)).toEqual(['\uFFFCHeader new text']);
  });

  test("a file's watermark: replaced and removed as the editor alone does, the old run whole; undo shows it again", async () => {
    for (const next of [DRAFT, null]) {
      const { room, ana, olive } = await room2(MARKED);
      const before = watermarks(await saved(ana));
      expect(before['word/header1.xml']?.[0]).toContain('string="OLD"');
      expect(setWatermark(olive, next).ok).toBe(true);
      await settle();
      converged(room, ana, olive);
      const after = watermarks(await saved(ana));
      expect(after).toEqual(watermarks(partsOf(await alone(MARKED, next))));
      expect(after['word/header1.xml']?.length).toBe(next ? 1 : 0);
      expect((await saved(ana))['word/header1.xml']).toContain('Header text');
      surface(olive).undo();
      await settle();
      converged(room, ana, olive);
      expect(watermarks(await saved(ana))).toEqual(before);
      closeAll();
    }
  });

  test('a document with no header: a new default header takes it; undo leaves no header part', async () => {
    const { room, ana, olive } = await room2(PLAIN);
    expect(setWatermark(olive, DRAFT).ok).toBe(true);
    await settle();
    converged(room, ana, olive);
    const made = watermarks(await saved(ana));
    const names = Object.keys(made);
    expect(names.length).toBe(1);
    expect(Object.values(made)).toEqual(
      Object.values(watermarks(partsOf(await alone(PLAIN, DRAFT))))
    );
    expect(watermarks(await saved(olive))).toEqual(made);
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect(watermarks(await saved(ana))).toEqual({});
  });
});

/** Two sections; the first has a different first page and no header of its own. */
const TITLED = docx(
  p(t('Section one.'), '<w:sectPr><w:titlePg/></w:sectPr>') +
    p(t('Section two.')) +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/></w:sectPr>',
  {
    parts: {
      'word/header1.xml':
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>',
    },
    relationships: `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes: `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>`,
  }
);

/** A watermark reading without the part names and ids where it sits. */
function withoutPlaces(reading: ReturnType<typeof readPackageWatermark>): unknown {
  if (reading?.kind !== 'picture') return reading;
  const { ownerPartName: _owner, relationshipId: _id, mediaPartName: _media, ...rest } = reading;
  return rest;
}

/** The watermark a save shows, and its picture's bytes. */
function pictureOf(bytes: Uint8Array) {
  const read = readOoxmlPackage(bytes, {});
  if (!read.ok) throw new Error(read.reason);
  const reading = readPackageWatermark(read.package);
  const media =
    reading?.kind === 'picture' && reading.mediaPartName
      ? read.package.partBytes.get(reading.mediaPartName)
      : undefined;
  return { reading, media };
}

describe('a picture watermark', () => {
  test('its bytes go to the room once; every copy shows what the editor alone writes; undo takes it out', async () => {
    const { room, ana, olive } = await room2();
    expect(setWatermark(olive, PICTURE).ok).toBe(true);
    await settle();
    converged(room, ana, olive);
    expect(room.blob(orderedBlobHash(RED_PNG))).toEqual(RED_PNG);
    const own = pictureOf(await alone(FILE, PICTURE));
    for (const person of [ana, olive]) {
      const shown = pictureOf(new Uint8Array(await person.editor.save()));
      expect(shown.reading).toMatchObject({ kind: 'picture', washout: true });
      expect(shown.media).toEqual(RED_PNG);
      // The same watermark as the editor alone's, wherever its parts are
      expect(withoutPlaces(shown.reading)).toEqual(withoutPlaces(own.reading));
    }
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect(pictureOf(new Uint8Array(await ana.editor.save()))).toEqual({
      reading: null,
      media: undefined,
    });
  });
});

describe('a watermark across sections', () => {
  test('each section gets the headers it can show, as the editor alone makes them', async () => {
    const { room, ana, olive } = await room2(TITLED);
    expect(setWatermark(olive, DRAFT).ok).toBe(true);
    await settle();
    converged(room, ana, olive);
    const shapes = (parts: Record<string, string>) => Object.values(watermarks(parts)).sort();
    const made = shapes(await saved(ana));
    // Two new headers for section one (default, first page), and section two's own
    expect(made.map((pictures) => pictures.length)).toEqual([1, 1, 1]);
    expect(made).toEqual(shapes(partsOf(await alone(TITLED, DRAFT))));
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect(shapes(await saved(ana))).toEqual([[]]);
  });
});

describe("a party's watermark", () => {
  test("is refused by the host's party rule: headers and footers are the organiser's alone", async () => {
    const { room, ana, olive } = await room2();
    setWatermark(ana, DRAFT);
    await settle();
    expect(ana.refusals.map((r) => `${r.code}:${r.reason}`)).toEqual(['ordered-refused:furniture']);
    converged(room, olive);
    expect(watermarks(await saved(ana))).toEqual({ 'word/header1.xml': [] });
  });
});
