// Paste of a chart and a SmartArt diagram in an ordered room (group 9.10): each copied from a
// Word file with its own parts (a chart's style, colors and embedded workbook; a diagram's data,
// layout, style, colors and drawing). The organiser's paste is a direct edit that undo and redo
// turn back and again; a party's is a tracked insertion that a reject takes out of the file
// with every part. Every copy equals the room, the room's body equals what the editor alone
// pastes, and the saved file holds the parts and reopens with the same view.

import { afterEach, describe, expect, test } from 'bun:test';
import type { OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { resolveInternalTarget } from '../../../../store/package/opc-names.ts';
import { serializeOoxmlPart, type OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { relationshipsOf } from '../../../../store/package/package-edit.ts';
import { orderedBlobHash } from '../../protocol/blobs.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { checkSaveAndLoad } from '../sessions/session.ts';
import { packageOf } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  FILE,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
} from './editor-harness.ts';
import { copiedObject, OBJECTS, WORKBOOK } from './paste-objects-harness.ts';
import {
  agreesAlone,
  allChecked,
  changeKinds,
  pastedAlone,
  resolvesAll,
  revisionCards,
  savedFile,
  type PasteRoom,
} from './paste-room-checks.ts';
import { reviewTestModule } from './review-module.ts';

const review = reviewTestModule();

afterEach(closeAll);

interface ObjectKind {
  readonly kind: keyof typeof OBJECTS;
  readonly name: string;
  /** What the pasted object shows in a copy's body. */
  readonly shows: RegExp;
  /** What its parts hold, all together, in the saved file. */
  readonly holds: readonly RegExp[];
  /** The media it brings, by bytes. */
  readonly media: readonly Uint8Array[];
}

const KINDS: readonly ObjectKind[] = [
  {
    kind: 'chart',
    name: 'a chart with its style, colors and workbook',
    shows: /<c:chart [^>]*r:id="[^"]+"/,
    holds: [
      /<c:chartSpace[ >].*Fees by quarter.*<c:v>250<\/c:v>.*<c:externalData r:id="/s,
      /<cs:chartStyle[ >]/,
      /<cs:colorStyle[ >]/,
    ],
    media: [WORKBOOK],
  },
  {
    kind: 'chartEx',
    name: 'an Office 2016 chart with its style, colors and workbook',
    shows: /<cx:chart [^>]*r:id="[^"]+"/,
    holds: [
      /<cx:chartSpace[ >].*<cx:externalData [^>]*r:id="[^"]+".*Fees spread/s,
      /<cs:chartStyle[ >]/,
      /<cs:colorStyle[ >]/,
    ],
    media: [WORKBOOK],
  },
  {
    kind: 'diagram',
    name: 'a SmartArt diagram with its five parts',
    shows:
      /<dgm:relIds (?=[^>]*r:dm="[^"]+")(?=[^>]*r:lo="[^"]+")(?=[^>]*r:qs="[^"]+")[^>]*r:cs="[^"]+"/,
    holds: [
      /<dgm:dataModel[ >].*Draft.*<dsp:dataModelExt [^>]*relId="/s,
      /<dgm:layoutDef[ >].*r:blip=""/s,
      /<dgm:styleDef[ >]/,
      /<dgm:colorsDef[ >]/,
      /<dsp:drawing[ >].*Sign/s,
    ],
    media: [],
  },
];

const BEFORE = ['The fee is 100 USD.', 'Due in 30 days.', 'Signed here.'];

/** The chart's and diagram's parts of a package, by name. */
const objectParts = (pkg: OoxmlPackage): string[] =>
  [...new Set([...pkg.parts.keys(), ...pkg.partBytes.keys()])]
    .map((name) => (name.startsWith('/') ? name : `/${name}`))
    .filter((name) => /^\/word\/(charts|diagrams|embeddings)\/[^_]/.test(name));

/** Every object part's XML of a saved file, joined. */
const objectXml = (pkg: OoxmlPackage): string =>
  objectParts(pkg)
    .map((name) => pkg.parts.get(name))
    .filter((part): part is OoxmlPart => part !== undefined)
    .map(serializeOoxmlPart)
    .join('\n');

/** The editor alone holds the object's parts too. */
const holdsParts = (pkg: OoxmlPackage): void => {
  expect(objectParts(pkg).length).toBeGreaterThan(0);
};

function savedHolds(room: PasteRoom, object: ObjectKind): void {
  const pkg = savedFile(room);
  expect(resolvesAll(pkg)).toBe(true);
  const xml = objectXml(pkg);
  for (const pattern of object.holds) expect(xml).toMatch(pattern);
  const media = [...pkg.partBytes].filter(([name]) => name.startsWith('/word/embeddings/'));
  expect(media.map(([, bytes]) => orderedBlobHash(bytes))).toEqual(
    object.media.map(orderedBlobHash)
  );
}

for (const object of KINDS)
  describe(`${object.name} pasted in an ordered room`, () => {
    test("the organiser's: every copy shows it as the editor alone does; undo and redo", async () => {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
      const ana = await join(room, 'Ana', { review });
      const flavours = copiedObject(object.kind);
      caretAt(olive, 1, 3);
      expect(surface(olive).pasteRich(flavours.text, flavours.html)).toBe(true);
      await settle(16);
      expect(changeKinds(room)).toEqual(['insert-fragment']);
      allChecked(room, [olive, ana]);
      for (const bytes of object.media) expect(room.blob(orderedBlobHash(bytes))).toEqual(bytes);
      expect(editorXml(ana)).toMatch(object.shows);
      expect(editorXml(ana)).not.toContain('<w:ins ');
      expect(objectXml(surface(ana).session.currentPackage())).toMatch(object.holds[0]!);
      savedHolds(room, object);
      expect(agreesAlone(room, pastedAlone(flavours, false, 1, 3, review, holdsParts), false)).toBe(
        'agrees'
      );
      checkSaveAndLoad(room.room.document, object.kind, packageOf(FILE), {
        blobs: (hash) => room.blob(hash),
      });

      expect(olive.handle.session.undo()).toBe(true);
      await settle(16);
      allChecked(room, [olive, ana]);
      expect(editorTexts(ana)).toEqual(BEFORE);
      expect(objectParts(savedFile(room))).toEqual([]);
      expect(olive.handle.session.redo()).toBe(true);
      await settle(16);
      allChecked(room, [olive, ana]);
      expect(editorXml(ana)).toMatch(object.shows);
      savedHolds(room, object);
    });

    test("a party's: a tracked insertion; a reject takes it and its parts out of the file", async () => {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
      const ana = await join(room, 'Ana', { review });
      const flavours = copiedObject(object.kind);
      caretAt(ana, 0, 4);
      expect(surface(ana).pasteRich(flavours.text, flavours.html)).toBe(true);
      await settle(16);
      expect(changeKinds(room)).toEqual(['insert-fragment']);
      allChecked(room, [olive, ana]);
      expect(editorXml(olive)).toMatch(/<w:ins [^>]*w:author="Ana"/);
      expect(editorXml(olive)).toMatch(object.shows);
      savedHolds(room, object);
      expect(agreesAlone(room, pastedAlone(flavours, true, 0, 4, review, holdsParts), true)).toBe(
        'agrees'
      );
      checkSaveAndLoad(room.room.document, object.kind, packageOf(FILE), {
        blobs: (hash) => room.blob(hash),
      });

      const [card] = revisionCards(olive);
      if (card === undefined) throw new Error('no card');
      expect(olive.editor.rejectReviewItem(card).ok).toBe(true);
      await settle(16);
      allChecked(room, [olive, ana]);
      expect(editorTexts(olive)).toEqual(BEFORE);
      expect(editorXml(olive)).not.toMatch(object.shows);
      expect(objectParts(savedFile(room))).toEqual([]);
    });
  });

/** Every `wp:docPr/@id` of the saved body, in order. */
const docPrIds = (pkg: OoxmlPackage): string[] =>
  [
    ...serializeOoxmlPart(pkg.parts.get('/word/document.xml')!).matchAll(
      /<wp:docPr [^>]*id="(\d+)"/g
    ),
  ].map((match) => match[1]!);

/** The part a relationship of `owner` names, by id. */
function target(pkg: OoxmlPackage, owner: string, id: string): string {
  const record = relationshipsOf(pkg, owner).find((r) => r.id === id);
  const resolved = record ? resolveInternalTarget(owner, record.rawTarget) : null;
  return resolved?.ok ? resolved.partName : `<none ${id}>`;
}

// Word gives a second paste of the same object its own parts (no part is shared), rewrites the
// diagram data's `dsp:dataModelExt/@relId` to its own drawing, and gives the pasted object a new
// `docPr` id when its id is taken (paste evidence 4, `~$paste-out-e4-chart-twice.docx`,
// `~$paste-out-e4-smartart-twice.docx`)
describe('the same chart and diagram pasted twice in an ordered room', () => {
  test('each paste has its own parts and its own drawing ids, as Word pastes', async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    for (const kind of ['chart', 'diagram', 'chart', 'diagram'] as const) {
      const flavours = copiedObject(kind);
      caretAt(olive, 1, 3);
      expect(surface(olive).pasteRich(flavours.text, flavours.html)).toBe(true);
      await settle(16);
    }
    allChecked(room, [olive, ana]);
    const pkg = savedFile(room);
    expect(resolvesAll(pkg)).toBe(true);
    const body = serializeOoxmlPart(pkg.parts.get('/word/document.xml')!);
    const charts = [...body.matchAll(/<c:chart [^>]*r:id="([^"]+)"/g)].map((m) =>
      target(pkg, '/word/document.xml', m[1]!)
    );
    expect(charts).toHaveLength(2);
    expect(new Set(charts).size).toBe(2);
    const workbooks = charts.map((chart) => {
      const record = relationshipsOf(pkg, chart).find((r) => r.type.endsWith('/package'));
      return record ? target(pkg, chart, record.id) : '<none>';
    });
    expect(new Set(workbooks).size).toBe(2);
    // The data part first, then the layout, style and colors parts
    const diagrams = [...body.matchAll(/<dgm:relIds ([^>]*)\/>/g)].map((m) =>
      ['dm', 'lo', 'qs', 'cs'].map((name) =>
        target(pkg, '/word/document.xml', new RegExp(`r:${name}="([^"]+)"`).exec(m[1]!)?.[1] ?? '')
      )
    );
    expect(diagrams).toHaveLength(2);
    expect(new Set(diagrams.flat()).size).toBe(8);
    // Each data part names its own drawing, through a relationship of the story part
    const drawings = diagrams.map(([data]) => {
      const relId = /dataModelExt [^>]*relId="([^"]+)"/.exec(
        serializeOoxmlPart(pkg.parts.get(data!)!)
      )?.[1];
      const record = relationshipsOf(pkg, '/word/document.xml').find((r) => r.id === relId);
      expect(record?.type).toBe(
        'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing'
      );
      return relId ? target(pkg, '/word/document.xml', relId) : '<none>';
    });
    expect(drawings.every((name) => pkg.parts.has(name))).toBe(true);
    expect(new Set(drawings).size).toBe(2);
    const ids = docPrIds(pkg);
    expect(ids).toHaveLength(4);
    expect(new Set(ids).size).toBe(4);
    checkSaveAndLoad(room.room.document, 'pasted twice', packageOf(FILE), {
      blobs: (hash) => room.blob(hash),
    });
  });
});
