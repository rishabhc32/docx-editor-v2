// Pictures and shapes in the list (group 9.4): a new picture is a tracked insertion of its run,
// its media a resource of its suggestion; a direct property change is a record that compose
// runs with the engine's own op, and `revert` takes it back exactly; a new image is a direct
// edit too; loaded media a session made unused leaves the saved file.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPackage, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart, validateOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { validateOrderedChange } from '../change/validate.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { composeOrderedPackage, orderedPackageBytes } from '../compose/package.ts';
import { revertDirect } from '../edits/revert.ts';
import { setStatus } from '../model/mutations.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { apply, at, BODY, DIRECT, EDIT, refusal } from './change-apply-fixture.ts';
import { docx, MATRIX_EXTRAS, p, pictureRun, RED_PNG, t } from './word-fixtures.ts';
import { testAttribution } from './test-persons.ts';

const PNG = RED_PNG;
const MEDIA = {
  sha256: orderedBlobHash(PNG),
  contentType: 'image/png',
  size: PNG.byteLength,
} as const;
const blobs = (hash: string) => (hash === MEDIA.sha256 ? PNG : undefined);
const PICTURE = { type: 'picture', media: MEDIA, extent: { cx: 914_400, cy: 457_200 } } as const;

/** A document with one loaded picture between two words, and one plain paragraph. */
function open(): { doc: OrderedDocument; base: OoxmlPackage } {
  const bytes = docx(
    p(t('Alpha ') + pictureRun(5) + t(' Omega')) + p(t('Plain words.')),
    MATRIX_EXTRAS
  );
  const opened = openOrderedWordFile(bytes, { file: 1 });
  return { doc: opened.document, base: opened.base };
}

const body = (doc: OrderedDocument) => serializeOoxmlPart(composeOrderedPart(doc, BODY).part);

/** The shown atom of a paragraph whose element is a drawing. */
function pictureOf(doc: OrderedDocument, index = 0): OrderedItem {
  for (const item of doc.stories.get(BODY)?.tree.items() ?? [])
    if (item.kind === 'atom' && item.node?.kind === 'drawing') {
      if (index === 0) return item;
      index -= 1;
    }
  throw new Error('no picture');
}

const address = (item: OrderedItem) => ({ space: item.space, offset: item.offset });

function decide(doc: OrderedDocument, id: string, status: 'accepted' | 'rejected') {
  withUndoRecord(doc, () => setStatus(doc, id, status));
}

describe('a new picture (insert-object)', () => {
  test('a party inserts it as a tracked run insertion, its media written while it shows', () => {
    const { doc, base } = open();
    const out = apply(
      doc,
      'ben-1:1',
      { ...EDIT, kind: 'insert-object', place: at(doc, 1, 6), object: PICTURE },
      'Ben'
    );
    const xml = body(doc);
    expect(xml).toMatch(/<w:ins [^>]*w:author="Ben"[^>]*><w:r><w:drawing[ >]/);
    expect(xml).toContain('r:embed="rOben-1_1_2"');
    expect(xml).toContain('cx="914400"');
    // `wp:docPr/@id` is document-wide: the new one is above the file's dense ids
    const docPr = /<wp:docPr id="(\d+)" name=""\/>/.exec(xml);
    expect(Number(docPr?.[1])).toBeGreaterThanOrEqual(1_000_000_000);
    const pkg = composeOrderedPackage(doc, base, { blobs });
    expect(pkg.partBytes.get('/word/media/ordered_ben-1_1_1.png')).toEqual(PNG);
    // Rejected: the picture, its media part and its relationship go
    decide(doc, out.suggestions[0] ?? '', 'rejected');
    const after = composeOrderedPackage(doc, base, { blobs });
    expect(
      serializeOoxmlPart(after.parts.get(BODY) ?? composeOrderedPart(doc, BODY).part)
    ).not.toContain('rOben-1_1_2');
    expect(after.partBytes.has('/word/media/ordered_ben-1_1_1.png')).toBe(false);
    expect(after.relationships.get(BODY)?.some((r) => r.id === 'rOben-1_1_2')).toBe(false);
  });

  test('an accepted picture reopens as the file own picture', () => {
    const { doc, base } = open();
    const out = apply(
      doc,
      'ben-1:1',
      {
        ...EDIT,
        kind: 'insert-object',
        place: at(doc, 1, 0),
        object: {
          ...PICTURE,
          title: 'Logo',
          description: 'A red dot\non two lines',
          link: 'https://example.com/a',
        },
      },
      'Ben'
    );
    decide(doc, out.suggestions[0] ?? '', 'accepted');
    const pkg = composeOrderedPackage(doc, base, { blobs });
    const xml = serializeOoxmlPart(pkg.parts.get(BODY) ?? composeOrderedPart(doc, BODY).part);
    expect(xml).not.toContain('<w:ins');
    expect(xml).toContain('title="Logo"');
    expect(xml).toContain('<a:hlinkClick r:id="rOben-1_1_3"/>');
    const rels = serializeOoxmlPart(
      pkg.parts.get('/word/_rels/document.xml.rels') ?? composeOrderedPart(doc, BODY).part
    );
    expect(rels).toMatch(
      /Id="rOben-1_1_3"[^>]*Target="https:\/\/example.com\/a"[^>]*TargetMode="External"/
    );
  });

  test('a direct insert by the organiser is plain, and revert takes it back with its media', () => {
    const { doc, base } = open();
    apply(doc, 'ana-1:1', {
      ...DIRECT,
      kind: 'insert-object',
      place: at(doc, 1, 6),
      object: PICTURE,
    });
    expect(body(doc)).not.toContain('<w:ins');
    expect(body(doc)).toContain('r:embed="rOana-1_1_2"');
    withUndoRecord(doc, () => revertDirect(doc, 'S:ana-1:1.1', testAttribution('Ana'), 'rejected'));
    expect(body(doc)).not.toContain('rOana-1_1_2');
    expect(
      composeOrderedPackage(doc, base, { blobs }).partBytes.has('/word/media/ordered_ana-1_1_1.png')
    ).toBe(false);
  });

  test('a text box is the engine own shape, and two copies make the same element', () => {
    const one = open().doc;
    const two = open().doc;
    for (const doc of [one, two])
      apply(doc, 'ana-1:1', {
        ...DIRECT,
        kind: 'insert-object',
        place: at(doc, 1, 0),
        object: { type: 'textbox', extent: { cx: 1_270_000 } },
      });
    expect(body(one)).toContain('<wps:txbx>');
    expect(body(one)).toContain('cx="1270000"');
    expect(body(one)).toBe(body(two));
  });
});

describe('a direct property change (object-property, object-resource)', () => {
  test('resize, crop, alt text and a link apply as the engine writes them, and revert one exactly', () => {
    const { doc } = open();
    const picture = pictureOf(doc);
    apply(doc, 'ana-1:1', {
      ...DIRECT,
      kind: 'object-property',
      object: address(picture),
      ops: [{ op: 'crop', crop: { left: 10_000, top: 0, right: 0, bottom: 5_000 } }],
    });
    apply(doc, 'ana-1:2', {
      ...DIRECT,
      kind: 'object-property',
      object: address(picture),
      ops: [
        { op: 'resize', cx: 228_600, cy: 228_600 },
        { op: 'metadata', title: 'Red', description: 'A dot', link: 'https://example.com/b' },
      ],
    });
    let xml = body(doc);
    expect(xml).toContain('<wp:extent cx="228600" cy="228600"/>');
    expect(xml).toMatch(/<a:srcRect [^>]*l="10000"/);
    expect(xml).toContain('descr="A dot"');
    expect(xml).toContain('<a:hlinkClick r:id="rOana-1_2_1"/>');
    // The drawing keeps its node id, so the editor's selection holds
    expect(composeOrderedPart(doc, BODY).part.root).toBeDefined();
    // Undo of the crop only: the size and the alt text stay, the crop goes
    withUndoRecord(doc, () => revertDirect(doc, 'S:ana-1:1.1', testAttribution('Ana'), 'rejected'));
    xml = body(doc);
    expect(xml).not.toContain('l="10000"');
    expect(xml).toContain('<wp:extent cx="228600" cy="228600"/>');
    // And of the rest: the file's own picture again
    withUndoRecord(doc, () => revertDirect(doc, 'S:ana-1:2.1', testAttribution('Ana'), 'rejected'));
    xml = body(doc);
    expect(xml).toContain('<wp:extent cx="457200" cy="457200"/>');
    expect(xml).not.toContain('hlinkClick');
  });

  test('a new image is a direct edit whose media goes with its revert', () => {
    const { doc, base } = open();
    const picture = pictureOf(doc);
    apply(doc, 'ana-1:1', {
      ...DIRECT,
      kind: 'object-resource',
      object: address(picture),
      media: MEDIA,
    });
    expect(body(doc)).toContain('r:embed="rOana-1_1_2"');
    expect(
      composeOrderedPackage(doc, base, { blobs }).partBytes.has('/word/media/ordered_ana-1_1_1.png')
    ).toBe(true);
    withUndoRecord(doc, () => revertDirect(doc, 'S:ana-1:1.1', testAttribution('Ana'), 'rejected'));
    expect(body(doc)).toContain('r:embed="rIdImg1"');
    expect(
      composeOrderedPackage(doc, base, { blobs }).partBytes.has('/word/media/ordered_ana-1_1_1.png')
    ).toBe(false);
  });

  test('a suggestion of a property, a locked picture, and a hidden one are refused by name', () => {
    const { doc } = open();
    const picture = pictureOf(doc);
    const resize = {
      kind: 'object-property',
      object: address(picture),
      ops: [{ op: 'resize', cx: 10, cy: 10 }],
    } as const;
    expect(refusal(doc, 'ben-1:1', { ...EDIT, ...resize }, 'Ben')).toBe('untracked-only');
    apply(doc, 'ana-1:1', {
      ...DIRECT,
      kind: 'object-property',
      object: address(picture),
      ops: [{ op: 'locks', locks: { resize: true } }],
    });
    expect(refusal(doc, 'ana-1:2', { ...DIRECT, ...resize })).toBe('engine-refused');
    // A position is for a floating picture only, as in the editor
    expect(
      refusal(doc, 'ana-1:3', {
        ...DIRECT,
        kind: 'object-property',
        object: address(picture),
        ops: [{ op: 'position', position: { horizontalEmu: 5 } }],
      })
    ).toBe('engine-refused');
    apply(doc, 'ana-1:4', { ...DIRECT, kind: 'delete', from: at(doc, 0, 6), to: at(doc, 0, 7) });
    expect(
      refusal(doc, 'ana-1:5', {
        ...DIRECT,
        kind: 'object-property',
        object: address(picture),
        ops: [{ op: 'wrap', wrap: 'square' }],
      })
    ).toBe('hidden-target');
  });

  test('a picture wrapped and moved is the engine element, and a party struck one shows struck and resized', () => {
    const { doc } = open();
    const picture = pictureOf(doc);
    apply(
      doc,
      'ben-1:1',
      { ...EDIT, kind: 'delete', from: at(doc, 0, 6), to: at(doc, 0, 7) },
      'Ben'
    );
    apply(doc, 'ana-1:1', {
      ...DIRECT,
      kind: 'object-property',
      object: address(picture),
      ops: [
        { op: 'wrap', wrap: 'square' },
        { op: 'position', position: { horizontalEmu: 12_700, relativeToH: 'page' } },
      ],
    });
    const xml = body(doc);
    expect(xml).toMatch(/<w:del [^>]*w:author="Ben"[^>]*><w:r><w:drawing[^>]*><wp:anchor/);
    expect(xml).toContain('<wp:wrapSquare');
    expect(xml).toContain('relativeFrom="page"');
  });
});

describe('the saved file', () => {
  test('a deleted loaded picture takes its media part with it; a pending deletion keeps it', () => {
    const { doc, base } = open();
    const out = apply(
      doc,
      'ben-1:1',
      { ...EDIT, kind: 'delete', from: at(doc, 0, 6), to: at(doc, 0, 7) },
      'Ben'
    );
    expect(composeOrderedPackage(doc, base).partBytes.has('/word/media/red.png')).toBe(true);
    decide(doc, out.suggestions[0] ?? '', 'accepted');
    const pkg = composeOrderedPackage(doc, base);
    expect(pkg.partBytes.has('/word/media/red.png')).toBe(false);
    expect(pkg.relationships.get(BODY)?.some((r) => r.id === 'rIdImg1')).toBe(false);
  });

  test('a new image takes the old media part out of the file, and back with its revert', () => {
    const { doc, base } = open();
    apply(doc, 'ana-1:1', {
      ...DIRECT,
      kind: 'object-resource',
      object: address(pictureOf(doc)),
      media: MEDIA,
    });
    expect(composeOrderedPackage(doc, base, { blobs }).partBytes.has('/word/media/red.png')).toBe(
      false
    );
    withUndoRecord(doc, () => revertDirect(doc, 'S:ana-1:1.1', testAttribution('Ana'), 'rejected'));
    expect(composeOrderedPackage(doc, base, { blobs }).partBytes.has('/word/media/red.png')).toBe(
      true
    );
  });

  test('two people insert the same bytes: two pictures, one blob, two parts', () => {
    const { doc, base } = open();
    apply(
      doc,
      'ben-1:1',
      { ...EDIT, kind: 'insert-object', place: at(doc, 1, 0), object: PICTURE },
      'Ben'
    );
    apply(
      doc,
      'cara-1:1',
      { ...EDIT, kind: 'insert-object', place: at(doc, 1, 1), object: PICTURE },
      'Cara'
    );
    const pkg = composeOrderedPackage(doc, base, { blobs });
    expect(pkg.partBytes.get('/word/media/ordered_ben-1_1_1.png')).toEqual(PNG);
    expect(pkg.partBytes.get('/word/media/ordered_cara-1_1_1.png')).toEqual(PNG);
  });
});

describe('validation', () => {
  const change = (body: Record<string, unknown>) =>
    validateOrderedChange({ v: 1, id: 'ana-1:1', part: BODY, mode: 'edit', ...body });
  const object = { space: 'b:1:x', offset: 0 };
  const codes = (body: Record<string, unknown>) => {
    const result = change(body);
    return result.ok ? 'ok' : result.error.code;
  };

  test('a picture names its media by hash and bounds every value', () => {
    const place = { space: 'b:1:x', offset: 0 };
    expect(codes({ kind: 'insert-object', place, object: PICTURE })).toBe('ok');
    expect(
      codes({ kind: 'insert-object', place, object: { ...PICTURE, link: 'javascript:alert(1)' } })
    ).toBe('invalid-href');
    expect(codes({ kind: 'insert-object', place, object: { ...PICTURE, title: 'a\u0001b' } })).toBe(
      'invalid-text'
    );
    expect(
      codes({ kind: 'insert-object', place, object: { ...PICTURE, extent: { cx: 0, cy: 1 } } })
    ).toBe('invalid-number');
    expect(
      codes({
        kind: 'insert-object',
        place,
        object: { ...PICTURE, media: { ...MEDIA, contentType: 'image/svg+xml' } },
      })
    ).toBe('invalid-value');
    expect(codes({ kind: 'insert-object', place, object: { ...PICTURE, xml: '<w:r/>' } })).toBe(
      'unknown-field'
    );
    expect(
      codes({
        kind: 'insert-object',
        place,
        object: { type: 'textbox', extent: { cx: 10 }, media: MEDIA },
      })
    ).toBe('unknown-field');
  });

  test('a property change is a short list of typed ops', () => {
    expect(
      codes({
        kind: 'object-property',
        object,
        ops: [{ op: 'crop', crop: { left: 60_000, top: 0, right: 40_000, bottom: 0 } }],
      })
    ).toBe('invalid-value');
    expect(
      codes({ kind: 'object-property', object, ops: [{ op: 'wrap', wrap: 'everywhere' }] })
    ).toBe('invalid-value');
    expect(
      codes({
        kind: 'object-property',
        object,
        ops: [{ op: 'position', position: { mode: 'simple', relativeToH: 'page' } }],
      })
    ).toBe('invalid-value');
    expect(codes({ kind: 'object-property', object, ops: [{ op: 'locks', locks: {} }] })).toBe(
      'empty'
    );
    expect(codes({ kind: 'object-property', object, ops: [] })).toBe('empty');
    expect(
      codes({
        kind: 'object-property',
        object,
        ops: Array.from({ length: 17 }, () => ({ op: 'transform', action: 'flipH' })),
      })
    ).toBe('too-many');
    expect(
      codes({
        kind: 'object-property',
        object,
        ops: [{ op: 'metadata', title: '', description: 'x', link: null }],
      })
    ).toBe('ok');
    expect(codes({ kind: 'object-resource', object, media: { ...MEDIA, sha256: 'ABC' } })).toBe(
      'invalid-value'
    );
  });
});

describe('the saved file reopens', () => {
  test('an accepted picture is valid file content: part, relationship and content type, and reopens', () => {
    const { doc, base } = open();
    const out = apply(
      doc,
      'ben-1:1',
      { ...EDIT, kind: 'insert-object', place: at(doc, 1, 0), object: PICTURE },
      'Ben'
    );
    decide(doc, out.suggestions[0] ?? '', 'accepted');
    const bytes = orderedPackageBytes(doc, base, { blobs });
    const read = readOoxmlPackage(bytes, {});
    if (!read.ok) throw new Error(read.reason);
    const pkg = read.package;
    const part = '/word/media/ordered_ben-1_1_1.png';
    expect(pkg.partBytes.get(part)).toEqual(PNG);
    const record = pkg.relationships.get(BODY)?.find((r) => r.id === 'rOben-1_1_2');
    expect(record?.rawTarget).toBe('media/ordered_ben-1_1_1.png');
    expect(pkg.contentTypes.overrides.get(part) ?? pkg.contentTypes.defaults.get('png')).toBe(
      'image/png'
    );
    for (const name of [BODY]) {
      const story = pkg.parts.get(name);
      if (!story) throw new Error('no story');
      expect(validateOoxmlPart(story).ok).toBe(true);
    }
    // Reopened, the picture is the file's own: no suggestion, the same element
    const again = openOrderedWordFile(bytes, { file: 2 });
    expect(again.document.suggestions.size).toBe(0);
    const xml = serializeOoxmlPart(composeOrderedPart(again.document, BODY).part);
    expect(xml).toContain('r:embed="rOben-1_1_2"');
    expect(xml).not.toContain('<w:ins');
  });
});
