// Resources in the saved file (task 9.2, design D14): each kind is written while its suggestion
// shows and something uses it, goes with a reject or a take-back, and reopens as file content;
// two copies that add resources in either order write the same package.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPackage, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import {
  canonicalOoxmlFingerprint,
  serializeOoxmlPart,
  type OoxmlElement,
  type OoxmlNode,
} from '../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { composeOrderedPackage, orderedPackageBytes } from '../compose/package.ts';
import { mediaBytesOf, withOrderedResources } from '../compose/resources.ts';
import {
  addList,
  addMedia,
  addRelationship,
  addSetting,
  addStory,
  addStyle,
} from '../edits/resources.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import { retractSuggestion } from '../edits/retract.ts';
import { withChildren } from '../ooxml-nodes.ts';
import { BODY, markOf } from './change-apply-fixture.ts';
import { applyAs, decide, media, openFixture, pngBytes, typePicture } from './resources-fixture.ts';
import { testAttribution } from './test-persons.ts';

const LEVEL = {
  level: 0,
  start: 3,
  format: 'upperRoman',
  text: '%1)',
  justification: 'left',
  indent: { left: 720, hanging: 360 },
} as const;

const xmlOf = (pkg: OoxmlPackage, name: string): string => {
  const part = pkg.parts.get(name);
  return part ? serializeOoxmlPart(part) : '<none>';
};

/** A new list, and the first paragraph on it: a list change as group 9.10 makes it. */
function listChange(doc: OrderedDocument, id: string, author = 'Ana'): string {
  let numId = '';
  let made = '';
  applyAs(doc, id, author, (run) => {
    numId = addList(run.ctx, run.suggestion, { levels: [LEVEL] }).numId;
    made = run.suggestion;
    run.run({
      kind: 'paragraph',
      at: markOf(doc, 0),
      ops: [{ op: 'setListNumbering', numId, level: 0 }],
    });
  });
  return `${made} ${numId}`;
}

/** A picture with its media and relationship: an insert as group 9.4 makes it. */
function pictureChange(doc: OrderedDocument, id: string, bytes: Uint8Array, author = 'Ben') {
  let suggestion = '';
  applyAs(doc, id, author, (run) => {
    const picture = addMedia(run.ctx, run.suggestion, media(bytes));
    const relationship = addRelationship(run.ctx, run.suggestion, 'image', {
      resource: picture.key,
    });
    typePicture(run, 1, relationship.id);
    suggestion = run.suggestion;
  });
  return suggestion;
}

describe('resources in the saved file (task 9.2)', () => {
  test('a new list definition is written while a paragraph uses it, reopens, and goes with a reject', () => {
    const { doc, base } = openFixture();
    const [suggestion = '', numId = ''] = listChange(doc, 'ana-1:1').split(' ');
    const numbering = xmlOf(composeOrderedPackage(doc, base), '/word/numbering.xml');
    expect(numbering).toContain(`w:numId="${numId}"`);
    expect(numbering).toContain('<w:numFmt w:val="upperRoman"/>');
    // The saved file reopens with the list as the file's own
    const reopened = openOrderedWordFile(orderedPackageBytes(doc, base), { file: 2 });
    expect(reopened.document.numbering?.has(numId)).toBe(true);
    expect(serializeOoxmlPart(composeOrderedPart(reopened.document, BODY).part)).toContain(
      `<w:numId w:val="${numId}"/>`
    );
    // A reject takes the list's paragraph property and the definition with it
    decide(doc, suggestion, 'rejected');
    expect(xmlOf(composeOrderedPackage(doc, base), '/word/numbering.xml')).not.toContain(numId);
  });

  test('media and its relationship are written while a picture uses them, and go with a reject or a take-back', () => {
    const { doc, base } = openFixture();
    const bytes = pngBytes(7);
    const blobs = (hash: string) => (hash === media(bytes).sha256 ? bytes : undefined);
    const suggestion = pictureChange(doc, 'ben-1:1', bytes);
    const pkg = composeOrderedPackage(doc, base, { blobs });
    const part = '/word/media/ordered_ben-1_1_1.png';
    expect(pkg.partBytes.get(part)).toEqual(bytes);
    expect(pkg.contentTypes.overrides.get(part)).toBe('image/png');
    expect(pkg.relationships.get(BODY)?.some((r) => r.id === 'rOben-1_1_2')).toBe(true);
    expect(xmlOf(pkg, BODY)).toContain('r:embed="rOben-1_1_2"');
    // It reopens as the file's own media
    const saved = orderedPackageBytes(doc, base, { blobs });
    const read = readOoxmlPackage(saved, {});
    expect(read.ok && read.package.partBytes.get(part)).toEqual(bytes);
    // Without the bytes a save is refused by name; a screen leaves the media out
    let reason = 'composed';
    try {
      composeOrderedPackage(doc, base);
    } catch (error) {
      if (error instanceof OrderedRefusal) reason = error.reason;
    }
    expect(reason).toBe('missing-blob');
    const screen = composeOrderedPackage(doc, base, { missing: 'skip' });
    expect(screen.partBytes.has(part)).toBe(false);
    expect(screen.relationships.get(BODY)?.some((r) => r.id === 'rOben-1_1_2')).toBe(false);
    // A reject takes the picture, its relationship and its media part
    decide(doc, suggestion, 'rejected');
    const rejected = composeOrderedPackage(doc, base, { blobs });
    expect(rejected.partBytes.has(part)).toBe(false);
    expect(rejected.relationships.get(BODY)?.some((r) => r.id === 'rOben-1_1_2')).toBe(false);
    // A take-back by its own author: nothing uses the media any more
    const other = openFixture();
    const taken = pictureChange(other.doc, 'ben-1:1', bytes);
    expect(composeOrderedPackage(other.doc, other.base, { blobs }).partBytes.has(part)).toBe(true);
    withUndoRecord(other.doc, () => retractSuggestion(other.doc, taken, testAttribution('Ben')));
    const retracted = composeOrderedPackage(other.doc, other.base, { blobs });
    expect(retracted.partBytes.has(part)).toBe(false);
    expect(retracted.relationships.get(BODY)?.some((r) => r.id === 'rOben-1_1_2')).toBe(false);
  });

  test('a new header part is written only while a relationship to it is used', () => {
    const { doc, base } = openFixture();
    let relationship = '';
    let part = '';
    applyAs(
      doc,
      'olive-1:1',
      'Olive',
      (run) => {
        const story = addStory(run.ctx, run.suggestion, 'header');
        part = story.resource.part;
        relationship = addRelationship(run.ctx, run.suggestion, 'header', {
          resource: story.resource.key,
        }).id;
      },
      true
    );
    expect(doc.stories.has(part)).toBe(true);
    // Unused: the story is not written, and neither is its relationship
    const unused = composeOrderedPackage(doc, base);
    expect(unused.parts.has(part)).toBe(false);
    expect(unused.relationships.get(BODY)?.some((r) => r.id === relationship)).toBe(false);
    // The section's reference (group 9.8 writes it): the part and its relationship are written
    const composed = composeOrderedPackage(doc, base);
    const main = composed.parts.get(BODY);
    if (!main) throw new Error('no main part');
    // A final section holding the reference, at the body's end
    const reference = (namespaceUri: string): OoxmlElement => {
      const attribute = (ns: string, prefix: string, localName: string, value: string) => ({
        kind: 'genericExtension' as const,
        namespaceUri: ns,
        localName,
        prefix,
        value,
      });
      const element = (id: string, localName: string, attributes: unknown[], children: unknown[]) =>
        ({
          id,
          kind: 'generic',
          namespaceUri,
          localName,
          prefix: 'w',
          namespaceBindings: [],
          attributes,
          children,
        }) as unknown as OoxmlElement;
      return element(
        'section',
        'sectPr',
        [],
        [
          element(
            'reference',
            'headerReference',
            [
              attribute(namespaceUri, 'w', 'type', 'default'),
              attribute(
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'r',
                'id',
                relationship
              ),
            ],
            []
          ),
        ]
      );
    };
    const withReference = (node: OoxmlNode): OoxmlNode => {
      if (node.kind === 'textValue') return node;
      if (node.localName === 'body')
        return withChildren(node, [...node.children, reference(node.namespaceUri)]);
      return withChildren(node, node.children.map(withReference));
    };
    const referenced: OoxmlPackage = {
      ...base,
      parts: new Map(base.parts).set(BODY, {
        ...main,
        root: withReference(main.root) as OoxmlElement,
      }),
    };
    const added = new Map([[part, composeOrderedPart(doc, part).part]]);
    const written = withOrderedResources(referenced, doc, added, mediaBytesOf({}));
    expect(written.parts.get(part)?.root.localName).toBe('hdr');
    expect(written.contentTypes.overrides.get(part)).toContain('header+xml');
    expect(written.relationships.get(BODY)?.find((r) => r.id === relationship)?.rawTarget).toBe(
      part.slice('/word/'.length)
    );
  });

  test('a style is written when something uses it and the file lacks it; the first by key wins', () => {
    const { doc, base } = openFixture();
    let suggestion = '';
    applyAs(doc, 'ana-1:1', 'Ana', (run) => {
      addStyle(run.ctx, run.suggestion, {
        styleId: 'OrderedNote',
        type: 'paragraph',
        name: 'Ordered note',
        basedOn: 'OrderedBase',
        run: [{ localName: 'b' }],
      });
      addStyle(run.ctx, run.suggestion, {
        styleId: 'OrderedBase',
        type: 'paragraph',
        name: 'Ordered base',
        paragraph: [{ localName: 'jc', attributes: { val: 'center' } }],
      });
      addStyle(run.ctx, run.suggestion, { styleId: 'Unused', type: 'paragraph', name: 'Unused' });
      run.run({
        kind: 'paragraph',
        at: markOf(doc, 0),
        ops: [
          {
            op: 'setParagraphProperties',
            properties: [{ localName: 'pStyle', attributes: { val: 'OrderedNote' } }],
          },
        ],
      });
      suggestion = run.suggestion;
    });
    const styles = xmlOf(composeOrderedPackage(doc, base), '/word/styles.xml');
    expect(styles).toContain('w:styleId="OrderedNote"');
    expect(styles).toContain('<w:basedOn w:val="OrderedBase"/>');
    // The style it draws on comes too; one nothing uses does not
    expect(styles).toContain('w:styleId="OrderedBase"');
    expect(styles).not.toContain('w:styleId="Unused"');
    decide(doc, suggestion, 'rejected');
    expect(xmlOf(composeOrderedPackage(doc, base), '/word/styles.xml')).not.toContain(
      'OrderedNote'
    );
  });

  test("a settings write applies once its suggestion is accepted, in the room's order", () => {
    const { doc, base } = openFixture();
    applyAs(doc, 'olive-1:1', 'Olive', ({ ctx, suggestion }) =>
      addSetting(ctx, suggestion, { op: 'setDocumentProtection', enforce: true })
    );
    const pending = composeOrderedPackage(doc, base);
    expect(
      [...pending.parts.values()].some((p) => serializeOoxmlPart(p).includes('documentProtection'))
    ).toBe(false);
    const direct = openFixture();
    applyAs(
      direct.doc,
      'olive-1:1',
      'Olive',
      ({ ctx, suggestion }) =>
        addSetting(ctx, suggestion, { op: 'setDocumentProtection', enforce: true }),
      true
    );
    const settings = [...composeOrderedPackage(direct.doc, direct.base).parts.values()].find(
      (p) => p.root.localName === 'settings'
    );
    expect(settings && serializeOoxmlPart(settings)).toContain('w:enforcement="1"');
  });

  test('two copies that add resources in either order write the same package', () => {
    const bytes = pngBytes(3);
    const blobs = () => bytes;
    const fingerprints = (doc: OrderedDocument, base: OoxmlPackage) => {
      const pkg = composeOrderedPackage(doc, base, { blobs });
      return {
        parts: [...pkg.parts.values()].map((p) => [p.name, canonicalOoxmlFingerprint(p)]).sort(),
        bytes: [...pkg.partBytes.keys()].filter((k) => k.includes('/media/')).sort(),
      };
    };
    const first = openFixture();
    listChange(first.doc, 'ana-1:1');
    pictureChange(first.doc, 'ben-1:1', bytes);
    const second = openFixture();
    const record = (() => {
      // The second copy applies Ben's change first, takes it back, and applies it again after
      // Ana's: a pending change on a client
      const taken = pictureChangeRecord(second.doc, bytes);
      takeBack(second.doc, taken);
      listChange(second.doc, 'ana-1:1');
      pictureChange(second.doc, 'ben-1:1', bytes);
      return true;
    })();
    expect(record).toBe(true);
    const third = openFixture();
    pictureChange(third.doc, 'ben-1:1', bytes);
    listChange(third.doc, 'ana-1:1');
    const expected = fingerprints(first.doc, first.base);
    expect(fingerprints(second.doc, second.base)).toEqual(expected);
    // Additions in the other order: the same names, the same parts, the same relationships
    const swapped = fingerprints(third.doc, third.base);
    expect(swapped.bytes).toEqual(expected.bytes);
    expect(swapped).toEqual(expected);
  });
});

/** Ben's picture change, as its record, for a copy that takes it back. */
function pictureChangeRecord(doc: OrderedDocument, bytes: Uint8Array) {
  return applyAs(doc, 'ben-1:1', 'Ben', (run) => {
    const picture = addMedia(run.ctx, run.suggestion, media(bytes));
    const relationship = addRelationship(run.ctx, run.suggestion, 'image', {
      resource: picture.key,
    });
    typePicture(run, 1, relationship.id);
  });
}
