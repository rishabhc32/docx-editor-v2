// A drawing slot that survives edits must not keep the part it was made from.
//
// A part's drawing slot is reused for as long as the edits leave its drawings alone, which in
// a long session is most of them. The slot compared each next part with the part it was made
// from, so it kept that first tree for its whole life, and with it the first version of every
// paragraph edited since: a long-lived editor kept one more copy of everything it changed.

import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  readOoxmlPackage,
  withPart,
  type ImageResourceLookup,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { strToU8, zipSync } from 'fflate';
import { createInlineDrawingLayoutBundle } from '../inline-drawing-source.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function openPackage(paragraphs: number): OoxmlPackage {
  const body = Array.from(
    { length: paragraphs },
    (_, index) => `<w:p><w:r><w:t>Paragraph ${index}</w:t></w:r></w:p>`
  ).join('');
  const loaded = readOoxmlPackage(
    zipSync({
      '[Content_Types].xml': strToU8(
        `<Types xmlns="${CT}">` +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
          '</Types>'
      ),
      '_rels/.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
      ),
      'word/document.xml': strToU8(
        `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
      ),
    })
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  return loaded.package;
}

const paragraphsOf = (pkg: OoxmlPackage): OoxmlNode[] => {
  const root = pkg.parts.get(pkg.mainDocumentPart)!.root;
  const body = root.children.find((node) => node.kind !== 'textValue' && node.localName === 'body');
  if (!body || body.kind === 'textValue') throw new Error('body missing');
  return body.children.filter((node) => node.kind === 'paragraph');
};

const missing = Object.freeze({ kind: 'missing' as const });
const resourceLookup: ImageResourceLookup = {
  resolveEmbedded: async () => missing,
  resolveLinked: () => missing,
  resolveForProjection: async () => missing,
  liveReferenceCount: () => 0,
  dispose: () => {},
};

test('a reused drawing slot releases the paragraphs edited since it was made', () => {
  let pkg = openPackage(24);
  let revision = 0;
  const session = {
    packageRevision: () => revision,
    currentPackage: () => pkg,
    part: () => pkg.parts.get(pkg.mainDocumentPart)!,
  };
  const bundle = createInlineDrawingLayoutBundle({
    session,
    decodePort: Object.freeze({
      decode: async () => {
        throw new Error('unused');
      },
    }),
    resourceLookup,
    onResourcesChanged: () => {},
  });
  const context = bundle.bodyContext;

  // Every paragraph is edited once, as a long session does; only weak links to the
  // first versions leave this function.
  const firstVersions = (() => {
    const refs: WeakRef<OoxmlNode>[] = [];
    for (const paragraph of paragraphsOf(pkg)) {
      refs.push(new WeakRef(paragraph));
      const part = session.part();
      const edited = applyTreeOp(part, {
        op: 'insertText',
        paragraphId: paragraph.id,
        offset: 0,
        text: 'x',
      });
      if (!edited.ok) throw new Error(edited.reason);
      pkg = withPart(pkg, edited.part);
      revision += 1;
      bundle.sync(session);
    }
    return refs;
  })();

  // The edits left the drawings alone, so the slot was reused, not made again.
  expect(bundle.bodyContext).toBe(context);
  Bun.gc(true);
  // Only the older half is checked: other short-lived memos and the collector's stack scan
  // may still reach the last few states, with or without a slot.
  const kept = firstVersions.slice(0, 12).filter((ref) => ref.deref() !== undefined);
  expect(kept).toHaveLength(0);
});
