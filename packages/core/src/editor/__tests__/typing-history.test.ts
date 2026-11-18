// The typing run's boundary rules, without a surface in the loop.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterAll, describe, expect, test } from 'bun:test';
import type { OoxmlPackage, TreeModelChange } from '@docx-editor.dev/core/store';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { readOoxmlPackage } from '../../store/package/ooxml-package.ts';
import { ORIGIN_IDS } from '../../store/registry/frozen-ids.ts';
import { TypingHistory } from '../typing-history.ts';
import { docx, mount, paragraph } from './paginated-surface-fixtures.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// A real package and paragraph id, so node lookups resolve the way they do in a session.
const fixture = mount(paragraph('one') + paragraph('two'));
const original = fixture.surface.session.currentPackage();
const [p1, p2] = fixture.surface.session.paragraphIds() as [string, string];
afterAll(() => {
  fixture.surface.destroy();
  fixture.container.remove();
});

const caret = (paragraphId: string, offset: number) => ({
  anchor: { paragraphId, offset },
  head: { paragraphId, offset },
});

function change(overrides: Partial<TreeModelChange>): TreeModelChange {
  return {
    change: 'model-change',
    fromRevision: 1,
    toRevision: 2,
    commitId: 'c',
    origin: ORIGIN_IDS.mutationRemote,
    dirty: [],
    created: [],
    deleted: [],
    splitJoin: [],
    dependencyKeys: [],
    impact: 'global',
    ...overrides,
  };
}

/** A run with one flush landed in the first paragraph at offset 2. */
function openRun(): { history: TypingHistory; group: symbol; swap: (pkg: OoxmlPackage) => void } {
  let current = original;
  const history = new TypingHistory(() => current);
  const group = history.groupAt(caret(p1, 1));
  history.landed(caret(p1, 2));
  return { history, group, swap: (pkg) => (current = pkg) };
}

describe('TypingHistory', () => {
  test('a flush from where the last one ended continues the run', () => {
    const { history, group } = openRun();
    expect(history.groupAt(caret(p1, 2))).toBe(group);
  });

  test('a flush from anywhere else starts a new run', () => {
    const { history, group } = openRun();
    expect(history.groupAt(caret(p1, 1))).not.toBe(group);
  });

  test("a foreign edit to the run's paragraph ends the run", () => {
    const { history, group } = openRun();
    history.noteForeignChange(change({ dirty: [p1], impact: 'text-local' }));
    expect(history.groupAt(caret(p1, 2))).not.toBe(group);
  });

  test('a foreign edit to another paragraph keeps the run', () => {
    const { history, group } = openRun();
    history.noteForeignChange(change({ dirty: [p2], impact: 'text-local' }));
    expect(history.groupAt(caret(p1, 2))).toBe(group);
  });

  /** The package with the main part rebuilt as new objects, its text optionally rewritten. */
  function rebuilt(rewrite?: (value: string) => string): OoxmlPackage {
    const main = original.parts.get(original.mainDocumentPart)!;
    const root = structuredClone(main.root);
    const walk = (node: { kind: string; value?: string; children?: unknown[] }): void => {
      if (node.kind === 'textValue' && rewrite) node.value = rewrite(node.value!);
      for (const child of node.children ?? []) walk(child as typeof node);
    };
    walk(root as never);
    const parts = new Map(original.parts);
    parts.set(original.mainDocumentPart, { ...main, root });
    return { ...original, parts };
  }

  test("a change that names no paragraphs keeps the run while the run's paragraph is unchanged", () => {
    // A wholesale install after an edit in another part can rebuild untouched nodes.
    const { history, group, swap } = openRun();
    swap(rebuilt());
    history.noteForeignChange(change({}));
    expect(history.groupAt(caret(p1, 2))).toBe(group);
  });

  test("an unchanged line holding a shape keeps the run: the choice's prefixes are the root's", () => {
    // Word declares `wps` on `w:document` and names it in the choice's `Requires`
    const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
    const shapes = readOoxmlPackage(
      zipSync({
        '[Content_Types].xml': unzipSync(docx(''))['[Content_Types].xml']!,
        '_rels/.rels': unzipSync(docx(''))['_rels/.rels']!,
        'word/document.xml': strToU8(
          `<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:wps="${WPS}"><w:body>` +
            '<w:p><w:r><w:t>one</w:t></w:r><w:r><mc:AlternateContent>' +
            '<mc:Choice Requires="wps"><w:t>a</w:t></mc:Choice><mc:Fallback><w:t>b</w:t></mc:Fallback>' +
            '</mc:AlternateContent></w:r></w:p></w:body></w:document>'
        ),
      })
    );
    if (!shapes.ok) throw new Error(shapes.reason);
    const pkg = shapes.package;
    const main = pkg.parts.get(pkg.mainDocumentPart)!;
    const body = main.root.children[0]!;
    const line = body.kind === 'textValue' ? undefined : body.children[0];
    if (!line) throw new Error('no line');
    let current = pkg;
    const history = new TypingHistory(() => current);
    const group = history.groupAt(caret(line.id, 1));
    history.landed(caret(line.id, 2));
    // A wholesale install rebuilds the line as new objects with the same content
    const parts = new Map(pkg.parts);
    parts.set(pkg.mainDocumentPart, { ...main, root: structuredClone(main.root) });
    current = { ...pkg, parts };
    history.noteForeignChange(change({}));
    expect(history.groupAt(caret(line.id, 2))).toBe(group);
  });

  test("a change that names no paragraphs ends the run when the run's paragraph changed", () => {
    const { history, group, swap } = openRun();
    swap(rebuilt((value) => (value === 'one' ? 'xone' : value)));
    history.noteForeignChange(change({}));
    expect(history.groupAt(caret(p1, 2))).not.toBe(group);
  });

  test("a change that names no paragraphs ends the run when the run's paragraph is gone", () => {
    const { history, group, swap } = openRun();
    const parts = new Map(original.parts);
    parts.delete(original.mainDocumentPart);
    swap({ ...original, parts });
    history.noteForeignChange(change({}));
    expect(history.groupAt(caret(p1, 2))).not.toBe(group);
  });

  test('projection and awareness commits are not edits', () => {
    const { history, group } = openRun();
    history.noteForeignChange(change({ origin: ORIGIN_IDS.projection, dirty: [p1] }));
    history.noteForeignChange(change({ origin: ORIGIN_IDS.awareness, dirty: [p1] }));
    expect(history.groupAt(caret(p1, 2))).toBe(group);
  });

  test('a flush that leaves a range selection ends the run', () => {
    const { history, group } = openRun();
    history.landed({
      anchor: { paragraphId: p1, offset: 0 },
      head: { paragraphId: p1, offset: 2 },
    });
    expect(history.groupAt(caret(p1, 2))).not.toBe(group);
  });
});
