// Rich paste in suggesting mode, and the collaboration gate a rich paste asks.
//
// Word, with Track Changes on, pastes WITH formatting and records the paste as one tracked
// insertion. The editor used to degrade every suggesting-mode paste to plain text. It now
// lands the fragment as the store's tracked `insertFragment`, and the review pane shows it.
//
// The gate half: a fragment paste merges styles, numbering, media and relationships into the
// package inside its transaction, past `applyOps`. The gate is asked there, with every op the
// paste commits and the package those ops apply to, so a collaboration session sees the whole
// write. A refusal stops the paste: falling back to plain text would put the same words past
// the same gate. The store judges the fragment first, so a fragment it lands nowhere falls back
// to plain text, which the gate then judges.
//
// Where a paste lands: never in a table of contents (refused, as typing there is), and at a
// control's prompt in place of the prompt.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import {
  serializeOoxmlPart,
  type OoxmlNode,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type {
  CollaborationGateContext,
  EditorCollaborationSession,
} from '../../collaboration/index.ts';
import { paragraphTextOf } from '../../store/store/paragraph-model-text.ts';
import { revisionItemsOf } from '../../store/store/review-reads.ts';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';
import type { ScaleMutableSurface } from '../paginated-surface-contract.ts';
import { ensureSeparatorStyle } from '../list-separator-style.ts';
import { stubCollaborationSession } from './collaboration-test-module.ts';
import { docx, mount, paragraph, putCaret } from './paginated-surface-fixtures.ts';

const opened: { surface: PaginatedSurface; container: HTMLElement }[] = [];

afterEach(() => {
  document.getSelection()?.removeAllRanges();
  for (const item of opened.splice(0)) {
    item.surface.destroy();
    item.container.remove();
  }
});

const RICH_BODY =
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
  '<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>plain</w:t></w:r></w:p>';

/** Word's own clipboard HTML for a centred bold line and a plain one, without a fragment. */
const WORD_HTML =
  '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word">' +
  '<body><!--StartFragment--><p class=MsoNormal align=center style="text-align:center">' +
  '<b><span style="font-size:14.0pt">Word bold</span></b></p>' +
  '<p class=MsoNormal>Word plain</p><!--EndFragment--></body></html>';

function copyRichFlavours(): { text: string; html: string | null } {
  const source = mount(RICH_BODY);
  opened.push(source);
  putCaret(source.surface, 0);
  source.surface.selectAll();
  return source.surface.copyFlavours();
}

function mountWith(
  body: string,
  options: { author?: string; suggest?: boolean; session?: EditorCollaborationSession } = {}
): PaginatedSurface {
  const container = document.createElement('div');
  const result = mountPaginatedSurface(container, docx(body), {
    scale: 1,
    ...(options.author ? { author: options.author } : {}),
    ...(options.suggest ? { editingMode: 'suggest' as const } : {}),
    ...(options.session ? { collaborationModel: { session: options.session } } : {}),
  });
  if (!result.ok) throw new Error(`${result.reason}: ${result.detail ?? ''}`);
  opened.push({ surface: result.surface, container });
  return result.surface;
}

const markup = (surface: PaginatedSurface): string => serializeOoxmlPart(surface.session.part());

/** The review cards the body's revisions make. */
const cards = (surface: PaginatedSurface) => revisionItemsOf(surface.session.part());

/** Every `w:id` the insertions of `author` carry. */
function insertionIds(xml: string, author: string): Set<string> {
  const ids = new Set<string>();
  for (const match of xml.matchAll(/<w:ins ([^>]*)>|<w:ins ([^>]*)\/>/g)) {
    const attributes = match[1] ?? match[2] ?? '';
    if (!attributes.includes(`w:author="${author}"`)) continue;
    const id = /w:id="([^"]*)"/.exec(attributes)?.[1];
    if (id !== undefined) ids.add(id);
  }
  return ids;
}

describe('rich paste in suggesting mode', () => {
  test("the editor's own clipboard lands with its formatting as one tracked insertion", () => {
    const flavours = copyRichFlavours();
    expect(flavours.html!).toContain('data-docx-fragment="');
    const surface = mountWith(paragraph(''), { author: 'Reviewer', suggest: true });
    putCaret(surface, 0);
    expect(surface.pasteRich(flavours.text, flavours.html)).toBe(true);

    const xml = markup(surface);
    expect(surface.session.bodyText()).toContain('bold');
    // The formatting travelled: the paragraph's centring and the run's bold.
    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml).toContain('<w:b/>');
    // As one proposal, with no property records: the paragraphs did not exist before.
    expect(xml).toContain('w:author="Reviewer"');
    expect(insertionIds(xml, 'Reviewer').size).toBe(1);
    expect(xml.includes('w:pPrChange')).toBe(false);
    expect(xml.includes('w:rPrChange')).toBe(false);
    expect(cards(surface).length).toBeGreaterThan(0);
    expect(new Set(cards(surface).map((card) => card.author))).toEqual(new Set(['Reviewer']));

    // Rejecting it gives back the empty paragraph; accepting it gives the edit-mode paste.
    surface.session.applyTreeOps([{ op: 'rejectAllRevisions' }]);
    expect(surface.session.bodyText()).toBe('');
    surface.undo();
    surface.session.applyTreeOps([{ op: 'acceptAllRevisions' }]);
    const plain = mountWith(paragraph(''));
    putCaret(plain, 0);
    plain.pasteRich(flavours.text, flavours.html);
    expect(markup(surface)).toBe(markup(plain));
    expect(cards(surface)).toHaveLength(0);
  });

  test("Word's HTML lands with its formatting as a tracked insertion", () => {
    const surface = mountWith(paragraph('Host'), { author: 'Reviewer', suggest: true });
    putCaret(surface, 4);
    expect(surface.pasteRich('Word bold\nWord plain', WORD_HTML)).toBe(true);

    const xml = markup(surface);
    expect(surface.session.bodyText()).toContain('Word bold');
    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml).toContain('<w:b/>');
    expect(insertionIds(xml, 'Reviewer').size).toBe(1);
    expect(cards(surface).length).toBeGreaterThan(0);
    surface.session.applyTreeOps([{ op: 'rejectAllRevisions' }]);
    expect(surface.session.bodyText()).toBe('Host');
  });

  // Word absorbs another author's insertion into the paste's own and keeps that author's
  // deletion nested in it (paste evidence 1b, `~$paste-out-e1b.docx`)
  test("another author's changes land as Word pastes them: insertion absorbed, deletion nested", () => {
    const source = mount(
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
        '<w:ins w:id="1" w:author="Ada"><w:r><w:t>proposed</w:t></w:r></w:ins>' +
        '<w:del w:id="2" w:author="Ada"><w:r><w:delText>old</w:delText></w:r></w:del></w:p>'
    );
    opened.push(source);
    putCaret(source.surface, 0);
    source.surface.selectAll();
    const flavours = source.surface.copyFlavours();
    expect(flavours.html).not.toBeNull();

    const surface = mountWith(paragraph(''), { author: 'Reviewer', suggest: true });
    putCaret(surface, 0);
    expect(surface.pasteRich(flavours.text, flavours.html)).toBe(true);
    const xml = markup(surface);
    expect(surface.session.bodyText()).toContain('proposed');
    expect(xml).toMatch(
      /<w:ins [^>]*w:author="Reviewer"[^>]*><w:r><w:t>proposed<\/w:t><\/w:r><w:del [^>]*w:author="Ada"/
    );
    expect(xml.includes('<w:ins w:author="Ada"')).toBe(false);
    expect(
      cards(surface)
        .map((card) => card.author)
        .sort()
    ).toEqual(['Ada', 'Reviewer']);
  });

  // Word copies the accepted final view from a document that tracks changes (paste evidence
  // 1d, `~$paste-out-e1e.docx`)
  test('a copy from a document in suggesting mode is the accepted final view', () => {
    const source = mountWith(
      '<w:p><w:ins w:id="1" w:author="Ada"><w:r><w:t>proposed</w:t></w:r></w:ins>' +
        '<w:del w:id="2" w:author="Ada"><w:r><w:delText>old</w:delText></w:r></w:del></w:p>',
      { author: 'Ada', suggest: true }
    );
    putCaret(source, 0);
    source.selectAll();
    const flavours = source.copyFlavours();
    const surface = mountWith(paragraph(''), { author: 'Reviewer' });
    putCaret(surface, 0);
    expect(surface.pasteRich(flavours.text, flavours.html)).toBe(true);
    const xml = markup(surface);
    expect(surface.session.bodyText().trim()).toBe('proposed');
    expect(xml.includes('<w:ins')).toBe(false);
    expect(xml.includes('<w:del')).toBe(false);
    expect(xml.includes('Ada')).toBe(false);
  });

  // Word turns the link into a HYPERLINK field across pasted paragraphs (paste evidence 2,
  // `~$paste-out-e2-link-2p.docx`): the engine cannot write that yet, and plain text would
  // only approximate it, so the paste is refused by name in both modes
  test.each([false, true])(
    'paragraphs pasted inside a link are refused by name (suggest %p)',
    (suggest) => {
      const flavours = copyRichFlavours();
      const link =
        '<w:p><w:hyperlink w:anchor="target"><w:r><w:t>LinkText</w:t></w:r></w:hyperlink></w:p>';
      const surface = mountWith(link, { author: 'Reviewer', suggest });
      putCaret(surface, 4);
      expect(surface.pasteRich(flavours.text, flavours.html)).toBe(false);
      expect(surface.session.bodyText()).toBe('LinkText');
    }
  );

  test('without an author the paste keeps to the plain lane, which refuses it', () => {
    const flavours = copyRichFlavours();
    const surface = mountWith(paragraph('Host'), { suggest: true });
    putCaret(surface, 4);
    surface.pasteRich(flavours.text, flavours.html);
    expect(surface.session.bodyText()).toBe('Host');
    expect(markup(surface).includes('w:jc')).toBe(false);
  });

  test('edit mode pastes untracked, as before', () => {
    const flavours = copyRichFlavours();
    const surface = mountWith(paragraph(''), { author: 'Reviewer' });
    putCaret(surface, 0);
    surface.pasteRich(flavours.text, flavours.html);
    const xml = markup(surface);
    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml.includes('<w:ins')).toBe(false);
    expect(cards(surface)).toHaveLength(0);
  });
});

interface GateCall {
  readonly ops: readonly TreeDocOp[];
  readonly context: CollaborationGateContext | undefined;
}

function gatedMount(
  body: string,
  answer: (ops: readonly TreeDocOp[]) => string | null = () => null
): { surface: PaginatedSurface; calls: GateCall[] } {
  const calls: GateCall[] = [];
  const session = stubCollaborationSession({
    gateOperations: ((ops, _scope, context) => {
      calls.push({ ops, context });
      return answer(ops);
    }) as EditorCollaborationSession['gateOperations'],
  });
  return { surface: mountWith(body, { session }), calls };
}

/** Every `@w:val` of the `w:numId` elements under the nodes. */
function namedNumIds(nodes: readonly OoxmlNode[], out = new Set<string>()): Set<string> {
  for (const node of nodes) {
    if (node.kind === 'textValue') continue;
    if (node.localName === 'numId') {
      const value = node.attributes.find((attribute) => attribute.localName === 'val')?.value;
      if (value !== undefined) out.add(value);
    }
    namedNumIds(node.children, out);
  }
  return out;
}

/** The `w:numId` values `numbering.xml` defines in the package. */
function definedNumIds(pkg: OoxmlPackage | undefined): Set<string> {
  const ids = new Set<string>();
  const part = pkg?.parts.get('/word/numbering.xml');
  for (const node of part?.root.kind === 'textValue' ? [] : (part?.root.children ?? [])) {
    if (node.kind === 'textValue' || node.localName !== 'num') continue;
    const id = node.attributes.find((attribute) => attribute.localName === 'numId')?.value;
    if (id !== undefined) ids.add(id);
  }
  return ids;
}

const LIST_HTML = '<ol><li>First pasted item</li><li>Second pasted item</li></ol>';

describe('the collaboration gate a rich paste asks', () => {
  test('a paste at a caret is asked once, with the insertFragment and the merged numbering', () => {
    const { surface, calls } = gatedMount(paragraph('Host'));
    putCaret(surface, 4);
    calls.length = 0;
    expect(surface.pasteRich('First pasted item\nSecond pasted item', LIST_HTML)).toBe(true);

    expect(calls).toHaveLength(1);
    const [call] = calls;
    expect(call!.ops.map((op) => op.op)).toEqual(['insertFragment']);
    const insert = call!.ops[0] as Extract<TreeDocOp, { op: 'insertFragment' }>;
    const named = namedNumIds(insert.blocks);
    expect(named.size).toBeGreaterThan(0);
    // The package holds the `w:num` the blocks name; the document before the paste did not.
    const defined = definedNumIds(call!.context?.package);
    for (const id of named) expect(defined.has(id)).toBe(true);
    expect(surface.session.bodyText()).toContain('First pasted item');
  });

  test('a paste over a selection is asked once, with the clearing ops first', () => {
    const { surface, calls } = gatedMount(paragraph('Host text'));
    const paragraphId = surface.session.paragraphIds()[0]!;
    surface.setSelection({
      anchor: { paragraphId, offset: 4 },
      head: { paragraphId, offset: 9 },
    });
    calls.length = 0;
    surface.pasteRich('First pasted item\nSecond pasted item', LIST_HTML);

    expect(calls).toHaveLength(1);
    const kinds = calls[0]!.ops.map((op) => op.op);
    expect(kinds.at(-1)).toBe('insertFragment');
    expect(kinds.slice(0, -1)).toContain('deleteText');
    expect(surface.session.bodyText()).not.toContain(' text');
  });

  test('a refusal leaves the document unchanged and does not fall back to plain text', () => {
    const { surface, calls } = gatedMount(paragraph('Host'), (ops) =>
      ops.some((op) => op.op === 'insertFragment') ? 'collaboration-session-not-ready' : null
    );
    putCaret(surface, 4);
    const before = markup(surface);
    calls.length = 0;
    expect(surface.pasteRich('First pasted item\nSecond pasted item', LIST_HTML)).toBe(false);

    expect(markup(surface)).toBe(before);
    expect(surface.session.bodyText()).toBe('Host');
    // Asked once: neither the plain lane nor another flavour tried again.
    expect(calls).toHaveLength(1);
  });

  test('a batch with package edits is asked with the edited package', () => {
    const { surface, calls } = gatedMount(paragraph('Host'));
    const paragraphId = surface.session.paragraphIds()[0]!;
    const current = surface.session.currentPackage();
    calls.length = 0;
    const result = (surface as ScaleMutableSurface).applyAutomationOps(
      () => [{ op: 'insertText', paragraphId, offset: 4, text: '!' }],
      { kind: 'body' },
      [ensureSeparatorStyle(null, 'GateProbe')]
    );
    expect(result.committed).toBe(true);

    expect(calls).toHaveLength(1);
    const gated = calls[0]!.context?.package;
    expect(gated).toBeDefined();
    expect(gated).not.toBe(current);
    const styles = gated!.parts.get('/word/styles.xml');
    expect(styles && serializeOoxmlPart(styles)).toContain('w:styleId="GateProbe"');
  });

  test('a batch without package edits is asked with the current package', () => {
    const seen: { gated: OoxmlPackage | undefined; live: OoxmlPackage }[] = [];
    const session = stubCollaborationSession({
      gateOperations: ((_ops, _scope, context) => {
        seen.push({ gated: context?.package, live: surface.session.currentPackage() });
        return null;
      }) as EditorCollaborationSession['gateOperations'],
    });
    const surface = mountWith(paragraph('Host'), { session });
    const paragraphId = surface.session.paragraphIds()[0]!;
    seen.length = 0;
    (surface as ScaleMutableSurface).applyAutomationOps(
      () => [{ op: 'insertText', paragraphId, offset: 4, text: '!' }],
      { kind: 'body' }
    );
    expect(seen).toHaveLength(1);
    expect(seen[0]!.gated).toBe(seen[0]!.live);
    expect(surface.session.bodyText()).toBe('Host!');
  });
});

describe('where a paste lands', () => {
  const TOC_BEGIN =
    '<w:r><w:fldChar w:fldCharType="begin"/><w:instrText>TOC \\o "1-1" \\h</w:instrText>' +
    '<w:fldChar w:fldCharType="separate"/></w:r>';
  const TOC_END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

  test('a fragment the store lands nowhere is never put to the gate: plain text goes', () => {
    // Inside a plain text control the store refuses a fragment; the editor pastes the text. A
    // session must judge that text, not refuse the fragment the editor never commits.
    const { surface, calls } = gatedMount(
      '<w:p><w:r><w:t xml:space="preserve">Alpha </w:t></w:r><w:sdt><w:sdtPr><w:text/></w:sdtPr>' +
        '<w:sdtContent><w:r><w:t>Rich</w:t></w:r></w:sdtContent></w:sdt></w:p>',
      (ops) => (ops.some((op) => op.op === 'insertFragment') ? 'collaboration-refused' : null)
    );
    putCaret(surface, 8);
    calls.length = 0;
    expect(surface.pasteRich('Word bold', WORD_HTML)).toBe(true);
    expect(calls.flatMap((call) => call.ops.map((op) => op.op))).not.toContain('insertFragment');
    // Word's Smart cut and paste puts a space on each side between two words
    expect(surface.session.bodyText()).toContain('Ri Word bold ch');
  });

  test('a paste in a table of contents is refused, as typing there is', () => {
    const surface = mountWith(
      `<w:p>${TOC_BEGIN}<w:r><w:t>Entry 1</w:t></w:r></w:p><w:p><w:r><w:t>Page 1</w:t></w:r>${TOC_END}</w:p>`
    );
    putCaret(surface, 2);
    const before = markup(surface);
    surface.pasteRich('Word bold', WORD_HTML);
    expect(markup(surface)).toBe(before);
    expect(surface.state().lastRejection).toBe('the table of contents is generated and read-only');
  });

  test("lines pasted at a control's prompt replace it, each cut where it ends", () => {
    const surface = mountWith(
      '<w:p><w:r><w:t xml:space="preserve">Alpha </w:t></w:r><w:sdt><w:sdtPr><w:showingPlcHdr/>' +
        '</w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>' +
        '<w:t>Click here to enter text.</w:t></w:r></w:sdtContent></w:sdt></w:p>'
    );
    // Inside the prompt, three characters past its start: the text lands where it begins
    putCaret(surface, 9);
    surface.pasteRich('One\nTwo\nThree', null);
    const part = surface.session.part();
    const texts = surface.session.paragraphIds().map((id) => paragraphTextOf(part, id));
    expect(texts).toEqual(['Alpha One', 'Two', 'Three']);
    expect(markup(surface)).not.toContain('showingPlcHdr');
  });
});
