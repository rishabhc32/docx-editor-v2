// The matrix of group 9.8 (task 9.11): every organiser edit of headers, footers and sections on
// several files, in an ordered room and in the editor alone. For each cell the room's save says
// what the editor's own save says (each section's properties in order, the text and watermark
// of each header and footer it names, the odd and even pages setting, the body text); undo in
// the room gives back the file; and the room's save reopens to itself. A cell the editor alone
// refuses is refused in the room too.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { collectSectionPropertyNodes } from '../../../../store/package/hf-references.ts';
import { metricsOfSection } from '../../../../store/store/tree-op-section-address.ts';
import { readOoxmlPackage, type OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import {
  serializeOoxmlPart,
  type OoxmlElement,
  type OoxmlNode,
} from '../../../../store/package/ooxml-tree.ts';
import { resolveRelationship } from '../../../../store/package/relationships.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import { docx, p, R, t, WML_CT } from '../word-fixtures.ts';
import {
  closeAll,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
  type Editor,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { converged, FILE, room2, TWO } from './sections-fixtures.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';

afterEach(closeAll);

// ---- the files ----

const HEADER_PART =
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>';

/** Two sections with columns, page numbers, line numbers and vertical alignment. */
const FANCY = docx(
  p(
    t('Two columns here.'),
    '<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
      '<w:cols w:num="2" w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>'
  ) +
    p(t('Then one column.')) +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/>' +
    '<w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>' +
    '<w:lnNumType w:countBy="1"/><w:pgNumType w:fmt="lowerRoman" w:start="3"/>' +
    '<w:vAlign w:val="center"/><w:docGrid w:linePitch="360"/></w:sectPr>',
  {
    parts: { 'word/header1.xml': HEADER_PART },
    relationships: `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes: `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>`,
  }
);

/** One paragraph, and no section the file states. */
const BARE = docx(p(t('Body text only.')));

const FILES: Readonly<Record<string, Uint8Array>> = { FILE, TWO, FANCY, BARE };

// ---- the edits ----

type Lifecycle = Parameters<
  NonNullable<ReturnType<typeof surface>['applyHeaderFooterLifecycle']>
>[0];

type Action = (editor: Editor) => boolean;

const lifecycleOf =
  (op: Lifecycle): Action =>
  (editor) =>
    editor.surface?.applyHeaderFooterLifecycle?.(op)?.ok === true;

const sectionBreak =
  (type: 'nextPage' | 'continuous'): Action =>
  (editor) => {
    const s = editor.surface;
    const [first] = s?.session.paragraphIds() ?? [];
    if (!s || first === undefined) return false;
    s.setSelection({
      anchor: { paragraphId: first, offset: 3 },
      head: { paragraphId: first, offset: 3 },
    });
    return s.insertSectionBreak(type);
  };

const ACTIONS: Readonly<Record<string, Action>> = {
  'landscape, whole document': (editor) =>
    editor.surface?.setSectionProperties({ orientation: 'landscape' }) === true,
  "margins, the first paragraph's section": (editor) => {
    const [first] = editor.surface?.session.paragraphIds() ?? [];
    return (
      editor.surface?.setSectionProperties({
        marginLeftTwips: 2160,
        marginTopTwips: 720,
        ...(first === undefined ? {} : { anchorParagraphId: first }),
      }) === true
    );
  },
  'a different first page': lifecycleOf({
    op: 'setSectionFurnitureOptions',
    sectionIndex: 0,
    titlePage: true,
  }),
  'header distance': lifecycleOf({
    op: 'setSectionFurnitureOptions',
    sectionIndex: 0,
    headerDistanceTwips: 500,
  }),
  'footer distance': lifecycleOf({
    op: 'setSectionFurnitureOptions',
    sectionIndex: 0,
    footerDistanceTwips: 540,
  }),
  'odd and even pages': lifecycleOf({ op: 'setSectionFurnitureOptions', evenAndOddHeaders: true }),
  'a new footer': lifecycleOf({
    op: 'createHeaderFooter',
    sectionIndex: 0,
    kind: 'footer',
    variant: 'default',
  }),
  'a new first-page header': lifecycleOf({
    op: 'createHeaderFooter',
    sectionIndex: 0,
    kind: 'header',
    variant: 'first',
    titlePage: true,
  }),
  'the last header removed': (editor) => {
    const sections = editor.surface?.session.headerFooterResolutionBySection() ?? [];
    const index = sections.findIndex((s) => s.headers.get('default')?.inherited === false);
    return (
      index >= 0 &&
      lifecycleOf({
        op: 'deleteHeaderFooter',
        sectionIndex: index,
        kind: 'header',
        variant: 'default',
      })(editor)
    );
  },
  'the second section unlinked': lifecycleOf({
    op: 'unlinkFromPrevious',
    sectionIndex: 1,
    kind: 'header',
    variant: 'default',
  }),
  'a next-page break': sectionBreak('nextPage'),
  'a continuous break': sectionBreak('continuous'),
  'a watermark': (editor) =>
    editor.exec({ type: 'setWatermark', watermark: { text: 'DRAFT' } }).ok === true,
};

// ---- what a save says ----

/** A node as text: its name, its attributes in order of name, and its children. */
function nodeText(node: OoxmlNode): string {
  if (node.kind === 'textValue') return node.value;
  const attributes = node.attributes
    .map((a) => `${a.localName}=${a.value}`)
    .sort()
    .join(' ');
  return `<${node.localName} ${attributes}>${node.children.map(nodeText).join('')}`;
}

/** The text of a part, and what its watermarks say. */
function storyText(pkg: OoxmlPackage, name: string): string {
  const part = pkg.parts.get(name);
  if (!part) return '<no part>';
  const xml = serializeOoxmlPart(part);
  const marks = [...xml.matchAll(/<v:textpath [^>]*string="([^"]*)"/g)].map((m) => m[1]);
  return `${xml.replace(/<[^>]+>/g, '')}${marks.length > 0 ? ` [watermark ${marks.join(',')}]` : ''}`;
}

/**
 * A section's page as it means: size, orientation and margins, Word's defaults for what it does
 * not state (ECMA-376 Part 1, §17.6.11 and §17.6.13).
 */
function pageOf(sectPr: OoxmlElement | null): string {
  const pgSz = sectPr?.children.find(
    (child) => child.kind !== 'textValue' && child.localName === 'pgSz'
  );
  const stated =
    pgSz && pgSz.kind !== 'textValue'
      ? pgSz.attributes.find((a) => a.localName === 'orient')?.value
      : undefined;
  const orient = stated ?? 'portrait';
  return `page ${JSON.stringify({ ...metricsOfSection(sectPr ?? null), orient })}`;
}

/**
 * Each section's properties in order, with the stories its references name; the settings; the
 * text. With `meaning`, each section's page is its effective size, orientation and margins.
 */
function signature(bytes: Uint8Array, meaning = false): unknown {
  const read = readOoxmlPackage(bytes, {});
  if (!read.ok) throw new Error(read.reason);
  const pkg = read.package;
  const mainName = pkg.mainDocumentPart;
  const main = pkg.parts.get(mainName);
  if (!main) throw new Error('no main part');
  const relationships = pkg.relationships.get(mainName) ?? [];
  const target = (id: string) => {
    const record = relationships.find((r) => r.id === id);
    const resolved = record ? resolveRelationship(record) : null;
    return resolved?.mode === 'Internal' && resolved.target.ok ? resolved.target.partName : '';
  };
  const sections = collectSectionPropertyNodes(main.root).map((sectPr) => [
    ...(meaning ? [pageOf(sectPr)] : []),
    ...(sectPr?.children ?? []).flatMap((child) => {
      if (child.kind === 'textValue') return [];
      if (meaning && (child.localName === 'pgSz' || child.localName === 'pgMar')) return [];
      if (child.localName !== 'headerReference' && child.localName !== 'footerReference')
        return [nodeText(child)];
      const type = child.attributes.find((a) => a.localName === 'type')?.value;
      const id = child.attributes.find((a) => a.localName === 'id' && a.prefix === 'r')?.value;
      return [`${child.localName}:${type} ${storyText(pkg, target(id ?? ''))}`];
    }),
  ]);
  const settings = [...pkg.parts.values()].find((part) => part.root.localName === 'settings');
  const evenAndOdd = settings ? serializeOoxmlPart(settings).includes('evenAndOddHeaders') : false;
  return { sections, evenAndOdd, text: storyText(pkg, mainName) };
}

/** The editor alone, with no room: the edit on the file, and its save; null when refused. */
async function alone(file: Uint8Array, action: Action | null): Promise<Uint8Array | null> {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: file, author: 'Olive', mode: 'edit' });
  try {
    if (action && !action(editor)) return null;
    return new Uint8Array(await editor.save());
  } finally {
    editor.destroy();
    container.remove();
  }
}

describe('headers, footers and sections: the room and the editor alone', () => {
  for (const [fileName, file] of Object.entries(FILES))
    for (const [name, action] of Object.entries(ACTIONS))
      test(`${fileName} | ${name}`, async () => {
        const own = await alone(file, action);
        const { room, ana, olive } = await room2(file);
        const done = action(olive.editor);
        await settle();
        // The room refuses exactly what the editor alone refuses
        expect(done).toBe(own !== null);
        if (own === null) {
          if (process.env.ORDERED_MATRIX_REPORT)
            console.warn(`refused by both: ${fileName} | ${name}`);
          return;
        }
        converged(room, ana, olive);
        const saved = new Uint8Array(await ana.editor.save());
        expect(signature(saved)).toEqual(signature(own));
        const before = await alone(file, null);
        if (before === null) throw new Error('the file does not open');
        // The cell means something: the edit changed what the save says
        expect(signature(saved)).not.toEqual(signature(before));
        // The save reopens to itself
        const reopened = openOrderedWordFile(saved, { file: 2 });
        expect(signature(orderedPackageBytes(reopened.document, reopened.base))).toEqual(
          signature(saved)
        );
        // Undo gives back the file
        surface(olive).undo();
        await settle();
        converged(room, ana, olive);
        expect(signature(new Uint8Array(await ana.editor.save()))).toEqual(signature(before));
      });
});

/** Decide the first revision card, as the organiser does. */
function decide(editor: Editor, decision: 'accept' | 'reject'): boolean {
  const [card] = editor.getReviewItems().filter((item) => item.kind === 'revision');
  if (!card) return false;
  const done =
    decision === 'accept' ? editor.acceptReviewItem(card.key) : editor.rejectReviewItem(card.key);
  return done.ok;
}

/** What happens to a suggestion: decided by the organiser, taken back by its author, or kept. */
type Outcome = 'accept' | 'reject' | 'undo' | 'pending';

/** The suggestions of the group: each is a tracked change in the editor alone. */
const SUGGESTED: Readonly<Record<string, Action>> = {
  'a next-page break': sectionBreak('nextPage'),
  'a continuous break': sectionBreak('continuous'),
  'landscape, whole document': ACTIONS['landscape, whole document'] ?? (() => false),
  "margins, the first paragraph's section":
    ACTIONS["margins, the first paragraph's section"] ?? (() => false),
  'a different first page': ACTIONS['a different first page'] ?? (() => false),
  'header distance': ACTIONS['header distance'] ?? (() => false),
  'footer distance': ACTIONS['footer distance'] ?? (() => false),
};

/** The editor alone: an edit suggested, then its outcome; null when the editor refuses it. */
async function aloneSuggested(
  file: Uint8Array,
  action: Action,
  outcome: Outcome
): Promise<Uint8Array | null> {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: file,
    author: 'Ana',
    mode: 'suggesting',
    modules: [reviewTestModule()],
  });
  try {
    if (!action(editor)) return null;
    if (outcome === 'undo') editor.surface?.undo();
    else if (outcome !== 'pending') {
      editor.surface?.setAuthor('Olive');
      editor.surface?.setEditingMode('edit');
      if (!decide(editor, outcome)) throw new Error('no card to decide');
    }
    return new Uint8Array(await editor.save());
  } finally {
    editor.destroy();
    container.remove();
  }
}

/** A save's signature without what names the record's author, date and id. */
const unattributed = (bytes: Uint8Array): unknown =>
  JSON.parse(JSON.stringify(signature(bytes)).replace(/<sectPrChange [^>]*>/g, '<sectPrChange>'));

/** Two saves mean the same document. */
const same = (a: Uint8Array, b: Uint8Array): boolean =>
  JSON.stringify(signature(a, true)) === JSON.stringify(signature(b, true));

/**
 * One cell: `author` suggests the edit in a room, the save while pending is the editor alone's,
 * then the outcome gives what it gives in the editor alone.
 */
async function suggestedCell(
  file: Uint8Array,
  action: Action,
  outcome: Exclude<Outcome, 'pending'>,
  author: 'party' | 'organiser'
): Promise<void> {
  const own = await aloneSuggested(file, action, outcome);
  const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
  const review = reviewTestModule();
  const ana = await join(room, 'Ana', { wordFile: file, review });
  const olive = await join(room, 'Olive', {
    mode: author === 'party' ? 'edit' : 'suggesting',
    login: ORGANISER,
    wordFile: file,
    review,
  });
  const by = author === 'party' ? ana : olive;
  const done = action(by.editor);
  await settle();
  // The room refuses what the editor alone refuses
  expect(done).toBe(own !== null);
  if (own === null) return;
  const before = await alone(file, null);
  const rejected = await aloneSuggested(file, action, 'reject');
  if (before === null || rejected === null) throw new Error('the file does not open');
  // The host's earlier check: a party's change holds only when rejecting it gives back the file. A
  // break's type for the next section is written untracked and stays: a direct edit. (A page
  // setup's reject in the editor alone states the earlier page in full, as its record does: the
  // same page, the room gives back the file as it was.)
  if (author === 'party' && !same(rejected, before)) {
    expect(ana.refusals.map((r) => r.reason)).toEqual(['package']);
    ana.refusals.length = 0;
    converged(room, ana, olive);
    expect(signature(new Uint8Array(await ana.editor.save()))).toEqual(signature(before));
    return;
  }
  converged(room, ana, olive);
  const pending = await aloneSuggested(file, action, 'pending');
  if (pending === null) throw new Error('the editor alone refused it');
  // The suggestion saves as the editor alone saves it, Word's form
  expect(unattributed(new Uint8Array(await ana.editor.save()))).toEqual(unattributed(pending));
  if (outcome === 'undo') surface(by).undo();
  else {
    olive.editor.surface?.setEditingMode('edit');
    expect(decide(olive.editor, outcome)).toBe(true);
  }
  await settle();
  converged(room, ana, olive);
  const meaning = outcome === 'reject';
  expect(signature(new Uint8Array(await ana.editor.save()), meaning)).toEqual(
    signature(own, meaning)
  );
}

describe("a party's suggestion: the room and the editor alone", () => {
  for (const [fileName, file] of Object.entries(FILES))
    for (const [name, action] of Object.entries(SUGGESTED))
      for (const outcome of ['accept', 'reject', 'undo'] as const)
        test(`${fileName} | ${name}, ${outcome}`, () =>
          suggestedCell(file, action, outcome, 'party'));
});

describe("the organiser's suggestion: the room and the editor alone", () => {
  for (const [fileName, file] of Object.entries(FILES))
    for (const [name, action] of Object.entries(SUGGESTED))
      for (const outcome of ['accept', 'reject', 'undo'] as const)
        test(`${fileName} | ${name}, ${outcome}`, () =>
          suggestedCell(file, action, outcome, 'organiser'));
});

// ---- a party's break inside another person's pending break ----

/** A next-page break at an offset of the paragraph at `index`. */
const breakAt =
  (index: number, offset: number): Action =>
  (editor) => {
    const s = editor.surface;
    const paragraphId = s?.session.paragraphIds()[index];
    if (!s || paragraphId === undefined) return false;
    s.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
    return s.insertSectionBreak('nextPage');
  };

/** Decide every card of one author, as the organiser does. */
function decideAll(editor: Editor, author: string, decision: 'accept' | 'reject'): void {
  const cards = editor
    .getReviewItems()
    .filter((item) => item.kind === 'revision' && item.author === author);
  expect(cards.length).toBeGreaterThan(0);
  for (const card of cards) {
    const done =
      decision === 'accept' ? editor.acceptReviewItem(card.key) : editor.rejectReviewItem(card.key);
    expect(done.ok).toBe(true);
  }
}

/**
 * The editor alone: Ben suggests a break in the second paragraph, Ana one in the section his
 * break ends, then the organiser decides Ben's.
 */
async function aloneBreakInBreak(outcome: 'accept' | 'reject' | 'pending'): Promise<Uint8Array> {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: FILE,
    author: 'Ben',
    mode: 'suggesting',
    modules: [reviewTestModule()],
  });
  try {
    if (!breakAt(1, 4)(editor)) throw new Error("Ben's break");
    editor.surface?.setAuthor('Ana');
    if (!breakAt(0, 5)(editor)) throw new Error("Ana's break");
    if (outcome !== 'pending') {
      editor.surface?.setAuthor('Olive');
      editor.surface?.setEditingMode('edit');
      decideAll(editor, 'Ben', outcome);
    }
    return new Uint8Array(await editor.save());
  } finally {
    editor.destroy();
    container.remove();
  }
}

describe("a party's break in another person's pending break: the room and the editor alone", () => {
  for (const outcome of ['accept', 'reject'] as const)
    test(`a decision on Ben's break leaves Ana's: ${outcome}`, async () => {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const review = reviewTestModule();
      const ana = await join(room, 'Ana', { wordFile: FILE, review });
      const ben = await join(room, 'Ben', { wordFile: FILE, review });
      const olive = await join(room, 'Olive', {
        mode: 'edit',
        login: ORGANISER,
        wordFile: FILE,
        review,
      });
      const all = [ana, ben, olive];
      expect(breakAt(1, 4)(ben.editor)).toBe(true);
      await settle();
      converged(room, ...all);
      // Ben's break has no section record of his for Ana's copy to take: her break is hers
      expect(breakAt(0, 5)(ana.editor)).toBe(true);
      await settle();
      converged(room, ...all);
      expect(unattributed(new Uint8Array(await olive.editor.save()))).toEqual(
        unattributed(await aloneBreakInBreak('pending'))
      );
      decideAll(olive.editor, 'Ben', outcome);
      await settle();
      converged(room, ...all);
      const saved = new Uint8Array(await ana.editor.save());
      expect(signature(saved)).toEqual(signature(await aloneBreakInBreak(outcome)));
      // Ana's break is still her suggestion, with the section Ben's break ended when she made it
      const left = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
      expect(left.map((item) => item.author)).toEqual(['Ana']);
      expect(surface(ana).session.headerFooterResolutionBySection().length).toBe(
        outcome === 'accept' ? 3 : 2
      );
      // The save reopens to itself
      const reopened = openOrderedWordFile(saved, { file: 2 });
      expect(signature(orderedPackageBytes(reopened.document, reopened.base))).toEqual(
        signature(saved)
      );
    });
});
