// The deprecated `EditOptions.placeholderText` still gives a caller the prompt it gave before.
//
// A server that checks suggestions applies one op at a time to one part, with no glossary of its
// own: it lends the glossary block's TEXT through `placeholderText`, and takes an author's
// suggestions back through the decision planner. `placeholderBlocks` and `retainPrompt` replaced
// that option for the editor, and the server must see the same result as before they came: the
// expected bodies below are what the store wrote before them, for the same calls. The untyped
// cases differ on purpose: a control that declares no type is rich text, as in Word, so taking
// back the words typed over its prompt shows its prompt, the glossary block's text or the type's,
// where it used to leave the control empty. A control whose prompt no glossary block keeps, and
// no reference names, differs too: the decision that empties it gives Word's blank prompt, five
// spaces under `w:showingPlcHdr` (Word 16.114 samples `~$i6w-*-out.docx`, `~$p4-enter-rejected.docx`).

import { describe, expect, test } from 'bun:test';
import {
  collectReviewItems,
  contentControlsIn,
  planRevisionBatch,
  readOoxmlPart,
  reviewItemKey,
  serializeOoxmlPart,
  type EditOptions,
  type OoxmlNode,
  type OoxmlPart,
} from '../index.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const revision = { author: 'Bea', date: '2026-10-16T00:00:00Z' };
const GREY = '<w:color w:val="808080"/>';
const STYLE = '<w:rStyle w:val="PlaceholderText"/>';
/** Word's blank prompt and the end properties Word writes with it. */
const BLANK = '<w:r><w:t xml:space="preserve">     </w:t></w:r>';
const BLANK_END = '<w:sdtEndPr/>';
const KEPT = '0123456789ABCDEF0123456789ABCDEF';

const line = (properties: string, content: string) =>
  `<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt>${properties}` +
  `<w:sdtContent>${content}</w:sdtContent></w:sdt></w:p>`;
const pr = (inner: string) => (inner ? `<w:sdtPr>${inner}</w:sdtPr>` : '<w:sdtPr/>');
const run = (face: string, text: string) => `<w:r><w:rPr>${face}</w:rPr><w:t>${text}</w:t></w:r>`;
const typed = (face: string, text: string) =>
  `<w:ins w:author="Bea" w:date="${revision.date}" w:id="0">` +
  (face ? run(face, text) : `<w:r><w:t>${text}</w:t></w:r>`) +
  '</w:ins>';
const ref = (name: string) => `<w:placeholder><w:docPart w:val="${name}"/></w:placeholder>`;
const LIST =
  '<w:listItem w:displayText="One" w:value="1"/><w:listItem w:displayText="Two" w:value="2"/>';

/** The glossary's words, as the server reads them from its own package. */
const placeholderText = (docPart: string): string | undefined =>
  ({ PartyName: 'Enter a name.', [KEPT]: 'Enter your full name' })[docPart];

interface Case {
  readonly name: string;
  readonly before: string;
  /** Typing over the prompt, or a pick in its list. */
  readonly edit: (part: OoxmlPart) => TreeDocOp;
  /** The pre-fix store's body after the edit, and after the take-back. */
  readonly typed: string;
  readonly back: string;
}

const typeOver = (part: OoxmlPart): TreeDocOp => ({
  op: 'insertText',
  paragraphId: (part.root.children[0] as { children: readonly OoxmlNode[] }).children[0]!.id,
  offset: 'Name: '.length,
  text: 'Ana',
  revision,
});

const CASES: readonly Case[] = [
  {
    name: 'a prompt no glossary block holds',
    before: line(pr('<w:showingPlcHdr/><w:text/>'), run(GREY, 'Type here')),
    edit: typeOver,
    typed: line(pr('<w:text/>'), typed(GREY, 'Ana')),
    back: line(pr('<w:showingPlcHdr/><w:text/>') + BLANK_END, BLANK),
  },
  {
    name: 'a prompt a Word glossary block holds',
    before: line(pr(`${ref('PartyName')}<w:showingPlcHdr/><w:text/>`), run(STYLE, 'Enter a name.')),
    edit: typeOver,
    typed: line(pr(`${ref('PartyName')}<w:text/>`), typed('', 'Ana')),
    back: line(pr(`${ref('PartyName')}<w:showingPlcHdr/><w:text/>`), run(STYLE, 'Enter a name.')),
  },
  {
    name: 'a prompt the editor kept in the glossary',
    before: line(
      pr(`${ref(KEPT)}<w:showingPlcHdr/><w:text/>`),
      run(STYLE + GREY, 'Enter your full name')
    ),
    edit: typeOver,
    typed: line(pr(`${ref(KEPT)}<w:text/>`), typed(GREY, 'Ana')),
    back: line(pr(`${ref(KEPT)}<w:showingPlcHdr/><w:text/>`), run(STYLE, 'Enter your full name')),
  },
  {
    name: 'an untyped control with no glossary block',
    before: line(pr('<w:showingPlcHdr/>'), run(`<w:b/>${GREY}`, 'Party')),
    edit: typeOver,
    typed: line(pr(''), typed(`<w:b/>${GREY}`, 'Ana')),
    // Rich text, as in Word: Word's blank prompt, as for a typed control with no block.
    back: line(pr('<w:showingPlcHdr/>') + BLANK_END, BLANK),
  },
  {
    // Word writes no type element for a rich text control, and keeps its prompt in the glossary.
    name: 'an untyped control whose prompt a Word glossary block holds',
    before: line(pr(`${ref('PartyName')}<w:showingPlcHdr/>`), run(STYLE, 'Enter a name.')),
    edit: typeOver,
    typed: line(pr(ref('PartyName')), typed('', 'Ana')),
    back: line(pr(`${ref('PartyName')}<w:showingPlcHdr/>`), run(STYLE, 'Enter a name.')),
  },
  {
    name: 'a list prompt',
    before: line(
      pr(`<w:showingPlcHdr/><w:dropDownList>${LIST}</w:dropDownList>`),
      run(GREY, 'Pick a term')
    ),
    edit: (part) => ({
      op: 'setContentControlValue',
      controlId: contentControlsIn(part.root)[0]!.node.id,
      value: '2',
      revision,
    }),
    typed: line(pr(`<w:dropDownList w:lastValue="2">${LIST}</w:dropDownList>`), typed(GREY, 'Two')),
    back: line(
      pr(`<w:showingPlcHdr/><w:dropDownList>${LIST}</w:dropDownList>`),
      run(STYLE, 'Choose an item.')
    ),
  },
];

function parse(body: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

const body = (part: OoxmlPart) =>
  /<w:body>(.*)<\/w:body>/s.exec(serializeOoxmlPart(part))?.[1] ?? '';

/** One op on one part, with the glossary's words lent: the server's call. */
function apply(part: OoxmlPart, op: TreeDocOp, options: EditOptions): OoxmlPart {
  const result = applyTreeOp(part, op, options);
  if (!result.ok) throw new Error(`refused: ${result.reason}`);
  return result.part;
}

/** The author's own suggestions taken back through the decision planner, op by op. */
function takeBack(part: OoxmlPart, options: EditOptions): OoxmlPart {
  const mine = collectReviewItems({ storyPart: part }).filter(
    (item) => item.kind === 'revision' && item.author === revision.author
  );
  const plan = planRevisionBatch(part, 'reject', mine.map(reviewItemKey));
  expect(plan.result.skipped).toEqual([]);
  return plan.ops.reduce((current, op) => apply(current, op, options), part);
}

describe('a caller that lends only the glossary text', () => {
  for (const sample of CASES) {
    test(`gets the same result as before for ${sample.name}`, () => {
      const part = parse(sample.before);
      const edited = apply(part, sample.edit(part), { placeholderText });
      expect(body(edited)).toBe(sample.typed);
      expect(body(takeBack(edited, { placeholderText }))).toBe(sample.back);
    });
  }
});

describe('a caller that lends both options', () => {
  const sample = CASES[1]!;
  const block = parse(`<w:p>${run(`<w:b/>${STYLE}`, 'Name the party')}</w:p>`);
  const paragraph = (block.root.children[0] as { children: readonly OoxmlNode[] }).children[0]!;

  test('gets the block when it names one', () => {
    const part = parse(sample.before);
    const edited = apply(part, sample.edit(part), { placeholderText });
    const placeholderBlocks = (name: string) => (name === 'PartyName' ? [paragraph] : undefined);
    const back = takeBack(edited, { placeholderText, placeholderBlocks });
    expect(body(back)).toContain(run(`<w:b/>${STYLE}`, 'Name the party'));
  });

  test('gets the text when no block holds the name', () => {
    const part = parse(sample.before);
    const edited = apply(part, sample.edit(part), { placeholderText });
    const back = takeBack(edited, { placeholderText, placeholderBlocks: () => undefined });
    expect(body(back)).toBe(sample.back);
  });
});
