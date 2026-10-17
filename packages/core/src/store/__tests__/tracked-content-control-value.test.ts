// A content control's value proposed in suggesting mode, at the store.
//
// Word records a pick, a toggle or a date as content inside the control: the old display text in
// a `w:del`, the new one in a `w:ins`. The control's own record of the value (`w:lastValue`,
// `w:fullDate`, `w14:checked`) has no tracked form and moves with the value, and a checkbox's
// `w14:checked` follows its glyph again when a decision changes which glyph is live. A prompt
// goes without a revision, as Word removes it, and a decision that empties a control shows the
// prompt again.

import { describe, expect, test } from 'bun:test';
import {
  contentControlsIn,
  readOoxmlPart,
  serializeOoxmlPart,
  type OoxmlElement,
  type OoxmlPart,
} from '../index.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const revision = { author: 'Bea', date: '2026-10-16T00:00:00Z' };

const DROPDOWN =
  '<w:dropDownList w:lastValue="1"><w:listItem w:displayText="One" w:value="1"/>' +
  '<w:listItem w:displayText="Two" w:value="2"/><w:listItem w:displayText="Three" w:value="3"/>' +
  '</w:dropDownList>';
const CHECKBOX =
  '<w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/>' +
  '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>';
const UNCHECKED = '<w:r><w:sym w:font="MS Gothic" w:char="2610"/></w:r>';

const sdt = (properties: string, content: string) =>
  `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const inline = (properties: string, content: string) => `<w:p>${sdt(properties, content)}</w:p>`;

function parse(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}</w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function apply(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`refused: ${result.reason}`);
  return result.part;
}

/** The body's XML with the attributes the assertions do not care about removed. */
function body(part: OoxmlPart): string {
  return serializeOoxmlPart(part)
    .replace(/^.*<w:body>/s, '')
    .replace(/<\/w:body>.*$/s, '')
    .replace(/ w14:(paraId|textId)="[^"]*"/g, '')
    .replace(/ w:date="[^"]*"/g, '')
    .replace(/ w:id="[^"]*"/g, '');
}

/** What the first control's `w:sdtContent` holds. */
function content(part: OoxmlPart): string {
  return /<w:sdtContent>(.*?)<\/w:sdtContent>/s.exec(body(part))?.[1] ?? '';
}

function controlId(part: OoxmlPart, index = 0): string {
  return contentControlsIn(part.root)[index]!.node.id;
}

function propose(part: OoxmlPart, value: string, index = 0): OoxmlPart {
  return apply(part, {
    op: 'setContentControlValue',
    controlId: controlId(part, index),
    value,
    revision,
  });
}

describe('a proposed value', () => {
  test('strikes the old item and inserts the new one inside the control', () => {
    const part = propose(parse(inline(DROPDOWN, '<w:r><w:t>One</w:t></w:r>')), '2');
    expect(body(part)).toContain(
      '<w:sdtContent><w:del w:author="Bea"><w:r><w:delText>One</w:delText></w:r></w:del>' +
        '<w:ins w:author="Bea"><w:r><w:t>Two</w:t></w:r></w:ins></w:sdtContent>'
    );
    // `w:lastValue` has no tracked form, so it moves with the value as Word moves it.
    expect(body(part)).toContain('<w:dropDownList w:lastValue="2">');

    // A decision reads `w:lastValue` back from the item the control shows.
    const rejected = body(apply(part, { op: 'rejectAllRevisions' }));
    expect(rejected).toContain('<w:sdtContent><w:r><w:t>One</w:t></w:r></w:sdtContent>');
    expect(rejected).toContain('<w:dropDownList w:lastValue="1">');
    const accepted = body(apply(part, { op: 'acceptAllRevisions' }));
    expect(accepted).toContain('<w:sdtContent><w:r><w:t>Two</w:t></w:r></w:sdtContent>');
    expect(accepted).toContain('<w:dropDownList w:lastValue="2">');
  });

  test('of a date: a decision reads w:fullDate back from the date the text shows, as Word does', () => {
    const date = (format: string, fullDate: string) =>
      `<w:date w:fullDate="${fullDate}"><w:dateFormat w:val="${format}"/><w:lid w:val="en-US"/></w:date>`;
    for (const [format, shown] of [
      ['M/d/yyyy', '1/2/2026'],
      ['dd-MMM-yy', '02-Jan-26'],
    ] as const) {
      const part = propose(
        parse(inline(date(format, '2026-01-02T00:00:00Z'), `<w:r><w:t>${shown}</w:t></w:r>`)),
        '2026-03-04'
      );
      expect(body(part)).toContain('w:fullDate="2026-03-04T00:00:00Z"');
      expect(body(apply(part, { op: 'rejectAllRevisions' }))).toContain(
        'w:fullDate="2026-01-02T00:00:00Z"'
      );
      expect(body(apply(part, { op: 'acceptAllRevisions' }))).toContain(
        'w:fullDate="2026-03-04T00:00:00Z"'
      );
    }
    // Text the format does not write names no date: the record stays as it is.
    const typed = parse(
      inline(date('M/d/yyyy', '2026-01-02T00:00:00Z'), '<w:r><w:t>soon</w:t></w:r>')
    );
    const struck = apply(typed, {
      op: 'deleteText',
      paragraphId: paragraphId(typed),
      start: 0,
      end: 2,
      revision,
    });
    expect(body(apply(struck, { op: 'acceptAllRevisions' }))).toContain(
      'w:fullDate="2026-01-02T00:00:00Z"'
    );
  });

  test('of a combo box: rejecting free text reads back the item the control shows again', () => {
    const combo =
      '<w:comboBox w:lastValue="1"><w:listItem w:displayText="One" w:value="1"/>' +
      '<w:listItem w:displayText="Two" w:value="2"/></w:comboBox>';
    const part = propose(parse(inline(combo, '<w:r><w:t>One</w:t></w:r>')), 'Other');
    expect(body(part)).toContain('<w:comboBox w:lastValue="Other">');
    expect(body(apply(part, { op: 'rejectAllRevisions' }))).toContain(
      '<w:comboBox w:lastValue="1">'
    );
    expect(body(apply(part, { op: 'acceptAllRevisions' }))).toContain(
      '<w:comboBox w:lastValue="Other">'
    );
  });

  test('a second pick replaces the author’s own pending pick instead of stacking on it', () => {
    let part = parse(inline(DROPDOWN, '<w:r><w:t>One</w:t></w:r>'));
    part = propose(propose(part, '2'), '3');
    expect(body(part)).toContain(
      '<w:sdtContent><w:del w:author="Bea"><w:r><w:delText>One</w:delText></w:r></w:del>' +
        '<w:ins w:author="Bea"><w:r><w:t>Three</w:t></w:r></w:ins></w:sdtContent>'
    );
    expect(body(apply(part, { op: 'rejectAllRevisions' }))).toContain(
      '<w:sdtContent><w:r><w:t>One</w:t></w:r></w:sdtContent>'
    );
  });

  test('a pick over another author’s pending pick strikes their insertion, not their deletion', () => {
    const pending =
      '<w:del w:id="5" w:author="Al"><w:r><w:delText>One</w:delText></w:r></w:del>' +
      '<w:ins w:id="6" w:author="Al"><w:r><w:t>Two</w:t></w:r></w:ins>';
    const part = propose(parse(inline(DROPDOWN, pending)), '3');
    expect(body(part)).toContain(
      '<w:del w:author="Al"><w:r><w:delText>One</w:delText></w:r></w:del>' +
        '<w:ins w:author="Al"><w:del w:author="Bea"><w:r><w:delText>Two</w:delText></w:r></w:del></w:ins>' +
        '<w:ins w:author="Bea"><w:r><w:t>Three</w:t></w:r></w:ins>'
    );
  });

  test('picking the item a pending pick already shows proposes nothing more', () => {
    const pending = propose(parse(inline(DROPDOWN, '<w:r><w:t>One</w:t></w:r>')), '2');
    expect(body(pending)).toContain('<w:ins w:author="Bea">');
    expect(body(propose(pending, '2'))).toBe(body(pending));
  });

  test('a control in the author’s own pending insertion takes the value directly', () => {
    const text = sdt('<w:text/>', '<w:r><w:t>Acme</w:t></w:r>');
    const xml = `<w:p><w:ins w:id="9" w:author="Bea">${text}</w:ins>${text}</w:p>`;
    const part = propose(propose(parse(xml), 'Beta', 0), 'Gamma', 1);
    // The first is inside the proposal already; the second is somebody's text to strike.
    expect(body(part)).toContain(
      '<w:ins w:author="Bea"><w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent>' +
        '<w:r><w:t>Beta</w:t></w:r></w:sdtContent></w:sdt></w:ins>'
    );
    expect(body(part)).toContain(
      '<w:sdtContent><w:del w:author="Bea"><w:r><w:delText>Acme</w:delText></w:r></w:del>' +
        '<w:ins w:author="Bea"><w:r><w:t>Gamma</w:t></w:r></w:ins></w:sdtContent>'
    );
  });

  test('a block-level text control proposes inside its one paragraph', () => {
    const part = propose(parse(sdt('<w:text/>', '<w:p><w:r><w:t>Acme</w:t></w:r></w:p>')), 'Beta');
    expect(body(part)).toContain(
      '<w:p><w:del w:author="Bea"><w:r><w:delText>Acme</w:delText></w:r></w:del>' +
        '<w:ins w:author="Bea"><w:r><w:t>Beta</w:t></w:r></w:ins></w:p>'
    );
  });

  test('over a prompt is proposed alone, and rejecting it shows the prompt again', () => {
    // Word removes a prompt without a revision: it is the control's state, not its text.
    const prompt =
      '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Choose an item.</w:t></w:r>';
    const part = propose(parse(inline(`<w:showingPlcHdr/>${DROPDOWN}`, prompt)), '2');
    expect(body(part)).not.toContain('showingPlcHdr');
    expect(content(part)).toBe('<w:ins w:author="Bea"><w:r><w:t>Two</w:t></w:r></w:ins>');

    // Showing its prompt again, the control holds no value.
    const rejected = apply(part, { op: 'rejectAllRevisions' });
    expect(body(rejected)).toContain('<w:showingPlcHdr/>');
    expect(body(rejected)).not.toContain('lastValue');
    expect(content(rejected)).toBe(prompt);
    expect(content(apply(part, { op: 'acceptAllRevisions' }))).toBe('<w:r><w:t>Two</w:t></w:r>');
  });

  for (const [shape, xml] of [
    [
      'a temporary control, which the value would remove',
      inline('<w:temporary/><w:text/>', '<w:r><w:t>Acme</w:t></w:r>'),
    ],
    [
      'a block-level control over two paragraphs',
      sdt(
        '<w:text w:multiLine="1"/>',
        '<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>'
      ),
    ],
  ] as const) {
    test(`is refused for ${shape}`, () => {
      const part = parse(xml);
      const op: TreeDocOp = {
        op: 'setContentControlValue',
        controlId: controlId(part),
        value: 'C',
        revision,
      };
      expect(applyTreeOp(part, op)).toMatchObject({ ok: false, reason: 'unsupported' });
    });
  }

  test('is refused with no author to attribute it to', () => {
    const part = parse(inline(DROPDOWN, '<w:r><w:t>One</w:t></w:r>'));
    const op: TreeDocOp = {
      op: 'setContentControlValue',
      controlId: controlId(part),
      value: '2',
      revision: { author: ' ' },
    };
    expect(applyTreeOp(part, op)).toMatchObject({ ok: false, reason: 'invalid-property-value' });
  });
});

describe('a proposed checkbox toggle', () => {
  test('swaps the glyph as a deletion and an insertion, and sets w14:checked', () => {
    const part = propose(parse(inline(CHECKBOX, UNCHECKED)), 'true');
    expect(body(part)).toContain('<w14:checked w14:val="1"/>');
    expect(body(part)).toMatch(
      /<w:del w:author="Bea"><w:r><w:sym w:char="2610" w:font="MS Gothic"\/><\/w:r><\/w:del><w:ins w:author="Bea"><w:r>(<w:rPr>.*<\/w:rPr>)?<w:sym w:char="2612" w:font="MS Gothic"\/><\/w:r><\/w:ins>/
    );
  });

  test('rejecting it puts the glyph and w14:checked back together', () => {
    const part = apply(propose(parse(inline(CHECKBOX, UNCHECKED)), 'true'), {
      op: 'rejectAllRevisions',
    });
    expect(body(part)).toContain('<w14:checked w14:val="0"/>');
    expect(content(part)).toBe('<w:r><w:sym w:char="2610" w:font="MS Gothic"/></w:r>');
  });

  test('accepting it keeps both at the new state', () => {
    const part = apply(propose(parse(inline(CHECKBOX, UNCHECKED)), 'true'), {
      op: 'acceptAllRevisions',
    });
    expect(body(part)).toContain('<w14:checked w14:val="1"/>');
    expect(content(part)).toContain('<w:sym w:char="2612" w:font="MS Gothic"/>');
    expect(content(part)).not.toContain('<w:del ');
  });

  test('a decision that leaves a glyph alone leaves its w14:checked as the file wrote it', () => {
    // The second checkbox is one the file left inconsistent, and no decision reaches it.
    const xml =
      inline(CHECKBOX, UNCHECKED) +
      inline(CHECKBOX.replace('w14:val="0"', 'w14:val="1"'), UNCHECKED);
    const part = apply(propose(parse(xml), 'true'), { op: 'rejectAllRevisions' });
    expect(body(part).match(/<w14:checked w14:val="\d"\/>/g)).toEqual([
      '<w14:checked w14:val="0"/>',
      '<w14:checked w14:val="1"/>',
    ]);
  });
});

/** "Name: " and a text control showing its bold prompt, then " end.". */
const PROMPT_RUN =
  '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/><w:b/></w:rPr><w:t>Enter a name.</w:t></w:r>';
const PROMPTED =
  '<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r>' +
  sdt(
    '<w:placeholder><w:docPart w:val="Name"/></w:placeholder><w:showingPlcHdr/><w:text/>',
    PROMPT_RUN
  ) +
  '<w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>';
const PROMPT_START = 'Name: '.length;
const PROMPT_END = 'Name: Enter a name.'.length;

/** The id of the body's first paragraph. */
function paragraphId(part: OoxmlPart): string {
  return ((part.root as OoxmlElement).children[0] as OoxmlElement).children[0]!.id;
}

function typeAt(part: OoxmlPart, offset: number, text: string, inside?: string): OoxmlPart {
  return apply(part, {
    op: 'insertText',
    paragraphId: paragraphId(part),
    offset,
    text,
    revision,
    ...(inside === undefined ? {} : { inside }),
  });
}

describe('a prompt in suggesting mode', () => {
  test('typing over it proposes only the words, in the prompt’s format without its style', () => {
    for (const offset of [PROMPT_START, PROMPT_END]) {
      const part = typeAt(parse(PROMPTED), offset, 'Ana');
      expect(body(part)).not.toContain('showingPlcHdr');
      expect(body(part)).not.toContain('Enter a name.');
      expect(content(part)).toBe(
        '<w:ins w:author="Bea"><w:r><w:rPr><w:b/></w:rPr><w:t>Ana</w:t></w:r></w:ins>'
      );
    }
  });

  test('the next keystroke at the end of the words extends them inside the control', () => {
    const typed = typeAt(parse(PROMPTED), PROMPT_START, 'A');
    const part = typeAt(typed, PROMPT_START + 1, 'n', controlId(typed));
    expect(content(part)).toBe(
      '<w:ins w:author="Bea"><w:r><w:rPr><w:b/></w:rPr><w:t>An</w:t></w:r></w:ins>'
    );
  });

  test('a deletion inside it proposes nothing, as Word keeps showing it', () => {
    const part = parse(PROMPTED);
    const deleted = apply(part, {
      op: 'deleteText',
      paragraphId: paragraphId(part),
      start: PROMPT_START,
      end: PROMPT_END,
      revision,
    });
    expect(body(deleted)).toBe(body(part));
  });

  test('rejecting the words shows the prompt again, and accepting keeps them as the value', () => {
    const part = typeAt(parse(PROMPTED), PROMPT_START, 'Ana');
    const rejected = apply(part, { op: 'rejectAllRevisions' });
    expect(body(rejected)).toContain('<w:showingPlcHdr/>');
    // The store reads no glossary here, so the type's own prompt stands in.
    expect(content(rejected)).toBe(
      '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Click here to enter text.</w:t></w:r>'
    );
    const accepted = apply(part, { op: 'acceptAllRevisions' });
    expect(body(accepted)).not.toContain('showingPlcHdr');
    expect(content(accepted)).toBe('<w:r><w:rPr><w:b/></w:rPr><w:t>Ana</w:t></w:r>');
  });
});

describe('a decision that empties a control', () => {
  const filled = () =>
    parse(
      '<w:p>' +
        sdt(
          '<w:placeholder><w:docPart w:val="Name"/></w:placeholder><w:text/>',
          '<w:r><w:rPr><w:b/></w:rPr><w:t>Acme</w:t></w:r>'
        ) +
        '</w:p>'
    );
  const struck = (part: OoxmlPart) =>
    apply(part, { op: 'deleteText', paragraphId: paragraphId(part), start: 0, end: 4, revision });

  test('shows its prompt again, in the glossary’s words when the transaction lends them', () => {
    const part = struck(filled());
    const accepted = applyTreeOp(
      part,
      { op: 'acceptAllRevisions' },
      { placeholderText: (name) => (name === 'Name' ? 'Enter a name.' : undefined) }
    );
    if (!accepted.ok) throw new Error(accepted.reason);
    expect(body(accepted.part)).toContain('<w:showingPlcHdr/>');
    expect(content(accepted.part)).toBe(
      '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Enter a name.</w:t></w:r>'
    );
  });

  test('leaves a control that keeps content alone', () => {
    const part = struck(filled());
    const rejected = apply(part, { op: 'rejectAllRevisions' });
    expect(body(rejected)).not.toContain('showingPlcHdr');
    expect(content(rejected)).toBe('<w:r><w:rPr><w:b/></w:rPr><w:t>Acme</w:t></w:r>');
  });
});
