// Accepting a suggestion typed over a content control's prompt.
//
// Earlier versions typed over a prompt in suggesting mode as a replacement, and files they saved
// still hold it: the prompt struck in a `w:del`, the new words in a `w:ins`, and the control
// keeping `w:showingPlcHdr` while the decision is open, so Reject brings the prompt back as it
// was. These tests build that shape directly. Accepting the new words leaves the control
// holding real content (§17.5.2.38), so the flag goes and the words lose Word's
// `PlaceholderText` style. They keep the prompt's other run properties, as a value written into
// the control does.

import { describe, expect, test } from 'bun:test';
import {
  contentControlTextOf,
  contentControlsIn,
  readOoxmlPart,
  serializeOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
} from '../index.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';
import { isShowingPlaceholder } from '../store/tree-op-nodes.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const DATE = '2026-10-28T10:00:00Z';
const PROMPT = 'Click here to enter text.';
const STYLED = '<w:rPr><w:rStyle w:val="PlaceholderText"/><w:b/></w:rPr>';
const STRUCK = `<w:del w:id="1" w:author="Ana" w:date="${DATE}"><w:r>${STYLED}<w:delText>${PROMPT}</w:delText></w:r></w:del>`;
const TYPED = `<w:ins w:id="2" w:author="Ana" w:date="${DATE}"><w:r>${STYLED}<w:t>Bob</w:t></w:r></w:ins>`;

function parse(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
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

function control(part: OoxmlPart, index = 0): OoxmlNode {
  const found = contentControlsIn(part.root)[index];
  if (!found) throw new Error('no control');
  return found.node;
}

/** The serialized `w:sdt`, the way a reader of the saved file sees it. */
function controlXml(part: OoxmlPart): string {
  const found = /<w:sdt>.*<\/w:sdt>/s.exec(serializeOoxmlPart(part))?.[0];
  if (found === undefined) throw new Error('no control XML');
  return found;
}

/** The serialized innermost `w:sdt`: the last one to open. */
function innermostControlXml(part: OoxmlPart): string {
  const xml = serializeOoxmlPart(part);
  const start = xml.lastIndexOf('<w:sdt>');
  return xml.slice(start, xml.indexOf('</w:sdt>', start) + '</w:sdt>'.length);
}

const inline = (properties: string, content: string) =>
  `<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="7"/>` +
  `<w:showingPlcHdr/>${properties}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt></w:p>`;
const block = (properties: string, content: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="7"/><w:showingPlcHdr/>${properties}</w:sdtPr>` +
  `<w:sdtContent><w:p>${content}</w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>after</w:t></w:r></w:p>`;

const ANA_DELETION = { id: '1', author: 'Ana', date: DATE };
const ANA_INSERTION = { id: '2', author: 'Ana', date: DATE };

describe('accepting the words typed over a prompt', () => {
  const shapes = [
    ['an inline plain text control', inline('<w:text/>', STRUCK + TYPED)],
    ['an inline rich text control', inline('<w:richText/>', STRUCK + TYPED)],
    ['an inline untyped control', inline('', STRUCK + TYPED)],
    ['a block rich text control', block('<w:richText/>', STRUCK + TYPED)],
    ['a block untyped control', block('', STRUCK + TYPED)],
    [
      'a drop-down list',
      inline(
        '<w:dropDownList><w:listItem w:displayText="One" w:value="1"/></w:dropDownList>',
        STRUCK + TYPED
      ),
    ],
    ['a date control', inline('<w:date><w:dateFormat w:val="M/d/yyyy"/></w:date>', STRUCK + TYPED)],
  ] as const;

  for (const [name, body] of shapes) {
    test(`Accept All leaves ${name} holding the words, not a prompt`, () => {
      const accepted = apply(parse(body), { op: 'acceptAllRevisions' });
      const xml = controlXml(accepted);
      expect(contentControlTextOf(control(accepted))).toBe('Bob');
      expect(isShowingPlaceholder(control(accepted))).toBe(false);
      expect(xml).not.toContain('showingPlcHdr');
      expect(xml).not.toContain('PlaceholderText');
      // The prompt's other run properties stay with the words.
      expect(xml).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t>Bob</w:t></w:r>');
    });

    test(`Reject All gives ${name} its prompt back exactly`, () => {
      const original = parse(
        body.replace(STRUCK + TYPED, `<w:r>${STYLED}<w:t>${PROMPT}</w:t></w:r>`)
      );
      const rejected = apply(parse(body), { op: 'rejectAllRevisions' });
      expect(controlXml(rejected)).toBe(controlXml(original));
      expect(isShowingPlaceholder(control(rejected))).toBe(true);
    });
  }

  test('accepting only the typed words ends the prompt and leaves the strike pending', () => {
    const part = parse(inline('<w:text/>', STRUCK + TYPED));
    const accepted = apply(part, {
      op: 'acceptRevision',
      revision: ANA_INSERTION,
      localName: 'ins',
    });
    const xml = controlXml(accepted);
    expect(isShowingPlaceholder(control(accepted))).toBe(false);
    // The struck prompt keeps its own face; the accepted words do not take the prompt style.
    expect(xml).toContain(`<w:del w:author="Ana" w:date="${DATE}" w:id="1"><w:r>${STYLED}`);
    expect(xml).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t>Bob</w:t></w:r>');
  });

  test('accepting only the strike ends the prompt and leaves the words proposed', () => {
    const part = parse(inline('<w:text/>', STRUCK + TYPED));
    const accepted = apply(part, {
      op: 'acceptRevision',
      revision: ANA_DELETION,
      localName: 'del',
    });
    const xml = controlXml(accepted);
    expect(isShowingPlaceholder(control(accepted))).toBe(false);
    expect(xml).not.toContain(PROMPT);
    expect(xml).toContain(
      `<w:ins w:author="Ana" w:date="${DATE}" w:id="2"><w:r><w:rPr><w:b/></w:rPr><w:t>Bob</w:t></w:r></w:ins>`
    );
    // Accepting the words afterwards leaves exactly what Accept All leaves.
    const both = apply(accepted, { op: 'acceptRevision', revision: ANA_INSERTION });
    expect(controlXml(both)).toBe(
      controlXml(apply(parse(inline('<w:text/>', STRUCK + TYPED)), { op: 'acceptAllRevisions' }))
    );
  });

  test('a prompt that is itself a tracked insertion stays a prompt when accepted', () => {
    // Word records a control inserted with tracking on as an insertion of its prompt runs.
    const body = inline(
      '<w:text/>',
      `<w:ins w:id="3" w:author="Bo" w:date="${DATE}"><w:r>${STYLED}<w:t>${PROMPT}</w:t></w:r></w:ins>`
    );
    const accepted = apply(parse(body), { op: 'acceptAllRevisions' });
    expect(isShowingPlaceholder(control(accepted))).toBe(true);
    expect(controlXml(accepted)).toContain(`<w:r>${STYLED}<w:t>${PROMPT}</w:t></w:r>`);
  });

  test('a control inside a tracked insertion keeps its prompt when the insertion is accepted', () => {
    const body =
      `<w:p><w:ins w:id="3" w:author="Bo" w:date="${DATE}"><w:sdt><w:sdtPr><w:showingPlcHdr/><w:text/></w:sdtPr>` +
      `<w:sdtContent><w:r>${STYLED}<w:t>${PROMPT}</w:t></w:r></w:sdtContent></w:sdt></w:ins></w:p>`;
    const accepted = apply(parse(body), { op: 'acceptAllRevisions' });
    expect(isShowingPlaceholder(control(accepted))).toBe(true);
    expect(controlXml(accepted)).toContain(`<w:r>${STYLED}<w:t>${PROMPT}</w:t></w:r>`);
  });

  test('a decision elsewhere leaves a pending replacement of a prompt alone', () => {
    const body =
      inline('<w:text/>', STRUCK + TYPED) +
      `<w:p><w:ins w:id="9" w:author="Cy" w:date="${DATE}"><w:r><w:t>elsewhere</w:t></w:r></w:ins></w:p>`;
    const part = parse(body);
    const accepted = apply(part, {
      op: 'acceptRevision',
      revision: { id: '9', author: 'Cy', date: DATE },
    });
    expect(controlXml(accepted)).toBe(controlXml(part));
    expect(isShowingPlaceholder(control(accepted))).toBe(true);
  });

  test("a nested control's own prompt is not the outer control's to end", () => {
    const nested =
      `<w:sdt><w:sdtPr><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent>` +
      `<w:r>${STYLED}<w:t>${PROMPT}</w:t></w:r></w:sdtContent></w:sdt>`;
    const accepted = apply(parse(block('<w:richText/>', STRUCK + TYPED + nested)), {
      op: 'acceptAllRevisions',
    });
    expect(isShowingPlaceholder(control(accepted, 0))).toBe(false);
    expect(isShowingPlaceholder(control(accepted, 1))).toBe(true);
    expect(innermostControlXml(accepted)).toContain(`<w:r>${STYLED}<w:t>${PROMPT}</w:t></w:r>`);
  });
});
