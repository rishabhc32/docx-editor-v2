// A control that declares no type is rich text, and shows its prompt like one.
//
// ECMA-376 makes the type element of `w:sdtPr` optional, and Word writes no type element for the
// rich text controls it inserts. So an untyped control shows its prompt when a file leaves it
// empty and when an edit empties it, exactly as a `w:richText` control does. A repeating section
// and its items also declare no ECMA-376 type (their type is a `w15` extension). They have no
// prompt of their own and stay as they were, as does every other control type.

import { describe, expect, test } from 'bun:test';
import {
  bodyStoryRoot,
  contentControlTextOf,
  contentControlsIn,
  readOoxmlPackage,
  readOoxmlPart,
  serializeOoxmlPart,
  storyParagraphs,
  type OoxmlNode,
  type OoxmlPart,
} from '../index.ts';
import { materializeGlossaryPlaceholders } from '../store/placeholder-materialize.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';
import { isShowingPlaceholder } from '../store/tree-op-nodes.ts';
import { zipDoc } from './canonical-primitive-journal-coverage-support.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const PROMPT = 'Click or tap here to enter text.';
const GLOSSARY_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument';
const GLOSSARY_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml';
const GLOSSARY =
  `<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr>` +
  '<w:name w:val="Clause"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/>' +
  '</w:category></w:docPartPr><w:docPartBody><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/>' +
  '</w:rPr><w:t>Type the clause.</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>';

function parse(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:w15="${W15}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

/** The main part as the open shows it: empty prompt-bearing controls filled. */
function open(body: string, glossary = false): OoxmlPart {
  const loaded = readOoxmlPackage(
    zipDoc({
      body: `${body}<w:sectPr/>`,
      extraXmlns: `xmlns:w15="${W15}"`,
      ...(glossary
        ? {
            rels: `<Relationship Id="rId9" Type="${GLOSSARY_TYPE}" Target="glossary/document.xml"/>`,
            overrides: `<Override PartName="/word/glossary/document.xml" ContentType="${GLOSSARY_CONTENT_TYPE}"/>`,
            extraXml: { 'word/glossary/document.xml': GLOSSARY },
          }
        : {}),
    })
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const main = loaded.package.parts.get(loaded.package.mainDocumentPart)!;
  return materializeGlossaryPlaceholders(loaded.package, main);
}

function apply(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`refused: ${result.reason}`);
  return result.part;
}

function firstParagraphId(part: OoxmlPart): string {
  const body = bodyStoryRoot(part);
  const paragraph = body ? storyParagraphs(body)[0] : undefined;
  if (!paragraph) throw new Error('no paragraph');
  return paragraph.id;
}

function control(part: OoxmlPart, index = 0): OoxmlNode {
  const found = contentControlsIn(part.root)[index];
  if (!found) throw new Error('no control');
  return found.node;
}

/** Every serialized `w:sdt`, with the one type element these tests vary taken out. */
function controlsXml(part: OoxmlPart): string {
  const xml = serializeOoxmlPart(part).replace('<w:richText/>', '');
  return (xml.match(/<w:sdt>.*<\/w:sdt>/s) ?? [''])[0];
}

const inlineFilled = (type: string) =>
  `<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="7"/>${type}` +
  `</w:sdtPr><w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>Hi</w:t></w:r></w:sdtContent></w:sdt></w:p>`;
const blockFilled = (type: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="7"/>${type}</w:sdtPr><w:sdtContent><w:p><w:r><w:t>Hi</w:t></w:r></w:p>` +
  '</w:sdtContent></w:sdt><w:p><w:r><w:t>after</w:t></w:r></w:p>';

/** Delete the two characters of "Hi", the control's whole content. */
function emptied(body: string, start: number): OoxmlPart {
  const part = parse(body);
  return apply(part, {
    op: 'deleteText',
    paragraphId: firstParagraphId(part),
    start,
    end: start + 2,
  });
}

describe('an untyped control shows its prompt like a rich text control', () => {
  for (const [name, shape, start] of [
    ['inline', inlineFilled, 6],
    ['block', blockFilled, 0],
  ] as const) {
    test(`an edit that empties an ${name} untyped control shows its prompt again`, () => {
      const untyped = emptied(shape(''), start);
      expect(isShowingPlaceholder(control(untyped))).toBe(true);
      expect(contentControlTextOf(control(untyped))).toBe(PROMPT);
      expect(controlsXml(untyped)).toBe(controlsXml(emptied(shape('<w:richText/>'), start)));
    });

    test(`the next keystroke replaces the restored prompt of an ${name} untyped control`, () => {
      const part = emptied(shape(''), start);
      const typed = apply(part, {
        op: 'insertText',
        paragraphId: firstParagraphId(part),
        offset: start,
        text: 'Yo',
      });
      expect(contentControlTextOf(control(typed))).toBe('Yo');
      expect(isShowingPlaceholder(control(typed))).toBe(false);
    });
  }

  for (const [name, body] of [
    ['inline', '<w:p><w:sdt><w:sdtPr><w:id w:val="7"/>TYPE</w:sdtPr><w:sdtContent/></w:sdt></w:p>'],
    ['block', '<w:sdt><w:sdtPr><w:id w:val="7"/>TYPE</w:sdtPr><w:sdtContent/></w:sdt>'],
  ] as const) {
    test(`an empty ${name} untyped control opens showing the type's prompt`, () => {
      const untyped = open(body.replace('TYPE', ''));
      expect(isShowingPlaceholder(control(untyped))).toBe(true);
      expect(contentControlTextOf(control(untyped))).toBe(PROMPT);
      expect(controlsXml(untyped)).toBe(controlsXml(open(body.replace('TYPE', '<w:richText/>'))));
    });

    test(`an empty ${name} untyped control opens showing the glossary prompt it names`, () => {
      const named = body.replace(
        'TYPE',
        '<w:placeholder><w:docPart w:val="Clause"/></w:placeholder>'
      );
      const untyped = open(named, true);
      expect(isShowingPlaceholder(control(untyped))).toBe(true);
      expect(contentControlTextOf(control(untyped))).toBe('Type the clause.');
    });
  }
});

describe('the change leaves every other control as it was', () => {
  const ITEM =
    '<w:sdt><w:sdtPr><w:id w:val="1"/><w15:repeatingSection/></w:sdtPr><w:sdtContent>' +
    '<w:sdt><w:sdtPr><w:id w:val="2"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent>' +
    '<w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt>' +
    '<w:p><w:r><w:t>after</w:t></w:r></w:p>';

  test('emptying a repeating section item shows no prompt', () => {
    const part = emptied(ITEM, 0);
    expect(contentControlsIn(part.root).map((entry) => isShowingPlaceholder(entry.node))).toEqual([
      false,
      false,
    ]);
    expect(serializeOoxmlPart(part)).not.toContain(PROMPT);
  });

  test('an empty repeating section or item opens as it was', () => {
    for (const body of [
      '<w:sdt><w:sdtPr><w15:repeatingSection/></w:sdtPr><w:sdtContent/></w:sdt>',
      '<w:sdt><w:sdtPr><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent/></w:sdt>',
    ]) {
      const part = open(body);
      expect(isShowingPlaceholder(control(part))).toBe(false);
      expect(serializeOoxmlPart(part)).not.toContain(PROMPT);
    }
  });

  test('an empty control opens with a prompt exactly when its type has one', () => {
    const TYPES: readonly (readonly [string, boolean])[] = [
      ['<w:text/>', true],
      ['<w:richText/>', true],
      ['', true],
      ['<w:date/>', true],
      ['<w:dropDownList/>', true],
      ['<w:comboBox/>', true],
      ['<w:docPartList><w:docPartGallery w:val="Quick Parts"/></w:docPartList>', true],
      ['<w14:checkbox><w14:checked w14:val="0"/></w14:checkbox>', false],
      ['<w:picture/>', false],
      ['<w:group/>', false],
      ['<w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj>', false],
      ['<w:citation/>', false],
      ['<w:bibliography/>', false],
      ['<w:equation/>', false],
    ];
    for (const [type, prompts] of TYPES) {
      const part = open(`<w:p><w:sdt><w:sdtPr>${type}</w:sdtPr><w:sdtContent/></w:sdt></w:p>`);
      expect({ type, prompts: isShowingPlaceholder(control(part)) }).toEqual({ type, prompts });
    }
  });
});
