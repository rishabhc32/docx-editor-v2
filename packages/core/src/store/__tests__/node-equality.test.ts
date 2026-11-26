// Two property sets are the same when they say the same thing. A save writes attributes in its own
// order, may give a namespace another prefix, and writes a text's `xml:space="preserve"` only where
// the text needs it, so a file read back from a save compares equal to the one it came from.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlNode } from '../index.ts';
import { equivalentNodes } from '../store/ooxml-node-equality.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** The children of the body of a document that holds `xml`, under the given namespace bindings. */
function nodes(xml: string, bindings = `xmlns:w="${W}"`): readonly OoxmlNode[] {
  const result = readOoxmlPart(`<w:document ${bindings}><w:body>${xml}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  });
  if (!result.ok) {
    throw new Error(result.reason);
  }

  const body = result.part.root.children[0];
  if (body === undefined || body.kind === 'textValue') {
    throw new Error('no body');
  }

  return body.children;
}

const cell = (width: string) => `<w:tc><w:tcPr><w:tcW ${width}/></w:tcPr><w:p/></w:tc>`;
const table = (width: string) => `<w:tbl><w:tblPr/><w:tblGrid/><w:tr>${cell(width)}</w:tr></w:tbl>`;
const run = (text: string) => `<w:p><w:r>${text}</w:r></w:p>`;

describe('equivalent nodes', () => {
  test('are equal with their attributes in another order', () => {
    expect(
      equivalentNodes(
        nodes(table('w:w="4000" w:type="dxa"')),
        nodes(table('w:type="dxa" w:w="4000"'))
      )
    ).toBe(true);
  });

  test('are equal with another prefix for the same namespace', () => {
    expect(
      equivalentNodes(
        nodes(table('w:w="4000" w:type="dxa"')),
        nodes(table('ns1:w="4000" ns1:type="dxa"').replace('<w:tc>', `<w:tc xmlns:ns1="${W}">`))
      )
    ).toBe(true);
  });

  test('are equal with and without a space attribute the text does not need', () => {
    expect(
      equivalentNodes(
        nodes(run('<w:t>Fee</w:t>')),
        nodes(run('<w:t xml:space="preserve">Fee</w:t>'))
      )
    ).toBe(true);
  });

  test('differ by their text', () => {
    expect(equivalentNodes(nodes(run('<w:t>Fee</w:t>')), nodes(run('<w:t>Fees</w:t>')))).toBe(
      false
    );
  });

  test('differ by a value, or by an attribute one side lacks', () => {
    const shown = nodes(table('w:w="4000" w:type="dxa"'));
    expect(equivalentNodes(shown, nodes(table('w:type="dxa" w:w="2000"')))).toBe(false);
    expect(equivalentNodes(shown, nodes(table('w:w="4000"')))).toBe(false);
  });
});
