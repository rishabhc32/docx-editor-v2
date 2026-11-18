// Notes' properties in the ordered model (group 9.9): the organiser's direct edit, for the
// document's settings or for one section. The list writes what the editor alone writes (the
// engine's own writer, in the schema's order), two copies agree, the undo (`revert`) gives the
// file back exactly, a save opens again, and Enter in the paragraph that ends the section keeps
// the write on the section.

import { describe, expect, test } from 'bun:test';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedIdMinter } from '../change/ids.ts';
import type { OrderedChange } from '../change/types.ts';
import { validateOrderedChange } from '../change/validate.ts';
import { sectionHolders } from '../client/ops/note-references.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { composeOrderedPackage, orderedPackageBytes } from '../compose/package.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { at, BODY, DATE, markOf } from './change-apply-fixture.ts';
import { docx, p, R, t, W, WML_CT } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

const SETTINGS = '/word/settings.xml';

/** Two sections: the first ends at a paragraph's section break, the last has a header. */
const FILE = docx(
  p(t('First section.'), '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>') +
    p(t('Second section.')) +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/>' +
    '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
  {
    parts: {
      'word/header1.xml': `<w:hdr xmlns:w="${W}">${p(t('Letterhead'))}</w:hdr>`,
      'word/settings.xml':
        `<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="720"/>` +
        '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' +
        '<w:compat/><w:rsids><w:rsidRoot w:val="00A1B2C3"/></w:rsids></w:settings>',
    },
    contentTypes:
      `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>` +
      `<Override PartName="/word/settings.xml" ContentType="${WML_CT}.settings+xml"/>`,
    relationships:
      `<Relationship Id="rIdHeader" Type="${R}/header" Target="header1.xml"/>` +
      `<Relationship Id="rIdSettings" Type="${R}/settings" Target="settings.xml"/>`,
  }
);

type Body = OrderedChange extends infer C
  ? C extends OrderedChange
    ? Omit<C, 'v' | 'id'>
    : never
  : never;

function open(bytes = FILE): { doc: OrderedDocument; base: OoxmlPackage } {
  const opened = openOrderedWordFile(bytes, { file: 1 });
  return { doc: opened.document, base: opened.base };
}

function apply(doc: OrderedDocument, id: string, body: Body): string[] {
  const change = { v: 1, id, ...body } as OrderedChange;
  return [
    ...withUndoRecord(doc, () =>
      applyOrderedChange(doc, change, {
        person: testPerson('Olive'),
        date: DATE,
        ids: orderedIdMinter(id),
      })
    ).value.suggestions,
  ];
}

/** The section at an index, as the change names it. */
function section(doc: OrderedDocument, index: number) {
  const holder = sectionHolders(doc)?.holders[index];
  if (!holder) throw new Error(`no section ${index}`);
  const address = { space: holder.space, offset: holder.offset };
  return holder.kind === 'zero' ? { body: address } : { mark: address };
}

const write = (doc: OrderedDocument, id: string, body: Partial<Body>, mode = 'edit') =>
  apply(doc, id, { kind: 'note-properties', part: BODY, mode, ...body } as Body)[0] ?? '';

const xmlOf = (pkg: OoxmlPackage, name: string): string => {
  const part = pkg.parts.get(name);
  if (!part) throw new Error(`no part ${name}`);
  return serializeOoxmlPart(part);
};

function shape(pkg: OoxmlPackage): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, part] of pkg.parts) out[name] = serializeOoxmlPart(part);
  return out;
}

describe("notes' properties (group 9.9)", () => {
  test("the document's: in the settings before w:compat, the same on two copies, and undone", () => {
    const one = open();
    const two = open();
    const before = shape(composeOrderedPackage(one.doc, one.base));
    const body = {
      footnote: { numFmt: 'lowerRoman', numStart: 3 },
      endnote: { position: 'sectEnd' },
    };
    const sid = write(one.doc, 'olive-1:1', body);
    write(two.doc, 'olive-1:1', body);
    const pkg = composeOrderedPackage(one.doc, one.base);
    expect(shape(pkg)).toEqual(shape(composeOrderedPackage(two.doc, two.base)));
    expect(xmlOf(pkg, SETTINGS)).toContain(
      '<w:footnotePr><w:numFmt w:val="lowerRoman"/><w:numStart w:val="3"/>' +
        '<w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' +
        '<w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr><w:compat/>'
    );
    // Saved and opened again
    const reopened = open(orderedPackageBytes(one.doc, one.base));
    expect(xmlOf(composeOrderedPackage(reopened.doc, reopened.base), SETTINGS)).toContain(
      '<w:numStart w:val="3"/>'
    );
    apply(one.doc, 'olive-1:2', { kind: 'revert', suggestions: [sid], status: 'rejected' } as Body);
    expect(shape(composeOrderedPackage(one.doc, one.base))).toEqual(before);
  });

  test("a section's: after the header reference, on the paragraph's or the body's section", () => {
    const { doc, base } = open();
    const before = shape(composeOrderedPackage(doc, base));
    const first = write(doc, 'olive-1:1', {
      section: section(doc, 0),
      footnote: { numRestart: 'eachSect' },
    });
    const last = write(doc, 'olive-1:2', {
      section: section(doc, 1),
      footnote: { numFmt: 'chicago', position: 'beneathText' },
    });
    const body = xmlOf(composeOrderedPackage(doc, base), BODY);
    expect(body).toContain(
      '<w:sectPr><w:footnotePr><w:numRestart w:val="eachSect"/></w:footnotePr><w:pgSz'
    );
    expect(body).toContain(
      '<w:headerReference r:id="rIdHeader" w:type="default"/><w:footnotePr>' +
        '<w:pos w:val="beneathText"/><w:numFmt w:val="chicago"/></w:footnotePr><w:pgSz'
    );
    // Undone last first, as the organiser's undo goes
    apply(doc, 'olive-1:3', { kind: 'revert', suggestions: [last], status: 'rejected' } as Body);
    apply(doc, 'olive-1:4', { kind: 'revert', suggestions: [first], status: 'rejected' } as Body);
    expect(shape(composeOrderedPackage(doc, base))).toEqual(before);
  });

  test('Enter in the paragraph that ends a section keeps the write on the section', () => {
    const { doc, base } = open();
    write(doc, 'olive-1:1', { section: section(doc, 0), footnote: { numStart: 7 } });
    apply(doc, 'olive-1:2', { kind: 'split', part: BODY, mode: 'edit', place: at(doc, 0, 5) });
    const body = xmlOf(composeOrderedPackage(doc, base), BODY);
    expect(body).toMatch(/<w:p [^>]*><w:pPr\/><w:r><w:t>First<\/w:t>/);
    expect(body).toContain(
      '<w:footnotePr><w:numStart w:val="7"/></w:footnotePr><w:pgSz w:h="15840" w:w="12240"/>' +
        '</w:sectPr></w:pPr><w:r><w:t xml:space="preserve"> section.</w:t>'
    );
  });

  test('refused: as a suggestion, a paragraph that ends no section, and values Word has not', () => {
    const { doc } = open();
    const reason = (body: () => unknown): string => {
      try {
        body();
      } catch (error) {
        if (error instanceof OrderedRefusal) return error.reason;
        throw error;
      }
      return 'applied';
    };
    expect(reason(() => write(doc, 'olive-1:1', { footnote: { numStart: 2 } }, 'suggest'))).toBe(
      'invalid-edit'
    );
    // The second paragraph ends no section: its section is the body's
    const plain = { mark: markOf(doc, 1) };
    expect(
      reason(() => write(doc, 'olive-1:2', { section: plain, footnote: { numStart: 2 } }))
    ).toBe('invalid-edit');
    const code = (body: Record<string, unknown>) => {
      const checked = validateOrderedChange(
        { v: 1, id: 'olive-1:3', kind: 'note-properties', part: BODY, mode: 'edit', ...body },
        { namespace: 'olive-1', parts: new Set([BODY]) }
      );
      return checked.ok ? 'ok' : checked.error.code;
    };
    expect(code({ footnote: { numFmt: 'lowerRoman' } })).toBe('ok');
    expect(code({ footnote: { numFmt: '<w:x/>' } })).toBe('invalid-value');
    expect(code({ endnote: { position: 'pageBottom' } })).toBe('invalid-value');
    expect(code({ footnote: { numStart: -1 } })).toBe('invalid-number');
    expect(code({})).toBe('missing-field');
    const both = { mark: markOf(doc, 0), body: markOf(doc, 0) };
    expect(code({ section: both, footnote: {} })).toBe('invalid-value');
  });
});
