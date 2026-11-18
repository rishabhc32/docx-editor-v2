// `setNoteProperties` writes `w:footnotePr` and `w:endnotePr` where the schema puts them. Word
// reads `settings.xml` and `w:sectPr` in schema order and reports a file whose members are out
// of sequence as damaged: in `CT_Settings` the note properties come before `w:compat`, and in
// `CT_SectPr` they come right after the header and footer references. A write also keeps the
// authored fields it does not give, and the separator references of the settings element.
// A new notes part lists its two separator notes in the settings, as Word writes them, when the
// file has settings.

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import {
  applyNoteLifecycleOp,
  readOoxmlPackage,
  serializeOoxmlPart,
  type OoxmlPackage,
} from '../package/index.ts';
import { settingsPartOf } from '../package/note-properties.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

/** Word's own settings members around the note properties, as Word writes them. */
const SETTINGS =
  `<w:settings xmlns:w="${W}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">` +
  '<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/>' +
  '<w:characterSpacingControl w:val="doNotCompress"/>' +
  '<w:footnotePr><w:numFmt w:val="decimal"/><w:footnote w:id="-1"/><w:footnote w:id="0"/>' +
  '</w:footnotePr>' +
  '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
  '<w:rsids><w:rsidRoot w:val="00A1B2C3"/></w:rsids>' +
  '<m:mathPr><m:mathFont m:val="Cambria Math"/></m:mathPr>' +
  '<w:decimalSymbol w:val="."/><w:listSeparator w:val=","/></w:settings>';

/** A section that has a header: its note properties go after the header reference. */
const BODY =
  '<w:p><w:r><w:t>Text</w:t></w:r></w:p>' +
  '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>' +
  '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" ' +
  'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

function file(settings: string | null): OoxmlPackage {
  const override = (name: string, type: string) =>
    `<Override PartName="/word/${name}" ContentType="${WML}.${type}+xml"/>`;
  const entries: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        override('document.xml', 'document.main') +
        override('header1.xml', 'header') +
        (settings === null ? '' : override('settings.xml', 'settings')) +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" ` +
        'Target="word/document.xml"/></Relationships>'
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}">` +
        `<Relationship Id="rIdH" Type="${R}/header" Target="header1.xml"/>` +
        (settings === null
          ? ''
          : `<Relationship Id="rIdS" Type="${R}/settings" Target="settings.xml"/>`) +
        '</Relationships>'
    ),
    'word/header1.xml': strToU8(`<w:hdr xmlns:w="${W}"><w:p/></w:hdr>`),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${BODY}</w:body></w:document>`
    ),
  };
  if (settings !== null) entries['word/settings.xml'] = strToU8(settings);
  const read = readOoxmlPackage(zipSync(entries));
  if (!read.ok) throw new Error(read.reason);
  return read.package;
}

function applied(pkg: OoxmlPackage, op: Parameters<typeof applyNoteLifecycleOp>[1]) {
  const result = applyNoteLifecycleOp(pkg, op);
  if (!result.ok) throw new Error(`${result.reason}: ${result.detail ?? ''}`);
  return result.package;
}

/** The element names of a part's root children, or of a section's, with their prefixes. */
function names(xml: string, parent: string): string[] {
  const inner = new RegExp(`<${parent}[^>]*>([\\s\\S]*)</${parent}>`).exec(xml)?.[1] ?? '';
  const out: string[] = [];
  let depth = 0;
  for (const match of inner.matchAll(/<(\/?)([\w:]+)[^>]*?(\/?)>/g)) {
    if (match[1] === '/') depth -= 1;
    else {
      if (depth === 0) out.push(match[2] ?? '');
      if (match[3] !== '/') depth += 1;
    }
  }
  return out;
}

function settingsXml(pkg: OoxmlPackage): string {
  const settings = settingsPartOf(pkg);
  if (!settings) throw new Error('no settings');
  return serializeOoxmlPart(settings);
}

function sectionXml(pkg: OoxmlPackage): string {
  const main = pkg.parts.get(pkg.mainDocumentPart);
  if (!main) throw new Error('no main part');
  return /<w:sectPr>[\s\S]*<\/w:sectPr>/.exec(serializeOoxmlPart(main))?.[0] ?? '';
}

describe('setNoteProperties keeps the schema order', () => {
  test('document scope: the note properties stay before w:compat, and keep the separators', () => {
    const pkg = applied(file(SETTINGS), {
      op: 'setNoteProperties',
      scope: 'document',
      footnote: { numStart: 3 },
      endnote: { numFmt: 'upperRoman' },
    });
    const xml = settingsXml(pkg);
    expect(names(xml, 'w:settings')).toEqual([
      'w:zoom',
      'w:defaultTabStop',
      'w:characterSpacingControl',
      'w:footnotePr',
      'w:endnotePr',
      'w:compat',
      'w:rsids',
      'm:mathPr',
      'w:decimalSymbol',
      'w:listSeparator',
    ]);
    // The given field is added in its place; the authored format and separators stay
    expect(names(xml, 'w:footnotePr')).toEqual([
      'w:numFmt',
      'w:numStart',
      'w:footnote',
      'w:footnote',
    ]);
    expect(xml).toContain('<w:numStart w:val="3"/>');
    expect(xml).toContain('<w:endnotePr><w:numFmt w:val="upperRoman"/></w:endnotePr>');
  });

  test('document scope on a file without settings makes the part with them', () => {
    const pkg = applied(file(null), {
      op: 'setNoteProperties',
      scope: 'document',
      footnote: { position: 'beneathText', numFmt: 'lowerLetter' },
    });
    expect(names(settingsXml(pkg), 'w:settings')).toEqual(['w:footnotePr']);
  });

  test('section scope: the note properties go after the header reference', () => {
    const once = applied(file(SETTINGS), {
      op: 'setNoteProperties',
      scope: 'section',
      sectionIndex: 0,
      footnote: { numRestart: 'eachPage', numFmt: 'lowerRoman' },
    });
    expect(names(sectionXml(once), 'w:sectPr')).toEqual([
      'w:headerReference',
      'w:footnotePr',
      'w:pgSz',
      'w:pgMar',
    ]);
    // A second write merges: the format stays, the new start goes between in schema order
    const twice = applied(once, {
      op: 'setNoteProperties',
      scope: 'section',
      sectionIndex: 0,
      footnote: { numStart: 5 },
      endnote: { position: 'sectEnd' },
    });
    const xml = sectionXml(twice);
    expect(names(xml, 'w:sectPr')).toEqual([
      'w:headerReference',
      'w:footnotePr',
      'w:endnotePr',
      'w:pgSz',
      'w:pgMar',
    ]);
    expect(names(xml, 'w:footnotePr')).toEqual(['w:numFmt', 'w:numStart', 'w:numRestart']);
  });
});

/** A new note of a kind at the start of the body's first paragraph. */
function withNewNote(pkg: OoxmlPackage, noteKind: 'footnote' | 'endnote'): OoxmlPackage {
  const main = pkg.parts.get(pkg.mainDocumentPart);
  const body = main?.root.children.find((child) => child.kind === 'body');
  const paragraph =
    body && body.kind !== 'textValue'
      ? body.children.find((child) => child.kind === 'paragraph')
      : undefined;
  if (!paragraph) throw new Error('no paragraph');
  return applied(pkg, { op: 'insertNote', noteKind, paragraphId: paragraph.id, offset: 0 });
}

describe('a new notes part lists its separators in the settings, as Word writes them', () => {
  test('a new endnotes part: w:endnotePr after w:footnotePr, before w:compat', () => {
    const xml = settingsXml(withNewNote(file(SETTINGS), 'endnote'));
    expect(names(xml, 'w:settings').slice(2, 6)).toEqual([
      'w:characterSpacingControl',
      'w:footnotePr',
      'w:endnotePr',
      'w:compat',
    ]);
    expect(xml).toContain('<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>');
  });

  test('a list the settings already hold stays as it is; authored fields stay first', () => {
    const listed = withNewNote(file(SETTINGS), 'footnote');
    expect(settingsXml(listed)).toBe(settingsXml(file(SETTINGS)));
    const fields = SETTINGS.replace('<w:footnote w:id="-1"/><w:footnote w:id="0"/>', '');
    const xml = settingsXml(withNewNote(file(fields), 'footnote'));
    expect(xml).toContain(
      '<w:footnotePr><w:numFmt w:val="decimal"/><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
    );
  });

  test('a file without settings keeps none: no part is made for the list', () => {
    // Two people adding a first note at once in a Yjs room would each make the part, which
    // shared state cannot merge, and Word reads the notes without the list
    expect(settingsPartOf(withNewNote(file(null), 'footnote'))).toBeNull();
  });
});
