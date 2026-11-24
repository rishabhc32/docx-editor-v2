// Each section change a person can make, suggested or direct, with and without another person's
// pending section record, then decided (design D14): a reject or a revert gives each section every
// setting it had, and the room, a save and reopen, and one and two folds agree on the file, on its
// accept and on its reject of each person's suggestions.

import { describe, expect, test } from 'bun:test';
import { openOrderedWordFile } from '../client/open.ts';
import type { OrderedDocument } from '../model/types.ts';
import { addressOf, apply, BODY, DIRECT, EDIT, markOf, partsOf } from './change-apply-fixture.ts';
import { checkFold, mediaStore } from './fold-check.ts';
import { checkSaveAndLoad } from './sessions/session.ts';
import { docx, p, R, t, WML_CT } from './word-fixtures.ts';

const SETTINGS =
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708"' +
  ' w:footer="708" w:gutter="0"/><w:pgNumType w:start="5"/><w:cols w:num="2" w:space="720"/>' +
  '<w:docGrid w:linePitch="360"/>';

// Two sections: the first names a header and starts the second on the same page
const FILE = docx(
  p(
    t('First section.'),
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/>' +
      `<w:pgSz w:w="12240" w:h="15840"/>${SETTINGS}</w:sectPr>`
  ) +
    p(t('Second section.')) +
    p(t('Its last paragraph.')) +
    `<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/>${SETTINGS}</w:sectPr>`,
  {
    parts: {
      'word/header1.xml':
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>',
    },
    relationships: `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes: `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>`,
  }
);

const bodySection = (doc: OrderedDocument) =>
  addressOf([...doc.stories.get(BODY)!.tree.items()].find((item) => item.block === true)!);

/** Each section of the main part as compose draws it, without ids. */
const sectionsOf = (doc: OrderedDocument): string[] =>
  (partsOf(doc)[0] ?? '')
    .match(/<w:sectPr>(?:(?!<\/w:sectPr>)[\s\S])*<\/w:sectPr>(?:<\/w:sectPrChange><\/w:sectPr>)?/g)
    ?.map((xml) => xml.replace(/ w:id="[^"]*"/g, '')) ?? [];

/** The settings no change here names, as each section states them now. */
const keptSettings = (doc: OrderedDocument): string[] =>
  sectionsOf(doc).map((xml) =>
    (
      xml
        .replace(/<w:sectPrChange\b.*/, '')
        .match(/<w:(pgMar|pgNumType|cols|docGrid)\b[^>]*\/>/g) ?? []
    ).join('')
  );

type Mode = typeof EDIT | typeof DIRECT;

type Change = (doc: OrderedDocument, mode: Mode) => void;

/** Olive's section changes Word records as a suggestion (`w:sectPrChange`), each in one mode. */
const CHANGES: Readonly<Record<string, Change>> = {
  'page size': (doc, mode) =>
    apply(
      doc,
      'olive:1',
      {
        kind: 'section-property',
        ...mode,
        sections: [],
        page: { pageWidthTwips: 11906, pageHeightTwips: 16838 },
      },
      'Olive'
    ),
  orientation: (doc, mode) =>
    apply(
      doc,
      'olive:1',
      { kind: 'section-property', ...mode, sections: [], page: { orientation: 'landscape' } },
      'Olive'
    ),
  "margins of one paragraph's section": (doc, mode) =>
    apply(
      doc,
      'olive:1',
      {
        kind: 'section-property',
        ...mode,
        sections: [],
        page: { marginTopTwips: 2000, marginLeftTwips: 1800 },
        pageAt: markOf(doc, 1),
      },
      'Olive'
    ),
  'a different first page': (doc, mode) =>
    apply(
      doc,
      'olive:1',
      {
        kind: 'section-property',
        ...mode,
        sections: [markOf(doc, 0)],
        furniture: { titlePage: true },
      },
      'Olive'
    ),
  'the header distance': (doc, mode) =>
    apply(
      doc,
      'olive:1',
      {
        kind: 'section-property',
        ...mode,
        sections: [bodySection(doc)],
        furniture: { headerDistanceTwips: 500 },
      },
      'Olive'
    ),
};

/** Olive's header and footer links, which Word records nowhere: direct only. */
const LINKS: Readonly<Record<string, Change>> = {
  'the header unlinked': (doc, mode) =>
    apply(
      doc,
      'olive:1',
      {
        kind: 'header-footer',
        ...mode,
        action: 'unlink',
        section: bodySection(doc),
        story: 'header',
        variant: 'default',
      },
      'Olive'
    ),
  'a new footer': (doc, mode) =>
    apply(
      doc,
      'olive:1',
      {
        kind: 'header-footer',
        ...mode,
        action: 'create',
        section: markOf(doc, 0),
        story: 'footer',
        variant: 'default',
      },
      'Olive'
    ),
};

/** Cy's margins pending on every section, or none. */
const CY = {
  'no other record': (_doc: OrderedDocument) => {},
  "Cy's margins pending": (doc: OrderedDocument) =>
    apply(
      doc,
      'cy:1',
      { kind: 'section-property', ...EDIT, sections: [], page: { marginRightTwips: 2160 } },
      'Cy'
    ),
};

describe('a section change decided, live and folded', () => {
  for (const [other, before] of Object.entries(CY)) {
    for (const [label, change] of Object.entries(CHANGES)) {
      test(`${label}, suggested (${other})`, async () => {
        const { document: doc, base } = openOrderedWordFile(FILE, { file: 1 });
        before(doc);
        const kept = keptSettings(doc);
        change(doc, EDIT);
        await checkFold(doc, base, mediaStore(base), `${label} (${other})`);
        checkSaveAndLoad(doc, label, base);
        apply(doc, 'olive:2', { kind: 'retract', suggestions: ['S:olive:1.1'] }, 'Olive');
        expect(keptSettings(doc)).toEqual(kept);
      });
    }

    for (const [label, change] of Object.entries({ ...CHANGES, ...LINKS })) {
      test(`${label}, direct and reverted (${other})`, async () => {
        const { document: doc, base } = openOrderedWordFile(FILE, { file: 1 });
        before(doc);
        const sections = sectionsOf(doc);
        change(doc, DIRECT);
        await checkFold(doc, base, mediaStore(base), `${label} (${other})`);
        apply(
          doc,
          'olive:2',
          { kind: 'revert', suggestions: ['S:olive:1.1'], status: 'rejected' },
          'Olive'
        );
        expect(sectionsOf(doc)).toEqual(sections);
        await checkFold(doc, base, mediaStore(base), `${label} reverted (${other})`);
      });
    }

    for (const [label, change] of Object.entries(LINKS)) {
      test(`${label}, suggested, is refused by name (${other})`, () => {
        const { document: doc } = openOrderedWordFile(FILE, { file: 1 });
        before(doc);
        expect(() => change(doc, EDIT)).toThrow('Word records no change');
      });
    }

    test(`a section break that starts a page, direct and reverted (${other})`, async () => {
      const { document: doc, base } = openOrderedWordFile(FILE, { file: 1 });
      before(doc);
      const sections = sectionsOf(doc);
      apply(
        doc,
        'olive:1',
        { kind: 'section-break', ...DIRECT, mark: markOf(doc, 1), type: 'nextPage' },
        'Olive'
      );
      expect(sectionsOf(doc)).toHaveLength(3);
      await checkFold(doc, base, mediaStore(base), `a section break (${other})`);
      apply(
        doc,
        'olive:2',
        { kind: 'revert', suggestions: ['S:olive:1.1'], status: 'rejected' },
        'Olive'
      );
      expect(sectionsOf(doc)).toEqual(sections);
      await checkFold(doc, base, mediaStore(base), `a section break reverted (${other})`);
      checkSaveAndLoad(doc, 'a section break reverted', base);
    });
  }
});
