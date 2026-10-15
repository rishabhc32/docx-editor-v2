// Links proposed in suggesting mode (`hyperlink-suggestion.ts`): a new link over existing
// words, a new target or tooltip for an existing link, and the removal of a link.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import {
  collectReviewItems,
  listBookmarks,
  paragraphTextOf,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';
import { diagnosticCodeOfRefusal } from '../diagnostic-codes.ts';
import { narrowReviewItems } from '../review-replacement-pairs.ts';
import { docx } from './paginated-surface-fixtures.ts';

const plainParagraph = (text: string) =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** An inline picture run, its namespaces declared on itself for the plain fixture document. */
const PICTURE =
  '<w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="228600" cy="114300"/>' +
  '<wp:docPr id="1" name="pic1"/><a:graphic>' +
  '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
  '<pic:nvPicPr><pic:cNvPr id="1" name=""/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:ext cx="228600" cy="114300"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr>' +
  '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

function withSurface(body: string, run: (surface: PaginatedSurface) => void): void {
  const container = document.createElement('div');
  document.body.append(container);
  const opened = mountPaginatedSurface(container, docx(body), { author: 'Ada Lovelace' });
  if (!opened.ok) throw new Error(opened.reason);
  try {
    opened.surface.setEditingMode('suggest');
    run(opened.surface);
  } finally {
    opened.surface.destroy();
    container.remove();
  }
}

function caretTo(surface: PaginatedSurface, index: number, offset: number): void {
  const paragraphId = surface.session.paragraphIds()[index]!;
  surface.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
}

function selectIn(surface: PaginatedSurface, index: number, start: number, end: number): void {
  const paragraphId = surface.session.paragraphIds()[index]!;
  surface.setSelection({
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  });
}

function paragraphTexts(surface: PaginatedSurface): string[] {
  return surface.session
    .paragraphIds()
    .map((paragraphId) => paragraphTextOf(surface.session.part(), paragraphId));
}

describe('a link this author is still proposing', () => {
  test('renaming your OWN pending link keeps its tooltip', () => {
    withSurface(plainParagraph('Alpha Beta'), (surface) => {
      caretTo(surface, 0, 'Alpha'.length);
      expect(
        surface.hyperlinks.applyHyperlink({
          url: 'https://example.com',
          text: 'docs',
          tooltip: 'Tip',
        })
      ).toBe(true);
      caretTo(surface, 0, 'Alpha do'.length);
      expect(surface.hyperlinks.applyHyperlink({ url: 'https://example.com', text: 'guide' })).toBe(
        true
      );
      // The rename proposes a fresh link over the copy; it must carry the tooltip over.
      expect(surface.hyperlinks.linkAtCaret()).toMatchObject({ text: 'guide', tooltip: 'Tip' });
    });
  });
});

/** "See Docs now" with "Docs" linked to example.com, suggesting with the caret in the link. */
function withLinkedDocs(run: (surface: PaginatedSurface) => void): void {
  withSurface(plainParagraph('See Docs now'), (surface) => {
    surface.setEditingMode('edit');
    selectIn(surface, 0, 'See '.length, 'See Docs'.length);
    expect(
      surface.hyperlinks.applyHyperlink({ url: 'https://example.com', tooltip: 'Old tip' })
    ).toBe(true);
    surface.setEditingMode('suggest');
    caretTo(surface, 0, 'See Do'.length);
    run(surface);
  });
}

/** Each link in the caret's paragraph as [live text, href, tooltip]. */
function links(surface: PaginatedSurface): (string | null | undefined)[][] {
  return surface.hyperlinks
    .linksInCaretParagraph()
    .map((link) => [link.text, link.href, link.tooltip]);
}

// OOXML has no revision element for a link's target, so a target changed in place could not
// be reviewed or rejected. The old link's words are struck and a copy in a link to the new
// target is proposed after them.
describe('a new link target in suggesting mode', () => {
  test('strikes the old link and proposes one replacement after it', () => {
    withLinkedDocs((surface) => {
      expect(surface.hyperlinks.applyHyperlink({ url: 'https://example.org', text: 'Guide' })).toBe(
        true
      );
      expect(links(surface)).toEqual([
        ['', 'https://example.com', 'Old tip'],
        ['Guide', 'https://example.org', 'Old tip'],
      ]);
      expect(serializeOoxmlPart(surface.session.part())).toMatch(
        /<w:hyperlink[^>]*><w:del[^>]*><w:r>(?:(?!<\/w:r>).)*<w:delText>Docs<\/w:delText><\/w:r><\/w:del><\/w:hyperlink><w:hyperlink[^>]*><w:ins[^>]*w:author="Ada Lovelace"[^>]*><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:t>Guide<\/w:t>/
      );
      // Struck first, copy after: the review queue reads the pair as one replacement.
      const items = narrowReviewItems(
        collectReviewItems({ storyPart: surface.session.part() }),
        undefined
      );
      expect(items).toHaveLength(1);
      expect(items[0]).toMatchObject({ revisionKind: 'replace', replacedText: 'Docs' });
    });
  });

  test('rejecting gives the old link back whole, and accepting keeps the new one', () => {
    for (const op of ['rejectAllRevisions', 'acceptAllRevisions'] as const) {
      withLinkedDocs((surface) => {
        // A target change alone still copies the words, so the replacement has a side to keep.
        expect(surface.hyperlinks.applyHyperlink({ url: 'https://example.org' })).toBe(true);
        expect(surface.session.applyTreeOps([{ op }]).committed).toBe(true);
        expect(paragraphTexts(surface)).toEqual(['See Docs now']);
        const href = op === 'rejectAllRevisions' ? 'https://example.com' : 'https://example.org';
        expect(links(surface)).toEqual([['Docs', href, 'Old tip']]);
      });
    }
  });

  test('a new tooltip alone is proposed the same way', () => {
    withLinkedDocs((surface) => {
      expect(
        surface.hyperlinks.applyHyperlink({ url: 'https://example.com', tooltip: 'New tip' })
      ).toBe(true);
      expect(links(surface)).toEqual([
        ['', 'https://example.com', 'Old tip'],
        ['Docs', 'https://example.com', 'New tip'],
      ]);
    });
  });

  test('the same target again proposes nothing', () => {
    withLinkedDocs((surface) => {
      expect(surface.hyperlinks.applyHyperlink({ url: 'https://example.com' })).toBe(true);
      expect(links(surface)).toEqual([['Docs', 'https://example.com', 'Old tip']]);
      expect(collectReviewItems({ storyPart: surface.session.part() })).toEqual([]);
    });
  });

  test('gives the copy the link’s own format, not the format after it', () => {
    withSurface(
      '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:hyperlink w:anchor="old">' +
        '<w:r><w:rPr><w:b/><w:color w:val="FF0000"/><w:lang w:val="fr-FR"/></w:rPr><w:t>Docs</w:t></w:r>' +
        '</w:hyperlink><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> now</w:t></w:r></w:p>',
      (surface) => {
        // The spell-check language changes no glyph, so it does not stop the copy.
        caretTo(surface, 0, 'See Do'.length);
        expect(surface.hyperlinks.applyHyperlink({ url: 'https://example.org' })).toBe(true);
        expect(surface.session.applyTreeOps([{ op: 'acceptAllRevisions' }]).committed).toBe(true);
        const link = /<w:hyperlink[^>]*>(.*?)<\/w:hyperlink>/.exec(
          serializeOoxmlPart(surface.session.part())
        )![1]!;
        expect(link).toContain('<w:b/>');
        expect(link).toContain('<w:color w:val="FF0000"/>');
        expect(link).not.toContain('<w:i/>');
      }
    );
  });

  test('refuses a link a copy cannot reproduce, and mints no relationship', () => {
    for (const link of [
      // Two formats, a tab, a picture, and a target frame the new link cannot take.
      '<w:hyperlink w:anchor="old"><w:r><w:t>Do</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>cs</w:t></w:r></w:hyperlink>',
      `<w:hyperlink w:anchor="old">${PICTURE}<w:r><w:t>Docs</w:t></w:r></w:hyperlink>`,
      '<w:hyperlink w:anchor="old"><w:r><w:t>Do</w:t><w:tab/><w:t>cs</w:t></w:r></w:hyperlink>',
      '<w:hyperlink w:anchor="old" w:tgtFrame="_blank"><w:r><w:t>Docs</w:t></w:r></w:hyperlink>',
    ]) {
      withSurface(
        `<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>${link}</w:p>`,
        (surface) => {
          caretTo(surface, 0, 'See D'.length);
          const before = serializeOoxmlPart(surface.session.part());
          expect(surface.hyperlinks.applyHyperlink({ url: 'https://example.org' })).toBe(false);
          expect(surface.state().lastRejection).toContain('plain text in one format');
          expect(serializeOoxmlPart(surface.session.part())).toBe(before);
          expect(surface.session.currentPackage().externalTargets).toEqual([]);
        }
      );
    }
  });

  test('a link inside a content control is refused, not changed in place', () => {
    withSurface(
      '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:sdt><w:sdtContent>' +
        '<w:hyperlink w:anchor="x"><w:r><w:t>Docs</w:t></w:r></w:hyperlink>' +
        '</w:sdtContent></w:sdt></w:p>',
      (surface) => {
        caretTo(surface, 0, 'See Do'.length);
        const before = serializeOoxmlPart(surface.session.part());
        expect(surface.hyperlinks.applyHyperlink({ url: 'https://example.org' })).toBe(false);
        expect(surface.state().lastRejection).toContain('suggesting mode');
        expect(serializeOoxmlPart(surface.session.part())).toBe(before);
      }
    );
  });
});

describe('a new link over existing words in suggesting mode', () => {
  test('keeps the words’ own format in the linked copy', () => {
    withSurface(
      '<w:p><w:r><w:t xml:space="preserve">Read the </w:t></w:r>' +
        '<w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>terms</w:t></w:r>' +
        '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> now</w:t></w:r></w:p>',
      (surface) => {
        selectIn(surface, 0, 'Read the '.length, 'Read the terms'.length);
        expect(surface.hyperlinks.applyHyperlink({ url: 'https://example.com' })).toBe(true);
        expect(surface.session.applyTreeOps([{ op: 'acceptAllRevisions' }]).committed).toBe(true);
        expect(paragraphTexts(surface)).toEqual(['Read the terms now']);
        const link = /<w:hyperlink[^>]*>(.*?)<\/w:hyperlink>/.exec(
          serializeOoxmlPart(surface.session.part())
        )![1]!;
        expect(link).toContain('<w:b/>');
        expect(link).toContain('<w:color w:val="FF0000"/>');
        expect(link).not.toContain('<w:i/>');
      }
    );
  });

  test('takes this author’s own pending words', () => {
    withSurface(plainParagraph('Alpha'), (surface) => {
      caretTo(surface, 0, 'Alpha'.length);
      for (const character of ' docs') surface.type(character);
      selectIn(surface, 0, 'Alpha '.length, 'Alpha docs'.length);
      expect(surface.hyperlinks.applyHyperlink({ url: 'https://example.com' })).toBe(true);
      expect(surface.hyperlinks.linksInCaretParagraph().map((link) => link.text)).toEqual(['docs']);
    });
  });

  test('refuses words a copy cannot reproduce, and mints no relationship', () => {
    for (const [body, start, end] of [
      // Two formats, a tab, a picture, and words already struck.
      [
        '<w:r><w:t xml:space="preserve">Read the </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>terms</w:t></w:r>',
        0,
        14,
      ],
      ['<w:r><w:t>Name</w:t><w:tab/><w:t>Link</w:t></w:r>', 0, 9],
      [`<w:r><w:t xml:space="preserve">See </w:t></w:r>${PICTURE}<w:r><w:t>now</w:t></w:r>`, 4, 5],
      [
        '<w:r><w:t xml:space="preserve">Keep </w:t></w:r><w:del w:id="5" w:author="Grace" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>',
        0,
        9,
      ],
    ] as const) {
      withSurface(`<w:p>${body}</w:p>`, (surface) => {
        selectIn(surface, 0, start, end);
        const before = serializeOoxmlPart(surface.session.part());
        expect(surface.hyperlinks.applyHyperlink({ url: 'https://example.com' })).toBe(false);
        expect(surface.state().lastRejection).toContain('plain text in one format');
        expect(serializeOoxmlPart(surface.session.part())).toBe(before);
        expect(surface.session.currentPackage().externalTargets).toEqual([]);
      });
    }
  });
});

/** "See Docs now" with "Docs" a bold red link, and the words after it in italic. */
const FORMATTED_LINK =
  '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:hyperlink w:anchor="old" w:tooltip="Tip">' +
  '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>Docs</w:t></w:r>' +
  '</w:hyperlink><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> now</w:t></w:r></w:p>';

// OOXML has no revision element for a link's removal either. The link's words stay struck in
// it, and the plain copy after them is what accepting keeps.
describe('removing a link in suggesting mode', () => {
  test('strikes the link and proposes a plain copy after it', () => {
    withLinkedDocs((surface) => {
      expect(surface.hyperlinks.removeHyperlink()).toBe(true);
      expect(links(surface)).toEqual([['', 'https://example.com', 'Old tip']]);
      // The copy drops the run's `Hyperlink` style, as an unlink in editing mode does.
      expect(serializeOoxmlPart(surface.session.part())).toMatch(
        /<w:hyperlink[^>]*><w:del[^>]*><w:r>(?:(?!<\/w:r>).)*<w:delText>Docs<\/w:delText><\/w:r><\/w:del><\/w:hyperlink><w:ins[^>]*w:author="Ada Lovelace"[^>]*><w:r><w:t>Docs<\/w:t><\/w:r><\/w:ins><w:r>/
      );
      const items = narrowReviewItems(
        collectReviewItems({ storyPart: surface.session.part() }),
        undefined
      );
      expect(items).toHaveLength(1);
      expect(items[0]).toMatchObject({ revisionKind: 'replace', replacedText: 'Docs' });
      expect(surface.state().selection.head.offset).toBe('See DocsDocs'.length);
    });
  });

  test('accepting matches an unlink in editing mode, and rejecting gives the link back', () => {
    let unlinked = '';
    withSurface(FORMATTED_LINK, (surface) => {
      surface.setEditingMode('edit');
      caretTo(surface, 0, 'See Do'.length);
      expect(surface.hyperlinks.removeHyperlink()).toBe(true);
      unlinked = serializeOoxmlPart(surface.session.part());
    });
    // The words keep their own format and lose the link's style.
    expect(unlinked).toContain(
      '<w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>Docs</w:t></w:r>'
    );
    for (const op of ['acceptAllRevisions', 'rejectAllRevisions'] as const) {
      withSurface(FORMATTED_LINK, (surface) => {
        caretTo(surface, 0, 'See Do'.length);
        const before = serializeOoxmlPart(surface.session.part());
        expect(surface.hyperlinks.removeHyperlink()).toBe(true);
        expect(surface.session.applyTreeOps([{ op }]).committed).toBe(true);
        expect(serializeOoxmlPart(surface.session.part())).toBe(
          op === 'acceptAllRevisions' ? unlinked : before
        );
      });
    }
  });

  test('refuses a link a copy cannot reproduce, and leaves the document as it was', () => {
    for (const link of [
      // Two formats, and a picture.
      '<w:hyperlink w:anchor="old"><w:r><w:t>Do</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>cs</w:t></w:r></w:hyperlink>',
      `<w:hyperlink w:anchor="old">${PICTURE}<w:r><w:t>Docs</w:t></w:r></w:hyperlink>`,
    ]) {
      withSurface(
        `<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>${link}</w:p>`,
        (surface) => {
          caretTo(surface, 0, 'See D'.length);
          const before = serializeOoxmlPart(surface.session.part());
          expect(surface.hyperlinks.removeHyperlink()).toBe(false);
          expect(surface.state().lastRejection).toContain('plain text in one format');
          expect(serializeOoxmlPart(surface.session.part())).toBe(before);
        }
      );
    }
  });

  test('editing mode still takes the link off in place', () => {
    withLinkedDocs((surface) => {
      surface.setEditingMode('edit');
      expect(surface.hyperlinks.removeHyperlink()).toBe(true);
      expect(links(surface)).toEqual([]);
      expect(paragraphTexts(surface)).toEqual(['See Docs now']);
      expect(collectReviewItems({ storyPart: surface.session.part() })).toEqual([]);
    });
  });

  test('your own pending link comes off in place, with its style', () => {
    withSurface(plainParagraph('See Docs now'), (surface) => {
      selectIn(surface, 0, 'See '.length, 'See Docs'.length);
      expect(surface.hyperlinks.applyHyperlink({ url: 'https://example.com' })).toBe(true);
      // In the linked copy, after the struck words.
      caretTo(surface, 0, 'See DocsDo'.length);
      expect(surface.hyperlinks.removeHyperlink()).toBe(true);
      expect(links(surface)).toEqual([]);
      // Still one proposal: the struck words, and the copy as plain text.
      expect(serializeOoxmlPart(surface.session.part())).toMatch(
        /<w:del[^>]*><w:r><w:delText>Docs<\/w:delText><\/w:r><\/w:del><w:ins[^>]*><w:r><w:t>Docs<\/w:t><\/w:r><\/w:ins>/
      );
      expect(surface.session.applyTreeOps([{ op: 'rejectAllRevisions' }]).committed).toBe(true);
      expect(paragraphTexts(surface)).toEqual(['See Docs now']);
      expect(serializeOoxmlPart(surface.session.part())).not.toMatch(/w:rStyle|w:hyperlink/);
    });
  });
});

/** "See Docs now" with `words` in the `_Ref1` bookmark, as Word marks a cross-reference target. */
const bookmarked = (words: string) =>
  '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:bookmarkStart w:id="7" w:name="_Ref1"/>' +
  `${words}<w:bookmarkEnd w:id="7"/><w:r><w:t xml:space="preserve"> now</w:t></w:r></w:p>`;

/** The text the `_Ref1` bookmark holds. */
function bookmarkedText(surface: PaginatedSurface): string | undefined {
  const part = surface.session.part();
  const { start, end } = listBookmarks(part)[0]!;
  return paragraphTextOf(part, start.paragraphId)?.slice(start.offset, end.offset);
}

// A bookmark often wraps a link or a phrase exactly, and text inserted at a bookmark's end lands
// outside it. The copy must stay in the bookmark, or accepting leaves the bookmark empty.
describe('a link change inside a bookmark in suggesting mode', () => {
  const link = bookmarked(
    '<w:hyperlink w:anchor="old"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>' +
      '<w:t>Docs</w:t></w:r></w:hyperlink>'
  );
  for (const [change, body, propose, accepted] of [
    [
      'a new target',
      link,
      (surface: PaginatedSurface) => {
        caretTo(surface, 0, 'See Do'.length);
        return surface.hyperlinks.applyHyperlink({ url: 'https://example.org' });
      },
      [['Docs', 'https://example.org', undefined]],
    ],
    [
      'a link removal',
      link,
      (surface: PaginatedSurface) => {
        caretTo(surface, 0, 'See Do'.length);
        return surface.hyperlinks.removeHyperlink();
      },
      [],
    ],
    [
      'a new link',
      bookmarked('<w:r><w:t>Docs</w:t></w:r>'),
      (surface: PaginatedSurface) => {
        selectIn(surface, 0, 'See '.length, 'See Docs'.length);
        return surface.hyperlinks.applyHyperlink({ url: 'https://example.org' });
      },
      [['Docs', 'https://example.org', undefined]],
    ],
  ] as const) {
    test(`${change} keeps the copy in the bookmark, as one card`, () => {
      for (const op of ['acceptAllRevisions', 'rejectAllRevisions'] as const) {
        withSurface(body, (surface) => {
          const before = serializeOoxmlPart(surface.session.part());
          expect(propose(surface)).toBe(true);
          const items = narrowReviewItems(
            collectReviewItems({ storyPart: surface.session.part() }),
            undefined
          );
          expect(items).toHaveLength(1);
          expect(items[0]).toMatchObject({ revisionKind: 'replace', replacedText: 'Docs' });

          expect(surface.session.applyTreeOps([{ op }]).committed).toBe(true);
          const after = serializeOoxmlPart(surface.session.part());
          if (op === 'rejectAllRevisions') {
            // No link is left over to hold the bookmark's end or a relationship.
            expect(after).toBe(before);
            return;
          }
          expect(paragraphTexts(surface)).toEqual(['See Docs now']);
          expect(bookmarkedText(surface)).toBe('Docs');
          caretTo(surface, 0, 0);
          expect(links(surface)).toEqual([...accepted]);
          // Only the accepted link names a relationship, and no link is empty.
          expect(after.match(/r:id="/g) ?? []).toHaveLength(accepted.length);
          expect(after).not.toMatch(/<w:hyperlink[^>]*>(<w:bookmark[^>]*\/>)*<\/w:hyperlink>/);
        });
      }
    });
  }

  test('a new target beside your own fresh strike still gets its link', () => {
    withSurface(link, (surface) => {
      // The link's strike joins this one, which ends where the link starts.
      selectIn(surface, 0, 0, 'See '.length);
      expect(surface.deleteSelection()).toBe(true);
      caretTo(surface, 0, 'See Do'.length);
      expect(surface.hyperlinks.applyHyperlink({ url: 'https://example.org' })).toBe(true);
      expect(surface.session.applyTreeOps([{ op: 'acceptAllRevisions' }]).committed).toBe(true);
      expect(paragraphTexts(surface)).toEqual(['Docs now']);
      expect(bookmarkedText(surface)).toBe('Docs');
      caretTo(surface, 0, 0);
      expect(links(surface)).toEqual([['Docs', 'https://example.org', undefined]]);
    });
  });

  test('removing a link beside your own fresh strike is refused with the reason', () => {
    withSurface(link, (surface) => {
      selectIn(surface, 0, 0, 'See '.length);
      expect(surface.deleteSelection()).toBe(true);
      const before = serializeOoxmlPart(surface.session.part());
      caretTo(surface, 0, 'See Do'.length);
      expect(surface.hyperlinks.removeHyperlink()).toBe(false);
      expect(surface.state().lastRejection).toContain('wait a minute');
      expect(diagnosticCodeOfRefusal(surface.state().lastRejection!)).toBe(
        'edit.tracked-own-pending-change'
      );
      expect(serializeOoxmlPart(surface.session.part())).toBe(before);
    });
  });
});
