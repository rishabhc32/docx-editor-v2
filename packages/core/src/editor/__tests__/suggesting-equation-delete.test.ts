// Deleting an EQUATION in suggesting mode, through the surface's own keys.
//
// An equation (`m:oMath`, or a display `m:oMathPara`) is one unit of paragraph text. Backspace
// over a range that holds one used to strike the words around it and leave the equation
// unmarked, with no refusal. The reviewer saw a proposal that took the equation away; Accept
// removed the words and kept the equation. The strike now covers the equation too, as
// `w:del` around it (`CT_RunTrackChange` → `EG_ContentRunContent` → `EG_RunLevelElts` →
// `EG_MathContent`), so Accept removes it and Reject keeps it.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';
import {
  canonicalOoxmlFingerprint,
  collectReviewItems,
  findNode,
  paragraphTextOf,
  readOoxmlPart,
  revisionItemsOf,
  serializeOoxmlPart,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import type { ReviewModuleContribution } from '../../contracts/modules.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const DOCUMENT = {
  name: '/word/document.xml',
  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
};

function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${DOCUMENT.contentType}"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:m="${M}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const inlineEquation = '<m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath>';
const displayEquation = '<m:oMathPara><m:oMath><m:r><m:t>y=2</m:t></m:r></m:oMath></m:oMathPara>';
/** "Alpha ", the equation at offset 6, then " Omega." */
const withEquation = (equation: string) =>
  `<w:p>${run('Alpha ')}${equation}${run(' Omega.')}</w:p>`;

/** The engine's own review derivation (core may not import the review module). */
const reviewModel: ReviewModuleContribution = {
  displayModes: ['all-markup', 'proposed', 'original'],
  collectReviewItems,
  revisionItemsOfParagraph: (part, paragraphId) => {
    const paragraph = findNode(part, paragraphId);
    if (!paragraph || paragraph.kind !== 'paragraph') return [];
    return revisionItemsOf({ ...part, root: paragraph });
  },
};

function withSurface(body: string, check: (surface: PaginatedSurface) => void): void {
  const container = document.createElement('div');
  document.body.append(container);
  const opened = mountPaginatedSurface(container, docx(body), {
    author: 'Ada Lovelace',
    reviewModel,
  });
  if (!opened.ok) throw new Error(opened.reason);
  try {
    opened.surface.setEditingMode('suggest');
    check(opened.surface);
  } finally {
    opened.surface.destroy();
    container.remove();
  }
}

function select(surface: PaginatedSurface, start: number, end: number): void {
  const paragraphId = surface.session.paragraphIds()[0]!;
  surface.setSelection({
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  });
}

function xmlOf(surface: PaginatedSurface): string {
  return serializeOoxmlPart(surface.session.part());
}

/** The math elements that sit inside a `w:del`, by local name, in document order. */
function struckMath(root: OoxmlNode): string[] {
  const out: string[] = [];
  const walk = (node: OoxmlNode, struck: boolean): void => {
    if (node.kind === 'textValue') return;
    if (struck && node.namespaceUri === M && /^oMath(Para)?$/.test(node.localName)) {
      out.push(node.localName);
      return;
    }
    const inside = struck || node.kind === 'revisionDelete';
    for (const child of node.children) walk(child, inside);
  };
  walk(root, false);
  return out;
}

function deletionCards(surface: PaginatedSurface) {
  return surface.session
    .reviewItems()
    .filter((item) => item.kind === 'revision' && item.revisionKind === 'delete');
}

const cases = [
  { name: 'an inline equation', equation: inlineEquation, element: 'oMath' },
  { name: 'a display equation', equation: displayEquation, element: 'oMathPara' },
] as const;

for (const { name, equation, element } of cases) {
  describe(`suggesting deletion over ${name}`, () => {
    test('a range across it strikes the equation with the words, in one deletion', () => {
      withSurface(withEquation(equation), (surface) => {
        // "ha " + equation + " Om".
        select(surface, 3, 10);
        surface.deleteBackward();
        expect(surface.state().lastRejection).toBeNull();
        expect(struckMath(surface.session.part().root)).toEqual([element]);
        const xml = xmlOf(surface);
        // ONE decision: the struck words and the equation share a single `w:del`.
        expect(xml.match(/<w:del /g)).toHaveLength(1);
        expect(xml).toMatch(
          new RegExp(
            `<w:del [^>]*w:author="Ada Lovelace"[^>]*><w:r><w:delText[^>]*>ha </w:delText></w:r>` +
              `<m:${element}>.*</m:${element}><w:r><w:delText[^>]*> Om</w:delText></w:r></w:del>`
          )
        );
        const cards = deletionCards(surface);
        expect(cards).toHaveLength(1);
        // The one card covers the equation's offset as well as the words.
        expect(cards[0]!.ranges).toHaveLength(1);
        expect(cards[0]!.ranges[0]).toMatchObject({ start: { offset: 3 }, end: { offset: 10 } });
      });
    });

    test('selecting only the equation strikes it', () => {
      withSurface(withEquation(equation), (surface) => {
        select(surface, 6, 7);
        surface.deleteBackward();
        expect(surface.state().lastRejection).toBeNull();
        expect(struckMath(surface.session.part().root)).toEqual([element]);
        expect(deletionCards(surface)).toHaveLength(1);
      });
    });

    test('Backspace with the caret after the equation strikes it', () => {
      withSurface(withEquation(equation), (surface) => {
        select(surface, 7, 7);
        surface.deleteBackward();
        expect(surface.state().lastRejection).toBeNull();
        expect(struckMath(surface.session.part().root)).toEqual([element]);
        // A second Backspace steps over the strike rather than nesting another one.
        surface.deleteBackward();
        expect(xmlOf(surface).match(new RegExp(`<m:${element}>`, 'g'))).toHaveLength(1);
        expect(struckMath(surface.session.part().root)).toEqual([element]);
      });
    });

    test('Delete with the caret before the equation strikes it', () => {
      withSurface(withEquation(equation), (surface) => {
        select(surface, 6, 6);
        surface.deleteForward();
        expect(surface.state().lastRejection).toBeNull();
        expect(struckMath(surface.session.part().root)).toEqual([element]);
      });
    });

    test('typing over a selection that holds it strikes it and lands the text after', () => {
      withSurface(withEquation(equation), (surface) => {
        select(surface, 3, 10);
        surface.type('Z');
        expect(surface.state().lastRejection).toBeNull();
        expect(struckMath(surface.session.part().root)).toEqual([element]);
        expect(xmlOf(surface)).toMatch(
          new RegExp(
            `</m:${element}><w:r><w:delText[^>]*> Om</w:delText></w:r></w:del>` +
              `<w:ins [^>]*w:author="Ada Lovelace"[^>]*><w:r><w:t>Z</w:t></w:r></w:ins>`
          )
        );
      });
    });

    test('Accept removes the equation, and Reject keeps it unmarked', () => {
      for (const action of ['acceptAllRevisions', 'rejectAllRevisions'] as const) {
        withSurface(withEquation(equation), (surface) => {
          select(surface, 3, 10);
          surface.deleteBackward();
          const result = surface.session.applyTreeOps([{ op: action }]);
          expect(result.committed).toBe(true);
          const xml = xmlOf(surface);
          expect(xml).not.toMatch(/<w:del /);
          expect(struckMath(surface.session.part().root)).toEqual([]);
          const part = surface.session.part();
          const text = paragraphTextOf(part, surface.session.paragraphIds()[0]!);
          if (action === 'acceptAllRevisions') {
            expect(xml).not.toContain(`<m:${element}>`);
            expect(text).toBe('Alpega.');
          } else {
            expect(xml.match(new RegExp(`<m:${element}>`, 'g'))).toHaveLength(1);
            expect(text).toBe('Alpha \uFFFC Omega.');
          }
          expect(deletionCards(surface)).toHaveLength(0);
        });
      }
    });

    test('the strike survives save and reopen', () => {
      withSurface(withEquation(equation), (surface) => {
        select(surface, 3, 10);
        surface.deleteBackward();
        const saved = serializeOoxmlPart(surface.session.part());
        const reopened = readOoxmlPart(saved, DOCUMENT);
        if (!reopened.ok) throw new Error(reopened.reason);
        expect(canonicalOoxmlFingerprint(reopened.part)).toBe(
          canonicalOoxmlFingerprint(surface.session.part())
        );
        expect(struckMath(reopened.part.root)).toEqual([element]);
      });
    });
  });
}

test('the equation in your own pending insertion goes outright', () => {
  // Your own suggestion was never anyone else's to see, so taking it back proposes nothing.
  withSurface(
    `<w:p>${run('Alpha ')}<w:ins w:id="5" w:author="Ada Lovelace" w:date="2026-01-01T00:00:00Z">` +
      `${inlineEquation}</w:ins>${run(' Omega.')}</w:p>`,
    (surface) => {
      select(surface, 6, 7);
      surface.deleteBackward();
      expect(surface.state().lastRejection).toBeNull();
      const xml = xmlOf(surface);
      expect(xml).not.toContain('<m:oMath>');
      expect(xml).not.toMatch(/<w:del /);
      expect(xml).not.toMatch(/<w:ins /);
    }
  );
});

test("another author's inserted equation is struck inside their insertion", () => {
  withSurface(
    `<w:p>${run('Alpha ')}<w:ins w:id="5" w:author="Grace Hopper" w:date="2026-01-01T00:00:00Z">` +
      `${inlineEquation}</w:ins>${run(' Omega.')}</w:p>`,
    (surface) => {
      select(surface, 6, 7);
      surface.deleteBackward();
      expect(surface.state().lastRejection).toBeNull();
      expect(xmlOf(surface)).toMatch(
        /<w:ins [^>]*w:author="Grace Hopper"[^>]*><w:del [^>]*w:author="Ada Lovelace"[^>]*><m:oMath>/
      );
    }
  );
});
