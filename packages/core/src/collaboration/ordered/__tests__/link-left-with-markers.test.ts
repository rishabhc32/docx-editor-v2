// A new link that a decision leaves with only markers (the random two-person session, seed 103):
// a party proposed a new target for a link, and a comment started inside the new link's copy.
// Word's Reject All keeps such a link around the comment's start, with its relationship and
// address; only a link left with nothing goes. The engine and a reopened room file did so
// already; the live room dropped the link with its rejected suggestion. Now it keeps it too, and
// the link's address with it (`compose/kept-links.ts`). A take-back is no decision: the link
// never was, and it goes with its address.

import { describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { setStatus } from '../model/mutations.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  APPLY_FILE,
  apply,
  at,
  BODY,
  DIRECT,
  EDIT,
  opens,
  textOf,
} from './change-apply-fixture.ts';
import { checkSaveAndLoad } from './sessions/session.ts';

const NEW_ADDRESS = 'https://example.com/new';

/** Ana's new target for the link of the second paragraph, and Olga's comment inside it. */
function proposedWithDiscussion() {
  const opened = openOrderedWordFile(APPLY_FILE, { file: 1 });
  const doc = opened.document;
  const [link] = opens(doc, 'hyperlink');
  if (!link) throw new Error('fixture');
  // "the terms" struck, a copy in a new link after it
  const proposed = apply(doc, 'ana-1:1', {
    kind: 'link-target',
    ...EDIT,
    link: { space: link.space, offset: link.offset },
    href: NEW_ADDRESS,
  });
  expect(textOf(doc, 1)).toBe('See the termsthe terms.');
  // The comment starts after "th" of the copy and ends after the paragraph's text
  apply(
    doc,
    'olga-1:1',
    { kind: 'comment', part: BODY, from: at(doc, 1, 15), to: at(doc, 1, 23), text: 'Which?' },
    'Olga'
  );
  return { doc, base: opened.base, suggestion: proposed.suggestions[0] ?? '' };
}

/** The saved file's second paragraph, and the links' relationships to the new address. */
function saved(doc: OrderedDocument, base: OoxmlPackage) {
  const files = unzipSync(orderedPackageBytes(doc, base));
  const xml = strFromU8(files['word/document.xml']!);
  const rels = strFromU8(files['word/_rels/document.xml.rels']!);
  const second = /<w:p\b(?:(?!<\/w:p>).)*See (?:(?!<\/w:p>).)*<\/w:p>/.exec(xml)?.[0] ?? '';
  const addresses = [...rels.matchAll(/<Relationship [^>]*\/>/g)]
    .map((match) => match[0])
    .filter((rel) => rel.includes(NEW_ADDRESS));
  return { second, addresses };
}

describe('a new link that a decision leaves with only a comment start', () => {
  test('stays around the start in the live document, with its address, as in Word', () => {
    const { doc, base, suggestion } = proposedWithDiscussion();
    setStatus(doc, suggestion, 'rejected');
    const live = serializeOoxmlPart(composeOrderedPart(doc, BODY).part);
    expect(live).not.toContain('the termsthe terms');
    const { second, addresses } = saved(doc, base);
    // The file's own link, and the rejected new link holding only the comment's start
    expect(second.match(/<w:hyperlink /g)?.length).toBe(2);
    const kept =
      /<w:hyperlink [^>]*r:id="([^"]+)"[^>]*><w:commentRangeStart [^>]*\/><\/w:hyperlink>/.exec(
        second
      );
    expect(kept).not.toBeNull();
    expect(addresses).toHaveLength(1);
    expect(addresses[0]).toContain(`Id="${kept?.[1]}"`);
    // The live copy and the reopened file show the same
    checkSaveAndLoad(doc, 'a rejected new link around a comment start', base);
  });

  test('saves with the start inside the new link, and reopens to decide the same', () => {
    const { doc, base } = proposedWithDiscussion();
    const xml = strFromU8(unzipSync(orderedPackageBytes(doc, base))['word/document.xml']!);
    const links = xml.match(/<w:hyperlink [^>]*>(?:(?!<\/w:hyperlink>).)*<\/w:hyperlink>/g) ?? [];
    expect(links.some((link) => link.includes('<w:commentRangeStart '))).toBe(true);
    checkSaveAndLoad(doc, 'a comment start inside a new link', base);
  });

  test("another person's text in a rejected new link stays in the link, with its address", () => {
    const { doc, base, suggestion } = proposedWithDiscussion();
    // Olga types inside Ana's copy: Word keeps the link around Olga's insertion
    apply(doc, 'olga-2:1', { kind: 'insert', ...EDIT, place: at(doc, 1, 16), text: 'X' }, 'Olga');
    setStatus(doc, suggestion, 'rejected');
    const { second, addresses } = saved(doc, base);
    expect(second).toMatch(/<w:hyperlink [^>]*r:id="rO[^>]*>(?:(?!<\/w:hyperlink>).)*>X<\/w:t>/);
    expect(addresses).toHaveLength(1);
    checkSaveAndLoad(doc, "another person's text in a rejected new link", base);
  });

  test('a take-back is no decision: the link goes with its address, the start stays', () => {
    const { doc, base, suggestion } = proposedWithDiscussion();
    apply(doc, 'ana-2:1', { kind: 'retract', suggestions: [suggestion] });
    const { second, addresses } = saved(doc, base);
    expect(second.match(/<w:hyperlink /g)?.length).toBe(1);
    expect(second).toContain('<w:commentRangeStart ');
    expect(second).not.toMatch(/<w:hyperlink [^>]*><w:commentRangeStart /);
    expect(addresses).toHaveLength(0);
    checkSaveAndLoad(doc, 'a taken-back new link around a comment start', base);
  });
});

describe('a rejected new link the document keeps is a link like any other', () => {
  /** Olga's text in Ana's rejected new link, and the link's open item. */
  function keptWithText() {
    const { doc, base, suggestion } = proposedWithDiscussion();
    apply(doc, 'olga-2:1', { kind: 'insert', ...EDIT, place: at(doc, 1, 16), text: 'X' }, 'Olga');
    setStatus(doc, suggestion, 'rejected');
    const story = doc.stories.get(BODY);
    const kept = [...(story?.tree.items() ?? [])].find(
      (item) => item.kind === 'open' && item.token === 'hyperlink' && item.by === suggestion
    );
    if (!kept) throw new Error('fixture');
    return { doc, base, link: { space: kept.space, offset: kept.offset } };
  }

  test('the organiser gives it a new address, as in a reopened file', () => {
    const { doc, base, link } = keptWithText();
    const href = 'https://example.com/third';
    apply(doc, 'org-1:1', { kind: 'link-target', ...DIRECT, link, href }, 'Org');
    const files = unzipSync(orderedPackageBytes(doc, base));
    const xml = strFromU8(files['word/document.xml']!);
    const rels = strFromU8(files['word/_rels/document.xml.rels']!);
    const id = /<w:hyperlink [^>]*r:id="([^"]+)"[^>]*>(?:(?!<\/w:hyperlink>).)*>X<\/w:t>/.exec(
      xml
    )?.[1];
    expect(rels).toMatch(new RegExp(`Id="${id}"[^>]*Target="${href}"`));
    checkSaveAndLoad(doc, 'a new address for a kept link', base);
  });

  test("a party's new address copies only the text that shows", () => {
    const { doc, base, link } = keptWithText();
    apply(doc, 'ben-1:1', { kind: 'link-target', ...EDIT, link, href: NEW_ADDRESS }, 'Ben');
    // Olga's "X" struck in the kept link, and a copy of it alone in Ben's new link
    expect(textOf(doc, 1)).toBe('See the termsXX.');
    checkSaveAndLoad(doc, "a party's new address for a kept link", base);
  });

  test('the organiser takes it off the text, as in a reopened file', () => {
    const { doc, base, link } = keptWithText();
    apply(doc, 'org-1:1', { kind: 'remove-link', ...DIRECT, link }, 'Org');
    const { second, addresses } = saved(doc, base);
    expect(second).not.toMatch(/<w:hyperlink [^>]*>(?:(?!<\/w:hyperlink>).)*>X<\/w:t>/);
    expect(second).toContain('>X</w:t>');
    expect(addresses).toHaveLength(0);
    checkSaveAndLoad(doc, 'a kept link taken off its text', base);
  });
});

describe("a party's new address for a link", () => {
  test('copies only the words that show, never text a decision took out', () => {
    const opened = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const doc = opened.document;
    // Olga's "Y" in the file's link, rejected
    const typed = apply(
      doc,
      'olga-1:1',
      { kind: 'insert', ...EDIT, place: at(doc, 1, 6), text: 'Y' },
      'Olga'
    );
    setStatus(doc, typed.suggestions[0] ?? '', 'rejected');
    const [link] = opens(doc, 'hyperlink');
    if (!link) throw new Error('fixture');
    apply(
      doc,
      'ben-1:1',
      {
        kind: 'link-target',
        ...EDIT,
        link: { space: link.space, offset: link.offset },
        href: NEW_ADDRESS,
      },
      'Ben'
    );
    expect(textOf(doc, 1)).toBe('See the termsthe terms.');
    checkSaveAndLoad(doc, 'a new address for a link with rejected text', opened.base);
  });
});
