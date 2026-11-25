// Whole-document save: save then load gives the same document, by view (every story's composed
// part) and by decisions (each review card accepted or rejected alone, and all of them), for
// the item-kind matrix fixtures and for a session's edits; what the session added outside the
// stories (link targets, hidden prompts, the room's comments, the comments part) is written
// where Word reads it.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPackage, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { relationshipsOf } from '../../../store/package/package-edit.ts';
import {
  canonicalOoxmlFingerprint,
  serializeOoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { composeOrderedPackage, orderedPackageBytes } from '../compose/package.ts';
import { orderedCards } from '../decide.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { setStatus } from '../model/mutations.ts';
import type { OrderedDocument } from '../model/types.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import {
  addressOf,
  apply,
  APPLY_FILE,
  at,
  BODY,
  DIRECT,
  EDIT,
  HEADER,
  load,
  markOf,
  opens,
} from './change-apply-fixture.ts';
import { normalizedBody, semanticBody } from './matrix/compare.ts';
import { unresolvedRelationships } from './oracle.ts';
import { ROOM_FILE } from './room-harness.ts';
import { bodyOf, docx, MATRIX_EXTRAS, MATRIX_ITEMS, packageOf } from './word-fixtures.ts';

/**
 * Every story's composed part: as the fidelity oracle sees it, or, after decisions, as the
 * matrix compares it (`semanticBody`: an emptied `w:rPr` the decision left is no difference).
 */
function view(
  doc: OrderedDocument,
  decide: readonly (readonly [string, 'accepted' | 'rejected'])[] = []
) {
  const { record } = withUndoRecord(doc, () => {
    for (const [id, status] of decide)
      if (doc.suggestions.get(id)?.status === 'pending') setStatus(doc, id, status);
  });
  const parts = [...doc.stories.keys()].map((name) => {
    const part = composeOrderedPart(doc, name).part;
    return [
      name,
      decide.length === 0
        ? canonicalOoxmlFingerprint(part)
        : semanticBody(normalizedBody(part, (id) => id)),
    ];
  });
  takeBack(doc, record);
  return parts;
}

/** The saved file, read and loaded again: each relationship id resolves in its own part. */
function reopened(doc: OrderedDocument, base: OoxmlPackage) {
  const read = readOoxmlPackage(orderedPackageBytes(doc, base), {});
  if (!read.ok) throw new Error(`the saved file does not open: ${read.reason}`);
  expect(unresolvedRelationships(read.package)).toEqual([]);
  return { pkg: read.package, again: loadOrderedPackage(read.package) };
}

/** The pending suggestions of each author, in creation order. */
function byAuthor(doc: OrderedDocument): Map<string, string[]> {
  const out = new Map<string, string[]>();
  for (const s of [...doc.suggestions.values()].sort((a, b) => a.seq - b.seq))
    if (s.status === 'pending')
      out.set(s.attribution.author, [...(out.get(s.attribution.author) ?? []), s.id]);
  return out;
}

/**
 * The same document after a save and a reopen: the view; every suggestion decided at once,
 * both ways; and each author's suggestions decided alone, both ways. With `cards`, each review
 * card decided alone too (cards in document order): a file's own revisions keep their cards, but
 * a session's change can be two revisions in Word form (a changed link target is the struck old
 * link and the new one), which the matrix also matches by author.
 */
function expectSameAfterReopen(
  label: string,
  doc: OrderedDocument,
  again: OrderedDocument,
  cards = false
): void {
  expect({ label, view: view(again) }).toEqual({ label, view: view(doc) });
  const live = byAuthor(doc);
  const later = byAuthor(again);
  expect({ label, authors: [...later.keys()].sort() }).toEqual({
    label,
    authors: [...live.keys()].sort(),
  });
  for (const status of ['accepted', 'rejected'] as const) {
    const each = (ids: readonly string[]) => ids.map((id) => [id, status] as const);
    expect({ label, status, view: view(again, each([...later.values()].flat())) }).toEqual({
      label,
      status,
      view: view(doc, each([...live.values()].flat())),
    });
    for (const [author, ids] of live)
      expect({ label, author, status, view: view(again, each(later.get(author) ?? [])) }).toEqual({
        label,
        author,
        status,
        view: view(doc, each(ids)),
      });
    if (!cards) continue;
    const liveCards = orderedCards(doc);
    const laterCards = orderedCards(again);
    expect({ label, cards: laterCards.length }).toEqual({ label, cards: liveCards.length });
    liveCards.slice(0, 6).forEach((card, k) => {
      const other = laterCards[k];
      expect({ label, card: k, status, view: view(again, each(other?.suggestions ?? [])) }).toEqual(
        { label, card: k, status, view: view(doc, each(card.suggestions)) }
      );
    });
  }
}

describe('whole-document save', () => {
  for (const item of MATRIX_ITEMS) {
    test(`the matrix fixture ${item.key}: save then load gives the same document`, () => {
      const base = packageOf(docx(bodyOf(item), MATRIX_EXTRAS));
      const doc = loadOrderedPackage(base);
      const { again } = reopened(doc, base);
      expectSameAfterReopen(item.key, doc, again, true);
    });
  }

  test("a session's edits survive the save, and what they added goes where Word reads it", () => {
    const base = packageOf(APPLY_FILE);
    const baseBody = base.parts.get(BODY);
    const doc = load();
    const [link] = opens(doc, 'hyperlink');
    const [control] = opens(doc, 'sdt');
    if (!link || !control) throw new Error('fixture');
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 4), text: 'new ' });
    apply(doc, 'ana:2', {
      kind: 'link-target',
      ...EDIT,
      link: addressOf(link),
      href: 'https://e.com/n',
    });
    apply(doc, 'ana:3', {
      kind: 'insert-note',
      ...EDIT,
      place: at(doc, 8, 3),
      note: 'footnote',
      text: ' Net of tax.',
    });
    apply(doc, 'ana:4', {
      kind: 'comment',
      part: BODY,
      from: at(doc, 0, 0),
      to: at(doc, 0, 3),
      text: 'Discuss this.',
    });
    apply(doc, 'ana:5', {
      kind: 'insert',
      part: HEADER,
      mode: 'suggest',
      place: at(doc, 0, 0, HEADER),
      text: 'Draft ',
    });
    apply(doc, 'ben:1', { kind: 'split', ...EDIT, place: at(doc, 9, 6) }, 'Ben');
    apply(
      doc,
      'ben:2',
      {
        kind: 'paragraph-property',
        ...EDIT,
        marks: [markOf(doc, 3)],
        properties: [{ localName: 'jc', attributes: { val: 'right' } }],
      },
      'Ben'
    );
    apply(
      doc,
      'ben:3',
      {
        kind: 'control-value',
        ...EDIT,
        control: addressOf(control),
        value: { kind: 'text', text: 'Beta Ltd' },
      },
      'Ben'
    );
    apply(doc, 'ben:4', { kind: 'insert-tab', ...EDIT, place: at(doc, 1, 0) }, 'Ben');
    const { pkg, again } = reopened(doc, base);
    expectSameAfterReopen('session', doc, again);
    // The new link's target is the story's relationship
    const linkId =
      [...doc.resources.added.values()].find((r) => r.kind === 'relationship')?.key === undefined
        ? ''
        : (
            [...doc.resources.added.values()].find((r) => r.kind === 'relationship') as {
              id: string;
            }
          ).id;
    expect(pkg.externalTargets.find((t) => t.id === linkId)?.rawTarget).toBe('https://e.com/n');
    // The room's comment is written after the file's
    const comments = pkg.parts.get('/word/comments.xml');
    expect(comments ? serializeOoxmlPart(comments) : '').toContain('Discuss this.');
    expect(comments ? serializeOoxmlPart(comments) : '').toContain('Check this.');
    // The base package is not changed
    expect(base.parts.get(BODY)).toBe(baseBody);
  });

  test('a filled prompt is kept in the glossary, so a reopened rejection shows it again', () => {
    const item = MATRIX_ITEMS.find((candidate) => candidate.key === 'placeholder');
    if (!item) throw new Error('no placeholder fixture');
    const base = packageOf(docx(bodyOf(item), MATRIX_EXTRAS));
    const doc = loadOrderedPackage(base);
    // Typing at the prompt replaces it (rule 7)
    const [control] = opens(doc, 'sdt');
    if (!control) throw new Error('no control');
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 1, 6), text: 'Acme' });
    const { pkg, again } = reopened(doc, base);
    expectSameAfterReopen('placeholder', doc, again);
    expect([...pkg.parts.keys()].some((name) => name.includes('glossary'))).toBe(true);
  });

  // Relationship ids are per part: the glossary's block names the organiser's link through a
  // relationship of the glossary's own, and the body keeps its own for a reject
  test("a filled prompt with the organiser's link keeps it in the glossary's relationships", () => {
    const item = MATRIX_ITEMS.find((candidate) => candidate.key === 'placeholder');
    if (!item) {
      throw new Error('no placeholder fixture');
    }

    const base = packageOf(docx(bodyOf(item), MATRIX_EXTRAS));
    const doc = loadOrderedPackage(base);
    const link = { from: at(doc, 1, 7), to: at(doc, 1, 9), href: 'https://example.com/fees' };
    apply(doc, 'olive:1', { kind: 'insert-link', ...DIRECT, ...link }, 'Olive');
    apply(doc, 'ana:2', { kind: 'insert', ...EDIT, place: at(doc, 1, 6), text: 'Acme' });
    const { pkg, again } = reopened(doc, base);
    expectSameAfterReopen('a link on the prompt', doc, again);
    const glossary = pkg.externalTargets.filter((t) => t.ownerPart.includes('glossary'));
    expect(glossary.map((t) => t.rawTarget)).toEqual(['https://example.com/fees']);
  });

  test('a file with no comments part gets one, with its content type and relationship', () => {
    const base = packageOf(ROOM_FILE);
    expect(base.parts.has('/word/comments.xml')).toBe(false);
    const doc = loadOrderedPackage(base);
    apply(
      doc,
      'olga:1',
      { kind: 'comment', part: BODY, from: at(doc, 1, 0), to: at(doc, 1, 3), text: 'Why?' },
      'Olga'
    );
    const saved = composeOrderedPackage(doc, base);
    const comments = saved.parts.get('/word/comments.xml');
    expect(comments ? serializeOoxmlPart(comments) : '').toContain('w:author="Olga"');
    expect(saved.contentTypes.overrides.get('/word/comments.xml')).toBe(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
    );
    expect(saved.contentTypes.overrides.get('/word/commentsextended.xml')).toBe(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml'
    );
    const related = relationshipsOf(saved, '/word/document.xml').map((r) => r.rawTarget);
    expect(related).toContain('comments.xml');
    expect(related).toContain('commentsExtended.xml');
    const { again } = reopened(doc, base);
    expectSameAfterReopen('comments part', doc, again);
    expect(base.parts.has('/word/comments.xml')).toBe(false);
  });
});
