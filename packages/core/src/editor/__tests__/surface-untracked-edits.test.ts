// Which attributed batches suggesting can show as suggestions, and the gate that refuses the rest
// when a review module asks for `untrackedEdits: 'refuse'`.

import { describe, expect, test } from 'bun:test';
import type { OoxmlNode, TreeDocOp } from '@docx-editor.dev/core/store';
import { readOoxmlPart } from '../../store/package/ooxml-tree.ts';
import {
  createUntrackedEditGate,
  untrackedEditOf,
  UNTRACKED_EDIT_REFUSAL,
} from '../surface-untracked-edits.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A pending insertion by `author` of one run. */
const inserted = (author: string, text: string) =>
  `<w:ins w:id="1" w:author="${author}" w:date="2026-10-16T00:00:00Z"><w:r><w:t>${text}</w:t></w:r></w:ins>`;

/**
 * "see " and then "site" linked to `rId5` with a tooltip. Then three pending links: one the
 * reviewer inserted, one another author inserted, and one that also holds a plain run.
 */
const read = readOoxmlPart(
  `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t xml:space="preserve">see </w:t></w:r>` +
    '<w:hyperlink r:id="rId5" w:tooltip="Tip"><w:r><w:t>site</w:t></w:r></w:hyperlink></w:p>' +
    `<w:p><w:hyperlink r:id="rId6">${inserted('Reviewer', 'own')}</w:hyperlink>` +
    `<w:hyperlink r:id="rId7">${inserted('Other', 'theirs')}</w:hyperlink>` +
    `<w:hyperlink r:id="rId8">${inserted('Reviewer', 'mixed')}<w:r><w:t>!</w:t></w:r></w:hyperlink>` +
    '</w:p></w:body></w:document>',
  { name: '/word/document.xml', contentType: 'application/xml' }
);
if (!read.ok) throw new Error(read.reason);
const part = read.part;
const partOf = (nodeId: string) => (nodeId.startsWith(`${part.name}#`) ? part : null);

/** The ids of every node of `kind`, in document order. */
function idsOf(kind: OoxmlNode['kind'], node: OoxmlNode = part.root): string[] {
  if (node.kind === kind) return [node.id];
  return node.kind === 'textValue' ? [] : node.children.flatMap((child) => idsOf(kind, child));
}
const [paragraphId] = idsOf('paragraph') as [string];
const [linkId, ownLinkId, theirLinkId, mixedLinkId] = idsOf('hyperlink') as [
  string,
  string,
  string,
  string,
];

const revision = { author: 'Reviewer', date: '2026-10-16T00:00:00Z' };
const margins: TreeDocOp = { op: 'setSectionProperties', marginLeftTwips: 720 };
const bold: TreeDocOp = {
  op: 'setRunProperties',
  paragraphId: 'p1',
  start: 0,
  end: 3,
  properties: [{ localName: 'b' }],
};

describe('untrackedEditOf', () => {
  test('a batch of tracked ops changes nothing untracked', () => {
    const ops: TreeDocOp[] = [
      { op: 'insertText', paragraphId: 'p1', offset: 0, text: 'a', revision },
      { op: 'deleteText', paragraphId: 'p1', start: 2, end: 4, revision },
      { ...bold, revision } as TreeDocOp,
      { op: 'proposeParagraphMerge', paragraphId: 'p2', revision },
    ];
    expect(untrackedEditOf(ops, true, partOf)).toBeNull();
  });

  test('one untracked op names the whole batch, whatever is tracked beside it', () => {
    const ops: TreeDocOp[] = [
      { op: 'insertText', paragraphId: 'p1', offset: 0, text: 'a', revision },
      margins,
    ];
    expect(untrackedEditOf(ops, true, partOf)).toBe(margins);
    expect(untrackedEditOf([bold], true, partOf)).toBe(bold);
  });

  test('review decisions, author changes, comments and protection are not edits to record', () => {
    const address = { id: '1', author: 'A', date: '2026-01-01T00:00:00Z' };
    const accept = { op: 'acceptRevision', revision: address } as TreeDocOp;
    expect(untrackedEditOf([accept], true, partOf)).toBeNull();
    expect(untrackedEditOf([{ op: 'rejectAllRevisions' } as TreeDocOp], true, partOf)).toBeNull();
    expect(
      untrackedEditOf(
        [
          {
            op: 'insertCommentMarker',
            paragraphId: 'p1',
            offset: 0,
            commentId: '3',
            marker: 'start',
          },
        ],
        true,
        partOf
      )
    ).toBeNull();
    const protection: TreeDocOp = { op: 'setDocumentProtection', enforce: true };
    expect(untrackedEditOf([protection], true, partOf)).toBeNull();
  });

  test('bookmarks are not revisions, and a batch with no ops has nothing to record', () => {
    const start = { paragraphId: 'p1', offset: 0 };
    const bookmarks: TreeDocOp[] = [
      { op: 'insertBookmark', name: 'Term', start, end: { ...start, offset: 3 } },
      { op: 'deleteBookmark', name: 'Term' },
      { op: 'deleteBookmark', name: 'Term', keepFirst: true },
    ];
    for (const bookmark of bookmarks) expect(untrackedEditOf([bookmark], true, partOf)).toBeNull();
    expect(untrackedEditOf([], true, partOf)).toBeNull();
  });

  test('formatting left untracked on purpose passes; other properties still do not', () => {
    expect(untrackedEditOf([bold], false, partOf)).toBeNull();
    expect(untrackedEditOf([bold, margins], false, partOf)).toBe(margins);
  });

  test('a split and a section mark ride on the paragraph mark the batch proposes', () => {
    const split: TreeDocOp = { op: 'splitParagraph', paragraphId: 'p1', offset: 2 };
    const sectionMark: TreeDocOp = { op: 'setSectionMark', paragraphId: 'p1' };
    const mark: TreeDocOp = {
      op: 'setParagraphMarkRevision',
      paragraphId: 'p1',
      kind: 'ins',
      revision,
    };
    expect(untrackedEditOf([split, mark, sectionMark], true, partOf)).toBeNull();
    expect(untrackedEditOf([split], true, partOf)).toBe(split);
    expect(untrackedEditOf([split, { ...mark, paragraphId: 'p2' }], true, partOf)).toBe(split);
  });

  describe('a suggested change to an existing link', () => {
    // The ops the link lane writes: the link's words struck, and a tracked copy of them aimed
    // at the strike's front edge, which the store lands after the struck words, in the link.
    const strike: TreeDocOp = { op: 'deleteText', paragraphId, start: 4, end: 8, revision };
    const copy: Extract<TreeDocOp, { op: 'insertText' }> = {
      op: 'insertText',
      paragraphId,
      offset: 4,
      text: 'site',
      revision,
    };
    const unlink: Extract<TreeDocOp, { op: 'removeHyperlink' }> = {
      op: 'removeHyperlink',
      linkId,
      range: { start: 8, end: 12 },
    };
    const byReviewer = (ops: readonly TreeDocOp[]) =>
      untrackedEditOf(ops, true, partOf, 'Reviewer');

    test('an unlink lets go of the copy where the store lands it, and of nothing else', () => {
      expect(byReviewer([strike, copy, unlink])).toBeNull();
      // The struck words where the copy was aimed, the whole link, or past the copy.
      const unlinks: TreeDocOp[] = [
        { ...unlink, range: { start: 4, end: 8 } },
        { op: 'removeHyperlink', linkId },
        { ...unlink, range: { start: 8, end: 13 } },
      ];
      for (const other of unlinks) expect(byReviewer([strike, copy, other])).toBe(other);
      // With no author, or a link the lookup cannot place, the copy is nobody's own.
      expect(untrackedEditOf([strike, copy, unlink], true, partOf)).toBe(unlink);
      expect(untrackedEditOf([strike, copy, unlink], true, () => null, 'Reviewer')).toBe(unlink);
    });

    test('a new link wraps the copy of the selected words where the store lands it', () => {
      const ops: TreeDocOp[] = [
        { op: 'deleteText', paragraphId, start: 0, end: 3, revision },
        { op: 'insertText', paragraphId, offset: 0, text: 'see', revision },
        { op: 'setRunProperties', paragraphId, start: 3, end: 6, properties: [], revision },
        { op: 'insertHyperlink', paragraphId, start: 3, end: 6, relationshipId: 'rId9' },
      ];
      expect(byReviewer(ops)).toBeNull();
      // A link over the struck words, where the copy was aimed, is no suggestion.
      const struck: TreeDocOp = { ...ops[3]!, start: 0, end: 3 } as TreeDocOp;
      expect(byReviewer([...ops.slice(0, 3), struck])).toBe(struck);
    });

    test('a new target re-links the copy aimed before the last struck character', () => {
      // Inside the struck words, the store lands the copy after them, in the link.
      const aimed = { ...copy, offset: 7 };
      const relink: TreeDocOp[] = [
        { op: 'setRunProperties', paragraphId, start: 8, end: 12, properties: [], revision },
        {
          op: 'insertHyperlink',
          paragraphId,
          start: 8,
          end: 12,
          relationshipId: 'rId9',
          tooltip: 'Tip',
        },
      ];
      // Each link op over the landed copy. The whole batch also unlinks the copy before it
      // re-links it, which needs a store that divides a link holding revisions.
      expect(byReviewer([strike, aimed, unlink])).toBeNull();
      expect(byReviewer([strike, aimed, ...relink])).toBeNull();
      // The same copy written by another author's insertion is not this suggestion.
      const theirs = { ...aimed, revision: { ...revision, author: 'Other' } };
      expect(byReviewer([strike, theirs, unlink])).toBe(unlink);
      expect(byReviewer([strike, theirs, ...relink])).toBe(relink[1]!);
    });

    test('a copy the batch does not write as a suggestion is never taken for one', () => {
      const plain: TreeDocOp = { op: 'insertText', paragraphId, offset: 4, text: 'site' };
      expect(byReviewer([strike, plain, unlink])).toBe(plain);
      // An op before the link that the replay does not know leaves the copy unknown.
      const marker: TreeDocOp = {
        op: 'insertCommentMarker',
        paragraphId,
        offset: 0,
        commentId: '3',
        marker: 'start',
      };
      expect(byReviewer([strike, copy, marker, unlink])).toBe(unlink);
    });

    test('a link re-aimed at its own target changes nothing; any other target is untracked', () => {
      const keep: Extract<TreeDocOp, { op: 'setHyperlinkTarget' }> = {
        op: 'setHyperlinkTarget',
        linkId,
        relationshipId: 'rId5',
      };
      // A suggested rename names the target it keeps beside the tracked strike and copy.
      expect(untrackedEditOf([keep, strike, copy], true, partOf)).toBeNull();
      expect(untrackedEditOf([{ ...keep, tooltip: 'Tip' }], true, partOf)).toBeNull();
      const retargets: TreeDocOp[] = [
        { ...keep, relationshipId: 'rId6' },
        { op: 'setHyperlinkTarget', linkId },
        { ...keep, anchor: 'Term' },
        { ...keep, tooltip: 'Other' },
        // A range splits the link, even with its own target.
        { ...keep, range: { start: 4, end: 6 } },
        { ...keep, linkId: paragraphId },
      ];
      for (const retarget of retargets) {
        expect(untrackedEditOf([retarget, strike, copy], true, partOf)).toBe(retarget);
      }
      expect(untrackedEditOf([keep], true, () => null)).toBe(keep);
    });

    test("re-aiming the author's own pending link is part of that proposal; no other link is", () => {
      const reaim = (id: string): Extract<TreeDocOp, { op: 'setHyperlinkTarget' }> => ({
        op: 'setHyperlinkTarget',
        linkId: id,
        relationshipId: 'rId9',
        tooltip: 'New',
      });
      expect(untrackedEditOf([reaim(ownLinkId)], true, partOf, 'Reviewer')).toBeNull();
      // Another author's insertion, a plain run beside one's own, or a link nobody proposed.
      for (const id of [theirLinkId, mixedLinkId, linkId]) {
        const op = reaim(id);
        expect(untrackedEditOf([op], true, partOf, 'Reviewer')).toBe(op);
      }
      // With no author, or with a range that splits the link, it is no longer the proposal.
      const own = reaim(ownLinkId);
      expect(untrackedEditOf([own], true, partOf)).toBe(own);
      const split = { ...own, range: { start: 0, end: 1 } };
      expect(untrackedEditOf([split], true, partOf, 'Reviewer')).toBe(split);
    });
  });
});

describe('createUntrackedEditGate', () => {
  const gate = (policy: 'apply' | 'refuse' | undefined, suggesting = true) =>
    createUntrackedEditGate({
      policy,
      suggesting: () => suggesting,
      formattingTracked: () => true,
      partOf,
      author: () => 'Reviewer',
    });

  test('refuses an untracked batch only for refuse, and only while suggesting', () => {
    expect(gate('refuse').refusal([margins])).toBe(UNTRACKED_EDIT_REFUSAL);
    expect(gate('refuse').refusal([{ ...bold, revision } as TreeDocOp])).toBeNull();
    expect(gate('refuse', false).refusal([margins])).toBeNull();
    expect(gate('apply').refusal([margins])).toBeNull();
    expect(gate(undefined).refusal([margins])).toBeNull();
  });

  test("passes the author's re-aim of their own pending link", () => {
    const reaim: TreeDocOp = {
      op: 'setHyperlinkTarget',
      linkId: ownLinkId,
      relationshipId: 'rId9',
    };
    expect(gate('refuse').refusal([reaim])).toBeNull();
    expect(gate('refuse').refusal([{ ...reaim, linkId: theirLinkId }])).toBe(
      UNTRACKED_EDIT_REFUSAL
    );
  });

  test('an exempt lane passes, and the gate closes again after it, even when it throws', () => {
    const strict = gate('refuse');
    expect(strict.exempt(() => strict.refusal([margins]))).toBeNull();
    expect(() =>
      strict.exempt(() => {
        throw new Error('lane failed');
      })
    ).toThrow('lane failed');
    expect(strict.refusal([margins])).toBe(UNTRACKED_EDIT_REFUSAL);
  });

  test('is in force for refuse while suggesting, outside an exempt lane', () => {
    const strict = gate('refuse');
    expect(strict.inForce()).toBe(true);
    expect(strict.exempt(() => strict.inForce())).toBe(false);
    expect(gate('refuse', false).inForce()).toBe(false);
    expect(gate('apply').inForce()).toBe(false);
    expect(gate(undefined).inForce()).toBe(false);
  });
});
