// A review module for core's ordered client tests that derives the engine's own review cards, as
// the Pro module does without its custom nodes, so the editor's accept, reject and comment
// writes run as in a product (core's tests may not import `@docx-editor.dev/pro`).

import type { EditorModule } from '../../../../contracts/modules.ts';
import { findNode } from '../../../../store/package/ooxml-edit.ts';
import type { OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { collectReviewItems, revisionItemsOf } from '../../../../store/store/review-reads.ts';

/** Revisions wholly inside one paragraph, read on a paragraph-root view of the part. */
function revisionItemsOfParagraph(part: OoxmlPart, paragraphId: string) {
  const paragraph = findNode(part, paragraphId);
  if (!paragraph || paragraph.kind !== 'paragraph') return [];
  return revisionItemsOf({
    id: part.id,
    name: part.name,
    contentType: part.contentType,
    root: paragraph,
  });
}

export function reviewTestModule(): EditorModule {
  return {
    id: 'review',
    review: {
      displayModes: ['all-markup', 'simple-markup', 'proposed', 'original'],
      collectReviewItems,
      revisionItemsOfParagraph,
    },
  };
}
