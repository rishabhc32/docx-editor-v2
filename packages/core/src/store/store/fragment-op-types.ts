import type { OoxmlNode } from '../package/ooxml-tree.ts';
import type { RevisionAttributionInput } from './tree-op-revision-attribution.ts';

/**
 * Insert clipboard-fragment blocks at a paragraph position, atomically.
 *
 * Blocks only: resources (styles, numbering, media, rels, note parts) merge through
 * the package-edit path in the same transaction, and the whole commit is promoted to
 * a package undo unit. A merged edge paragraph takes the properties of the paragraph
 * mark that ends it: the fragment's first mark leading, the host's original mark
 * trailing. Nodes receive fresh ids at apply.
 */
export type InsertFragmentOp = {
  readonly op: 'insertFragment';
  readonly paragraphId: string;
  readonly offset: number;
  readonly blocks: readonly OoxmlNode[];
  /** True when the fragment's last paragraph mark travelled (its paragraph stays whole). */
  readonly lastMarkCovered?: boolean;
  /**
   * Write the fragment as ONE tracked insertion by this author and date, the way Word writes a
   * paste with Track Changes on: every pasted run inside a `w:ins`, every pasted paragraph
   * mark and every new mark the landing writes as `w:pPr/w:rPr/w:ins`, every pasted table row
   * as `w:trPr/w:ins`. Paragraph, table and cell properties stay plain, and the resources the
   * paste merges are not revisions. All of it takes one `w:id`. Other authors' changes in the
   * fragment are kept as Word keeps them: their insertions become part of this one, and their
   * deletions and property records stay nested in it with their own author and date. Only a
   * custom XML element's tracked range is refused (`fragment-holds-revisions`).
   */
  readonly revision?: RevisionAttributionInput;
};
