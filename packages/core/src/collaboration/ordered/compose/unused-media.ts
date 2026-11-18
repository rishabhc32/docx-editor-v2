// Loaded media a session made unused leaves the saved file (group 9.4): a deleted picture's
// bytes must not stay in the file someone receives.

import { cleanupOrphanImageMedia } from '../../../store/package/drawing-package-edit.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { isRelationshipReference } from '../../../store/package/canonical-media-references.ts';
import type { OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import {
  IMAGE_RELATIONSHIP_TYPE,
  resolveImageRelationship,
} from '../../../store/package/relationships.ts';
import type { OrderedDocument } from '../model/types.ts';
import { eachElement } from '../ooxml-nodes.ts';

/** Relationship ids a node and everything under it name, in one bounded walk. */
function relationshipsUsed(root: OoxmlNode): Set<string> {
  const out = new Set<string>();
  eachElement(root, (node) => {
    for (const attribute of node.attributes)
      if (isRelationshipReference(attribute)) out.add(attribute.value);
  });
  return out;
}

/**
 * `pkg` without the image relationships (and their media parts, when nothing else names them)
 * that a loaded story of `doc` used and its composed part in `pkg` no longer uses, as the editor
 * alone removes them (`cleanupOrphanImageMedia`). An orphan the file came with stays, and the
 * engine's own check keeps a part any other relationship or drawing still names.
 */
export function withoutUnusedLoadedMedia(pkg: OoxmlPackage, doc: OrderedDocument): OoxmlPackage {
  let next = pkg;
  for (const [name, story] of doc.stories) {
    const records = pkg.relationships.get(name)?.filter((r) => r.type === IMAGE_RELATIONSHIP_TYPE);
    const composed: OoxmlPart | undefined = pkg.parts.get(name);
    if (!records || records.length === 0 || !composed || composed === story.part) continue;
    const before = relationshipsUsed(story.part.root);
    const now = relationshipsUsed(composed.root);
    for (const record of records) {
      if (!before.has(record.id) || now.has(record.id)) continue;
      const resolved = resolveImageRelationship(records, name, record.id);
      const media = resolved.mode === 'internal' ? resolved.partName : null;
      next = cleanupOrphanImageMedia(next, name, media, record.id);
    }
  }

  return next;
}
