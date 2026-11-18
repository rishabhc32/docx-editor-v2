// Loaded header and footer parts a session stops using (task 9.8).

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { collectSectionPropertyNodes } from '../../../store/package/hf-references.ts';
import { withoutPart } from '../../../store/package/package-edit.ts';
import { loadedFurnitureOf } from '../load/sections.ts';
import type { OrderedDocument } from '../model/types.ts';
import { isWml } from '../ooxml-nodes.ts';

/** The relationship ids a main part's sections name for their headers and footers. */
function referencesOf(root: OoxmlNode): Set<string> {
  const ids = new Set<string>();
  for (const sectPr of collectSectionPropertyNodes(root)) {
    for (const child of sectPr?.children ?? []) {
      if (child.kind === 'textValue') continue;
      if (!isWml(child, 'headerReference') && !isWml(child, 'footerReference')) continue;
      const id = child.attributes.find((a) => a.localName === 'id' && a.prefix === 'r');
      if (id) ids.add(id.value);
    }
  }

  return ids;
}

/**
 * The package without the loaded header and footer parts no section names any more. The engine
 * removes a header or footer part, its relationship and its content type when a delete or a link
 * takes the last section reference to it (`hf-lifecycle.ts`, `garbageCollectStory`). A part the
 * file never named stays, as the engine leaves it.
 */
export function withoutDroppedFurniture(pkg: OoxmlPackage, doc: OrderedDocument): OoxmlPackage {
  const loaded = loadedFurnitureOf(doc).parts;
  if (loaded.size === 0) return pkg;
  const story = doc.stories.get(pkg.mainDocumentPart);
  const main = pkg.parts.get(pkg.mainDocumentPart);
  if (!story || !main || story.part.root === main.root) return pkg;
  const before = referencesOf(story.part.root);
  const now = referencesOf(main.root);
  const used = new Set<string>();
  for (const id of now) {
    const part = loaded.get(id)?.part;
    if (part !== undefined) used.add(part);
  }

  let next = pkg;
  for (const id of before) {
    const part = loaded.get(id)?.part;
    if (part === undefined || used.has(part) || !next.parts.has(part)) continue;
    const removed = withoutPart(next, part);
    if (removed.ok) next = removed.pkg;
  }

  return next;
}
