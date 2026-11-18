// The Word round trip for the matrix: compose every story into the package, keep the prompts
// a suggestion hid in the glossary, write the file, read it again.

import {
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '../../../../store/package/ooxml-package.ts';
import { relationshipsOf } from '../../../../store/package/package-edit.ts';
import { withRetainedPrompts } from '../../../../store/store/content-control-prompt-glossary.ts';
import { composeOrderedPart, loadOrderedPackage, type OrderedDocument } from '../../index.ts';
import { mediaBytesOf, withOrderedResources } from '../../compose/resources.ts';
import { isTextBoxStoryName } from '../../model/text-boxes.ts';
import type { RelationshipResolver } from './compare.ts';

/** The package with the model's stories composed into it, written and read again. */
export function saved(pkg: OoxmlPackage, doc: OrderedDocument): OoxmlPackage {
  let next = pkg;
  for (const name of doc.stories.keys()) {
    // A text box's story is drawn into its object, in the body's part
    if (isTextBoxStoryName(name)) continue;
    const { part, composed } = composeOrderedPart(doc, name);
    next = withPart(next, part);
    next = withRetainedPrompts(next, composed.prompts, new Set());
  }
  // What the edits added outside the stories: a new link's address. A new picture's media
  // stays out, as on each side the new image reads by its order of appearance (`compare.ts`)
  next = withOrderedResources(next, doc, new Map(), mediaBytesOf({ missing: 'skip' }));
  const read = readOoxmlPackage(writeOoxmlPackage(next), {});
  if (!read.ok) throw new Error(`the saved file does not open: ${read.reason}`);
  return read.package;
}

/** The model, loaded again from a saved file. */
export function reopened(pkg: OoxmlPackage): OrderedDocument {
  return loadOrderedPackage(pkg);
}

/** Relationship targets of a story part: the package's, then the model's own new ones. */
export function resolverOf(
  pkg: OoxmlPackage,
  name: string,
  doc?: OrderedDocument
): RelationshipResolver {
  const records = relationshipsOf(pkg, name);
  const added = new Map<string, string>();
  for (const resource of doc?.resources.added.values() ?? [])
    if (
      resource.kind === 'relationship' &&
      resource.owner === name &&
      'external' in resource.target
    )
      added.set(resource.id, resource.target.external);
  return (id) => added.get(id) ?? records.find((r) => r.id === id)?.rawTarget ?? null;
}
