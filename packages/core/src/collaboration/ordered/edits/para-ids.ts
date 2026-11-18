// A new paragraph's Word identity. Word gives every paragraph a `w14:paraId` (and `w14:textId`)
// when the story binds the `w14` namespace, and the editor gives one to any paragraph that lacks
// it when it opens a file. A paragraph an edit makes gets its ids at once, from the change's id,
// so every copy writes the same file and a save opened again composes to the same file.

import { W14_NAMESPACE_URI, WML_NAMESPACE_URI } from '../../../store/package/ooxml-shared.ts';
import type { OoxmlAttribute, OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { usedParaIds, w14RootPrefix } from '../../../store/package/para-id.ts';
import type { OrderedMintedParagraph } from '../change/ids.ts';
import { recordUndo } from '../model/undo-record.ts';
import { withAttributes, withChildren } from '../ooxml-nodes.ts';
import type { Ctx } from './context.ts';

/** Whether an attribute is one of a paragraph's Word ids (`w14:paraId`, `w14:textId`). */
export const isParagraphIdentity = (attribute: OoxmlAttribute): boolean =>
  attribute.namespaceUri === W14_NAMESPACE_URI &&
  (attribute.localName === 'paraId' || attribute.localName === 'textId');

/**
 * A new paragraph's mark space and Word ids, past every id `taken` names and every id this
 * session's new paragraphs hold. The session keeps the id until the edit's undo.
 */
export function mintParagraph(
  ctx: Pick<Ctx, 'doc' | 'ids'>,
  taken: (paraId: string) => boolean = () => false
): OrderedMintedParagraph {
  const minted = ctx.ids.paragraph((id) => taken(id) || ctx.doc.paraIds.has(id));
  const ids = ctx.doc.paraIds;
  ids.add(minted.paraId);
  recordUndo(ctx.doc, () => ids.delete(minted.paraId));
  return minted;
}

/** The `w14:paraId` and `w14:textId` attributes of minted Word ids, under the story's prefix. */
export function paragraphIdAttributes(
  prefix: string,
  minted: { readonly paraId: string; readonly textId: string }
): OoxmlAttribute[] {
  const attribute = (localName: string, value: string): OoxmlAttribute => ({
    kind: 'genericExtension',
    namespaceUri: W14_NAMESPACE_URI,
    localName,
    prefix,
    value,
  });
  return [attribute('paraId', minted.paraId), attribute('textId', minted.textId)];
}

/** A paragraph with these Word ids in place of its own. */
export function withParagraphIdentity(
  paragraph: OoxmlElement,
  identity: readonly OoxmlAttribute[]
): OoxmlElement {
  const kept = paragraph.attributes.filter((attribute) => !isParagraphIdentity(attribute));
  return withAttributes(paragraph, [...kept, ...identity]);
}

/** A minter of new paragraph identities, past every id the story and this session hold. */
function minter(ctx: Ctx, prefix: string) {
  const used = usedParaIds(ctx.story.part.root);
  return () => {
    const minted = mintParagraph(ctx, (id) => used.has(id));
    return { space: minted.space, attributes: paragraphIdAttributes(prefix, minted) };
  };
}

/**
 * A new paragraph's identity: its mark's space, which is also its node id, and, when the story
 * binds the `w14` namespace, a `w14:paraId` and `w14:textId` past every id the story and this
 * session's new paragraphs hold. All of it comes from the change's id, so every copy makes the
 * same paragraph.
 */
export function newParagraphIdentity(ctx: Ctx): {
  space: string;
  attributes: OoxmlAttribute[];
} {
  const prefix = w14RootPrefix(ctx.story.part.root);
  if (prefix === null) return { space: ctx.ids.paragraph().space, attributes: [] };
  return minter(ctx, prefix)();
}

const isParagraph = (node: OoxmlElement) =>
  node.namespaceUri === WML_NAMESPACE_URI && node.localName === 'p';

/**
 * An engine template (a new note, a new table) with a new `w14:paraId` and `w14:textId` on each
 * of its paragraphs, when the story binds the `w14` namespace. An id the template already holds
 * goes too: the engine minted it on a scratch part, against that part's few paragraphs and from
 * a seed another template can share, so two new tables could take the same ids. Templates are a
 * few nodes deep, so the walk is short.
 */
export function withParagraphIds(ctx: Ctx, template: OoxmlElement): OoxmlElement {
  const prefix = w14RootPrefix(ctx.story.part.root);
  if (prefix === null) return template;
  const mint = minter(ctx, prefix);
  const visit = (node: OoxmlElement): OoxmlElement => {
    const children = node.children.map((child) =>
      child.kind === 'textValue' ? child : visit(child)
    );
    const changed = children.some((child, k) => child !== node.children[k]);
    const kept = changed ? withChildren(node, children) : node;
    if (!isParagraph(node)) return kept;
    return withParagraphIdentity(kept, mint().attributes);
  };
  return visit(template);
}
