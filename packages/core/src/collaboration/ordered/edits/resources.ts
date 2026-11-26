// The edits that add resources outside the story text (design D14), for the groups' change
// kinds to call inside their apply. Every key and Word name comes from the change's id minter,
// so every copy adds the same resource under the same name. A resource belongs to the
// suggestion that added it: another suggestion's edit may not use it while it is pending
// (`foreign-resource`), so a reject of the one never breaks the other.

import {
  RELATIONSHIPS_NAMESPACE_URI,
  W14_NAMESPACE_URI,
} from '../../../store/package/ooxml-shared.ts';
import {
  readOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import { sanitizeHref } from '../../../store/package/sinks.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import type {
  OrderedChangeList,
  OrderedChangeMedia,
  OrderedChangeSetting,
  OrderedChangeStyle,
} from '../change/resources.ts';
import { flatProperties } from '../change/apply-properties.ts';
import { addOrderedStory, createOrderedItem, labelUnindex } from '../model/document.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import {
  addOrderedResource,
  orderedListNumber,
  orderedMediaPart,
  orderedRelationshipId,
  orderedStoryPart,
  type OrderedListBullet,
  type OrderedListLevelResource,
  type OrderedListResource,
  type OrderedMediaResource,
  type OrderedRelationshipResource,
  type OrderedRelationshipType,
  type OrderedResource,
  type OrderedSettingResource,
  type OrderedStoryResource,
  type OrderedStoryRole,
  type OrderedStyleResource,
} from '../model/resources.ts';
import type { OrderedStory } from '../model/types.ts';
import { recordUndo } from '../model/undo-record.ts';
import type { Ctx } from './context.ts';
import { mintParagraph, paragraphIdAttributes } from './para-ids.ts';

/**
 * Media a change names by hash, for the change's suggestion, under `key` (a key the edit minted
 * earlier, when its element names it first).
 */
export function addMedia(
  ctx: Ctx,
  suggestion: string,
  media: OrderedChangeMedia,
  key: string = ctx.ids.resource()
): OrderedMediaResource {
  const resource: OrderedMediaResource = {
    kind: 'media',
    key,
    suggestion,
    part: orderedMediaPart(key, media.contentType),
    contentType: media.contentType,
    sha256: media.sha256,
    size: media.size,
  };
  addOrderedResource(ctx.doc, resource);
  return resource;
}

/**
 * A resource an edit uses: one its own suggestion added, or one whose suggestion is accepted
 * (the document's now). Another person's pending resource is refused.
 */
export function resourceFor(ctx: Ctx, key: string, suggestion: string): OrderedResource {
  const found = ctx.doc.resources.added.get(key);
  if (!found) throw new OrderedRefusal('unknown-address', `no resource ${key.slice(0, 160)}`);
  const status = ctx.doc.suggestions.get(found.suggestion)?.status;
  if (found.suggestion !== suggestion && status !== 'accepted')
    throw new OrderedRefusal('foreign-resource', key.slice(0, 160));
  return found;
}

/** Whether a link's address is one `sanitizeHref` keeps as it is. */
function isSafeLinkAddress(href: string): boolean {
  const projection = sanitizeHref(href);
  return projection.ok && projection.href === href;
}

/**
 * A relationship from a story part (the edit's own by default) to a resource the same
 * suggestion added, or for a link to an address that `sanitizeHref` keeps as it is.
 */
export function addRelationship(
  ctx: Ctx,
  suggestion: string,
  type: OrderedRelationshipType,
  target: OrderedRelationshipResource['target'],
  owner: string = ctx.story.name,
  /** The owner is a chart's or a diagram's part the same edit adds next (`kind: 'part'`). */
  newPart = false
): OrderedRelationshipResource {
  if (!newPart && !ctx.doc.stories.has(owner))
    throw new OrderedRefusal('unknown-story', owner.slice(0, 160));
  if ('resource' in target) {
    if (type === 'hyperlink') throw new OrderedRefusal('invalid-edit', 'a link points outside');
    resourceFor(ctx, target.resource, suggestion);
  } else {
    // Only a link's address points outside: never a picture's, which Word loads when it opens
    // the file (design D18)
    if (type !== 'hyperlink' || !isSafeLinkAddress(target.external)) {
      throw new OrderedRefusal('invalid-edit', 'only a safe link address points outside');
    }
  }

  return addRelationshipResource(ctx, ctx.ids.resource(), suggestion, owner, type, target);
}

/**
 * A relationship resource under `key`, with no checks: for a target the caller added or the
 * validator kept (a picture's image, a link's address, a new notes part).
 */
export function addRelationshipResource(
  ctx: Ctx,
  key: string,
  suggestion: string,
  owner: string,
  type: OrderedRelationshipType,
  target: OrderedRelationshipResource['target']
): OrderedRelationshipResource {
  const resource: OrderedRelationshipResource = {
    kind: 'relationship',
    key,
    suggestion,
    owner,
    id: orderedRelationshipId(key),
    type,
    target,
  };
  addOrderedResource(ctx.doc, resource);
  return resource;
}

const STORY_ROOTS: Readonly<Record<'header' | 'footer', string>> = {
  header: 'hdr',
  footer: 'ftr',
};

const STORY_TYPES: Readonly<Record<'header' | 'footer', string>> = {
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
};

/**
 * An empty story root that binds the namespaces Word's story parts bind, parsed from constants
 * only.
 */
export function emptyStoryRoot(localName: string, name: string, contentType: string): OoxmlPart {
  const read = readOoxmlPart(
    `<w:${localName} xmlns:w="${WML_NAMESPACE_URI}" xmlns:r="${RELATIONSHIPS_NAMESPACE_URI}" ` +
      `xmlns:w14="${W14_NAMESPACE_URI}"/>`,
    { name, contentType }
  );
  if (!read.ok) throw new OrderedRefusal('engine-refused', `a ${localName} root: ${read.reason}`);
  return read.part;
}

/**
 * A new story part's resource, named from a key the change mints, so every copy adds the same
 * part. Refused when a story already has the name.
 */
export function addStoryResource(
  ctx: Ctx,
  suggestion: string,
  role: OrderedStoryRole,
  relationshipsOf?: string
): OrderedStoryResource {
  const key = ctx.ids.resource();
  const name = orderedStoryPart(key, role);
  const resource: OrderedStoryResource = {
    kind: 'story',
    key,
    suggestion,
    part: name,
    role,
    ...(relationshipsOf === undefined ? {} : { relationshipsOf }),
  };
  addOrderedResource(ctx.doc, resource);
  if (ctx.doc.stories.has(name)) throw new OrderedRefusal('name-taken', name);
  return resource;
}

/**
 * A new header or footer story with one empty paragraph, which the suggestion inserted. Its
 * part is written only while a relationship to it is used (a section's reference, group 9.8).
 */
export function addStory(
  ctx: Ctx,
  suggestion: string,
  role: 'header' | 'footer'
): { readonly resource: OrderedStoryResource; readonly story: OrderedStory } {
  const doc = ctx.doc;
  const resource = addStoryResource(ctx, suggestion, role);
  const name = resource.part;
  const minted = mintParagraph(ctx);
  const shell: OoxmlElement = build(
    minted.space,
    'paragraph',
    'p',
    paragraphIdAttributes('w14', minted),
    []
  );
  const mark = createOrderedItem({
    space: minted.space,
    kind: 'mark',
    node: shell,
    by: suggestion,
  });
  const root = emptyStoryRoot(STORY_ROOTS[role], name, STORY_TYPES[role]);
  const story = addOrderedStory(doc, name, root, [mark]);
  recordUndo(doc, () => {
    doc.stories.delete(name);
    labelUnindex(doc, mark);
  });
  return { resource, story };
}

/** A new list definition: its `w:abstractNum` and `w:num`, numbered from the key. */
export function addList(
  ctx: Ctx,
  suggestion: string,
  list: OrderedChangeList
): OrderedListResource {
  const key = ctx.ids.resource();
  const resource: OrderedListResource = {
    kind: 'list',
    key,
    suggestion,
    numId: orderedListNumber(key, 'num'),
    abstractNumId: orderedListNumber(key, 'abstract'),
    levels: list.levels,
  };
  addOrderedResource(ctx.doc, resource);
  return resource;
}

/**
 * A list level a list's definition does not declare, declared at compose in Word's default
 * format.
 */
export function addListLevel(
  ctx: Ctx,
  suggestion: string,
  { numId, level, format }: Pick<OrderedListLevelResource, 'numId' | 'level' | 'format'>
): OrderedListLevelResource {
  const resource: OrderedListLevelResource = {
    kind: 'list-level',
    key: ctx.ids.resource(),
    suggestion,
    numId,
    level,
    format,
  };
  addOrderedResource(ctx.doc, resource);
  return resource;
}

/** A new style, written only when the file has none with its id and something uses it. */
export function addStyle(
  ctx: Ctx,
  suggestion: string,
  style: OrderedChangeStyle
): OrderedStyleResource {
  const resource: OrderedStyleResource = {
    kind: 'style',
    key: ctx.ids.resource(),
    suggestion,
    styleId: style.styleId,
    type: style.type,
    name: style.name,
    ...(style.basedOn === undefined ? {} : { basedOn: style.basedOn }),
    ...(style.next === undefined ? {} : { next: style.next }),
    ...(style.link === undefined ? {} : { link: style.link }),
    ...(style.uiPriority === undefined ? {} : { uiPriority: style.uiPriority }),
    ...(style.unhideWhenUsed === undefined ? {} : { unhideWhenUsed: style.unhideWhenUsed }),
    paragraph: flatProperties(style.paragraph ?? []),
    run: flatProperties(style.run ?? []),
  };
  addOrderedResource(ctx.doc, resource);
  return resource;
}

/** A settings write, replayed on the settings part in the order changes made them. */
export function addSetting(
  ctx: Ctx,
  suggestion: string,
  setting: OrderedChangeSetting
): OrderedSettingResource {
  const resource: OrderedSettingResource = {
    kind: 'setting',
    key: ctx.ids.resource(),
    suggestion,
    op: setting,
  };
  addOrderedResource(ctx.doc, resource);
  return resource;
}

/**
 * A pasted list definition, whole: its `w:abstractNum` and the overrides of its `w:num`,
 * numbered from the key as every new list is, and the picture bullets its levels name.
 */
export function addListDefinition(
  ctx: Ctx,
  suggestion: string,
  definition: { readonly abstract: OoxmlElement; readonly overrides: readonly OoxmlElement[] },
  key: string = ctx.ids.resource(),
  bullets: readonly OrderedListBullet[] = []
): OrderedListResource {
  const resource: OrderedListResource = {
    kind: 'list',
    key,
    suggestion,
    numId: orderedListNumber(key, 'num'),
    abstractNumId: orderedListNumber(key, 'abstract'),
    levels: [],
    definition,
    ...(bullets.length > 0 ? { bullets } : {}),
  };
  addOrderedResource(ctx.doc, resource);
  return resource;
}

/** The `w:val` of a style's child, such as its `w:basedOn` or its `w:name`. */
function childValue(style: OoxmlElement, name: string): string | undefined {
  for (const child of style.children)
    if (
      child.kind !== 'textValue' &&
      child.namespaceUri === WML_NAMESPACE_URI &&
      child.localName === name
    )
      return child.attributes.find((a) => a.localName === 'val')?.value;
  return undefined;
}

/** A pasted style, whole: written only when the file has none with its id and something uses it. */
export function addStyleElement(
  ctx: Ctx,
  suggestion: string,
  element: OoxmlElement
): OrderedStyleResource {
  const resource = styleElementResource(ctx.ids.resource(), suggestion, element);
  if (resource === null) {
    throw new OrderedRefusal('invalid-edit', 'a pasted style needs an id and a known type');
  }

  addOrderedResource(ctx.doc, resource);
  return resource;
}

/** A whole `w:style` as a style resource; null with no id or no known type. */
function styleElementResource(
  key: string,
  suggestion: string,
  element: OoxmlElement
): OrderedStyleResource | null {
  const attribute = (name: string) => element.attributes.find((a) => a.localName === name)?.value;
  const styleId = attribute('styleId');
  const type = attribute('type') ?? 'paragraph';
  if (styleId === undefined || (type !== 'paragraph' && type !== 'character' && type !== 'table')) {
    return null;
  }

  const basedOn = childValue(element, 'basedOn');
  const next = childValue(element, 'next');
  const link = childValue(element, 'link');
  return {
    kind: 'style',
    key,
    suggestion,
    styleId,
    type,
    name: childValue(element, 'name') ?? styleId,
    ...(basedOn === undefined ? {} : { basedOn }),
    ...(next === undefined ? {} : { next }),
    ...(link === undefined ? {} : { link }),
    paragraph: [],
    run: [],
    element,
  };
}
