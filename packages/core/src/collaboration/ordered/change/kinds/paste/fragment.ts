// The `insert-fragment` change kind (group 9.10): a paste with formatting.

import { readOoxmlPart, WML_NAMESPACE_URI } from '../../../../../store/package/ooxml-tree.ts';
import { sanitizeFragmentBlocks } from '../../../../../store/store/clipboard-fragment-sanitize.ts';
import { REVISION_MARKUP } from '../../../../../store/store/tracked-fragment.ts';
import type { OoxmlElement } from '../../../../../store/package/ooxml-tree.ts';
import { orderedViewParagraph } from '../../../compose/compose.ts';
import { contextOf, suggestionFor } from '../../../edits/context.ts';
import { settle } from '../../../edits/direct.ts';
import {
  fragmentNodes,
  fragmentPartNode,
  fragmentResourceNodes,
  type OrderedFragmentResolver,
} from '../../../edits/fragment-nodes.ts';
import { declarePrefixes } from '../../../edits/namespaces.ts';
import { addNotesStory, noteIdsFor, notesStoryOf } from '../../../edits/notes.ts';
import { pasteFragment } from '../../../edits/paste/fragment.ts';
import {
  addListDefinition,
  addMedia,
  addRelationship,
  addStyleElement,
} from '../../../edits/resources.ts';
import { templateItems } from '../../../edits/templates.ts';
import { insertItems } from '../../../model/mutations.ts';
import { OrderedRefusal } from '../../../model/refusal.ts';
import {
  addOrderedResource,
  isOrderedMediaFor,
  ORDERED_MEDIA_TYPES,
  ORDERED_OLE_OBJECT_TYPE,
  ORDERED_WORKBOOK_TYPE,
  orderedListNumber,
  orderedListOf,
  orderedObjectPart,
  orderedRelationshipId,
  type OrderedListBullet,
  type OrderedMediaUse,
  type OrderedRelationshipType,
} from '../../../model/resources.ts';
import type { OrderedItem } from '../../../model/types.ts';
import { withChildren } from '../../../ooxml-nodes.ts';
import type { Ctx } from '../../../edits/context.ts';
import {
  fragmentBudget,
  fragmentElement,
  fragmentElements,
  fragmentParts,
  ORDERED_FRAGMENT_LIMITS as LIMITS,
  type OrderedFragmentBudget,
  type OrderedFragmentElement,
  type OrderedFragmentList,
  type OrderedFragmentNote,
  type OrderedFragmentPart,
  type OrderedFragmentReferences,
} from '../../fragment.ts';
import {
  FRAGMENT_PART_CONTENT_TYPES,
  type OrderedFragmentObjectUse,
  type OrderedFragmentPartType,
} from '../../fragment-names-objects.ts';
import { resolveChangePlace } from '../../change-place.ts';
import { at, bool, fieldPath, type Fields, fieldsOf, list, oneOf, required } from '../../read.ts';
import { changeMedia, type OrderedChangeMedia } from '../../resources.ts';
import type { OrderedChangePlace, OrderedEditBase } from '../../types.ts';
import { href, place } from '../../values.ts';
import { orderedKindSpec } from '../spec.ts';

/**
 * Paste blocks with their formatting before a place (from Word, a browser or the editor's own
 * clipboard): paragraphs, runs, lists, tables, links, fields, controls, bookmarks, equations
 * and notes, as the engine's `insertFragment` lands them. `whole` says the fragment's last
 * paragraph mark travelled (its last paragraph stays a paragraph of its own). A `w:numId` of
 * `#<k>` names `lists[k]`, a link's `r:id` of `#<k>` names `links[k]`, a picture's `r:embed` of
 * `#<k>` names `media[k]`, a note reference's `w:id` of `#<k>` names `notes[k]`; `styles` are
 * the styles the blocks name that the document does not have. A chart's `r:id` and a
 * diagram's `r:dm`, `r:lo`, `r:qs` and `r:cs` of `#<k>` name `parts[k]`, a part of that type.
 *
 * The resources are the paste suggestion's own (design D14), so a reject or a take-back takes
 * them out of the file with it; media travel by hash. Word writes a tracked paste as an
 * insertion with plain properties and records no change of the numbering or styles parts
 * (ECMA-376 Part 1, §17.13.5), which is what compose writes.
 *
 * @public
 */
export interface OrderedInsertFragmentChange extends OrderedEditBase {
  readonly kind: 'insert-fragment';
  readonly place: OrderedChangePlace;
  readonly blocks: readonly OrderedFragmentElement[];
  readonly whole?: boolean;
  readonly lists?: readonly OrderedFragmentList[];
  readonly styles?: readonly OrderedFragmentElement[];
  readonly links?: readonly string[];
  readonly media?: readonly OrderedChangeMedia[];
  readonly notes?: readonly OrderedFragmentNote[];
  readonly parts?: readonly OrderedFragmentPart[];
  /**
   * The picture bullets a list level names (`w:lvlPicBulletId` of `#<j>`): each a
   * `w:numPicBullet` without its id, its image (`v:imagedata/@r:id`, `a:blip/@r:embed`) of `#<k>`
   * into `media`.
   */
  readonly bullets?: readonly OrderedFragmentElement[];
}

/** A list of the change's resources: empty when absent, within the resource cap. */
const resourceList = (fields: Fields, key: string) =>
  fields.has(key) ? list(fields.get(key), fieldPath(fields, key), 1, LIMITS.resources) : [];

/** A paste's list definitions: each a `w:abstractNum` and the overrides of its `w:num`. */
function readLists(
  lists: readonly unknown[],
  path: string,
  references: OrderedFragmentReferences,
  budget: OrderedFragmentBudget
): OrderedFragmentList[] {
  return lists.map((entry, k): OrderedFragmentList => {
    const where = at(path, k);
    const own = fieldsOf(entry, where, new Set(['abstract', 'overrides']));
    const abstract = fragmentElement(
      required(own, 'abstract'),
      at(where, 'abstract'),
      'abstractNum',
      references,
      budget
    );
    if (!own.has('overrides')) {
      return { abstract };
    }

    const overrides = fragmentElements(
      own.get('overrides'),
      at(where, 'overrides'),
      'lvlOverride',
      references,
      budget
    );
    return { abstract, overrides };
  });
}

/** A paste's notes: each its kind and its blocks. */
function readNotes(
  notes: readonly unknown[],
  path: string,
  references: OrderedFragmentReferences,
  budget: OrderedFragmentBudget
): OrderedFragmentNote[] {
  return notes.map((entry, k): OrderedFragmentNote => {
    const where = at(path, k);
    const own = fieldsOf(entry, where, new Set(['kind', 'blocks']));
    return {
      kind: oneOf(required(own, 'kind'), at(where, 'kind'), ['footnote', 'endnote'] as const),
      blocks: fragmentElements(
        required(own, 'blocks'),
        at(where, 'blocks'),
        'block',
        references,
        budget
      ),
    };
  });
}

export const INSERT_FRAGMENT_SPEC = orderedKindSpec<OrderedInsertFragmentChange>({
  kind: 'insert-fragment',
  fields: [
    'place',
    'blocks',
    'whole',
    'lists',
    'styles',
    'links',
    'media',
    'notes',
    'parts',
    'bullets',
  ],
  read(fields) {
    // Every path starts at the change, so a refusal names its place in the message
    const lists = resourceList(fields, 'lists');
    const links = resourceList(fields, 'links');
    const media = resourceList(fields, 'media');
    const notes = resourceList(fields, 'notes');
    const parts = resourceList(fields, 'parts');
    const bullets = resourceList(fields, 'bullets');
    const references: OrderedFragmentReferences = {
      lists: lists.length,
      links: links.length,
      notes: notes.length,
      media: media.length,
      parts: parts.length,
      bullets: bullets.length,
    };
    const budget = fragmentBudget();
    const out: Record<string, unknown> = {
      place: place(required(fields, 'place'), fieldPath(fields, 'place')),
      blocks: fragmentElements(
        required(fields, 'blocks'),
        fieldPath(fields, 'blocks'),
        'block',
        references,
        budget
      ),
    };
    if (fields.has('whole')) out.whole = bool(fields.get('whole'), fieldPath(fields, 'whole'));
    if (bullets.length > 0)
      out.bullets = fragmentElements(
        bullets,
        fieldPath(fields, 'bullets'),
        'numPicBullet',
        references,
        budget
      );
    if (lists.length > 0)
      out.lists = readLists(lists, fieldPath(fields, 'lists'), references, budget);
    if (fields.has('styles'))
      out.styles = fragmentElements(
        fields.get('styles'),
        fieldPath(fields, 'styles'),
        'style',
        references,
        budget
      );
    if (links.length > 0)
      out.links = links.map((link, k) => href(link, at(fieldPath(fields, 'links'), k)));
    if (media.length > 0)
      out.media = media.map((entry, k) =>
        changeMedia(entry, at(fieldPath(fields, 'media'), k), ORDERED_MEDIA_TYPES)
      );
    if (parts.length > 0)
      out.parts = fragmentParts(parts, fieldPath(fields, 'parts'), references, budget);
    if (notes.length > 0)
      out.notes = readNotes(notes, fieldPath(fields, 'notes'), references, budget);
    return out;
  },
  touches(change, reader) {
    reader.caret(change.place);
  },
  apply(change, { story, edit }) {
    refuseWrongNoteKinds(change);
    refuseTrackedChanges(change);
    const ctx = contextOf(edit, []);
    const sid = suggestionFor(ctx).id;
    const resolve = resolver(ctx, sid, change);
    addListsAndStyles(ctx, sid, change, resolve);

    // The note bodies first: their references in the blocks name the ids they took
    addNotes(ctx, sid, change.notes ?? [], resolve);

    // The clipboard's own trust rules again: dangerous field verbs unlinked to their results,
    // no section break, no comment markers (`clipboard-fragment-sanitize.ts`)
    const blocks = sanitized(fragmentNodes(change.blocks, resolve));
    declarePrefixes(ctx.doc, story, blocks);
    const at = resolveChangePlace(ctx.doc, story.name, change.place);
    const p = orderedViewParagraph(ctx.doc, story.name, at.index);
    if (!p) throw new OrderedRefusal('malformed-story', `no paragraph ${at.index}`);
    pasteFragment(ctx, p, at.offset, blocks, change.whole === true, sid);
    if (ctx.direct) {
      settle(ctx, sid);
    }

    return { suggestions: [sid], replaced: [] };
  },
  policy: { text: true, properties: false },
  // A paste can hold tables and block controls: a block edge moves
  publish: 'story',
  // The pasted content goes before the place's unit, so the caret stays before it: after it
  caret: (change) => change.place,
  types: true,
  blobs: (change) => (change.media ?? []).map((entry) => entry.sha256),
});

/** The `k` of a `#<k>` reference; -1 when the value is not one. */
function referenceIndex(value: string): number {
  const found = /^#(\d{1,4})$/.exec(value);
  return found === null ? -1 : Number(found[1]);
}

/** The value a `#<k>` reference names in `values`. */
function byIndex(values: readonly string[], value: string, what: string): string {
  const name = values[referenceIndex(value)];
  if (name === undefined) {
    throw new OrderedRefusal('unknown-address', `no ${what} ${value}`);
  }

  return name;
}

/** The relationship type of an embedded object's part by its use and its media's type. */
function objectRelationshipType(
  use: OrderedFragmentObjectUse,
  type: string | undefined
): 'ink' | 'oleObject' | 'package' {
  if (use === 'ink') {
    return 'ink';
  }

  if (type === ORDERED_OLE_OBJECT_TYPE) {
    return 'oleObject';
  }

  return 'package';
}

/** A media resource's key by a `#<k>` reference, made when something first names it. */
type MediaKey = (value: string, use?: OrderedMediaUse) => string;

/**
 * How the blocks' `#<k>` references name the resources a paste adds, as the suggestion's own.
 * A decimal list number must be a list the document defines, or one its author added (pending)
 * or that is accepted. A picture's image relationship belongs to the story that holds the
 * picture (`owner`: the blocks', or a pasted note's).
 */
function resolver(ctx: Ctx, sid: string, change: OrderedInsertFragmentChange): PasteResolver {
  // A media resource, made once, when something first names it: a picture's image, a chart's
  // embedded workbook, an embedded object's part or an ink part, each only of its own types
  const media = new Map<number, string>();
  const mediaKey: MediaKey = (value, use = 'picture') => {
    const k = referenceIndex(value);
    const entry = change.media?.[k];
    if (entry === undefined) {
      throw new OrderedRefusal('unknown-address', `no media ${value}`);
    }

    if (!isOrderedMediaFor(entry.contentType, use)) {
      throw new OrderedRefusal('invalid-edit', `media ${value} is not of a ${use}`);
    }

    let key = media.get(k);
    if (key === undefined) {
      key = addMedia(ctx, sid, entry).key;
      media.set(k, key);
    }

    return key;
  };

  // A relationship belongs to the story whose content names it: one per story, made once
  const owned = new Map<string, string>();
  const ownRelationship = (
    at: string,
    type: OrderedRelationshipType,
    target: { readonly external: string } | { readonly resource: string }
  ): string => {
    const key = `${self.owner}|${at}`;
    let id = owned.get(key);
    if (id === undefined) {
      id = addRelationship(ctx, sid, type, target, self.owner).id;
      owned.set(key, id);
    }

    return id;
  };

  // A diagram's drawing belongs to the story that holds the diagram, as Word writes it
  const objectPart = objectParts(ctx, sid, change, mediaKey, (value) =>
    self.part(value, 'diagramDrawing')
  );
  const self: PasteResolver = {
    notes: new Map<string, string>(),
    numbers: [],
    listBullet: null,
    owner: ctx.story.name,
    mediaKey,
    bullet(value) {
      if (self.listBullet === null) {
        throw new OrderedRefusal('invalid-edit', 'a picture bullet outside a list');
      }

      return self.listBullet(value);
    },
    numId(value) {
      if (value.startsWith('#')) {
        return byIndex(self.numbers, value, 'list');
      }

      const known = ctx.doc.numbering;
      if (known !== null && !known.has(value) && !orderedListOf(ctx.doc, value, ctx.attribution)) {
        throw new OrderedRefusal('not-covered', `a list number the document does not define`);
      }

      return value;
    },
    relationship(value) {
      const target = byIndex(change.links ?? [], value, 'link');
      return ownRelationship(`link${value}`, 'hyperlink', { external: target });
    },
    media: (value) => ownRelationship(`media${value}`, 'image', { resource: mediaKey(value) }),
    // An embedded object's part (an OLE part or an Office file) or an ink part, by its type
    object(value, use) {
      const resource = mediaKey(value, use);
      const type = change.media?.[referenceIndex(value)]?.contentType;
      return ownRelationship(`media${value}`, objectRelationshipType(use, type), { resource });
    },
    part(value, type) {
      return ownRelationship(`part${value}`, type, {
        resource: objectPart(referenceIndex(value), type),
      });
    },
    // A drawing's id is document-wide: a decimal from the change, as a new picture's
    drawingId: () => String(ctx.ids.decimal()),
    note(value) {
      const id = self.notes.get(value);
      if (id === undefined) {
        throw new OrderedRefusal('unknown-address', `no note ${value}`);
      }

      return id;
    },
  };
  return self;
}

/**
 * The maker of a paste's chart and diagram parts, each made once, when something first names
 * it: its own relationships first (to later parts, to media), then the part with their minted
 * ids. `drawing` names a diagram's drawing.
 */
function objectParts(
  ctx: Ctx,
  sid: string,
  change: OrderedInsertFragmentChange,
  mediaKey: MediaKey,
  drawing: (value: string) => string
): (k: number, type: OrderedFragmentPartType) => string {
  const keys = new Map<number, string>();
  const objectPart = (k: number, type: OrderedFragmentPartType): string => {
    const entry = change.parts?.[k];
    if (entry === undefined || entry.type !== type) {
      throw new OrderedRefusal('unknown-address', `no ${type} part #${k}`);
    }

    const known = keys.get(k);
    if (known !== undefined) {
      return known;
    }

    const key = ctx.ids.resource();
    keys.set(k, key);
    const name = orderedObjectPart(key, type);
    const local = (entry.rels ?? []).map((rel) => {
      if ('part' in rel) {
        const child = change.parts?.[rel.part];
        if (child === undefined) {
          throw new OrderedRefusal('unknown-address', `no part #${rel.part}`);
        }

        const target = { resource: objectPart(rel.part, child.type) };
        return addRelationship(ctx, sid, child.type, target, name, true).id;
      }

      const workbook = change.media?.[rel.media]?.contentType === ORDERED_WORKBOOK_TYPE;
      if (workbook && type !== 'chart' && type !== 'chartEx') {
        throw new OrderedRefusal('invalid-edit', `a workbook in a ${type} part`);
      }

      const target = { resource: mediaKey(`#${rel.media}`, workbook ? 'workbook' : 'picture') };
      return addRelationship(ctx, sid, workbook ? 'package' : 'image', target, name, true).id;
    });
    const root = fragmentPartNode(entry.root, type, name, FRAGMENT_PART_CONTENT_TYPES[type], {
      relationship: (value) => byIndex(local, value, 'relationship'),
      drawing,
    });
    addOrderedResource(ctx.doc, { kind: 'part', key, suggestion: sid, part: name, type, root });
    return key;
  };
  return objectPart;
}

/**
 * The paste's lists, then its styles, as the suggestion's own: styles may name the lists, and a
 * list never names a new style's list. A list's picture bullets are its own, numbered from its
 * key, made when a level names one.
 */
function addListsAndStyles(
  ctx: Ctx,
  sid: string,
  change: OrderedInsertFragmentChange,
  resolve: PasteResolver
): void {
  for (const entry of change.lists ?? []) {
    const key = ctx.ids.resource();
    const bullets = new Map<number, OrderedListBullet>();
    resolve.listBullet = (value) => {
      const j = referenceIndex(value);
      let bullet = bullets.get(j);
      if (bullet === undefined) {
        const id = orderedListNumber(key, 'bullet', j);
        bullet = listBullet(ctx, change, j, id, resolve.mediaKey);
        bullets.set(j, bullet);
      }

      return bullet.id;
    };
    const [abstract] = fragmentResourceNodes([entry.abstract], 'abstractNum', resolve);
    if (!abstract) {
      throw new OrderedRefusal('engine-refused', 'a list without its definition');
    }

    const overrides = fragmentResourceNodes(entry.overrides ?? [], 'lvlOverride', resolve);
    resolve.listBullet = null;
    const definition = { abstract, overrides };
    resolve.numbers.push(addListDefinition(ctx, sid, definition, key, [...bullets.values()]).numId);
  }

  for (const style of fragmentResourceNodes(change.styles ?? [], 'style', resolve)) {
    addStyleElement(ctx, sid, style);
  }
}

/**
 * A list's picture bullet `#<j>` of the change, under the id `id`: its images are relationships
 * of the numbering part, which compose writes with the list (design D18).
 */
function listBullet(
  ctx: Ctx,
  change: OrderedInsertFragmentChange,
  j: number,
  id: string,
  mediaKey: (value: string) => string
): OrderedListBullet {
  const tree = change.bullets?.[j];
  if (tree === undefined) throw new OrderedRefusal('unknown-address', `no picture bullet #${j}`);
  const images: { id: string; media: string }[] = [];
  const none = (what: string) => (): never => {
    throw new OrderedRefusal('invalid-edit', `a picture bullet names ${what}`);
  };
  const [element] = fragmentResourceNodes([tree], 'numPicBullet', {
    numId: none('a list'),
    relationship: none('a link'),
    note: none('a note'),
    part: none('a part'),
    object: none('an object'),
    media(value) {
      const image = { id: orderedRelationshipId(ctx.ids.resource()), media: mediaKey(value) };
      images.push(image);
      return image.id;
    },
    drawingId: () => String(ctx.ids.decimal()),
  });
  if (!element) throw new OrderedRefusal('engine-refused', 'a picture bullet');
  return { id, element, images };
}

/** A paste's resolver, with the ids it minted and the story that holds what it writes. */
interface PasteResolver extends OrderedFragmentResolver {
  /** The note ids the pasted notes took, by `#<k>`. */
  readonly notes: Map<string, string>;
  /** The list numbers the pasted lists took, by index. */
  readonly numbers: string[];
  /** The picture bullet of the list being added, by `#<j>`; null outside a list. */
  listBullet: ((value: string) => string) | null;
  owner: string;
  readonly mediaKey: MediaKey;
}

const W = WML_NAMESPACE_URI;

/** Engine nodes through the clipboard's sanitizer, elements only. */
function sanitized(nodes: readonly OoxmlElement[]): OoxmlElement[] {
  const out: OoxmlElement[] = [];
  for (const node of sanitizeFragmentBlocks(nodes)) if (node.kind !== 'textValue') out.push(node);
  return out;
}

/** Visit every element of the trees, each once, in no set order. */
function eachElement(
  roots: readonly OrderedFragmentElement[],
  visit: (node: OrderedFragmentElement) => void
): void {
  const stack = [...roots];
  for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
    visit(node);
    for (const child of node.c ?? []) {
      if (typeof child !== 'string') {
        stack.push(child);
      }
    }
  }
}

/** The kind of note a note reference element names. */
const NOTE_KIND_OF: Readonly<Record<string, 'footnote' | 'endnote'>> = {
  'w:footnoteReference': 'footnote',
  'w:endnoteReference': 'endnote',
};

/** A note reference names a note of its own kind: a footnote's a footnote. */
function refuseWrongNoteKinds(change: OrderedInsertFragmentChange): void {
  const notes = change.notes ?? [];
  eachElement([...change.blocks, ...notes.flatMap((n) => n.blocks)], (node) => {
    const kind = Object.hasOwn(NOTE_KIND_OF, node.e) ? NOTE_KIND_OF[node.e] : undefined;
    const id = node.a?.['w:id'];
    if (kind === undefined || id === undefined) {
      return;
    }

    const named = notes[referenceIndex(id)];
    if (named === undefined || named.kind !== kind) {
      throw new OrderedRefusal('invalid-edit', `a ${kind} reference names another kind of note`);
    }
  });
}

/**
 * A paste holds no tracked change. The server stamps every change with its sender, and a pasted
 * record would bring another person's name and date into the room: the client sends the
 * fragment's accepted final view (design D18).
 */
function refuseTrackedChanges(change: OrderedInsertFragmentChange): void {
  const roots = [
    ...change.blocks,
    ...(change.notes ?? []).flatMap((note) => note.blocks),
    ...(change.styles ?? []),
    ...(change.lists ?? []).flatMap((list) => [list.abstract, ...(list.overrides ?? [])]),
    ...(change.bullets ?? []),
  ];
  eachElement(roots, (node) => {
    if (node.e.startsWith('w:') && REVISION_MARKUP.has(node.e.slice(2))) {
      throw new OrderedRefusal('invalid-edit', `a paste holds a tracked change (${node.e})`);
    }
  });
}

/** An empty note element of a kind and id, parsed from engine literals. */
function noteShell(kind: 'footnote' | 'endnote', id: number): OoxmlElement {
  const root = kind === 'footnote' ? 'footnotes' : 'endnotes';
  const read = readOoxmlPart(`<w:${root} xmlns:w="${W}"><w:${kind} w:id="${id}"/></w:${root}>`, {
    name: '/word/ordered-note.xml',
    contentType: 'application/xml',
  });
  const first = read.ok ? read.part.root.children.find((c) => c.kind !== 'textValue') : undefined;
  if (!first || first.kind === 'textValue')
    throw new OrderedRefusal('engine-refused', 'a note template did not parse');
  return first;
}

/**
 * The pasted notes, at the end of their notes story, in the paste's suggestion: covered by
 * their references' marker, as a new note is (`insertNote`).
 */
function addNotes(
  ctx: Ctx,
  sid: string,
  notes: readonly OrderedFragmentNote[],
  resolve: PasteResolver
): void {
  const ids = new Map<string, () => number>();
  notes.forEach((note, k) => {
    // A file with no notes part of the kind gets one, as the engine's paste makes it
    const story = notesStoryOf(ctx.doc, note.kind) ?? addNotesStory(ctx, note.kind, sid);
    const next = ids.get(story.name) ?? noteIdsFor(story);
    ids.set(story.name, next);
    const id = next();
    resolve.notes.set(`#${k}`, String(id));
    const shell = noteShell(note.kind, id);

    // A picture in the note is the notes part's: its image relationship is owned there
    resolve.owner = story.name;
    const element = withChildren(shell, sanitized(fragmentNodes(note.blocks, resolve)));
    resolve.owner = ctx.story.name;
    declarePrefixes(ctx.doc, story, [element]);
    const items = templateItems({ ...ctx, story }, element, sid).map(
      (item): OrderedItem =>
        item.kind === 'open' || item.kind === 'close'
          ? { ...item, by: sid }
          : { ...item, by: sid, quiet: true }
    );
    insertItems(ctx.doc, story, story.tree.last(), items);
  });
}
