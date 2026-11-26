// The screen follows the list (design D7).

import {
  canonicalTreeDifference,
  namespaceDeclarationsMatch,
} from '../../../store/package/ooxml-difference.ts';
import { withPart, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import {
  DEFAULT_FINGERPRINT_BINDINGS,
  extendedBindings,
} from '../../../store/package/ooxml-serialize.ts';
import type { OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import {
  withPromptRelationships,
  withRetainedPrompts,
} from '../../../store/store/content-control-prompt-glossary.ts';
import { withPackageShellOf } from '../../../store/store/tree-package-remote.ts';
import { storyPart } from '../compose/compose.ts';
import type { OrderedRetainedPrompt } from '../compose/context.ts';
import {
  orderedCommentRoots,
  withOrderedCommentRoots,
  type OrderedCommentRoots,
} from '../compose/comments.ts';
import { laterRelationshipsOf } from '../compose/relationship-targets.ts';
import {
  isAddedStory,
  mediaBytesOf,
  withOrderedResources,
  type OrderedMediaBytes,
  type OrderedResourceOptions,
} from '../compose/resources.ts';
import type { OrderedDocument, OrderedItemAddress } from '../model/types.ts';
import { withChildren } from '../ooxml-nodes.ts';
import { withFileMediaBytes, withFileMediaPlaces } from '../compose/file-media.ts';
import { withoutDroppedFurniture } from '../compose/furniture.ts';
import { withDecidedPartRevisions } from '../compose/part-revisions.ts';
import { withParagraphsComposed } from './publish-region.ts';
import { withTextBoxHosts } from './publish-text-boxes.ts';

/**
 * What a write changed, for the next publish: stories to compose whole, and paragraphs (by
 * their marks' addresses, shown or not) to compose with the paragraphs around them. `stale`
 * says the editor wrote what the list does not hold (a paragraph under the engine's own id):
 * the next publish goes out even when it names nothing, and takes it back. `comments` says a
 * write changed the comments: their parts are composed again.
 */
export class OrderedDirty {
  readonly stories = new Set<string>();
  readonly marks = new Map<string, Map<string, OrderedItemAddress>>();
  stale = false;
  comments = false;

  get empty(): boolean {
    return this.stories.size === 0 && this.marks.size === 0 && !this.stale && !this.comments;
  }

  story(name: string): void {
    this.stories.add(name);
    this.marks.delete(name);
  }

  mark(story: string, address: OrderedItemAddress): void {
    if (this.stories.has(story)) return;
    let marks = this.marks.get(story);
    if (!marks) {
      marks = new Map();
      this.marks.set(story, marks);
    }

    marks.set(`${address.space}|${address.offset}`, address);
  }

  add(other: OrderedDirty): void {
    for (const name of other.stories) {
      this.story(name);
    }

    for (const [story, marks] of other.marks) {
      for (const address of marks.values()) {
        this.mark(story, address);
      }
    }

    if (other.stale) {
      this.stale = true;
    }

    if (other.comments) {
      this.comments = true;
    }
  }

  clear(): void {
    this.stories.clear();
    this.marks.clear();
    this.stale = false;
    this.comments = false;
  }
}

/**
 * `fresh` with every block of `old` it says the same as kept as `old`'s node: paragraphs by
 * content, containers (a body, a table, a row, a cell) child by child under the same ids. The
 * editor's store compares the package with what it shows object by object, so it sees only what
 * changed.
 * `scope` is the namespace bindings both inherit: a shape's `mc:Choice` in a paragraph names
 * prefixes only the root declares. A container that declares other prefixes keeps nothing.
 */
function adopted(
  old: OoxmlNode,
  fresh: OoxmlNode,
  scope: ReadonlyMap<string, string> = DEFAULT_FINGERPRINT_BINDINGS,
  depth = 0
): OoxmlNode {
  if (old === fresh || old.id !== fresh.id) return fresh;
  if (old.kind === 'textValue' || fresh.kind === 'textValue' || depth > 64) return fresh;
  const bindings = { bindings: scope };
  if (fresh.kind === 'paragraph')
    return canonicalTreeDifference(old, fresh, bindings).equal ? old : fresh;
  if (!namespaceDeclarationsMatch(old, fresh)) return fresh;
  const inner = extendedBindings(scope, old);
  const before = new Map<string, OoxmlNode>();
  for (const child of old.children) before.set(child.id, child);
  let moved = false;
  const children = fresh.children.map((child) => {
    const earlier = before.get(child.id);
    const kept = earlier ? adopted(earlier, child, inner, depth + 1) : child;
    if (kept !== child) moved = true;
    return kept;
  });
  if (!moved) return fresh;
  const next = withChildren(fresh, children);
  return canonicalTreeDifference(old, next, bindings).equal &&
    children.length === old.children.length &&
    children.every((child, k) => child === old.children[k])
    ? old
    : next;
}

/**
 * `fresh` with each part that says what the part of that name in `old` said, under the same ids,
 * kept as `old`'s part. A change's resources are written again at each publish: its list in
 * `numbering.xml`, a style, a setting, a new header or notes part. A new part object with the same
 * content told the editor's store that a part outside the main one changed, so each keystroke
 * laid the whole document out again and built a new numbering index.
 */
function withUnchangedPartsKept(old: OoxmlPackage, fresh: OoxmlPackage): OoxmlPackage {
  let next = withPackageShellOf(fresh, old);
  for (const [name, part] of fresh.parts) {
    const earlier = old.parts.get(name);
    if (!earlier || earlier === part || earlier.id !== part.id) continue;
    if (earlier.contentType !== part.contentType) continue;
    const difference = canonicalTreeDifference(earlier.root, part.root);
    if (difference.equal && difference.idsPreserved) next = withPart(next, earlier);
  }

  return next;
}

export class OrderedPublisher {
  readonly #doc: OrderedDocument;
  readonly #media: OrderedMediaBytes;
  /** The package as the editor was last given it. */
  #pkg: OoxmlPackage;
  /**
   * The same without the resources a change added outside the story text: they are written on
   * it at each publish, so one that nothing uses any more goes (design D14).
   */
  #plain: OoxmlPackage;
  /** The stories a change added, as last composed: written only while something uses them. */
  readonly #added = new Map<string, OoxmlPart>();
  #prompts = new Map<string, OrderedRetainedPrompt>();
  /** The comment parts' roots as last composed: composed again when a write changed comments. */
  #commentRoots: OrderedCommentRoots | null = null;

  /**
   * `resources.blobs` gives the bytes of added media the copy holds, and of the file's media a
   * checkpoint names; media without its bytes is left out until they arrive (`missing: 'skip'`).
   */
  constructor(doc: OrderedDocument, base: OoxmlPackage, resources: OrderedResourceOptions = {}) {
    this.#doc = doc;
    this.#pkg = base;

    // The file's media a checkpoint names keeps its place while its bytes are not read yet
    this.#plain = withFileMediaPlaces(base, doc);
    this.#media = mediaBytesOf({ ...resources, missing: 'skip' });
  }

  /** The package the editor was last given. */
  get current(): OoxmlPackage {
    return this.#pkg;
  }

  /**
   * The whole list as a package: every story composed afresh, keeping no node of the last
   * package (the first publish, and the one after a refused publish).
   */
  full(): OoxmlPackage {
    const dirty = new OrderedDirty();
    for (const name of this.#doc.stories.keys()) dirty.story(name);
    dirty.comments = true;
    return this.next(dirty, false);
  }

  /** The package with the list's version of what `dirty` names. It becomes `current`. */
  next(named: OrderedDirty, keep = true): OoxmlPackage {
    // A text box's story shows in the paragraph that holds its object
    const dirty = withTextBoxHosts(this.#doc, named);
    const regions = this.#composeRegions(this.#plain, dirty);
    const composed = this.#composeStories(regions.pkg, regions.stories, keep);
    this.#pkg = this.#finished(composed, keep, named.comments);
    return this.#pkg;
  }

  /**
   * The paragraphs `dirty` names composed with the paragraphs around them (`publish-region.ts`),
   * and the stories to compose whole: the ones it names, and the ones whose region does not
   * compose alone (a block edge moved: a table, a row, a note, a move).
   */
  #composeRegions(
    plain: OoxmlPackage,
    dirty: OrderedDirty
  ): { readonly pkg: OoxmlPackage; readonly stories: Set<string> } {
    let pkg = plain;
    const stories = new Set(dirty.stories);
    for (const [name, marks] of dirty.marks) {
      if (stories.has(name)) {
        continue;
      }

      const story = this.#doc.stories.get(name);
      const part = pkg.parts.get(name);
      const region =
        story && part ? withParagraphsComposed(this.#doc, story, part, [...marks.values()]) : null;
      if (region === null) {
        stories.add(name);
        continue;
      }

      for (const prompt of region.prompts) {
        this.#prompts.set(prompt.name, prompt);
      }

      if (region.part !== part) {
        pkg = withPart(pkg, region.part);
      }
    }

    return { pkg, stories };
  }

  /**
   * Whole stories composed afresh. With `keep`, each block that says what the editor shows keeps
   * the shown node; a story a change added is kept aside, written only while something uses it.
   */
  #composeStories(
    withRegions: OoxmlPackage,
    stories: ReadonlySet<string>,
    keep: boolean
  ): OoxmlPackage {
    let pkg = withRegions;
    for (const name of stories) {
      // A story a change added and a refusal or a take-back removed again is not written
      if (!this.#doc.stories.has(name)) {
        continue;
      }

      const { part, prompts } = storyPart(this.#doc, name);
      for (const prompt of prompts) {
        this.#prompts.set(prompt.name, prompt);
      }

      if (isAddedStory(this.#doc, name)) {
        this.#added.set(name, part);
        continue;
      }

      const shown = pkg.parts.get(name);
      const kept = shown && keep ? adopted(shown.root, part.root) : part.root;
      if (kept === shown?.root) {
        continue;
      }

      const root = kept.kind === 'textValue' ? part.root : kept;
      pkg = withPart(pkg, root === part.root ? part : { ...part, root });
    }

    return pkg;
  }

  /**
   * The composed package with the prompts, the comments, the decided change records outside the
   * stories, the added resources and the furniture; with `keep`, each part that says the same
   * keeps the editor's part. `comments` composes the comment parts again.
   */
  #finished(composed: OoxmlPackage, keep: boolean, comments: boolean): OoxmlPackage {
    let pkg = composed;
    if (this.#prompts.size > 0) {
      pkg = withRetainedPrompts(
        pkg,
        [...this.#prompts.values()],
        new Set(),
        laterRelationshipsOf(this.#doc)
      );
    }

    this.#plain = pkg;

    // The comment parts and the decided change records outside the stories are written from the
    // file's parts each time; the comment parts are composed again only when comments changed
    if (comments || this.#commentRoots === null) {
      this.#commentRoots = orderedCommentRoots(this.#doc, pkg);
    }

    const decided = withDecidedPartRevisions(
      withOrderedCommentRoots(pkg, this.#commentRoots),
      this.#doc
    );

    // Taken back with its change: a story the list no longer holds is not written
    for (const name of [...this.#added.keys()]) {
      if (!this.#doc.stories.has(name)) {
        this.#added.delete(name);
      }
    }

    const resourced = withoutDroppedFurniture(
      this.#doc.resources.added.size === 0
        ? decided
        : withOrderedResources(decided, this.#doc, this.#added, this.#media),
      this.#doc
    );

    // The glossary's own copy of each relationship a kept prompt names, now that the room's are
    // written
    const published = withFileMediaBytes(
      withPromptRelationships(resourced, [...this.#prompts.values()]),
      this.#doc,
      this.#media
    );
    return keep ? withUnchangedPartsKept(this.#pkg, published) : published;
  }
}
