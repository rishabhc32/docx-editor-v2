// A prompt that a suggestion replaces is kept where Word keeps every prompt: the glossary.

import { buildingBlocksOf, glossaryPartOf } from '../package/building-blocks.ts';
import { orderedContentControlProperties } from '../package/content-control-nodes.ts';
import { relativeTarget } from '../package/custom-xml-part.ts';
import {
  createNodeIdAllocator,
  insertChildren,
  replaceNode,
  type EditOptions,
  type RetainedPrompt,
} from '../package/ooxml-edit.ts';
import {
  recordPutRelationship,
  runWithoutJournalCapture,
} from '../package/canonical-primitive-capture.ts';
import { withPart, type OoxmlPackage } from '../package/ooxml-package.ts';
import { resolveInternalTarget, validateExternalTarget } from '../package/opc-names.ts';
import {
  relationshipsOf,
  relsPartNameFor,
  withNewPart,
  withRelationship,
  withRelationshipsPartFor,
} from '../package/package-edit.ts';
import type { RelationshipRecord } from '../package/relationships.ts';
import { fnv1a32 } from '../package/para-id.ts';
import { W14_NAMESPACE_URI, WML_NAMESPACE_URI } from '../package/ooxml-shared.ts';
import {
  readOoxmlPart,
  type OoxmlElement,
  type OoxmlNamespaceBinding,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { relationshipIdsIn } from './clipboard-fragment-closure.ts';
import { isInlineControl, isWml } from './content-control-checkbox.ts';
import { wmlElement } from './tree-op-content-controls.ts';
import {
  cloneWithNewIds,
  contentControlContentOf,
  contentControlPropertiesOf,
  isShowingPlaceholder,
} from './tree-op-nodes.ts';

const GLOSSARY_PART_NAME = '/word/glossary/document.xml';
const GLOSSARY_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml';
const GLOSSARY_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** Bounds on the walk that names a prompt: its content comes from the file. */
const MAX_DEPTH = 64;
const MAX_NODES = 20_000;

/** The `w:docPart/@w:val` a control names, whether or not the glossary holds that block. */
function referencedDocPart(control: OoxmlNode): string | undefined {
  const sdtPr = contentControlPropertiesOf(control);
  const placeholder = sdtPr?.children.find((child) => isWml(child, 'placeholder'));
  if (!placeholder || placeholder.kind === 'textValue') return undefined;
  const docPart = placeholder.children.find((child) => isWml(child, 'docPart'));
  if (!docPart || docPart.kind === 'textValue') return undefined;
  const name = docPart.attributes.find((attribute) => attribute.localName === 'val')?.value.trim();
  return name && name.length <= 512 ? name : undefined;
}

/**
 * One entry per node, without ids or paragraph ids, every string quoted: equal prompts give
 * equal text, so they share one block. Null for content past the walk's bounds.
 */
function promptDigestSource(nodes: readonly OoxmlNode[]): string | null {
  const out: string[] = [];
  let visited = 0;
  const visit = (node: OoxmlNode, depth: number): boolean => {
    visited += 1;
    if (visited > MAX_NODES || depth > MAX_DEPTH) return false;
    if (node.kind === 'textValue') {
      out.push(JSON.stringify(node.value));
      return true;
    }

    const attributes = node.attributes
      .filter((attribute) => attribute.localName !== 'paraId' && attribute.localName !== 'textId')
      .map((attribute) => [attribute.namespaceUri ?? '', attribute.localName, attribute.value]);
    out.push(`<${JSON.stringify([node.namespaceUri ?? '', node.localName, attributes])}`);
    for (const child of node.children) {
      if (!visit(child, depth + 1)) return false;
    }

    out.push('>');
    return true;
  };
  for (const node of nodes) {
    if (!visit(node, 0)) return null;
  }

  return out.join('');
}

/** A 32-digit name in the shape Word gives its own placeholder blocks. */
function blockNameFor(source: string): string {
  return [0, 1, 2, 3]
    .map((seed) =>
      fnv1a32(`${String(seed)}:${source}`)
        .toString(16)
        .toUpperCase()
        .padStart(8, '0')
    )
    .join('');
}

/**
 * The prompt a suggestion is about to remove from `control`, to keep in the glossary, or null
 * when nothing needs keeping: the caller cannot write the glossary, the control does not show a
 * prompt, or the block the control names already holds it.
 *
 * `w:showingPlcHdr` says a control's content IS its prompt (§17.5.2.39), and `w:placeholder`
 * names the glossary block shown once the control is empty again (§17.5.2.25). Suggesting mode
 * removes a prompt without a revision, as Word does, so a reject empties the control. A prompt
 * with no other copy becomes a block in Word's placeholder gallery, and the control names it, so
 * reject, undo, a save and reopen, and Word show the prompt the control had.
 */
export function promptToRetain(
  part: OoxmlPart,
  control: OoxmlNode,
  options?: EditOptions
): RetainedPrompt | null {
  if (!options?.retainPrompt || !isShowingPlaceholder(control)) return null;
  const named = referencedDocPart(control);
  if (named !== undefined && options.placeholderBlocks?.(named) !== undefined) return null;
  const content = contentControlContentOf(control);
  if (!content) return null;
  const inline = isInlineControl(part, control.id);
  const nodes = content.children.filter((child) => child.kind !== 'textValue');
  if (nodes.length === 0) return null;
  const source = promptDigestSource(nodes);
  if (source === null) return null;
  return {
    name: named ?? blockNameFor(`${inline ? 'inline' : 'block'}\n${source}`),
    inline,
    nodes,
    owner: part.name,
  };
}

/**
 * `control` naming `name` as its placeholder block, in schema order. A control that already
 * names a block keeps that reference: the block is written under the name it gives.
 */
export function withPlaceholderReference(
  control: OoxmlElement,
  name: string,
  nextId: () => string
): OoxmlElement {
  if (referencedDocPart(control) !== undefined) return control;
  const sdtPr = contentControlPropertiesOf(control);
  if (!sdtPr) return control;
  const reference = wmlElement(nextId, 'placeholder', {
    children: [wmlElement(nextId, 'docPart', { attributes: [['val', name]] })],
  });
  const children = orderedContentControlProperties([...sdtPr.children, reference]);
  const properties = { ...sdtPr, children } as OoxmlNode;
  return {
    ...control,
    children: control.children.map((child) => (child.id === sdtPr.id ? properties : child)),
  } as OoxmlElement;
}

/** The blocks of the glossary entry named `name`, or undefined when the package has none. */
export function placeholderBlocksIn(
  pkg: OoxmlPackage,
  name: string
): readonly OoxmlNode[] | undefined {
  return buildingBlocksOf(pkg).find((block) => block.name === name)?.blocks;
}

/**
 * `pkg` with a new glossary document holding `prompts`, related from the main part, or `pkg`
 * itself when one cannot be added. The part is built whole before it joins the package, and
 * nothing is written unless the main part can take the relationship.
 */
function withNewGlossary(pkg: OoxmlPackage, prompts: readonly RetainedPrompt[]): OoxmlPackage {
  const main = pkg.mainDocumentPart;
  const rels = withRelationshipsPartFor(pkg, main);
  if (!rels.parts.has(relsPartNameFor(main))) return pkg;
  const parsed = readOoxmlPart(
    `<w:glossaryDocument xmlns:w="${WML_NAMESPACE_URI}"><w:docParts/></w:glossaryDocument>`,
    { name: GLOSSARY_PART_NAME, contentType: GLOSSARY_CONTENT_TYPE }
  );
  if (!parsed.ok) return pkg;
  const root = parsed.part.root;
  const docParts = root.children.find((child) => isWml(child, 'docParts'));
  if (!docParts || docParts.kind === 'textValue') return pkg;
  const nextId = createNodeIdAllocator(parsed.part);
  const blocks = prompts.map((p) => placeholderBlock(p, nextId));
  const bindings = bindingsFor(root, blocks);
  if (bindings === null) {
    return pkg;
  }

  const namespaceBindings = [...root.namespaceBindings, ...bindings];
  const filled = { ...docParts, children: blocks };
  const whole = { ...root, namespaceBindings, children: [filled] } as OoxmlElement;
  const created = withNewPart(rels, GLOSSARY_PART_NAME, whole, GLOSSARY_CONTENT_TYPE);
  if (!created.parts.has(GLOSSARY_PART_NAME)) return pkg;
  const related = withRelationship(
    created,
    main,
    GLOSSARY_RELATIONSHIP_TYPE,
    relativeTarget(main, GLOSSARY_PART_NAME)
  );
  return related.ok ? related.pkg : pkg;
}

/**
 * A copy of `node` for the glossary, with fresh node ids and without `w14:paraId` and
 * `w14:textId`: those name paragraphs of the story, and a restore mints its own.
 */
function glossaryCopy(node: OoxmlNode, nextId: () => string, depth = 0): OoxmlNode {
  if (node.kind === 'textValue' || depth > MAX_DEPTH) return cloneWithNewIds(node, nextId);
  const attributes = node.attributes.filter(
    (attribute) =>
      attribute.namespaceUri !== W14_NAMESPACE_URI ||
      (attribute.localName !== 'paraId' && attribute.localName !== 'textId')
  );
  return {
    ...node,
    id: nextId(),
    attributes,
    children: node.children.map((child) => glossaryCopy(child, nextId, depth + 1)),
  } as OoxmlNode;
}

/** Word's own form of a placeholder block: the placeholder gallery, inserted as content. */
function placeholderBlock(prompt: RetainedPrompt, nextId: () => string): OoxmlElement {
  const name = prompt.name;
  const hex = /^[0-9A-F]{32}$/.test(name) ? name : blockNameFor(name);
  const guid = `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}}`;
  const copies = prompt.nodes.map((node) => glossaryCopy(node, nextId));
  const body = prompt.inline
    ? [wmlElement(nextId, 'p', { kind: 'paragraph' as OoxmlNode['kind'], children: copies })]
    : copies;
  const value = (localName: string, val: string) =>
    wmlElement(nextId, localName, { attributes: [['val', val]] });
  return wmlElement(nextId, 'docPart', {
    children: [
      wmlElement(nextId, 'docPartPr', {
        children: [
          value('name', name),
          wmlElement(nextId, 'category', {
            children: [value('name', 'General'), value('gallery', 'placeholder')],
          }),
          wmlElement(nextId, 'types', { children: [value('type', 'bbPlcHdr')] }),
          wmlElement(nextId, 'behaviors', { children: [value('behavior', 'content')] }),
          value('guid', guid),
        ],
      }),
      wmlElement(nextId, 'docPartBody', { children: body }),
    ],
  });
}

/** `pkg` with `prompts` appended to the glossary it already has, or `pkg` itself. */
function withBlocksAdded(pkg: OoxmlPackage, prompts: readonly RetainedPrompt[]): OoxmlPackage {
  const glossary = glossaryPartOf(pkg);
  const docParts = glossary?.root.children.find((child) => isWml(child, 'docParts'));
  if (!glossary || !docParts || docParts.kind === 'textValue') return pkg;
  const nextId = createNodeIdAllocator(glossary);
  const blocks = prompts.map((prompt) => placeholderBlock(prompt, nextId));
  const bindings = bindingsFor(glossary.root, blocks);
  if (bindings === null) {
    return pkg;
  }

  // The root binds the blocks' prefixes in the same edit, through the edit primitives
  const options = { deferValidation: true };
  const root = {
    ...glossary.root,
    namespaceBindings: [...glossary.root.namespaceBindings, ...bindings],
  };
  const bound =
    bindings.length === 0
      ? { ok: true as const, part: glossary }
      : replaceNode(glossary, root.id, root, options);
  if (!bound.ok) {
    return pkg;
  }

  const written = insertChildren(
    bound.part,
    docParts.id,
    docParts.children.length,
    blocks,
    options
  );
  return written.ok ? withPart(pkg, written.part) : pkg;
}

/**
 * The bindings that the prefixes of `blocks` need and `root` lacks, or null when `root` binds
 * one of them to another namespace. A block keeps the prefixes of the part that showed it (`r`
 * on a link), and a prefix must be bound where it is used; a prefix a block binds itself is
 * its own.
 */
function bindingsFor(
  root: OoxmlElement,
  blocks: readonly OoxmlNode[]
): OoxmlNamespaceBinding[] | null {
  const bound = new Map(root.namespaceBindings.map((binding) => [binding.prefix, binding]));
  const added: OoxmlNamespaceBinding[] = [];
  let clash = false;
  const need = (prefix: string | undefined, namespaceUri: string, own: ReadonlySet<string>) => {
    if (prefix === undefined || prefix === 'xml' || own.has(prefix)) {
      return;
    }

    const binding = bound.get(prefix);
    if (binding === undefined) {
      const minted = { prefix, namespaceUri };
      bound.set(prefix, minted);
      added.push(minted);
    } else if (binding.namespaceUri !== namespaceUri) {
      clash = true;
    }
  };
  const visit = (node: OoxmlNode, own: ReadonlySet<string>, depth: number): void => {
    if (node.kind === 'textValue' || depth > MAX_DEPTH) {
      return;
    }

    const inScope =
      node.namespaceBindings.length === 0
        ? own
        : new Set([...own, ...node.namespaceBindings.map((binding) => binding.prefix)]);
    need(node.prefix, node.namespaceUri, inScope);
    for (const attribute of node.attributes) {
      need(attribute.prefix, attribute.namespaceUri, inScope);
    }

    for (const child of node.children) {
      visit(child, inScope, depth + 1);
    }
  };
  for (const block of blocks) {
    visit(block, new Set(), 0);
  }

  return clash ? null : added;
}

/**
 * `pkg` with each prompt as a glossary block, creating the glossary when the package has none,
 * and each part that changed named in `touched`. A block of the same name is not written twice.
 * Returns `pkg` itself when the glossary cannot be written, so a decision falls back to the
 * type's own prompt rather than refusing the edit.
 *
 * The op that replaces a prompt only reports the block (`EditOptions.retainPrompt`); the store
 * transaction that ran the op writes it here, in the same package edit, so one undo step and one
 * collaboration journal hold both.
 */
export function withRetainedPrompts(
  pkg: OoxmlPackage,
  prompts: readonly RetainedPrompt[],
  touched: Set<string>
): OoxmlPackage {
  const pending = prompts.filter(
    (prompt, index) =>
      prompts.findIndex((other) => other.name === prompt.name) === index &&
      placeholderBlocksIn(pkg, prompt.name) === undefined
  );
  if (pending.length === 0) return pkg;
  const written = glossaryPartOf(pkg)
    ? withBlocksAdded(pkg, pending)
    : withNewGlossary(pkg, pending);
  const next = withPromptRelationships(written, pending);
  for (const [name, part] of next.parts) {
    if (pkg.parts.get(name) !== part) touched.add(name);
  }

  return next;
}

/**
 * `pkg` with each relationship the kept prompts name copied from the part that showed them to
 * the glossary, under the same id. Ids are per part (OPC), so a link or a picture in a block
 * resolves in the glossary's own relationships, and a restore puts the block back where the id
 * already resolves. An id the glossary already holds, or the part does not, stays as it is. The
 * room calls this again once it has written its own relationships.
 */
export function withPromptRelationships(
  pkg: OoxmlPackage,
  prompts: readonly RetainedPrompt[]
): OoxmlPackage {
  const glossary = glossaryPartOf(pkg);
  if (!glossary) {
    return pkg;
  }

  let next = pkg;
  for (const prompt of prompts) {
    for (const id of relationshipIdsIn(prompt.nodes)) {
      const record = relationshipOf(next, prompt.owner, id);
      if (record && relationshipOf(next, glossary.name, id) === undefined) {
        next = withCopiedRelationship(next, glossary.name, record);
      }
    }
  }

  return next;
}

/** The relationship `id` of `owner`, internal or external, or undefined. */
function relationshipOf(
  pkg: OoxmlPackage,
  owner: string,
  id: string
): RelationshipRecord | undefined {
  const record = relationshipsOf(pkg, owner).find((entry) => entry.id === id);
  if (record) {
    return record;
  }

  const external = pkg.externalTargets.find(
    (entry) => entry.ownerPart === owner && entry.id === id
  );
  return external === undefined
    ? undefined
    : { ...external, targetMode: 'External', order: relationshipsOf(pkg, owner).length };
}

/**
 * `pkg` with `record` added to `owner`'s relationships under its own id, an internal target
 * pointing at the same part from `owner`, or `pkg` itself when it cannot be written.
 */
function withCopiedRelationship(
  pkg: OoxmlPackage,
  owner: string,
  record: RelationshipRecord
): OoxmlPackage {
  const external = record.targetMode === 'External';
  const resolved = external ? null : resolveInternalTarget(record.ownerPart, record.rawTarget);
  if (resolved !== null && !resolved.ok) {
    return pkg;
  }

  const rawTarget = resolved === null ? record.rawTarget : relativeTarget(owner, resolved.partName);
  const prepared = withRelationshipsPartFor(pkg, owner);
  const relsName = relsPartNameFor(owner);
  const rels = prepared.parts.get(relsName);
  if (!rels) {
    return pkg;
  }

  const attribute = (localName: string, value: string) => ({
    kind: 'genericExtension' as const,
    namespaceUri: '',
    localName,
    value,
  });
  const node = {
    id: createNodeIdAllocator(rels)(),
    kind: 'generic',
    namespaceUri: REL,
    localName: 'Relationship',
    namespaceBindings: [],
    attributes: [
      attribute('Id', record.id),
      attribute('Type', record.type),
      attribute('Target', rawTarget),
      ...(external ? [attribute('TargetMode', 'External')] : []),
    ],
    children: [],
  } as unknown as OoxmlElement;
  const inserted = runWithoutJournalCapture(() =>
    insertChildren(rels, rels.root.id, rels.root.children.length, [node])
  );
  if (!inserted.ok) {
    return pkg;
  }

  const owned = relationshipsOf(prepared, owner);
  const copy = { ...record, ownerPart: owner, rawTarget, order: owned.length };
  recordPutRelationship(owner, copy);
  const target = { ownerPart: owner, id: record.id, type: record.type, rawTarget };
  return Object.freeze({
    ...withPart(prepared, inserted.part),
    relationships: new Map([...prepared.relationships, [owner, [...owned, copy]]]),
    externalTargets: external
      ? [...prepared.externalTargets, { ...target, sinkSafe: validateExternalTarget(rawTarget).ok }]
      : prepared.externalTargets,
  });
}
