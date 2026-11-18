// A prompt that a suggestion replaces is kept where Word keeps every prompt: the glossary.

import { buildingBlocksOf, glossaryPartOf } from '../package/building-blocks.ts';
import { orderedContentControlProperties } from '../package/content-control-nodes.ts';
import { relativeTarget } from '../package/custom-xml-part.ts';
import {
  createNodeIdAllocator,
  insertChildren,
  type EditOptions,
  type RetainedPrompt,
} from '../package/ooxml-edit.ts';
import { withPart, type OoxmlPackage } from '../package/ooxml-package.ts';
import {
  relsPartNameFor,
  withNewPart,
  withRelationship,
  withRelationshipsPartFor,
} from '../package/package-edit.ts';
import { fnv1a32 } from '../package/para-id.ts';
import { W14_NAMESPACE_URI, WML_NAMESPACE_URI } from '../package/ooxml-shared.ts';
import {
  readOoxmlPart,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
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
  const filled = { ...docParts, children: prompts.map((p) => placeholderBlock(p, nextId)) };
  const whole = { ...root, children: [filled] } as OoxmlElement;
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
  const written = insertChildren(glossary, docParts.id, docParts.children.length, blocks, {
    deferValidation: true,
  });
  return written.ok ? withPart(pkg, written.part) : pkg;
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
  const next = glossaryPartOf(pkg) ? withBlocksAdded(pkg, pending) : withNewGlossary(pkg, pending);
  for (const [name, part] of next.parts) {
    if (pkg.parts.get(name) !== part) touched.add(name);
  }

  return next;
}
