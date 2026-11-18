// Repeating section items (`w15:repeatingSection`, MS-DOCX 2.6.x): add an item, remove one.

import {
  createNodeIdAllocator,
  findNode,
  replaceNode,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { mintBuildingBlockIdentities } from './building-block-identities.ts';
import { W15_NAMESPACE_URI } from './content-control-prompt-kinds.ts';
import { withRestoredPrompt } from './content-control-prompt-restore.ts';
import { withParagraphDiff } from './content-control-value-content.ts';
import { contentControlEffect } from './tree-op-content-controls.ts';
import {
  cloneWithNewIds,
  contentControlContentOf,
  contentControlPropertiesOf,
  effectiveLockOf,
  findContentControl,
  fromEdit,
  isBoundContentControl,
  isParagraphPropertiesNode,
  isRunPropertiesNode,
} from './tree-op-nodes.ts';
import type { TreeDocOp, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';

type ItemOp = Extract<
  TreeDocOp,
  { readonly op: 'addRepeatingSectionItem' | 'removeRepeatingSectionItem' }
>;

const isW15 = (node: OoxmlNode, localName: string): node is OoxmlElement =>
  node.kind !== 'textValue' &&
  node.namespaceUri === W15_NAMESPACE_URI &&
  node.localName === localName;

const isControl = (node: OoxmlNode): node is OoxmlElement =>
  node.kind === 'contentControl' ||
  (node.kind !== 'textValue' &&
    node.namespaceUri === WML_NAMESPACE_URI &&
    node.localName === 'sdt');

/** The `w15:` child of an element, if any. */
function w15Child(element: OoxmlElement | undefined, localName: string): OoxmlElement | undefined {
  for (const child of element?.children ?? []) {
    if (isW15(child, localName)) return child;
  }

  return undefined;
}

/** The `w15:` element a control's properties hold, if any. */
const w15Property = (control: OoxmlElement, localName: string): OoxmlElement | undefined =>
  w15Child(contentControlPropertiesOf(control), localName);

/** Whether an on/off `w15:` element is on (no `w15:val`, or a true one). */
function isOn(element: OoxmlElement | undefined): boolean {
  if (!element) return false;
  const value = element.attributes.find((a) => a.localName === 'val')?.value;
  return value === undefined || !(value === '0' || value === 'false' || value === 'off');
}

/** Whether a control is a repeating section (`w15:repeatingSection` in its properties). */
export function isRepeatingSection(control: OoxmlNode): boolean {
  return isControl(control) && w15Property(control, 'repeatingSection') !== undefined;
}

/**
 * A section's items, in order: its content's controls marked `w15:repeatingSectionItem`. A
 * repeating section holds only such controls (MS-DOCX).
 */
function itemsOf(section: OoxmlElement): OoxmlElement[] {
  const children = contentControlContentOf(section)?.children ?? [];
  return children.filter(
    (child): child is OoxmlElement =>
      isControl(child) && w15Property(child, 'repeatingSectionItem') !== undefined
  );
}

/**
 * The repeating section an op names, or the engine's refusal. The section's
 * `w15:doNotAllowInsertDeleteSection` turns both item ops off.
 */
function sectionFor(part: OoxmlPart, op: ItemOp): OoxmlElement | TreeOpRejection {
  if (typeof op.controlId !== 'string') return 'unknown-control';
  const section = findContentControl(part, op.controlId);
  if (!section) return 'unknown-control';
  const marker = w15Property(section, 'repeatingSection');
  if (!marker) return 'typeMismatch';
  if (isOn(w15Child(marker, 'doNotAllowInsertDeleteSection'))) return 'locked';
  if (isBoundContentControl(section)) return 'bound';
  if (effectiveLockOf(part, section).content) return 'locked';
  return section;
}

/** The engine's refusal of an item op, or null. */
export function validateRepeatingSectionItem(part: OoxmlPart, op: ItemOp): TreeOpRejection | null {
  const section = sectionFor(part, op);
  if (typeof section === 'string') return section;
  const items = itemsOf(section);
  if (op.op === 'addRepeatingSectionItem') {
    if (items.length === 0) return 'unsupported';
    const index = op.index ?? items.length;
    return Number.isInteger(index) && index >= 0 && index <= items.length ? null : 'invalidArgs';
  }

  if (!Number.isInteger(op.index) || op.index < 0 || op.index >= items.length) return 'invalidArgs';

  // A section keeps one item: Word removes the last one only with the section
  return items.length > 1 ? null : 'invalidArgs';
}

/** The first run under a node, depth first. */
function firstRun(node: OoxmlNode, depth = 0): OoxmlElement | undefined {
  if (node.kind === 'textValue' || depth > 64) return undefined;
  for (const child of node.children) {
    if (child.kind === 'run') return child;
    const found = firstRun(child, depth + 1);
    if (found) return found;
  }

  return undefined;
}

/**
 * A control with nothing in it: its first run kept with only its face, in its first paragraph
 * (with that paragraph's properties) for a block control, so its prompt comes back in them.
 */
function emptied(control: OoxmlElement): OoxmlElement {
  const content = contentControlContentOf(control);
  if (!content) return control;
  const run = firstRun(content);
  const face = run?.children.filter((child) => isRunPropertiesNode(child)) ?? [];
  const bare = run ? [{ ...run, children: face } as OoxmlNode] : [];
  const paragraph = content.children.find((child) => child.kind === 'paragraph');
  const children: OoxmlNode[] =
    paragraph && paragraph.kind !== 'textValue'
      ? [
          {
            ...paragraph,
            children: [...paragraph.children.filter((c) => isParagraphPropertiesNode(c)), ...bare],
          } as OoxmlNode,
        ]
      : bare;
  return {
    ...control,
    children: control.children.map((child) =>
      child === content ? ({ ...content, children } as OoxmlNode) : child
    ),
  } as OoxmlElement;
}

/** A new item with the controls to reset emptied, and their ids. */
function withControlsEmptied(item: OoxmlElement): { item: OoxmlElement; ids: string[] } {
  const ids: string[] = [];
  const visit = (node: OoxmlNode, depth: number): OoxmlNode => {
    if (node.kind === 'textValue' || depth > 64) return node;
    const children = node.children.map((child) => {
      if (!isControl(child)) return visit(child, depth + 1);
      if (isBoundContentControl(child) || w15Property(child, 'repeatingSection')) return child;
      ids.push(child.id);
      return emptied(child);
    });
    return { ...node, children } as OoxmlNode;
  };
  const visited = visit(item, 0);
  return { item: visited.kind === 'textValue' ? item : visited, ids };
}

/** Each emptied control shows its prompt again, the glossary's own when it has one. */
function withPromptsShown(
  part: OoxmlPart,
  ids: readonly string[],
  options?: EditOptions
): OoxmlPart {
  let current = part;
  for (const id of ids) {
    current = withRestoredPrompt(current, id, options);
  }

  return current;
}

/** A section content's new children, and the ids of the controls whose prompts come back. */
interface ItemUpdate {
  readonly children: OoxmlNode[];
  readonly resetIds: string[];
}

/** The section's content without the item at `index`. */
function withItemRemoved(
  content: OoxmlElement,
  items: readonly OoxmlElement[],
  index: number
): ItemUpdate {
  const gone = items[index]?.id;
  return { children: content.children.filter((child) => child.id !== gone), resetIds: [] };
}

/**
 * The section's content with a new item at `index` that repeats the one before it: its words
 * are copied, and the controls in it are reset to their prompts, except a nested repeating
 * section and a control mapped to custom XML ("Content controls in Word", Microsoft Learn).
 */
function withItemAdded(
  part: OoxmlPart,
  content: OoxmlElement,
  items: readonly OoxmlElement[],
  index: number
): ItemUpdate | TreeOpRejection {
  const source = items[Math.max(0, index - 1)];
  if (!source) return 'unsupported';
  const fresh = mintBuildingBlockIdentities(part)(
    cloneWithNewIds(source, createNodeIdAllocator(part, 'paste'))
  );
  if (!fresh || fresh.kind === 'textValue') {
    return 'fragment-resource-budget';
  }

  const reset = withControlsEmptied(fresh);
  const where = content.children.findIndex((child) => child.id === source.id);
  const at = index === 0 ? where : where + 1;
  return {
    children: [...content.children.slice(0, at), reset.item, ...content.children.slice(at)],
    resetIds: reset.ids,
  };
}

/**
 * Add an item that repeats the one before `index`, or remove the item at `index`. Word records
 * neither with a tracked change of its own, and the engine writes them with tracking off, as
 * other control edits.
 */
export function applyRepeatingSectionItem(
  part: OoxmlPart,
  op: ItemOp,
  options?: EditOptions
): TreeOpResult {
  const refused = validateRepeatingSectionItem(part, op);
  if (refused) return { ok: false, reason: refused };
  const section = findContentControl(part, op.controlId);
  const content = section ? contentControlContentOf(section) : undefined;
  if (!section || !content) return { ok: false, reason: 'unknown-control' };
  const items = itemsOf(section);
  const update =
    op.op === 'removeRepeatingSectionItem'
      ? withItemRemoved(content, items, op.index)
      : withItemAdded(part, content, items, op.index ?? items.length);
  if (typeof update === 'string') return { ok: false, reason: update };
  const { children, resetIds } = update;
  const nextSection = {
    ...section,
    children: section.children.map((child) =>
      child === content ? { ...content, children } : child
    ),
  } as OoxmlNode;
  const written = replaceNode(part, section.id, nextSection, options);
  if (!written.ok) return fromEdit(written, contentControlEffect(section.id, 'flow-structural'));
  const result = withPromptsShown(written.part, resetIds, options);
  const after = findNode(result, section.id) ?? nextSection;
  return {
    ok: true,
    part: result,
    effect: withParagraphDiff(contentControlEffect(section.id, 'flow-structural'), section, after),
  };
}
