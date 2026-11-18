// What control edits share (task 9.6): the engine's op on the view, items from the element it
// wrote, and a record of the `w:sdtPr` children an edit changed.

import { canonicalTreeDifference } from '../../../store/package/ooxml-difference.ts';
import { findNode } from '../../../store/package/ooxml-edit.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import { namespaceBindingsAt } from '../../../store/package/wml-namespace.ts';
import { applyTreeOp } from '../../../store/store/tree-op-apply.ts';
import { editedProperties } from '../../../store/store/tree-op-content-controls.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import type { TreeDocOp } from '../../../store/store/tree-op-types.ts';
import { propertyKey } from '../compose/properties.ts';
import { createOrderedItem } from '../model/document.ts';
import { nextSeq } from '../model/mutations.ts';
import { closeOf, itemsInside } from '../placement.ts';
import { engineRefusal, OrderedRefusal } from '../model/refusal.ts';
import { recordUndo } from '../model/undo-record.ts';
import {
  isUnit,
  type OrderedElementRecord,
  type OrderedItem,
  type OrderedWithin,
} from '../model/types.ts';
import { childNamed, isControlContent, isWml, withChildren, withId } from '../ooxml-nodes.ts';
import { holdWithin, strikeItems, withNewIds, type Ctx } from './context.ts';
import { templateItems, viewPart } from './templates.ts';

/** The key of `w:showingPlcHdr`: whether a prompt shows is compose's to say, never a record's. */
const PROMPT_FLAG = `${WML_NAMESPACE_URI} showingPlcHdr`;

/**
 * The engine's op on the story as the view draws it: refused by name as the engine's own
 * validator refuses it (a lock, a binding, a range across a container), else its result.
 */
export function engineOnView(ctx: Ctx, op: TreeDocOp): { view: OoxmlPart; result: OoxmlPart } {
  const view = viewPart(ctx);
  return { view, result: engineApply(view, op) };
}

/** The engine's op on a part, refused by name as the engine refuses it (its own validation). */
export function engineApply(view: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const applied = applyTreeOp(view, op, { deferValidation: true });
  if (!applied.ok) {
    throw engineRefusal(op.op, applied.reason);
  }

  return applied.part;
}

/** Every control element under a node, by id. */
function controlsUnder(node: OoxmlNode, out = new Map<string, OoxmlElement>(), depth = 0) {
  if (node.kind === 'textValue' || depth > 64) return out;
  if (node.kind === 'contentControl' || isWml(node, 'sdt')) out.set(node.id, node);
  for (const child of node.children) controlsUnder(child, out, depth + 1);
  return out;
}

/** The control the engine added to a paragraph: the one the view's paragraph did not hold. */
export function addedControl(view: OoxmlPart, result: OoxmlPart, paragraph: string): OoxmlElement {
  const before = findNode(view, paragraph);
  const after = findNode(result, paragraph);
  if (!before || !after) throw new OrderedRefusal('malformed-story', `${paragraph} went`);
  const old = controlsUnder(before);
  const added = [...controlsUnder(after).values()].filter((c) => !old.has(c.id));
  const [control] = added;
  if (!control || added.length !== 1)
    throw new OrderedRefusal('engine-refused', 'the engine added no control');
  return control;
}

/**
 * A control's `w:id` from the change's id minter, so every copy writes the same: the engine's
 * own choice depends on the part it runs on, which another person's change can make differ.
 */
function withMintedControlId(ctx: Ctx, shell: OoxmlElement): OoxmlElement {
  const sdtPr = childNamed(shell, 'sdtPr');
  if (!sdtPr) return shell;
  const minted = editedProperties(sdtPr, { id: ctx.ids.decimal() }, () => ctx.ids.node());
  return withChildren(
    shell,
    shell.children.map((c) => (c === sdtPr ? minted : c))
  );
}

/** What a new control's items carry. */
export interface ControlLabels {
  readonly suggestion: string;
  readonly within: OrderedWithin;
}

/**
 * The open and close items of a new inline control, its shell from the element the engine
 * wrote (its properties, no content), under the change's ids. The full element is kept as
 * the control's original, so compose can keep its prompt in the glossary once text fills it.
 */
export function controlEdges(
  ctx: Ctx,
  control: OoxmlElement,
  labels: ControlLabels,
  prompt: boolean
): { open: OrderedItem; close: OrderedItem } {
  const full = withNewIds(ctx, control);
  const space = ctx.ids.container();
  const bare = withChildren(
    full,
    full.children.filter((c) => !isControlContent(c))
  );
  const minted = withMintedControlId(ctx, bare);

  // A container a change made has its space as its node id, as a new paragraph has its mark's:
  // the editor names it by that id, and the list finds it by that id (`openOfNode`)
  const shell = withId(minted, space);
  const original = withChildren(full, [
    ...shell.children,
    ...full.children.filter(isControlContent),
  ]);
  ctx.doc.originals.set(shell.id, original);
  recordUndo(ctx.doc, () => ctx.doc.originals.delete(shell.id));
  const common = {
    space,
    token: 'sdt' as const,
    pair: space,
    by: labels.suggestion,
    within: labels.within,
  };
  return {
    open: createOrderedItem({
      ...common,
      kind: 'open',
      node: shell,
      ...(prompt ? { prompt: true } : {}),
    }),
    close: createOrderedItem({ ...common, offset: 1, kind: 'close' }),
  };
}

/**
 * The units of an inline control's content as the engine wrote it, loaded as a paragraph would
 * be (a check box's symbol is an object), inserted by the suggestion.
 */
export function contentUnits(
  ctx: Ctx,
  control: OoxmlElement,
  labels: ControlLabels,
  prompt: boolean
): OrderedItem[] {
  const content = control.children.find(isControlContent);
  const runs = content && content.kind !== 'textValue' ? content.children : [];
  if (runs.length === 0) return [];
  const paragraph = build(ctx.ids.node(), 'paragraph', 'p', [], runs);
  return templateItems(ctx, paragraph, labels.suggestion)
    .filter(isUnit)
    .map(
      (item): OrderedItem => ({
        ...item,
        by: labels.suggestion,
        within: labels.within,
        ...(prompt ? { prompt: true } : {}),
      })
    );
}

/** The children of a properties element by key; the first of each. */
function byKey(element: OoxmlElement | undefined): Map<string, OoxmlElement> {
  const out = new Map<string, OoxmlElement>();
  for (const child of element?.children ?? [])
    if (child.kind !== 'textValue' && !out.has(propertyKey(child)))
      out.set(propertyKey(child), child);
  return out;
}

/**
 * The `w:sdtPr` children's keys an edit changed, `before` to `after`; never the prompt's flag.
 * `part` holds `after`: the children compare under its scope, where an ancestor may declare a
 * prefix one of them names.
 */
export function changedProperties(
  before: OoxmlElement | undefined,
  after: OoxmlElement,
  part: OoxmlPart
): string[] {
  const old = byKey(before);
  const now = byKey(after);
  const keys = new Set<string>();
  const scope = { bindings: namespaceBindingsAt(part, after) };
  for (const [key, child] of now) {
    const was = old.get(key);
    if (!was || !canonicalTreeDifference(was, child, scope).equal) keys.add(key);
  }

  for (const key of old.keys()) if (!now.has(key)) keys.add(key);
  keys.delete(PROMPT_FLAG);
  return [...keys].sort();
}

/**
 * A record that sets the `w:sdtPr` children an edit changed (`before` to `after`, which `part`
 * holds), and only those; null when it changed none. The prompt's flag is never in it.
 * Word has no change record for `w:sdtPr` (ECMA-376 `CT_SdtPr` has no `*Change` child), so the
 * record shows with none, and a later record of other children never takes back an earlier one.
 */
export function controlPropertiesRecord(
  ctx: Ctx,
  suggestion: string,
  before: OoxmlElement | undefined,
  after: OoxmlElement,
  part: OoxmlPart
): OrderedElementRecord | null {
  const merge = changedProperties(before, after, part);
  return merge.length === 0 ? null : sdtPrRecord(ctx, suggestion, after, merge);
}

/** A record that sets the children of `after` that `merge` names (no change record). */
export function sdtPrRecord(
  ctx: Ctx,
  suggestion: string,
  after: OoxmlElement,
  merge: string[]
): OrderedElementRecord {
  const keys = new Set(merge);
  const given = after.children.filter((c) => c.kind !== 'textValue' && keys.has(propertyKey(c)));
  return {
    kind: 'element',
    suggestion,
    seq: nextSeq(ctx.doc),
    target: 'sdtPr',
    element: withNewIds(ctx, withChildren(after, given)),
    merge,
    untracked: true,
  };
}

/**
 * Strike a block container with the edit's suggestion (a direct edit settles it): its edges, and
 * everything in it through `within`, as a struck row is. Words typed in it later go with it too.
 */
export function strikeBlock(ctx: Ctx, open: OrderedItem, suggestion: string): void {
  strikeItems(ctx, ctx.story, [open, closeOf(ctx.story, open, ctx.doc)], suggestion);
  for (const item of itemsInside(ctx.story, open, ctx.doc)) {
    holdWithin(ctx, ctx.story, item, suggestion);
  }
}
