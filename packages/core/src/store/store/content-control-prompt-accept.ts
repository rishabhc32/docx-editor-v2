// An accept that leaves real content in a content control ends the control's prompt.

import { contentControlsIn, isContentControlWrapper } from '../package/content-control-nodes.ts';
import { findNode, replaceNode, type EditOptions } from '../package/ooxml-edit.ts';
import { WML_NAMESPACE_URI } from '../package/ooxml-shared.ts';
import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { withoutPlaceholderStyle } from './content-control-checkbox.ts';
import {
  contentControlContentOf,
  contentControlPropertiesOf,
  isParagraphPropertiesNode,
  isRunPropertiesNode,
  isShowingPlaceholder,
} from './tree-op-nodes.ts';

/** Stops a hostile document's deep nesting: the walks go no deeper than this. */
const MAX_DEPTH = 32;

function isStrike(node: OoxmlNode): boolean {
  return node.kind === 'revisionDelete' || node.kind === 'revisionMoveFrom';
}

function isProposal(node: OoxmlNode): boolean {
  return node.kind === 'revisionInsert' || node.kind === 'revisionMoveTo';
}

/** Whether a walk of a control's own content stops at `node`: a nested control or a mark. */
function isOpaque(node: OoxmlNode): boolean {
  return isContentControlWrapper(node) || isParagraphPropertiesNode(node);
}

/** The ids of the strikes and proposals in a control's own content. */
function revisionWrapperIds(content: OoxmlNode): {
  readonly strikes: readonly string[];
  readonly proposals: readonly string[];
} {
  const strikes: string[] = [];
  const proposals: string[] = [];
  const walk = (node: OoxmlNode, depth: number): void => {
    if (node.kind === 'textValue' || depth > MAX_DEPTH || isOpaque(node)) return;
    if (isStrike(node)) {
      strikes.push(node.id);
    } else if (isProposal(node)) {
      proposals.push(node.id);
    }

    for (const child of node.children) {
      walk(child, depth + 1);
    }
  };
  for (const child of content.kind === 'textValue' ? [] : content.children) {
    walk(child, 0);
  }

  return { strikes, proposals };
}

/** Whether a control's own content holds anything that is not a pending strike. */
function holdsWords(content: OoxmlNode): boolean {
  const walk = (node: OoxmlNode, depth: number): boolean => {
    if (node.kind === 'textValue' || depth > MAX_DEPTH || isStrike(node)) return false;
    if (isContentControlWrapper(node)) return true;
    if (isParagraphPropertiesNode(node)) return false;
    if (node.kind === 'run') return node.children.some((child) => !isRunPropertiesNode(child));
    return node.children.some((child) => walk(child, depth + 1));
  };
  return content.kind !== 'textValue' && content.children.some((child) => walk(child, 0));
}

/**
 * `node` with Word's `PlaceholderText` style taken off every run outside a pending strike. Word
 * never styles typed words as a prompt, so they keep only the prompt's other run properties, as
 * a value written into the control does (`valueContent`). A struck prompt still pending keeps
 * its face.
 */
function withoutPromptStyle(node: OoxmlNode, depth: number): OoxmlNode {
  if (node.kind === 'textValue' || depth > MAX_DEPTH || isStrike(node) || isOpaque(node)) {
    return node;
  }

  const children =
    node.kind === 'run'
      ? node.children.flatMap((child) => {
          if (!isRunPropertiesNode(child)) return [child];
          const kept = withoutPlaceholderStyle(child);
          return kept ? [kept] : [];
        })
      : node.children.map((child) => withoutPromptStyle(child, depth + 1));
  const same =
    children.length === node.children.length &&
    children.every((child, index) => child === node.children[index]);
  return same ? node : ({ ...node, children } as OoxmlNode);
}

/**
 * `control` holding its content as real content: no `w:showingPlcHdr` (§17.5.2.38), which would
 * make the next keystroke replace the accepted words whole, and no prompt style.
 */
function withPromptEnded(control: OoxmlNode): OoxmlNode {
  if (control.kind === 'textValue') return control;
  const properties = contentControlPropertiesOf(control);
  const content = contentControlContentOf(control);
  const children = control.children.map((child) => {
    if (child === properties) {
      return {
        ...properties,
        children: properties.children.filter(
          (entry) =>
            entry.kind === 'textValue' ||
            entry.namespaceUri !== WML_NAMESPACE_URI ||
            entry.localName !== 'showingPlcHdr'
        ),
      } as OoxmlNode;
    }

    return child === content ? withoutPromptStyle(content, 0) : child;
  });
  return { ...control, children } as OoxmlNode;
}

/**
 * End the prompt of every control in `before` whose pending replacement of its prompt an
 * accept decision resolved, in `after`, the part that decision produced. Returns `after` itself
 * when no such control exists.
 *
 * A replacement is a strike AND a proposal inside a control that shows its prompt: the shape
 * older editor versions wrote for typing over a prompt in suggesting mode (this one replaces a
 * prompt without a revision, as Word does). A prompt inside a tracked insertion with no strike is
 * a control inserted with tracking on, and accepting that insertion keeps the prompt. A nested
 * control's prompt is its own.
 */
export function endAcceptedPrompts(
  before: OoxmlPart,
  after: OoxmlPart,
  options?: EditOptions
): OoxmlPart {
  let next = after;
  for (const { node } of contentControlsIn(before.root)) {
    if (!isShowingPlaceholder(node)) continue;
    const content = contentControlContentOf(node);
    if (!content) continue;
    const { strikes, proposals } = revisionWrapperIds(content);
    if (strikes.length === 0 || proposals.length === 0) continue;
    const current = findNode(next, node.id);

    // The same node: the decision did not reach this control.
    if (!current || current === node || current.kind !== 'contentControl') continue;
    const currentContent = contentControlContentOf(current);
    if (!currentContent || !holdsWords(currentContent)) continue;
    const remaining = revisionWrapperIds(currentContent);
    const remainingIds = new Set([...remaining.strikes, ...remaining.proposals]);
    if ([...strikes, ...proposals].every((id) => remainingIds.has(id))) continue;
    const written = replaceNode(next, current.id, withPromptEnded(current), options);
    if (written.ok) next = written.part;
  }

  return next;
}
