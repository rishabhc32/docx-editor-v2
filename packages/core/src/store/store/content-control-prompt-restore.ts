// A control a content edit or a revision decision emptied shows its prompt again.
//
// Word's rule: delete the last character of a text, date or list control and the
// placeholder comes back under `w:showingPlcHdr`, ready for the next keystroke to replace
// it whole. Split from `tree-op-content-controls.ts` to keep that file under its line cap.

import {
  contentControlPropertiesContainerOf,
  contentControlPropertiesOf,
  type ContentControlKind,
} from '../package/content-control-nodes.ts';
import { actorScopedSeed } from '../package/actor-scoped-ids.ts';
import {
  createNodeIdAllocator,
  findNode,
  replaceNode,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import { MAX_INLINE_CONTAINER_DEPTH } from '../package/ooxml-shared.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { usedParaIds, w14PrefixInScopeAt, withFreshParaIds } from '../package/para-id.ts';
import { isInlineControl, isWml } from './content-control-checkbox.ts';
import { showsPromptWhenEmpty } from './content-control-prompt-kinds.ts';
import { isCommentAnchorRun, valueContent } from './content-control-value-content.ts';
import { placeholderRunProperties } from './placeholder-materialize.ts';
import { editedProperties, promptFor, textRun, wmlElement } from './tree-op-content-controls.ts';
import {
  cloneWithNewIds,
  contentControlContentOf,
  findContentControl,
  isRunPropertiesNode,
} from './tree-op-nodes.ts';
import { paragraphLength } from './tree-op-segments.ts';
import type { TreeOpResult } from './tree-op-types.ts';

function holdsContent(node: OoxmlNode, depth: number): boolean {
  if (node.kind === 'textValue' || depth > MAX_INLINE_CONTAINER_DEPTH) return false;

  // A comment's anchor is no content: it stays beside the prompt, as beside a value
  if (node.kind === 'run') {
    return !isCommentAnchorRun(node) && node.children.some((child) => !isRunPropertiesNode(child));
  }

  if (node.localName === 'sdtPr' || node.localName === 'pPr') return false;
  return node.children.some((child) => holdsContent(child, depth + 1));
}

/** Whether a control holds any content, struck text included. */
export function controlHoldsContent(control: OoxmlNode): boolean {
  const content = contentControlContentOf(control);
  return content !== undefined && holdsContent(content, 0);
}

/**
 * Restore the prompt of a control the edit in `result` has just emptied. The wrapper stays,
 * the prompt goes back under `w:showingPlcHdr`, and the next keystroke replaces it whole.
 * Without it, deleting the last character leaves a zero-width control the caret cannot enter.
 * `withRestoredPrompt` names the controls that leave `result` unchanged.
 */
export function restoreEmptiedPlaceholder(
  result: TreeOpResult,
  controlId: string,
  options?: EditOptions
): TreeOpResult {
  if (!result.ok) return result;
  const part = withRestoredPrompt(result.part, controlId, options);
  return part === result.part ? result : { ...result, part };
}

/**
 * `part` with the prompt of an empty control back under `w:showingPlcHdr`, or `part` itself
 * when the control holds content, already shows its prompt, or has none to show.
 *
 * The prompt is the glossary block the control's `w:placeholder` names (§17.5.2.25), with its
 * own runs and run properties, when the transaction can read the glossary
 * (`EditOptions.placeholderBlocks`). A suggestion over a prompt puts that block there first
 * (`content-control-prompt-glossary.ts`), so the control shows the prompt it had. Without a
 * block, the prompt is the block's text when the caller lends only that (the deprecated
 * `EditOptions.placeholderText`), else the type's own prompt, in one run in the placeholder
 * style. The `w:placeholder` reference stays for Word and for the next open, where
 * `materializeGlossaryPlaceholders` resolves it.
 */
export function withRestoredPrompt(
  part: OoxmlPart,
  controlId: string,
  options?: EditOptions
): OoxmlPart {
  const control = findContentControl(part, controlId);
  if (!control) return part;
  const summary = contentControlPropertiesOf(control);
  if (summary.showingPlaceholder) return part;
  const content = contentControlContentOf(control);
  if (!content || holdsContent(content, 0)) return part;
  const docPart = summary.placeholderDocPart;
  const blocks = docPart === undefined ? undefined : options?.placeholderBlocks?.(docPart);
  const nextId = createNodeIdAllocator(part);
  const inline = isInlineControl(part, control.id);
  const fromBlock = blocks
    ? promptBlockContent(part, control, content, blocks, nextId, inline)
    : null;
  // A block restores whatever the control is. Without one, only a kind that has a prompt of its
  // own shows one: an untyped control is rich text, as in Word, and a checkbox, a picture, a
  // group or a repeating section has none.
  if (!fromBlock && !showsPromptWhenEmpty(control, summary.type)) return part;
  // No block to restore: the block's text when the caller lends only that, then the type's own.
  const text = () =>
    (docPart === undefined ? undefined : options?.placeholderText?.(docPart)) ??
    promptFor(summary.type);
  const children =
    fromBlock ??
    valueContent(
      content,
      (face) => textRun(nextId, text(), promptProperties(face, nextId)),
      nextId,
      inline
    );
  if (!children) return part;
  const sdtPr = contentControlPropertiesContainerOf(control);
  const properties = editedProperties(sdtPr, { showingPlaceholder: true }, nextId);
  const rebuilt = {
    ...control,
    children: [
      properties,
      ...control.children.filter((child) => child.id !== sdtPr?.id && child.id !== content.id),
      { ...content, children } as OoxmlNode,
    ],
  } as OoxmlNode;
  const written = replaceNode(part, control.id, rebuilt, options);
  return written.ok ? written.part : part;
}

/** The kinds a decision empties into Word's blank prompt when they name no prompt. */
const BLANK_PROMPT_KINDS: ReadonlySet<ContentControlKind> = new Set<ContentControlKind>([
  'plainText',
  'richText',
  'untyped',
]);

/** The text of Word's blank prompt: five spaces. */
export const BLANK_PROMPT_TEXT = '     ';

/**
 * `part` with a control a revision decision emptied showing its prompt, as Word's Accept All and
 * Reject All leave it: a reject of the value typed over a prompt, or an accept of the deletion
 * of a control's last characters. A control that names a prompt (`w:placeholder`), or whose
 * kind is not text, shows its prompt as an emptying edit does (`withRestoredPrompt`).
 *
 * A text, rich text or untyped control that names no prompt shows Word's blank prompt: Word
 * 16.114 writes `w:showingPlcHdr` and an empty `w:sdtEndPr`, drops `w:richText`, and fills the
 * content with one run of five spaces in no style. Word adds no glossary block (inline and
 * block controls, Accept of a deletion and Reject of an insertion, compatibility and native
 * mode).
 */
export function withDecisionPrompt(
  part: OoxmlPart,
  controlId: string,
  options?: EditOptions
): OoxmlPart {
  const control = findContentControl(part, controlId);
  if (!control) return part;
  const summary = contentControlPropertiesOf(control);
  if (
    summary.placeholderDocPart !== undefined ||
    !BLANK_PROMPT_KINDS.has(summary.type) ||
    !showsPromptWhenEmpty(control, summary.type)
  ) {
    return withRestoredPrompt(part, controlId, options);
  }
  const content = contentControlContentOf(control);
  if (summary.showingPlaceholder || !content || holdsContent(content, 0)) return part;
  const nextId = createNodeIdAllocator(part);
  const blank = () => textRun(nextId, BLANK_PROMPT_TEXT, undefined);
  const children = valueContent(content, blank, nextId, isInlineControl(part, control.id));
  if (!children) return part;
  const sdtPr = contentControlPropertiesContainerOf(control);
  const flagged = editedProperties(sdtPr, { showingPlaceholder: true }, nextId);
  const properties = {
    ...flagged,
    children: flagged.children.filter((child) => !isWml(child, 'richText')),
  } as OoxmlNode;
  const endProperties = control.children.find((child) => isWml(child, 'sdtEndPr'));
  const rebuilt = {
    ...control,
    children: [
      properties,
      endProperties ??
        wmlElement(nextId, 'sdtEndPr', {
          kind: 'contentControlEndProperties' as OoxmlNode['kind'],
        }),
      ...control.children.filter(
        (child) => child.id !== sdtPr?.id && child !== endProperties && child.id !== content.id
      ),
      { ...content, children } as OoxmlNode,
    ],
  } as OoxmlNode;
  const written = replaceNode(part, control.id, rebuilt, options);
  return written.ok ? written.part : part;
}

const isParagraph = (node: OoxmlNode): boolean => node.kind === 'paragraph' || isWml(node, 'p');

function containsNode(node: OoxmlNode, id: string, depth: number): boolean {
  if (node.id === id) return true;
  if (node.kind === 'textValue' || depth > MAX_INLINE_CONTAINER_DEPTH) return false;
  return node.children.some((child) => containsNode(child, id, depth + 1));
}

/** `nodes` with the node `id` replaced by `replacement`, rebuilding only its ancestors. */
function spliced(
  nodes: readonly OoxmlNode[],
  id: string,
  replacement: readonly OoxmlNode[],
  depth: number
): OoxmlNode[] {
  return nodes.flatMap((node) => {
    if (node.id === id) return [...replacement];
    if (node.kind === 'textValue' || !containsNode(node, id, depth)) return [node];
    return [{ ...node, children: spliced(node.children, id, replacement, depth + 1) } as OoxmlNode];
  });
}

/**
 * The emptied content holding the glossary block's prompt, or null to fall back to the type's
 * prompt. The block's first paragraph goes where a value would, into the control's display
 * paragraph, which keeps its id and its markers, with the runs exactly as the block has them. At
 * block level the block's other paragraphs follow it, as the prompt had them.
 */
function promptBlockContent(
  part: OoxmlPart,
  control: OoxmlElement,
  content: OoxmlElement,
  blocks: readonly OoxmlNode[],
  nextId: () => string,
  inline: boolean
): readonly OoxmlNode[] | null {
  const [first, ...rest] = blocks;
  if (!first || first.kind === 'textValue' || !isParagraph(first)) return null;
  const runs = first.children
    .filter((child) => !isWml(child, 'pPr'))
    .map((child) => cloneWithNewIds(child, nextId));
  if (runs.length === 0) return null;
  const marker = textRun(nextId, '', undefined);
  const filled = valueContent(content, () => marker, nextId, inline);
  if (!filled) return null;
  const children = spliced(filled, marker.id, runs, 0);
  if (inline || rest.length === 0) return children;
  return withFollowingPromptParagraphs(part, control, children, runs[0]!.id, first, rest, nextId);
}
/**
 * `children` with a prompt of several paragraphs: the display paragraph (the one that holds
 * `firstRunId`) takes the first block paragraph's properties, and the other block paragraphs
 * follow it with paragraph ids of their own.
 */
function withFollowingPromptParagraphs(
  part: OoxmlPart,
  control: OoxmlElement,
  children: OoxmlNode[],
  firstRunId: string,
  first: OoxmlElement,
  rest: readonly OoxmlNode[],
  nextId: () => string
): OoxmlNode[] {
  const at = children.findIndex(
    (child) => isParagraph(child) && containsNode(child, firstRunId, 0)
  );
  const display = children[at];
  if (!display || display.kind === 'textValue') return children;
  const properties = first.children.find((child) => isWml(child, 'pPr'));
  children[at] = {
    ...display,
    children: [
      ...(properties ? [cloneWithNewIds(properties, nextId)] : []),
      ...display.children.filter((child) => !isWml(child, 'pPr')),
    ],
  } as OoxmlNode;
  const paraIds = new Set(usedParaIds(part.root as OoxmlElement));
  const counter = { value: 0 };
  const prefix = w14PrefixInScopeAt(part, control);
  const following = rest.map((block, index) =>
    withFreshParaIds(
      cloneWithNewIds(block, nextId),
      paraIds,
      actorScopedSeed(`${control.id}:prompt:${String(index)}`),
      counter,
      prefix
    )
  );
  children.splice(at + 1, 0, ...following);
  return children;
}

/**
 * The emptied content's face in Word's placeholder style, as the open styles a prompt. A face
 * with a character style of its own keeps it: a run takes one.
 */
function promptProperties(face: OoxmlNode | undefined, nextId: () => string): OoxmlNode {
  if (!face || face.kind === 'textValue') return placeholderRunProperties(nextId);
  if (face.children.some((child) => isWml(child, 'rStyle'))) return face;
  const style = placeholderRunProperties(nextId);
  const rStyle = style.kind === 'textValue' ? [] : style.children;
  return { ...face, children: [...rStyle, ...face.children] } as OoxmlNode;
}

/**
 * Where the caller's characters go once a prompt has been emptied.
 *
 * Prompt replacement uses the prompt's start position. Clamp the position to the shortened
 * paragraph before inserting the replacement text.
 */
export function promptInsertionOffset(
  part: OoxmlPart,
  paragraphId: string,
  planned: number
): number {
  const paragraph = findNode(part, paragraphId);
  if (!paragraph || paragraph.kind !== 'paragraph') return planned;
  return Math.min(planned, paragraphLength(paragraph));
}
