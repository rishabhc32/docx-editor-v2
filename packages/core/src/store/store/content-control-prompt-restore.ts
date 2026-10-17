// A control a content edit or a revision decision emptied shows its prompt again.
//
// Word's rule: delete the last character of a text, date or list control and the
// placeholder comes back under `w:showingPlcHdr`, ready for the next keystroke to replace
// it whole. A decision empties a control the same way: rejecting the value typed over a
// prompt, or accepting the deletion of a control's last characters. The prompt is the text
// of the glossary block the control's `w:placeholder` names when the transaction can read
// the glossary (`EditOptions.placeholderText`), and the type's own prompt otherwise, in the
// placeholder style the open gives it; the reference itself survives for Word and for the
// next open, where `materializeGlossaryPlaceholders` resolves it. Split from
// `tree-op-content-controls.ts` to keep that file under its line cap.

import type { ContentControlKind } from '../package/content-control-nodes.ts';
import {
  contentControlPropertiesContainerOf,
  contentControlPropertiesOf,
} from '../package/content-control-nodes.ts';
import {
  createNodeIdAllocator,
  findNode,
  replaceNode,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import { MAX_INLINE_CONTAINER_DEPTH } from '../package/ooxml-shared.ts';
import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { isInlineControl, isWml } from './content-control-checkbox.ts';
import { valueContent } from './content-control-value-content.ts';
import { placeholderRunProperties } from './placeholder-materialize.ts';
import { editedProperties, promptFor, textRun } from './tree-op-content-controls.ts';
import {
  contentControlContentOf,
  findContentControl,
  isRunPropertiesNode,
} from './tree-op-nodes.ts';
import { paragraphLength } from './tree-op-segments.ts';
import type { TreeOpResult } from './tree-op-types.ts';

/** The kinds whose empty content shows a prompt. A checkbox or a picture has none. */
const PROMPT_KINDS: ReadonlySet<ContentControlKind> = new Set<ContentControlKind>([
  'plainText',
  'richText',
  'date',
  'dropDownList',
  'comboBox',
  'docPartList',
]);

function holdsContent(node: OoxmlNode, depth: number): boolean {
  if (node.kind === 'textValue' || depth > MAX_INLINE_CONTAINER_DEPTH) return false;
  if (node.kind === 'run') return node.children.some((child) => !isRunPropertiesNode(child));
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
 * Deleting the last character used to leave a zero-width control the caret could not enter,
 * which read as the control having vanished. A control that still holds content, already
 * shows its prompt, or has no prompt to show leaves `result` unchanged.
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
 */
export function withRestoredPrompt(
  part: OoxmlPart,
  controlId: string,
  options?: EditOptions
): OoxmlPart {
  const control = findContentControl(part, controlId);
  if (!control) return part;
  const summary = contentControlPropertiesOf(control);
  if (summary.showingPlaceholder || !PROMPT_KINDS.has(summary.type)) return part;
  const content = contentControlContentOf(control);
  if (!content || holdsContent(content, 0)) return part;
  const nextId = createNodeIdAllocator(part);
  const inline = isInlineControl(part, control.id);
  const docPart = summary.placeholderDocPart;
  const prompt =
    (docPart === undefined ? undefined : options?.placeholderText?.(docPart)) ??
    promptFor(summary.type);
  const children = valueContent(
    content,
    (properties) => textRun(nextId, prompt, promptProperties(properties, nextId)),
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
