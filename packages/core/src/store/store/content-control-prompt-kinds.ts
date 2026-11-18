// Which content controls show a prompt when they hold nothing.

import {
  contentControlPropertiesContainerOf,
  type ContentControlKind,
} from '../package/content-control-nodes.ts';
import type { OoxmlNode } from '../package/ooxml-tree.ts';

/** `w15` wordml-2012, where the repeating section extensions live. */
export const W15_NAMESPACE_URI = 'http://schemas.microsoft.com/office/word/2012/wordml';

/** The kinds whose empty content shows a prompt. A checkbox, a picture or a group has none. */
const PROMPT_KINDS: ReadonlySet<ContentControlKind> = new Set<ContentControlKind>([
  'plainText',
  'richText',
  'untyped',
  'date',
  'dropDownList',
  'comboBox',
  'docPartList',
]);

/** The `w15` extensions that type a control which has no ECMA-376 type element. */
const W15_CONTROL_TYPES: ReadonlySet<string> = new Set([
  'repeatingSection',
  'repeatingSectionItem',
]);

/**
 * Whether `control`, of kind `kind`, shows a prompt when its content is empty. The open
 * (`placeholder-materialize.ts`) and a content edit that empties a control
 * (`content-control-prompt-restore.ts`) both ask here, so they cannot drift.
 *
 * The type element of `w:sdtPr` is optional (ECMA-376 `CT_SdtPr`), and Word writes none for the
 * rich text controls it inserts, so a control with no type shows its prompt as rich text does.
 * A repeating section and its items have no ECMA-376 type either: their type is a `w15`
 * extension, and they have no prompt of their own.
 */
export function showsPromptWhenEmpty(control: OoxmlNode, kind: ContentControlKind): boolean {
  if (!PROMPT_KINDS.has(kind)) return false;
  if (kind !== 'untyped') return true;
  const properties = contentControlPropertiesContainerOf(control);
  const children: readonly OoxmlNode[] = properties?.children ?? [];
  return !children.some(
    (child) =>
      child.kind !== 'textValue' &&
      child.namespaceUri === W15_NAMESPACE_URI &&
      W15_CONTROL_TYPES.has(child.localName)
  );
}
