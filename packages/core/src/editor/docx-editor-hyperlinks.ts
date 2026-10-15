// The link commands' own argument and refusal rules: a tooltip for `insertHyperlink`, a
// refused `removeHyperlink` told apart from a missing link, and the refusals `can` reports.

import type { EditorCommand, ExecResult } from '../contracts/editor.ts';
import { isValidXmlText } from '../store/package/sinks.ts';
import { MAX_HYPERLINK_ATTRIBUTE_LENGTH } from '../store/store/tree-op-validate.ts';
import type { PaginatedSurface } from './paginated-surface-contract.ts';

type Refusal = Exclude<ExecResult, { ok: true }>;

/**
 * Why `insertHyperlink` cannot write `tooltip`, or null when it can.
 *
 * Written straight into `w:tooltip`, so checked before the target is minted. An empty one is
 * refused rather than read as "no tooltip": on an existing link it would ask to clear one,
 * which the write cannot do.
 */
export function hyperlinkTooltipRefusal(tooltip: unknown): string | null {
  if (
    tooltip === undefined ||
    (typeof tooltip === 'string' &&
      tooltip.length > 0 &&
      tooltip.length <= MAX_HYPERLINK_ATTRIBUTE_LENGTH &&
      isValidXmlText(tooltip))
  )
    return null;
  return 'insertHyperlink requires a tooltip of 1-512 characters that XML can hold';
}

/**
 * Why `removeHyperlink` committed nothing while the caret is in a link, or null when there is
 * no link at the caret. A link that stays is a refusal, e.g. one suggesting mode cannot
 * propose removing.
 */
export function removeHyperlinkRefusal(mounted: PaginatedSurface): Refusal | null {
  if (!mounted.hyperlinks.linkAtCaret()) return null;
  return {
    ok: false,
    code: 'invalidArgs',
    reason: mounted.state().lastRejection ?? 'the link removal was refused',
  };
}

/** Why `removeHyperlink` refuses when the caret is in no link. */
export const NO_HYPERLINK_AT_SELECTION = 'there is no hyperlink at the selection';

/**
 * The link `insertHyperlink` asks the surface for. `#name` is a bookmark in this document;
 * anything else is an external target.
 */
export function hyperlinkInputOf(command: Extract<EditorCommand, { type: 'insertHyperlink' }>): {
  readonly url?: string;
  readonly anchor?: string;
  readonly text?: string;
  readonly tooltip?: string;
} {
  const internal = command.href.startsWith('#');
  return {
    ...(internal ? { anchor: command.href.slice(1) } : { url: command.href }),
    ...(command.text !== undefined ? { text: command.text } : {}),
    ...(command.tooltip !== undefined ? { tooltip: command.tooltip } : {}),
  };
}

/**
 * Why `exec` would refuse a link command, as far as can be told without writing, or null: no
 * link at the caret to remove, or a link change suggesting mode cannot propose. `can` reports
 * it, so the two agree.
 */
export function hyperlinkCommandRefusal(
  command: EditorCommand,
  mounted: PaginatedSurface
): Refusal | null {
  let reason: string | null = null;
  if (command.type === 'insertHyperlink') {
    reason = mounted.hyperlinks.applyRefusal(hyperlinkInputOf(command));
  } else if (command.type === 'removeHyperlink') {
    if (!mounted.hyperlinks.linkAtCaret()) {
      return { ok: false, code: 'notFound', reason: NO_HYPERLINK_AT_SELECTION };
    }
    reason = mounted.hyperlinks.removeRefusal();
  }
  return reason === null ? null : { ok: false, code: 'invalidArgs', reason };
}
