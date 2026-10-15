// The link commands' own argument and refusal rules: a tooltip for `insertHyperlink`, and a
// refused `removeHyperlink` told apart from a missing link.

import type { ExecResult } from '../contracts/editor.ts';
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
