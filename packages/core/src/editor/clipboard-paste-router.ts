import { fragmentContainsClipboardObject } from './clipboard-object-selection.ts';
// Paste flavour routing (rich-clipboard-fidelity task 4.1).
//
// Fidelity order: internal fragment, then external `text/html`, then `text/plain` — and
// the degrade is CONTINUOUS: a payload that fails decoding, fails the bounded package
// read, or has a recoverable apply refusal falls to the next flavour. Unsupported object
// content refuses without fallback, and so does a refusal from the collaboration gate, which
// would refuse the plain text too. Other failures do not leave a no-op paste. Suggesting
// mode lands a fragment as one tracked insertion; without an author, and in non-body
// stories, the plain lane runs; the drop lane never routes through here.

import { fragmentFromHtml } from './clipboard-fragment-codec.ts';
import { projectExternalHtml } from './clipboard-html-read.ts';

export interface PasteRouteTarget {
  /** True when a fragment landing is even possible: body story, edit or suggesting mode. */
  readonly richLaneOpen: boolean;
  /**
   * False permits fallback; unsupported-content refuses the complete paste, and so does
   * refused, the answer when the write gate refused a paste that already reported why.
   */
  pasteFragment(
    bytes: Uint8Array,
    lastMarkCovered: boolean
  ): boolean | 'unsupported-content' | 'refused';
  insertPlainText(text: string): void;
}

export interface PasteRouteInput {
  readonly html: string | null;
  readonly text: string;
  /** Cmd+Shift+V or the pasteWithoutFormatting command: plain lane, no questions. */
  readonly forcePlain: boolean;
}

export type PasteRouteLane =
  | 'fragment'
  | 'external-html'
  | 'plain'
  | 'none'
  | 'unsupported-content'
  | 'refused';

/** Route one paste payload; reports the lane that actually landed. */
export function routePaste(target: PasteRouteTarget, input: PasteRouteInput): PasteRouteLane {
  const plain = (): PasteRouteLane => {
    if (input.text.length === 0) return 'none';
    target.insertPlainText(input.text);
    return 'plain';
  };

  if (input.forcePlain) return plain();
  const html = input.html;
  if (html === null || html.length === 0) return plain();

  const embedded = fragmentFromHtml(html);
  if (!target.richLaneOpen) {
    if (embedded && fragmentContainsClipboardObject(embedded.bytes)) return 'unsupported-content';
    return plain();
  }
  if (embedded) {
    const result = target.pasteFragment(embedded.bytes, embedded.lastMarkCovered);
    if (result === 'unsupported-content' || result === 'refused') return result;
    if (result) return 'fragment';
  }

  const projected = projectExternalHtml(html);
  if (projected.ok && projected.truncated && input.text.length > 0) return plain();
  if (projected.ok) {
    const result = target.pasteFragment(projected.fragmentBytes, projected.lastMarkCovered);
    if (result === 'unsupported-content' || result === 'refused') return result;
    if (result) return 'external-html';
  }

  return plain();
}
