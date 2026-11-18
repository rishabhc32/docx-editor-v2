// Edits Word itself refuses, which stay refused in ordered rooms (`tasks.md` section 9): the
// list refuses them by name, as Word does. Where the editor alone differs from Word, the editor
// is fixed to match Word (task 9.13), not the list to match the editor.

import type { TreeDocOpKind } from '../../../../store/store/tree-op-types.ts';
import type { OrderedRefusalReason } from '../../model/refusal.ts';

/** One refusal Word makes: the op, the list's reason, and what Word does. */
export interface OrderedWordRefusal {
  readonly op: TreeDocOpKind;
  readonly reason: OrderedRefusalReason;
  readonly word: string;
}

export const ORDERED_WORD_REFUSALS: readonly OrderedWordRefusal[] = [
  {
    op: 'splitParagraph',
    reason: 'single-line-control',
    word: "Enter in a control that holds one line (plain text without multiple lines, a list, a date, a check box, a picture) does nothing, and so does Enter inside an inline rich text control (Word 16.114, 2026-10-30); the editor's store refuses both, with the same name (9.13, 9.14)",
  },
  {
    op: 'splitParagraph',
    reason: 'break-in-container',
    word: "Enter in an inline control's prompt: the prompt is no content to split",
  },
  {
    op: 'insertNote',
    reason: 'note-outside-body',
    word: 'A note in a text box, a comment, a footnote or an endnote: Word ignores a note reference in a text box or a comment ([MS-OI29500] 17.11.14(a), 17.11.7(b)), and one in a note makes the file non-conformant (ECMA-376 17.11.14, 17.11.7); the editor refuses each of them too (9.9)',
  },
];
