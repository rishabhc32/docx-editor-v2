// The block content controls a whole-story replacement takes out, as Word does.

import { findNode } from '../store/package/ooxml-edit.ts';
import { contentControlContentNodeOf } from '../store/package/content-control-nodes.ts';
import type { OoxmlPart } from '../store/package/ooxml-tree.ts';
import type { TreeDocOp } from '../store/store/tree-ops.ts';
import type { AutomationBlockRead } from './reads.ts';

/** What a whole-story replacement does with the story's block controls. */
export interface StoryControlPlan {
  /** Ops before the text is cleared: the control that leaves its paragraphs goes. */
  readonly first: readonly TreeDocOp[];
  /** Ops after it: that control's other paragraphs go, and every other control with its content. */
  readonly then: readonly TreeDocOp[];
  /** The control blocks this plan takes out. */
  readonly handled: readonly AutomationBlockRead[];
  /** The paragraph that holds the text when a control leaves it; null otherwise. */
  readonly kept: string | null;
  /** The paragraphs that leave the story. */
  readonly gone: readonly string[];
}

/** Whether a control holds only paragraphs, so that removing it leaves them as they are. */
function holdsOnlyParagraphs(part: OoxmlPart, block: AutomationBlockRead): boolean {
  const control = findNode(part, block.id);
  const content = control ? contentControlContentNodeOf(control) : undefined;
  const children = content?.children.filter((child) => child.kind !== 'textValue') ?? [];
  return children.length > 0 && children.every((child) => child.kind === 'paragraph');
}

/**
 * The plan for the block controls of a story that a replacement empties.
 *
 * Word, setting a story's whole text to nothing, removes every block control and leaves one
 * empty paragraph (Word 16 on macOS, 2026-10-30: with a paragraph outside the controls, and
 * with every paragraph inside one control). Each control goes with its content; with no
 * paragraph outside controls, the first control goes and leaves its paragraphs, and the first
 * of them holds the text. A control locked against deletion refuses the whole replacement
 * (`removeContentControl`). With no paragraph outside controls and a first control that holds
 * more than paragraphs, that control stays and holds the text (this host creates no paragraph
 * at the top level).
 */
export function storyControlPlan(
  part: OoxmlPart,
  blocks: readonly AutomationBlockRead[]
): StoryControlPlan {
  const controls = blocks.filter((block) => block.kind === 'control');
  const hasOutsideParagraph = blocks.some((block) => block.kind === 'paragraph');
  const firstControl = controls[0];
  const unwrapped =
    !hasOutsideParagraph && firstControl && holdsOnlyParagraphs(part, firstControl)
      ? firstControl
      : undefined;
  if (!hasOutsideParagraph && !unwrapped) {
    return { first: [], then: [], handled: [], kept: null, gone: [] };
  }

  const removed = controls.filter((block) => block !== unwrapped);
  const left = unwrapped?.paragraphIds.slice(1) ?? [];
  return {
    first: unwrapped
      ? [{ op: 'removeContentControl', controlId: unwrapped.id, keepContent: true }]
      : [],
    then: [
      ...left.map((blockId): TreeDocOp => ({ op: 'deleteBlock', blockId })),
      ...removed.map(
        (block): TreeDocOp => ({
          op: 'removeContentControl',
          controlId: block.id,
          keepContent: false,
        })
      ),
    ],
    handled: controls,
    kept: unwrapped?.paragraphIds[0] ?? null,
    gone: [...left, ...removed.flatMap((block) => block.paragraphIds)],
  };
}
