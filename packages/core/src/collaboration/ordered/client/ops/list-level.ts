import {
  ensureNumberingLevel,
  numberingLevelFormat,
} from '../../../../store/package/numbering-part.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import { addressOf, type ListParagraph } from '../paragraphs.ts';
import { currentList } from '../properties.ts';
import type { OrderedChangeBody, OrderedTranslateBatch } from './types.ts';

type SetListLevelOp = Extract<TreeDocOp, { readonly op: 'setListLevel' }>;

/** The list a paragraph moves within, and the format of the level it moves from. */
function listOf(
  batch: OrderedTranslateBatch,
  p: ListParagraph
): { numId: string; format: 'bullet' | 'ordered' } | null {
  const list = currentList(batch.doc, p.story, p.mark);
  const shown = batch.shown();
  if (!list?.numId || !shown) return null;
  const from = numberingLevelFormat(shown, list.numId, list.level);
  return { numId: list.numId, format: from === 'bullet' ? 'bullet' : 'ordered' };
}

/**
 * The `list-level` change of a batch's first op that moves a paragraph to a level its list does
 * not declare (Increase Indent), with every paragraph of the batch that moves to that level;
 * nothing for the batch's later ops to the same level; null for an op whose level the list
 * declares, which is a paragraph press as any other. The editor declares the level before the
 * batch in Word's default format for that depth (`ensureNumberingLevel`).
 */
export function listLevelChanges(
  op: SetListLevelOp,
  batch: OrderedTranslateBatch
): OrderedChangeBody[] | null {
  const shown = batch.shown();
  const p = batch.paragraph(op.paragraphId);
  const list = listOf(batch, p);
  if (!shown || !list) return null;

  // The level is declared when the engine's declaration leaves the package as it is
  const declared = ensureNumberingLevel(shown, list.numId, op.level, list.format);
  if (declared === null || declared === shown) return null;
  const moves = (other: TreeDocOp): other is SetListLevelOp =>
    other.op === 'setListLevel' &&
    other.level === op.level &&
    listOf(batch, batch.paragraph(other.paragraphId))?.numId === list.numId;

  // Only the batch's first op to the level sends the change, for all of them
  if (batch.ops.find(moves) !== op) return [];
  const same = batch.ops.filter(moves);
  return [
    {
      kind: 'list-level',
      part: p.story.name,
      numId: list.numId,
      level: op.level,
      format: list.format,
      marks: same.map((other) => addressOf(batch.paragraph(other.paragraphId).mark)),
    },
  ];
}
