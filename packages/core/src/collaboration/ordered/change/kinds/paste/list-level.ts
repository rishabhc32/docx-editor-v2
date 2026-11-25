// The `list-level` change kind (group 9.10): Increase Indent to a level a list does not declare.

import { contextOf, suggestionFor } from '../../../edits/context.ts';
import { addListLevel } from '../../../edits/resources.ts';
import { at, fieldPath, int, Invalid, list, oneOf, required, str } from '../../read.ts';
import {
  ORDERED_CHANGE_LIMITS,
  type OrderedChangeAddress,
  type OrderedEditBase,
} from '../../types.ts';
import { NUM_ID } from '../../apply-properties.ts';
import { address } from '../../values.ts';
import { orderedKindSpec } from '../spec.ts';
import { pressKeeps, pressParagraphs, type ParagraphPress } from './paragraph-presses.ts';

/**
 * Declare a list level that a list's definition lacks, in Word's default format for that depth
 * (`ensureNumberingLevel`), and move paragraphs of that list to it. The declaration is a
 * resource of the change's suggestion (design D14); Word records no change in the numbering
 * part, and a tracked paragraph keeps its earlier `w:numPr` in its `w:pPrChange`.
 *
 * @public
 */
export interface OrderedListLevelChange extends OrderedEditBase {
  readonly kind: 'list-level';
  /** `w:num/@w:numId`: the list whose definition takes the level. */
  readonly numId: string;
  /** 0 to 8. */
  readonly level: number;
  /** Which of Word's default formats the level takes: a bullet, or a number. */
  readonly format: 'bullet' | 'ordered';
  /** The marks of the paragraphs that move to the level. */
  readonly marks: readonly OrderedChangeAddress[];
}

export const LIST_LEVEL_SPEC = orderedKindSpec<OrderedListLevelChange>({
  kind: 'list-level',
  fields: ['numId', 'level', 'format', 'marks'],
  read(fields) {
    const numId = str(required(fields, 'numId'), fieldPath(fields, 'numId'), 9);
    if (!NUM_ID.test(numId))
      throw new Invalid('invalid-value', fieldPath(fields, 'numId'), 'a list number');
    return {
      numId,
      level: int(required(fields, 'level'), fieldPath(fields, 'level'), 0, 8),
      format: oneOf(required(fields, 'format'), fieldPath(fields, 'format'), [
        'bullet',
        'ordered',
      ] as const),
      marks: list(
        required(fields, 'marks'),
        fieldPath(fields, 'marks'),
        1,
        ORDERED_CHANGE_LIMITS.targets
      ).map((mark, k) => address(mark, at(fieldPath(fields, 'marks'), k))),
    };
  },
  touches(change, reader) {
    for (const mark of change.marks) reader.item(mark, 'mark');
  },
  apply(change, { edit }) {
    const ctx = contextOf(edit, []);
    const presses = change.marks.map(
      (mark): ParagraphPress => ({ mark, op: { op: 'setListLevel', level: change.level } })
    );

    // Paragraphs at the level already change nothing: no declaration and no record
    if (presses.every((press) => pressKeeps(ctx, press))) {
      return { suggestions: [], replaced: [] };
    }

    const sid = suggestionFor(ctx).id;
    addListLevel(ctx, sid, change);
    return pressParagraphs(edit, ctx, sid, presses);
  },
  policy: { text: false, properties: true },
  publish: 'paragraphs',
});
