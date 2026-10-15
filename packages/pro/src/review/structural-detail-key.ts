/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The packaged sentence for one structural change, shared so the React and Vue twins cannot
// drift.

import type { TranslationKey } from '@docx-editor.dev/i18n';
import type { ReviewRevisionItem } from '@docx-editor.dev/core/contracts/editor';

/**
 * The packaged sentence for one structural change. Cell changes that make up whole table
 * columns (`structuralColumns`) read as the columns, which is how a reviewer thinks of a
 * proposed column.
 */
export function structuralDetailKey(
  kind: NonNullable<ReviewRevisionItem['structuralChanges']>[number],
  columns?: number
): TranslationKey {
  if (columns !== undefined && kind === 'cellDelete') {
    return columns > 1
      ? 'revisions.structuralDetails.columnsDelete'
      : 'revisions.structuralDetails.columnDelete';
  }
  if (columns !== undefined && kind === 'cellInsert') {
    return columns > 1
      ? 'revisions.structuralDetails.columnsInsert'
      : 'revisions.structuralDetails.columnInsert';
  }
  return `revisions.structuralDetails.${kind}`;
}
