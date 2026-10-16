/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/**
 * Table rows, cells and grids in the concurrent singleton rules of `singleton-children.ts`.
 *
 * Two people sizing one table row or shading one cell each give it a `w:trPr` or `w:tcPr`
 * when it had none, or each add a `w:shd` to the `w:tcPr` it had. The tree keeps those
 * containers generic, so validation accepts the copies, but later edits of that table refuse
 * with `duplicate-property-container` or leave the other copy behind, and the saved document
 * is not valid for Word. The rules show one copy of each, as they do for a paragraph's `w:pPr`.
 */

import { WML_NAMESPACE_URI } from '@docx-editor.dev/core/store';
import type { ChildShell } from './singleton-children.ts';

/**
 * The kind the singleton rules know a node by. The tree keeps most WML property containers
 * generic (`w:trPr`, `w:tcPr`), so a generic WML element goes by `wml:` and its local name,
 * the key its children group under too.
 */
export function ruleKindOf(node: Pick<ChildShell, 'kind' | 'namespaceUri' | 'localName'>): string {
  return node.kind === 'generic' && node.namespaceUri === WML_NAMESPACE_URI
    ? `wml:${node.localName}`
    : node.kind;
}

/** `ruleKindOf` by id. Only a generic node's shell is read: a typed node goes by its kind. */
export function ruleKindReader<Id>(
  kindOf: (id: Id) => string | null,
  shellOf: (id: Id) => ChildShell | null
): (id: Id) => string | null {
  return (id) => {
    const kind = kindOf(id);
    const shell = kind === 'generic' ? shellOf(id) : null;
    return shell ? ruleKindOf(shell) : kind;
  };
}

/** The table entries of the rule tables in `singleton-children.ts`, by `ruleKindOf`. */
export const TABLE_RULES = {
  /**
   * `SINGLETON_KINDS`. ECMA-376 `CT_Row`, `CT_Tc` and `CT_TblGrid`: two peers sizing a row,
   * or shading a cell, that had no properties element each give it one.
   */
  singletonKinds: {
    tableRow: ['wml:tblPrEx', 'wml:trPr'],
    tableCell: ['wml:tcPr'],
    tableGrid: ['wml:tblGridChange'],
  },
  /**
   * `UNIQUE_PROPERTY_PARENTS` and `MERGED_KINDS`. Each direct WML child appears once
   * (ECMA-376 `CT_TblPrEx`, `CT_TcPr`, `CT_TcBorders`), and concurrent copies merge: a fill
   * and a vertical alignment given to one cell are both. `CT_TrPrBase` is a repeating choice,
   * but layout and the editor's own edits read and replace only the first of each, so a
   * second `w:trHeight` is a value nobody sees.
   */
  propertyContainers: ['wml:tblPrEx', 'wml:trPr', 'wml:tcPr', 'wml:tcBorders'],
  /** `LEADING_KIND`: the properties of a row and a cell lead them, as `leadingPrecedes` says. */
  leadingKinds: { tableRow: 'wml:trPr', tableCell: 'wml:tcPr' },
} as const;

/**
 * Whether the singleton that leads a `parentKind` has to precede `child`. A paragraph's
 * `w:pPr` and a run's `w:rPr` precede everything. A row's `w:trPr` precedes its cells and a
 * cell's `w:tcPr` its paragraphs and tables (ECMA-376 `CT_Row`, `CT_Tc`), which is where a
 * local edit writes them: a `w:tblPrEx` or a bookmark ahead of one stays where it is, as it
 * is in the author's tree.
 */
export function leadingPrecedes(parentKind: string, child: ChildShell): boolean {
  switch (parentKind) {
    case 'tableRow':
      return child.kind === 'tableCell';
    case 'tableCell':
      return child.kind === 'paragraph' || child.kind === 'table';
    default:
      return true;
  }
}
