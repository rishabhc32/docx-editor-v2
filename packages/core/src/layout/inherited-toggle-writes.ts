// A toggle press that takes a format suggestion back leaves no suggestion behind.
//
// Toggling a run property off writes an explicit off value (`w:b w:val="0"`), because the
// property may come from a style, and removing the direct one would bring the style's back.
// Where no style gives it, the off value changes nothing visible. On text whose `w:rPrChange`
// records no such property, it still differs from that original, so Underline pressed twice in
// suggesting mode left a format suggestion that changed nothing. There, the press removes the
// direct property instead, as Word does: the run lands back on its original and the record
// goes (`withPropertyChangeRecord`). The same holds for a toggle on that would restate what a
// style gives over a recorded off value. Without a pending record the off value stays.

import { findNode } from '../store/package/ooxml-edit.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../store/package/ooxml-tree.ts';
import { withComplexScriptCompanions } from '../store/store/complex-script-companions.ts';
import {
  formattingRunAt,
  mergedProperties,
  propertyContainer,
} from '../store/store/direct-properties.ts';
import type {
  FormattingDisplayMode,
  FormattingRevisionAuthorFilter,
} from '../store/store/formattable-runs.ts';
import type { OoxmlProperty, TreeDocOp } from '../store/store/tree-op-types.ts';
import { createFormattingCellStyles } from './complex-script-table-context.ts';
import { propertiesOf } from './paragraph-flow.ts';
import { resolveRunStyle, runStylesEqual } from './run-style.ts';
import {
  buildStyleCascadeTable,
  cascadeParagraphFormatting,
  cascadeRunProperties,
  type StyleCascadeTable,
} from './style-cascade.ts';

type FormattingCellStyles = ReturnType<typeof createFormattingCellStyles>;

/** The original properties a pending `w:rPrChange` in this `w:rPr` records, or null. */
function recordedOriginal(container: OoxmlNode | undefined): readonly OoxmlProperty[] | null {
  if (!container || container.kind === 'textValue') return null;
  const record = container.children.find(
    (child) =>
      child.kind !== 'textValue' &&
      child.localName === 'rPrChange' &&
      child.namespaceUri === WML_NAMESPACE_URI
  );
  return record ? propertiesOf(propertyContainer(record, 'runProperties', 'rPr')) : null;
}

/**
 * `ops`, a toggle of `localName` over the selection, without the toggled property on each run
 * or paragraph mark whose pending format record does not name it, where the styles alone give
 * the same look.
 *
 * Each `setRunProperties` and `setParagraphMarkProperties` op is resolved through the style
 * cascade twice, with and without the property and its complex-script twin. Other ops pass
 * through as they are.
 */
export function withoutInheritedToggles(
  part: OoxmlPart,
  ops: readonly TreeDocOp[],
  localName: string,
  styles: OoxmlElement | null,
  settings: OoxmlElement | null,
  displayMode: FormattingDisplayMode,
  authorFilter: FormattingRevisionAuthorFilter | undefined
): TreeDocOp[] {
  const names = new Set(
    withComplexScriptCompanions(true, [{ localName }]).map((property) => property.localName)
  );
  // Built at the first op over a pending record, so a press elsewhere costs nothing.
  let cascade: { table: StyleCascadeTable; cellStyleOf: FormattingCellStyles } | undefined;
  return ops.map((op) => {
    if (op.op !== 'setRunProperties' && op.op !== 'setParagraphMarkProperties') return op;
    if (!op.properties.some((property) => names.has(property.localName))) return op;
    const paragraph = findNode(part, op.paragraphId);
    if (paragraph?.kind !== 'paragraph') return op;
    const pPr = propertyContainer(paragraph, 'paragraphProperties', 'pPr');
    let owner: OoxmlNode | null | undefined = pPr;
    if (op.op === 'setRunProperties') {
      // An edit covers one run's share of the range, so the run its end closes is that run.
      const target = op.targetRunIds?.[0];
      owner = target
        ? findNode(part, target)
        : formattingRunAt(paragraph, op.end, displayMode, authorFilter);
    }
    const container = propertyContainer(owner, 'runProperties', 'rPr');
    const original = recordedOriginal(container);
    if (!original || original.some((property) => names.has(property.localName))) return op;
    if (!cascade) {
      const table = buildStyleCascadeTable(styles, undefined, settings);
      cascade = { table, cellStyleOf: createFormattingCellStyles(table) };
    }
    const { table, cellStyleOf } = cascade;
    const inherited = cascadeParagraphFormatting(
      table,
      pPr,
      cellStyleOf(part, paragraph, displayMode, authorFilter)
    ).runProperties;
    // The write replaces the properties it names; the rest of the container, `w:rStyle`
    // among them, stays and joins the cascade.
    const written = mergedProperties(propertiesOf(container), op.properties);
    const unwritten = written.filter((property) => !names.has(property.localName));
    const same = runStylesEqual(
      resolveRunStyle(cascadeRunProperties(inherited, written, table)),
      resolveRunStyle(cascadeRunProperties(inherited, unwritten, table))
    );
    if (!same) return op;
    return {
      ...op,
      properties: op.properties.filter((property) => !names.has(property.localName)),
    };
  });
}
