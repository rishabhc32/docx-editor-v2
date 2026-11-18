// Properties a file records a change of (`w:rPrChange`, `w:pPrChange`, `w:tblPrChange`, ...):
// the current properties are the suggestion's, the record holds the old ones. The item keeps
// the old ones as its own and carries the suggestion as a format or property record.

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { validNumberingRevision } from '../../../store/store/revision-property-records.ts';
import { REVISION_ROLE } from '../model/revision-ids.ts';
import type {
  OrderedFormatRecord,
  OrderedPropertyRecord,
  OrderedPropertyTarget,
} from '../model/types.ts';
import {
  childNamed,
  firstElement,
  isRunProperties,
  isWml,
  withChildren,
  withoutMarkers,
} from '../ooxml-nodes.ts';
import type { LoadState } from './state.ts';

/** A run's (or a mark's) face, and the format record a file gives it. */
export function faceOf(
  state: LoadState,
  rPr: OoxmlElement | undefined
): { face: OoxmlElement | null; fmts: OrderedFormatRecord[] } {
  if (!rPr) return { face: null, fmts: [] };
  const current = withoutMarkers(rPr);
  const record = childNamed(rPr, 'rPrChange');
  if (!record) return { face: current, fmts: [] };
  const suggestion = state.revisions.suggestionFor(record, REVISION_ROLE.format);
  const old = firstElement(record);
  return {
    face: old ? withChildren(current, old.children) : null,
    fmts: [
      {
        kind: 'face',
        id: `f:${record.id}`,
        suggestion,
        seq: state.revisions.seqOf(suggestion),
        face: current,
        fileRecord: record,
      },
    ],
  };
}

const TARGETS: ReadonlySet<string> = new Set<OrderedPropertyTarget>([
  'pPr',
  'tblPr',
  'tblGrid',
  'tblPrEx',
  'trPr',
  'tcPr',
]);

const isTarget = (name: string): name is OrderedPropertyTarget => TARGETS.has(name);

/**
 * A container's shell (the children `keep` selects), with each kept properties child's change
 * record turned into a property record. `markersOf` names the child whose revision markers
 * became the container's labels (`w:trPr`, `w:tcPr`): they leave its shell. `gridGesture` is
 * the record whose suggestion a table's `w:tblGridChange` joins (`gridGestureOf`).
 */
export function containerShell(
  state: LoadState,
  element: OoxmlElement,
  keep: (child: OoxmlNode) => boolean,
  markersOf?: string,
  gridGesture?: OoxmlElement
): { shell: OoxmlElement; props: OrderedPropertyRecord[] } {
  const props: OrderedPropertyRecord[] = [];
  const children: OoxmlNode[] = [];
  for (const child of element.children) {
    if (!keep(child)) continue;
    if (child.kind === 'textValue') {
      children.push(child);
      continue;
    }

    const name = child.localName;
    const record = isTarget(name) ? childNamed(child, `${name}Change`) : undefined;
    if (!record || !isTarget(name)) {
      if (markersOf !== name) {
        children.push(child);
        continue;
      }

      // Properties that held only the container's revision markers go with them; compose
      // writes them again while a marker shows
      const kept = withoutMarkers(child);
      if (kept !== child && !kept.children.some((c) => c.kind !== 'textValue')) continue;
      children.push(kept);
      continue;
    }

    const suggestion = state.revisions.suggestionFor(
      name === 'tblGrid' && gridGesture ? gridGesture : record,
      REVISION_ROLE.tableProperties
    );
    const current = withoutMarkers(
      withChildren(
        child,
        child.children.filter((c) => c !== record)
      )
    );
    props.push({
      kind: 'element',
      suggestion,
      seq: state.revisions.seqOf(suggestion),
      target: name,
      element: current,
      fileRecord: record,
      recordUnchanged: true,
    });

    // The container keeps the old properties as its own
    const old = firstElement(record);
    children.push(old ? withChildren(child, old.children) : withChildren(child, []));
  }

  return { shell: withChildren(element, children), props };
}

/**
 * A file's own numbering records in a paragraph's `w:numPr`, the ones the engine decides
 * (`validNumberingRevision`): a tracked list membership (`w:ins`) and a first-edition
 * `w:numberingChange`. Each is a suggestion whose accept and reject both remove the record and
 * keep the list, as Word retains `w:numPr` either way (Word for Mac on a `w:numberingChange`,
 * 2026-10-29). The record stays in the paragraph's properties, so a later list edit keeps,
 * drops or copies it as the engine's does; compose removes it once its suggestion is decided
 * (`withoutDecidedNumbering`). A malformed one stays file content, as the engine leaves it.
 */
function numberingRecords(state: LoadState, pPr: OoxmlElement): OrderedPropertyRecord[] {
  const numPr = childNamed(pPr, 'numPr');
  if (!numPr) return [];
  return numPr.children.flatMap((record): OrderedPropertyRecord[] => {
    if (record.kind === 'textValue' || !validNumberingRevision(record, numPr)) return [];
    const suggestion = state.revisions.suggestionFor(record, REVISION_ROLE.numbering);
    return [
      {
        kind: 'element',
        suggestion,
        seq: state.revisions.seqOf(suggestion),
        target: 'numPr',
        element: null,
        fileRecord: record,
        recordUnchanged: true,
      },
    ];
  });
}

/**
 * A paragraph's own `w:pPr` (the old properties when a file records a change of them, with the
 * current mark run properties and section), and the records as property records.
 */
export function paragraphProperties(
  state: LoadState,
  given: OoxmlElement | undefined
): { own: OoxmlElement | undefined; props: OrderedPropertyRecord[] } {
  if (!given) return { own: undefined, props: [] };
  const pPr = given;
  const numbering = numberingRecords(state, pPr);
  const record = childNamed(pPr, 'pPrChange');
  if (!record) return { own: pPr, props: numbering };
  const suggestion = state.revisions.suggestionFor(record, REVISION_ROLE.paragraphProperties);
  const current = withChildren(
    pPr,
    pPr.children.filter((c) => c !== record)
  );
  const old = firstElement(record);
  const own = withChildren(pPr, [
    ...(old?.children ?? []),
    ...pPr.children.filter((c) => isRunProperties(c) || isWml(c, 'sectPr')),
  ]);
  return {
    own,
    props: [
      ...numbering,
      {
        kind: 'element',
        suggestion,
        seq: state.revisions.seqOf(suggestion),
        target: 'pPr',
        element: current,
        fileRecord: record,
        recordUnchanged: true,
      },
    ],
  };
}
