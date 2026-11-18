// A TRACKED `insertFragment`: a paste with Track Changes on, written the way Word writes one.
//
// Word records the whole paste as ONE insertion (ECMA-376 Part 1, §17.13.5):
//   - every pasted run sits in a `w:ins` (§17.13.5.18) and keeps its own `w:rPr` plain;
//   - every pasted paragraph mark is `w:pPr/w:rPr/w:ins` (§17.13.5.20), and the paragraph's
//     own `w:pPr` (style, numbering) stays plain, with no `w:pPrChange`: the paragraph did not
//     exist before, so it had no earlier properties to record;
//   - every pasted table row is `w:trPr/w:ins` (§17.13.5.17), with `w:tblPr`, `w:tblGrid` and
//     `w:tcPr` plain, and the cell paragraphs are marked like any other pasted paragraph.
// Media, styles, numbering and relationships are package resources, not revisions.

import { createMarkContainerIdAllocator } from '../package/ooxml-edit.ts';
import { isInertMarker } from './revision-marker-content.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { build, revisionAttributes } from './tree-op-tracked-builders.ts';
import { patchTrPrChild } from './tree-op-table-properties.ts';
import { isWmlNamed } from './tree-op-tracked.ts';
import type { RevisionAttributionInput } from './tree-op-revision-attribution.ts';
import type { TreeOpRejection } from './tree-op-types.ts';

/** Every WordprocessingML element that records a tracked change or a part of one. */
export const REVISION_MARKUP: ReadonlySet<string> = new Set([
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'delText',
  'delInstrText',
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
  'customXmlInsRangeStart',
  'customXmlInsRangeEnd',
  'customXmlDelRangeStart',
  'customXmlDelRangeEnd',
  'customXmlMoveFromRangeStart',
  'customXmlMoveFromRangeEnd',
  'customXmlMoveToRangeStart',
  'customXmlMoveToRangeEnd',
  'rPrChange',
  'pPrChange',
  'trPrChange',
  'tcPrChange',
  'tblPrChange',
  'tblGridChange',
  'tblPrExChange',
  'sectPrChange',
  'numberingChange',
  'cellIns',
  'cellDel',
  'cellMerge',
]);

/**
 * Run-level changes that the paste's insertion absorbs: their content becomes the paster's.
 *
 * A fragment copied from a document that did not track changes brings other people's changes
 * with it, and Word keeps them under the paste's insertion (Word 16.114 for Mac, samples
 * `~$paste-out-e1b.docx` and `~$paste-out-e1g.docx` of the paste evidence). Another author's
 * insertion or move destination is absorbed, and so are a paragraph mark's, a list number's
 * and a cell's insertion. Property records stay where they are.
 */
const ABSORBED: ReadonlySet<string> = new Set(['ins', 'moveTo']);

/** A move destination's range markers, which go with the destination it absorbs. */
const ABSORBED_MARKERS: ReadonlySet<string> = new Set(['moveToRangeStart', 'moveToRangeEnd']);

/**
 * Run-level changes that stay nested in the paste's insertion, with their own author and date:
 * another author's deletion, and a move's source.
 */
const NESTED: ReadonlySet<string> = new Set(['del', 'moveFrom']);

/**
 * Revision markers no tracked paste can place: a custom XML element's tracked range. Word
 * writes them only around tracked custom XML, which it no longer edits (since Word 2010).
 */
const UNPLACEABLE_MARKUP: ReadonlySet<string> = new Set(
  [...REVISION_MARKUP].filter((name) => name.startsWith('customXml'))
);

/**
 * Paragraph-content containers whose runs take the wrapper inside them. `w:ins` is
 * `CT_RunTrackChange`, whose content is RUN content, and a hyperlink or a simple field is
 * paragraph content, so the wrapper goes around their runs. A run-level content control is
 * run content: the wrapper takes it whole, the shape a tracked content-control insertion
 * writes, so rejecting the paste removes the control and not only its text.
 */
const RUN_CONTAINERS: ReadonlySet<string> = new Set([
  'hyperlink',
  'fldSimple',
  'smartTag',
  'customXml',
  'dir',
  'bdo',
]);

/** Property children of those containers, which stay where they are. */
const CONTAINER_PROPERTIES: ReadonlySet<string> = new Set(['fldData', 'smartTagPr', 'customXmlPr']);

/**
 * Block-level containers whose blocks are tracked one by one.
 *
 * A block-level content control or custom XML element is not: no revision markup can carry
 * the wrapper itself, so rejecting the paste would empty it and leave it behind between the
 * two halves of the host paragraph. Such a fragment is refused rather than proposed half way.
 */
const BLOCK_CONTAINERS: ReadonlySet<string> = new Set(['tbl', 'tc']);

/** Block-level properties, which stay plain. */
const BLOCK_PROPERTIES: ReadonlySet<string> = new Set(['tblPr', 'tblGrid', 'tblPrEx', 'tcPr']);

const isWml = (node: OoxmlNode, names: ReadonlySet<string>): boolean =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && names.has(node.localName);

/** Whether any node under `node`, itself included, is a WordprocessingML element of `names`. */
function holdsMarkup(node: OoxmlNode, names: ReadonlySet<string>): boolean {
  if (node.kind === 'textValue') return false;
  if (isWml(node, names)) return true;
  return node.children.some((child) => holdsMarkup(child, names));
}

/** Whether any node of `blocks`, or under them, records a tracked change. */
export function holdsRevisionMarkup(blocks: readonly OoxmlNode[]): boolean {
  return blocks.some((block) => holdsMarkup(block, REVISION_MARKUP));
}

/**
 * Paragraph content with other authors' insertions absorbed: an insertion's or a move
 * destination's content takes its place, and the destination's range markers go with it.
 */
function absorbed(children: readonly OoxmlNode[]): OoxmlNode[] {
  const out: OoxmlNode[] = [];
  for (const child of children) {
    if (isWml(child, ABSORBED_MARKERS)) continue;
    if (!isWml(child, ABSORBED)) {
      out.push(child);
      continue;
    }

    // One at a time: a hostile fragment can hold a long run of children.
    for (const inner of absorbed(child.kind === 'textValue' ? [] : child.children)) {
      out.push(inner);
    }
  }

  return out;
}

/** How the fragment's ids and attribution reach the writers below. */
export interface FragmentWriter {
  readonly revisionId: string;
  readonly revision: RevisionAttributionInput;
  /** Ids for the wrappers and markers the transform adds. */
  readonly nextId: () => string;
  /** Ids for the `w:pPr` and `w:rPr` a mark stamp creates, which resolving the mark removes. */
  readonly markId: () => string;
}

type Tracked<T> = { readonly ok: true; readonly value: T } | { readonly ok: false };

const UNTRACKABLE: Tracked<never> = { ok: false };

/** A `w:ins` over run content. */
function insertion(writer: FragmentWriter, children: readonly OoxmlNode[]): OoxmlNode {
  return build(
    writer.nextId(),
    'revisionInsert',
    'ins',
    revisionAttributes(writer.revisionId, writer.revision),
    children
  );
}

/**
 * Paragraph content with its runs in `w:ins`. Runs, run-level content controls and other
 * authors' deletions are wrapped, and other authors' insertions absorbed; position markers
 * between them travel inside the same wrapper, so rejecting the paste takes a pasted bookmark's
 * two ends together. A wrapper is written only around real content.
 */
function trackedInline(
  children: readonly OoxmlNode[],
  writer: FragmentWriter | null
): Tracked<readonly OoxmlNode[]> {
  const out: OoxmlNode[] = [];
  let pending: OoxmlNode[] = [];
  let pendingContent = false;
  const flush = (): void => {
    if (pending.length === 0) return;
    if (pendingContent && writer) {
      out.push(insertion(writer, pending));
    } else {
      // One at a time: a hostile fragment can hold a long run of markers.
      for (const node of pending) {
        out.push(node);
      }
    }

    pending = [];
    pendingContent = false;
  };
  for (const child of absorbed(children)) {
    if (child.kind === 'textValue' || isInertMarker(child)) {
      pending.push(child);
      continue;
    }

    // Runs, run-level controls and other authors' deletions go in the wrapper whole
    if (isWmlNamed(child, 'r') || isWmlNamed(child, 'sdt') || isWml(child, NESTED)) {
      pending.push(child);
      pendingContent = true;
      continue;
    }

    if (isWml(child, RUN_CONTAINERS)) {
      flush();
      const content = child.children.filter((node) => !isWml(node, CONTAINER_PROPERTIES));
      const properties = child.children.filter((node) => isWml(node, CONTAINER_PROPERTIES));
      const inner = trackedInline(content, writer);
      if (!inner.ok) return UNTRACKABLE;
      out.push({ ...child, children: [...properties, ...inner.value] } as OoxmlNode);
      continue;
    }

    // A subdocument, an equation, or an element this engine does not model: no `w:ins`
    // here can carry it, and leaving it bare would land it untracked inside the proposal.
    return UNTRACKABLE;
  }

  flush();
  return { ok: true, value: out };
}

/** A property container without the records of `names`: insertions the paste absorbs. */
function withoutRecords(node: OoxmlElement, names: ReadonlySet<string>): OoxmlElement {
  if (!node.children.some((child) => isWml(child, names))) return node;
  return {
    ...node,
    children: node.children.filter((child) => !isWml(child, names)),
  } as OoxmlElement;
}

/** A list number's insertion (`w:numPr/w:ins`), absorbed by the pasted paragraph's. */
const LIST_INSERTION: ReadonlySet<string> = new Set(['ins']);

/** `w:pPr` with its list number's insertion absorbed; the same node when it has none. */
function withoutListInsertion(pPr: OoxmlElement): OoxmlElement {
  const children = pPr.children.map((child) => {
    if (isWmlNamed(child, 'numPr') && child.kind !== 'textValue') {
      return withoutRecords(child, LIST_INSERTION);
    }

    return child;
  });
  const changed = children.some((child, k) => child !== pPr.children[k]);
  return changed ? ({ ...pPr, children } as OoxmlElement) : pPr;
}

/** A cell's insertion (`w:tcPr/w:cellIns`), absorbed by the pasted row's. */
const CELL_INSERTION: ReadonlySet<string> = new Set(['cellIns']);

/** The `CT_PPr` children that may follow `w:rPr`. */
const isTrailingParagraphProperty = (node: OoxmlNode): boolean =>
  isWmlNamed(node, 'sectPr') || isWmlNamed(node, 'pPrChange');

/**
 * `w:pPr` with the paragraph's mark stamped inserted: `w:pPr/w:rPr/w:ins`, first in the
 * `w:rPr` as `EG_ParaRPrTrackChanges` requires, and the `w:rPr` after the base properties as
 * `CT_PPr` requires. An insertion already on the mark is replaced; the mark's formatting
 * stays. Containers the stamp creates take the `mark` id family, so accepting removes them.
 */
export function withInsertedMark(pPr: OoxmlNode | undefined, writer: FragmentWriter): OoxmlElement {
  const marker = build(
    writer.nextId(),
    'generic',
    'ins',
    revisionAttributes(writer.revisionId, writer.revision),
    []
  );
  const properties = pPr && pPr.kind !== 'textValue' ? pPr : undefined;
  const rPr = properties?.children.find((child) => isWmlNamed(child, 'rPr'));
  const markRun =
    rPr && rPr.kind !== 'textValue'
      ? ({
          ...rPr,
          // A mark's own insertion or move destination is absorbed by the paste's
          children: [marker, ...rPr.children.filter((child) => !isWml(child, ABSORBED))],
        } as OoxmlElement)
      : build(writer.markId(), 'runProperties', 'rPr', [], [marker]);
  const rest = properties?.children.filter((child) => child !== rPr) ?? [];
  const children = [
    ...rest.filter((child) => !isTrailingParagraphProperty(child)),
    markRun,
    ...rest.filter(isTrailingParagraphProperty),
  ];
  return properties
    ? ({ ...properties, children } as OoxmlElement)
    : build(writer.markId(), 'paragraphProperties', 'pPr', [], children);
}

function trackedParagraph(
  paragraph: OoxmlElement,
  writer: FragmentWriter | null
): Tracked<OoxmlNode> {
  const pPr = paragraph.children.find((child) => child.kind === 'paragraphProperties');
  const inline = trackedInline(
    paragraph.children.filter((child) => child !== pPr),
    writer
  );
  if (!inline.ok) return UNTRACKABLE;
  if (!writer) return { ok: true, value: paragraph };

  // The paragraph is new: a list number another author inserted on it is the paste's now
  const ownPPr = pPr && pPr.kind !== 'textValue' ? withoutListInsertion(pPr) : pPr;
  return {
    ok: true,
    value: {
      ...paragraph,
      children: [withInsertedMark(ownPPr, writer), ...inline.value],
    } as OoxmlNode,
  };
}

/** A row proposed as inserted: `w:trPr/w:ins`, in `CT_TrPr` order. */
function trackedRowProperties(row: OoxmlElement, writer: FragmentWriter): Tracked<OoxmlElement> {
  const marker = build(
    writer.nextId(),
    'generic',
    'ins',
    revisionAttributes(writer.revisionId, writer.revision),
    []
  );
  const existing = row.children.find((child) => isWmlNamed(child, 'trPr'));
  const trPr =
    existing && existing.kind !== 'textValue'
      ? existing
      : build(writer.nextId(), 'generic', 'trPr', [], []);
  const patched = patchTrPrChild(trPr, marker);
  if (!patched.ok) return UNTRACKABLE;

  // `CT_Row` opens with `w:tblPrEx`, then `w:trPr`, then the cells.
  const children = existing
    ? row.children.map((child) => (child === existing ? patched.container : child))
    : [
        ...row.children.filter((child) => isWmlNamed(child, 'tblPrEx')),
        patched.container,
        ...row.children.filter((child) => !isWmlNamed(child, 'tblPrEx')),
      ];
  return { ok: true, value: { ...row, children } as OoxmlElement };
}

function trackedBlock(node: OoxmlNode, writer: FragmentWriter | null): Tracked<OoxmlNode> {
  if (writer && isWmlNamed(node, 'tcPr') && node.kind !== 'textValue') {
    return { ok: true, value: withoutRecords(node, CELL_INSERTION) };
  }

  if (node.kind === 'textValue' || isInertMarker(node) || isWml(node, BLOCK_PROPERTIES)) {
    return { ok: true, value: node };
  }

  if (node.kind === 'paragraph') return trackedParagraph(node, writer);
  const row = isWmlNamed(node, 'tr');
  if (!row && !isWml(node, BLOCK_CONTAINERS)) return UNTRACKABLE;

  // A row's own properties are tracked here, and its cells below like any other block.
  const stamped = row && writer ? trackedRowProperties(node, writer) : { ok: true, value: node };
  if (!stamped.ok) return UNTRACKABLE;
  const children: OoxmlNode[] = [];
  for (const child of stamped.value.children) {
    if (row && isWmlNamed(child, 'trPr')) {
      children.push(child);
      continue;
    }

    const tracked = trackedBlock(child, writer);
    if (!tracked.ok) return UNTRACKABLE;
    children.push(tracked.value);
  }

  return writer ? { ok: true, value: { ...stamped.value, children } as OoxmlNode } : stamped;
}

/**
 * Why a fragment cannot be written as a tracked insertion, or null when it can.
 *
 * Run by validation, after the shape and budget checks have bounded the walk.
 */
export function trackedFragmentRefusal(blocks: readonly OoxmlNode[]): TreeOpRejection | null {
  if (blocks.some((block) => holdsMarkup(block, UNPLACEABLE_MARKUP))) {
    return 'fragment-holds-revisions';
  }

  for (const block of blocks) {
    if (!trackedBlock(block, null).ok) return 'fragment-untrackable-content';
  }

  return null;
}

/** The writer for one tracked fragment: one revision id, and the part's id families. */
export function fragmentWriter(
  part: OoxmlPart,
  revisionId: string,
  revision: RevisionAttributionInput,
  nextId: () => string
): FragmentWriter {
  return { revisionId, revision, nextId, markId: createMarkContainerIdAllocator(part) };
}

/**
 * The fragment's blocks as ONE tracked insertion, or null when one of them cannot be. The
 * transform runs BEFORE the blocks land. The landing then places the marks by its usual rule:
 * the first result paragraph ends on the fragment's first (inserted) mark, and the host's
 * original mark ends the last one, so rejecting every inserted mark joins the host paragraph
 * back together under its own properties.
 */
export function trackedFragmentBlocks(
  blocks: readonly OoxmlNode[],
  writer: FragmentWriter
): readonly OoxmlNode[] | null {
  const out: OoxmlNode[] = [];
  for (const block of blocks) {
    const tracked = trackedBlock(block, writer);
    if (!tracked.ok) return null;
    out.push(tracked.value);
  }

  return out;
}
