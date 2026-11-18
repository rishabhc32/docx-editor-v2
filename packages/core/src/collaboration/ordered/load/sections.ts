// Sections at load (task 9.8, `model/sections.ts`), and what the file says about headers and
// footers outside the story text.

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { resolveRelationship } from '../../../store/package/relationships.ts';
import { restoredSectionProperties } from '../../../store/store/revision-property-records.ts';
import { createOrderedItem } from '../model/document.ts';
import { REVISION_ROLE } from '../model/revision-ids.ts';
import { implicitSectionSpace } from '../model/sections.ts';
import type { OrderedDocument, OrderedElementRecord, OrderedItem } from '../model/types.ts';
import {
  attributeOf,
  childWhere,
  firstElement,
  isParagraphProperties,
  isWml,
  withChildren,
} from '../ooxml-nodes.ts';
import type { LoadState } from './state.ts';

const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SETTINGS = `${OFFICE}/settings`;

/** What the loaded file says about headers and footers, outside the story text. */
export interface OrderedLoadedFurniture {
  /** The main part's header and footer relationships: id to part name. */
  readonly parts: ReadonlyMap<
    string,
    { readonly part: string; readonly kind: 'header' | 'footer' }
  >;
  /** Whether the file's settings separate even and odd pages (`w:evenAndOddHeaders`). */
  readonly evenAndOddHeaders: boolean;
}

const FURNITURE = new WeakMap<OrderedDocument, OrderedLoadedFurniture>();
const NONE: OrderedLoadedFurniture = { parts: new Map(), evenAndOddHeaders: false };

/** What the loaded file says about headers and footers (nothing for a document with no package). */
export function loadedFurnitureOf(doc: OrderedDocument): OrderedLoadedFurniture {
  return FURNITURE.get(doc) ?? NONE;
}

const ON = new Set(['true', '1', 'on']);
const FURNITURE_KINDS: ReadonlyMap<string, 'header' | 'footer'> = new Map([
  [`${OFFICE}/header`, 'header'],
  [`${OFFICE}/footer`, 'footer'],
]);

/**
 * Record the main part's header and footer relationships and the file's even and odd pages,
 * for the header and footer edits: a section's reference names a part through them.
 */
export function recordFurniture(doc: OrderedDocument, pkg: OoxmlPackage): void {
  FURNITURE.set(doc, {
    parts: furnitureParts(doc, pkg),
    evenAndOddHeaders: evenAndOddHeadersOf(pkg),
  });
}

/** The main part's header and footer relationships to a loaded story: id to part name. */
function furnitureParts(doc: OrderedDocument, pkg: OoxmlPackage): OrderedLoadedFurniture['parts'] {
  const parts = new Map<string, { part: string; kind: 'header' | 'footer' }>();
  for (const record of pkg.relationships.get(pkg.mainDocumentPart) ?? []) {
    const kind = FURNITURE_KINDS.get(record.type);
    const resolved = resolveRelationship(record);
    if (kind === undefined || resolved.mode !== 'Internal' || !resolved.target.ok) {
      continue;
    }

    if (doc.stories.has(resolved.target.partName)) {
      parts.set(record.id, { part: resolved.target.partName, kind });
    }
  }

  return parts;
}

/** Whether the file's settings separate even and odd pages (`w:evenAndOddHeaders`). */
function evenAndOddHeadersOf(pkg: OoxmlPackage): boolean {
  let evenAndOddHeaders = false;
  for (const record of pkg.relationships.get(pkg.mainDocumentPart) ?? []) {
    const resolved = resolveRelationship(record);
    if (record.type !== SETTINGS || resolved.mode !== 'Internal' || !resolved.target.ok) {
      continue;
    }

    const settings = pkg.parts.get(resolved.target.partName);
    const flag = settings?.root.children.find((c) => isWml(c, 'evenAndOddHeaders'));
    if (flag && flag.kind !== 'textValue') {
      const value = attributeOf(flag, 'val');
      evenAndOddHeaders = value === undefined || ON.has(value);
    }
  }

  return evenAndOddHeaders;
}

/**
 * A section with its file's change record (`w:sectPrChange`) as a suggestion: the section shows
 * the properties before it (as the engine's reject leaves them, `restoredSectionProperties`),
 * and the record on the item gives the properties after it, with the file's record written back
 * while it stays alone. A decision is then a status, as for every other change record.
 */
function splitRecord(
  state: LoadState,
  sectPr: OoxmlElement
): { readonly before: OoxmlElement; readonly record: OrderedElementRecord } | null {
  const change = childWhere(sectPr, (c) => isWml(c, 'sectPrChange'));
  if (!change) return null;
  const current = withChildren(
    sectPr,
    sectPr.children.filter((c) => c !== change)
  );
  const recorded = firstElement(change);
  const before = withChildren(
    sectPr,
    restoredSectionProperties(current.children, recorded?.children ?? [])
  );
  const suggestion = state.revisions.suggestionFor(change, REVISION_ROLE.sectionProperties);
  return {
    before,
    record: {
      kind: 'element',
      suggestion,
      seq: state.revisions.seqOf(suggestion),
      target: 'sectPr',
      element: current,
      fileRecord: change,
      recordUnchanged: true,
    },
  };
}

/** A mark whose paragraph ends a section with a file's change record. */
function withMarkRecord(state: LoadState, item: OrderedItem): OrderedItem {
  const shell = item.node;
  if (!shell || shell.kind === 'textValue') return item;
  const pPr = childWhere(shell, isParagraphProperties);
  const sectPr = pPr ? childWhere(pPr, (c) => isWml(c, 'sectPr')) : undefined;
  const split = pPr && sectPr ? splitRecord(state, sectPr) : null;
  if (!pPr || !sectPr || !split) return item;
  const swap = (n: OoxmlNode): OoxmlNode => (n === sectPr ? split.before : n);
  const nextPPr = withChildren(pPr, pPr.children.map(swap));
  const node = withChildren(
    shell,
    shell.children.map((c) => (c === pPr ? nextPPr : c))
  );
  return { ...item, node, props: [...item.props, split.record] };
}

/**
 * The items of a story with its sections' file records as suggestions, and the main story's
 * body section item when the file wrote none, at a fixed address, so a change to the last
 * section names it alike on every copy.
 */
export function withLoadedSections(
  state: LoadState,
  part: OoxmlPart,
  items: readonly OrderedItem[]
): OrderedItem[] {
  let hasBodySection = false;
  const out = items.map((item) => {
    if (item.kind === 'mark') return withMarkRecord(state, item);
    if (item.kind !== 'zero' || item.block !== true || !item.node || !isWml(item.node, 'sectPr')) {
      return item;
    }

    hasBodySection = true;
    const split = item.node.kind === 'textValue' ? null : splitRecord(state, item.node);
    return split ? { ...item, node: split.before, props: [...item.props, split.record] } : item;
  });
  if (!hasBodySection && part.root.localName === 'document') {
    out.push(
      createOrderedItem({
        space: implicitSectionSpace(state.doc.file),
        kind: 'zero',
        block: true,
        node: null,
      })
    );
  }

  return out;
}
