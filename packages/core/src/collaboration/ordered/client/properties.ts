// The engine's property ops REPLACE a whole set (`setRunProperties`, `setParagraphProperties`,
// `setParagraphMarkProperties`): the editor sends what the run or paragraph has with its press
// merged in.

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import {
  AUTHORABLE_PARAGRAPH_PROPERTIES,
  AUTHORABLE_RUN_PROPERTIES,
  authoredProperties,
} from '../../../store/store/direct-properties.ts';
import type { OoxmlProperty } from '../../../store/store/tree-op-types.ts';
import { ComposeContext, type OrderedViewParagraph } from '../compose/context.ts';
import { faceAfter, paragraphPropertiesAfter } from '../compose/properties.ts';
import type { OrderedChangeProperty } from '../change/types.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from '../model/types.ts';
import {
  attributeOf,
  childWhere,
  isParagraphProperties,
  isWml,
  withChildren,
} from '../ooxml-nodes.ts';
import { OrderedClientRefusal } from './refusal.ts';

function scratch(doc: OrderedDocument, story: OrderedStory): ComposeContext {
  let next = 0;
  return new ComposeContext(doc, story, {}, () => `ordered-client-scratch#${(next += 1)}`);
}

/** The element without its change record: what the item shows now. */
function shown(element: OoxmlElement | null, record: string): OoxmlElement | null {
  if (!element) return null;
  return withChildren(
    element,
    element.children.filter((child) => !isWml(child, record))
  );
}

/** The run properties an item shows now, under its accepted and pending format records. */
export function currentFace(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem
): OoxmlElement | null {
  return shown(faceAfter(scratch(doc, story), item.fmts, item.face), 'rPrChange');
}

/** The paragraph properties a paragraph's mark shows now, under its records. */
export function currentParagraphProperties(
  doc: OrderedDocument,
  story: OrderedStory,
  mark: OrderedItem
): OoxmlElement | null {
  const shell = mark.node;
  const base =
    shell && shell.kind !== 'textValue' ? (childWhere(shell, isParagraphProperties) ?? null) : null;
  return shown(paragraphPropertiesAfter(scratch(doc, story), mark, base), 'pPrChange');
}

const sameProperty = (a: OoxmlProperty, b: OoxmlProperty): boolean => {
  if (a.localName !== b.localName) return false;
  const x = a.attributes ?? {};
  const y = b.attributes ?? {};
  const keys = Object.keys(x);
  return keys.length === Object.keys(y).length && keys.every((key) => x[key] === y[key]);
};

/** The properties that the current set does not already hold. */
function changedOnly(
  current: readonly OoxmlProperty[],
  properties: readonly OoxmlProperty[]
): OoxmlProperty[] {
  return properties.filter(
    (property) => !current.some((existing) => sameProperty(existing, property))
  );
}

/**
 * The authored properties an op leaves out, as removals. A change MERGES its properties into
 * what the item shows, so it names the ones the op leaves out too (`remove: true`): the
 * engine's op takes them off (a
 * toggle off of one's own pending bold drops `w:b`; Clear All Formatting drops a font). A run
 * whose style turns a property on gets its off value from the editor itself, as a property the
 * op names.
 */
function takenOff(
  current: readonly OoxmlProperty[],
  incoming: readonly OoxmlProperty[]
): OrderedChangeProperty[] {
  const named = new Set(incoming.map((property) => property.localName));
  const out: OrderedChangeProperty[] = [];
  for (const { localName } of current) {
    if (named.has(localName)) continue;
    named.add(localName);
    out.push({ localName, remove: true });
  }

  return out;
}

/** A change's properties: flat, with only the attributes the op gave. */
function changeProperties(properties: readonly OoxmlProperty[]): OrderedChangeProperty[] {
  return properties.map((property) =>
    property.attributes === undefined || Object.keys(property.attributes).length === 0
      ? { localName: property.localName }
      : { localName: property.localName, attributes: { ...property.attributes } }
  );
}

/**
 * The properties a `format` change writes for a `setRunProperties` op over these units: the
 * op's own set, and what it takes off any face in the range.
 */
export function runFormatProperties(
  doc: OrderedDocument,
  story: OrderedStory,
  view: OrderedViewParagraph,
  start: number,
  end: number,
  properties: readonly OoxmlProperty[]
): OrderedChangeProperty[] {
  const seen = new Set<OrderedItem>();
  const current: OoxmlProperty[] = [];
  for (let k = start; k < end; k += 1) {
    const item = view.chars[k]?.item;
    if (!item || seen.has(item)) continue;
    seen.add(item);
    const face = currentFace(doc, story, item);
    current.push(...authoredProperties(face ?? undefined, AUTHORABLE_RUN_PROPERTIES));
  }

  return [...changeProperties(properties), ...takenOff(current, properties)];
}

/**
 * The properties a `format` change writes for a `setRunProperties` op that names one run of a
 * field's result (`targetRunIds`), whose properties the editor shows as `face`: the ones the op
 * changes there. The op names the run's own set with the press merged in, and the change's
 * record reaches every result run of the field, so it carries only what the press changes.
 */
export function resultRunFormatProperties(
  face: OoxmlElement | null,
  properties: readonly OoxmlProperty[]
): OrderedChangeProperty[] {
  const own = shown(face, 'rPrChange');
  const current = authoredProperties(own ?? undefined, AUTHORABLE_RUN_PROPERTIES);
  return [...changeProperties(changedOnly(current, properties)), ...takenOff(current, properties)];
}

/** The properties a `mark-format` change writes for a `setParagraphMarkProperties` op. */
export function markFormatProperties(
  doc: OrderedDocument,
  story: OrderedStory,
  mark: OrderedItem,
  properties: readonly OoxmlProperty[]
): OrderedChangeProperty[] {
  const face = currentFace(doc, story, mark);
  const current = authoredProperties(face ?? undefined, AUTHORABLE_RUN_PROPERTIES);
  return [...changeProperties(properties), ...takenOff(current, properties)];
}

/**
 * The properties a `paragraph-property` change writes for a `setParagraphProperties` op: the
 * ones the op changes, and the ones it takes off. Unchanged entries are left out, so a list's
 * `w:numPr` or a paragraph's `w:tabs` the editor sends back as it found them are never written
 * again as flat properties (they have their own ops). Null when the op changes nothing.
 */
export function paragraphPropertyDelta(
  doc: OrderedDocument,
  story: OrderedStory,
  mark: OrderedItem,
  properties: readonly OoxmlProperty[]
): OrderedChangeProperty[] | null {
  const pPr = currentParagraphProperties(doc, story, mark);
  const current = authoredProperties(pPr ?? undefined, AUTHORABLE_PARAGRAPH_PROPERTIES);
  const changed = changedOnly(current, properties);
  for (const property of changed)
    if (property.localName === 'numPr' || property.localName === 'tabs')
      throw new OrderedClientRefusal(
        'ordered-unsupported-op',
        'paragraph-property-structured',
        property.localName
      );
  const written = [...changeProperties(changed), ...takenOff(current, properties)];
  return written.length === 0 ? null : written;
}

/** The list a paragraph shows now (`w:numPr`): its number, if named, and its level. */
export function currentList(
  doc: OrderedDocument,
  story: OrderedStory,
  mark: OrderedItem
): { numId: string | null; level: number } | null {
  const pPr = currentParagraphProperties(doc, story, mark);
  const numPr = pPr ? childWhere(pPr, (child) => isWml(child, 'numPr')) : undefined;
  if (!numPr) return null;
  const val = (name: string) => {
    const child = childWhere(numPr, (c) => isWml(c, name));
    return child ? attributeOf(child, 'val') : undefined;
  };
  const level = Number(val('ilvl') ?? 0);
  return {
    numId: val('numId') ?? null,
    level: Number.isInteger(level) && level >= 0 && level <= 8 ? level : 0,
  };
}

/** The list level a paragraph shows now (`w:numPr/w:ilvl`), or null outside a list. */
export function currentListLevel(
  doc: OrderedDocument,
  story: OrderedStory,
  mark: OrderedItem
): number | null {
  return currentList(doc, story, mark)?.level ?? null;
}
