// Tracked section property writes: `w:sectPrChange` (ECMA-376 Part 1, §17.13.5.32), which
// Word writes for a page setup change and for a section break that changes how the next section
// starts ("Formatted: Section start") while it tracks changes.

import { WML_NAMESPACE_URI, type OoxmlElement, type OoxmlNode } from '../package/ooxml-tree.ts';
import { equivalentNodes } from './ooxml-node-equality.ts';
import { SECTION_PROPERTY_SEQUENCE } from './revision-property-records.ts';
import { attributeValueOf, cloneWithNewIds } from './tree-op-nodes.ts';
import { recordedProperties } from './tree-op-tracked-properties.ts';
import { build, isWmlNamed, revisionAttributes } from './tree-op-tracked.ts';
import type { RevisionAttributionInput } from './tree-op-types.ts';

/**
 * A section property's attributes when the section does not state them (§17.6). A write that
 * adds one of these properties, or one of their attributes, is recorded at this value. A writer
 * that adds another section property must add its absent value here.
 */
const ABSENT_VALUES: ReadonlyMap<string, Readonly<Record<string, string>>> = new Map<
  string,
  Readonly<Record<string, string>>
>([
  ['type', { val: 'nextPage' }],
  ['pgSz', { w: '12240', h: '15840', orient: 'portrait' }],
  [
    'pgMar',
    {
      top: '1440',
      right: '1440',
      bottom: '1440',
      left: '1440',
      header: '720',
      footer: '720',
      gutter: '0',
    },
  ],
  ['titlePg', { val: '0' }],
]);

/** On/off section properties (`CT_OnOff`, §17.17.4): present with no or a true `w:val`, on. */
const ON_OFF: ReadonlySet<string> = new Set(['titlePg']);

/** The `ST_OnOff` values that say off. */
const OFF: ReadonlySet<string> = new Set(['0', 'false', 'off']);

/** `CT_SectPrBase`'s sequence as names: the order the record states its properties in. */
const SEQUENCE: readonly string[] = SECTION_PROPERTY_SEQUENCE;

/** A node's place in `SEQUENCE`; a name it does not list goes after the listed ones. */
function rank(node: OoxmlNode): number {
  if (node.kind === 'textValue') {
    return SEQUENCE.length + 2;
  }

  const index = SEQUENCE.indexOf(node.localName);
  if (index === -1) {
    return SEQUENCE.length + 1;
  }

  return index + 1;
}

const wAttribute = (localName: string, value: string) => ({
  kind: 'genericExtension' as const,
  namespaceUri: WML_NAMESPACE_URI,
  localName,
  prefix: 'w',
  value,
});

/** A section's properties as a record can state them: no references, no record. */
function recordable(children: readonly OoxmlNode[]): OoxmlElement[] {
  return children.filter(
    (child): child is OoxmlElement =>
      child.kind !== 'textValue' &&
      !isWmlNamed(child, 'headerReference') &&
      !isWmlNamed(child, 'footerReference') &&
      !isWmlNamed(child, 'sectPrChange')
  );
}

/**
 * The properties without what states a value the section has when it is absent (a record
 * states `w:orient="portrait"`, the section may leave it out): two sets that say the same.
 */
function withoutAbsentValues(nodes: readonly OoxmlElement[]): OoxmlElement[] {
  return nodes.flatMap((node) => {
    // An on/off property is absent only when it says off; present, it says on with no value too
    if (ON_OFF.has(node.localName)) {
      const value = attributeValueOf(node, 'val', WML_NAMESPACE_URI);
      if (value !== undefined && OFF.has(value)) return [];
      const attributes = node.attributes.filter((a) => a.localName !== 'val');
      return [{ ...node, attributes } as OoxmlElement];
    }

    const absent = ABSENT_VALUES.get(node.localName);
    if (absent === undefined) return [node];
    const attributes = node.attributes.filter((a) => absent[a.localName] !== a.value);
    if (attributes.length === 0 && node.children.length === 0) return [];
    return [{ ...node, attributes } as OoxmlElement];
  });
}

/**
 * Whether two sets of section properties say the same: the same properties, read with the values
 * a section has when it does not state them. References and a change record are not properties.
 */
export function sameSectionProperties(
  left: readonly OoxmlNode[],
  right: readonly OoxmlNode[]
): boolean {
  return equivalentNodes(
    withoutAbsentValues(recordable(left)),
    withoutAbsentValues(recordable(right))
  );
}

/**
 * What a `w:sectPrChange` records for a write from `prior` to `next`: every property `prior`
 * states, and each property and attribute of `ABSENT_VALUES` that `next` adds at its value when
 * absent, in schema order, under new ids. The header and footer references stay out:
 * `CT_SectPrBase` cannot hold them.
 *
 * A reject, in Word as in the engine, lays the recorded properties over the current ones and
 * keeps what the record does not state (`restoredSectionProperties`), so an added value is
 * recorded absent for the reject to take it back too. Word for Mac records only the properties
 * its change touched (`w:pgMar` for a header distance, 2026-10-30, `~$hdrdist-tracked.docx`),
 * which it reads the same way. An on/off property the write turned on is recorded off:
 * `<w:titlePg w:val="0"/>`, as Word records a different first page (`~$titlepg-tracked.docx`).
 */
export function sectionRecordChildren(
  prior: readonly OoxmlNode[],
  next: readonly OoxmlNode[],
  mint: () => string
): OoxmlNode[] {
  const before = recordable(prior);
  const after = recordable(next);
  const out: OoxmlNode[] = [];
  for (const had of before) {
    const now = after.find((child) => child.localName === had.localName);
    const absent = ABSENT_VALUES.get(had.localName) ?? {};
    const added = (now?.attributes ?? []).flatMap((attribute) => {
      const value = absent[attribute.localName];
      return value === undefined || had.attributes.some((a) => a.localName === attribute.localName)
        ? []
        : [wAttribute(attribute.localName, value)];
    });

    // The same element with the attributes it gained: an attribute list is any element's
    out.push(
      cloneWithNewIds({ ...had, attributes: [...had.attributes, ...added] } as OoxmlElement, mint)
    );
  }

  for (const now of after) {
    const absent = ABSENT_VALUES.get(now.localName);
    if (absent === undefined || before.some((child) => child.localName === now.localName)) continue;

    // An on/off property states its value: present with no `w:val` it would say on
    const attributes = Object.entries(absent)
      .filter(
        ([name]) => ON_OFF.has(now.localName) || now.attributes.some((a) => a.localName === name)
      )
      .map(([name, value]) => wAttribute(name, value));
    out.push(build(mint(), 'generic', now.localName, attributes, []));
  }

  return out.sort((a, b) => rank(a) - rank(b));
}

/**
 * A section's rewritten children with the write recorded as `w:sectPrChange`, its last child
 * (`CT_SectPr`). `prior` is what the section held before the write (its record, if one stands,
 * included), `next` what the write made of it.
 *
 * The record is the ORIGINAL, so a second write over a pending record keeps what it holds and
 * re-attributes it; and a write that arrives back at the original leaves no record of this
 * author's: the rules of every property record (`tree-op-tracked-properties.ts`).
 */
export function withSectionChangeRecord(options: {
  readonly prior: readonly OoxmlNode[];
  readonly next: readonly OoxmlNode[];
  readonly revision: RevisionAttributionInput;
  readonly mint: () => string;
  readonly nextRevisionId: () => string;
}): readonly OoxmlNode[] {
  const { prior, next, revision, mint, nextRevisionId } = options;
  const existing = prior.find((child) => isWmlNamed(child, 'sectPrChange'));

  // The ORIGINAL: what a standing record holds, else the section this write replaces
  const original = recordable((existing ? recordedProperties(existing) : null) ?? prior);
  const body = next.filter((child) => !isWmlNamed(child, 'sectPrChange'));
  if (sameSectionProperties(body, original)) {
    // Back at the original: this author's own record goes, anyone else's stays
    const owned =
      existing && attributeValueOf(existing, 'author', WML_NAMESPACE_URI) === revision.author;
    return existing !== undefined && !owned ? [...body, existing] : body;
  }

  const record = build(
    mint(),
    'generic',
    'sectPr',
    [],
    sectionRecordChildren(original, body, mint)
  );
  const wrapper = build(
    mint(),
    'generic',
    'sectPrChange',
    revisionAttributes(nextRevisionId(), revision),
    [record]
  );
  return [...body, wrapper];
}
