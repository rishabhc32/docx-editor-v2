// Document/section note-property writes for `setNoteProperties`.

import { createNodeIdAllocator, replaceChildren } from './ooxml-edit.ts';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from './ooxml-tree.ts';
import type { OoxmlPackage } from './ooxml-package.ts';
import { withPart } from './ooxml-package.ts';
import { WML_NAMESPACE_URI } from './ooxml-shared.ts';
import { collectSectionPropertyNodes } from './hf-references.ts';
import { freeRelationshipId, withContentTypeOverride } from './note-lifecycle-shell.ts';
import { withStoryRelationship } from './hf-lifecycle-shell.ts';
import { settingsPartOf } from './note-properties.ts';

const W = WML_NAMESPACE_URI;
const SETTINGS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';
const SETTINGS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';
const SETTINGS_PART = '/word/settings.xml';

export type SetNotePropertiesOp = {
  readonly op: 'setNoteProperties';
  readonly scope: 'document' | 'section';
  readonly sectionIndex?: number;
  readonly footnote?: {
    readonly numFmt?: string;
    readonly numRestart?: string;
    readonly position?: string;
    readonly numStart?: number;
  };
  readonly endnote?: {
    readonly numFmt?: string;
    readonly numRestart?: string;
    readonly position?: string;
    readonly numStart?: number;
  };
};

export function writeDocumentNoteProperties(
  pkg: OoxmlPackage,
  op: SetNotePropertiesOp
): OoxmlPackage | null {
  let next = pkg;
  let settings = settingsPartOf(next);
  if (!settings) {
    const created = ensureSettingsPart(next);
    if (!created) return null;
    next = created;
    settings = settingsPartOf(next);
  }
  if (!settings) return null;

  const rebuilt = rebuildSettings(settings, op);
  if (!rebuilt) return null;
  return withPart(next, rebuilt);
}

function ensureSettingsPart(pkg: OoxmlPackage): OoxmlPackage | null {
  if (pkg.parts.has(SETTINGS_PART)) return pkg;
  const xml = `<w:settings xmlns:w="${W}"></w:settings>`;
  const read = readOoxmlPart(xml, { name: SETTINGS_PART, contentType: SETTINGS_CT });
  if (!read.ok) return null;
  let next = withPart(pkg, read.part);
  const rId = freeRelationshipId(next);
  const related = withStoryRelationship(next, rId, SETTINGS_REL, 'settings.xml');
  if (!related) return null;
  next = related;
  const typed = withContentTypeOverride(next, SETTINGS_PART, SETTINGS_CT);
  return typed;
}

function rebuildSettings(settings: OoxmlPart, op: SetNotePropertiesOp): OoxmlPart | null {
  const nextId = createNodeIdAllocator(settings);
  const children = withNoteProperties(settings.root.children, op, nextId, followsInSettings);
  const replaced = replaceChildren(settings, settings.root.id, children);
  return replaced.ok ? replaced.part : null;
}

export function writeSectionNoteProperties(
  pkg: OoxmlPackage,
  sectionIndex: number,
  op: SetNotePropertiesOp
): OoxmlPackage | null {
  const main = pkg.parts.get(pkg.mainDocumentPart);
  if (!main) return null;
  const sectPrNodes = collectSectionPropertyNodes(main.root);
  if (sectionIndex >= sectPrNodes.length) return null;
  const target = sectPrNodes[sectionIndex];
  if (!target) return null;

  const nextId = createNodeIdAllocator(main);
  const children = withNoteProperties(target.children, op, nextId, followsInSection);
  const replaced = replaceChildren(main, target.id, children);
  if (!replaced.ok) return null;
  return withPart(pkg, replaced.part);
}

/**
 * A section's `w:sectPr` children with note properties written, as `writeSectionNoteProperties`
 * writes them, for a caller that holds the element itself (the ordered model's compose).
 */
export function sectionChildrenWithNoteProperties(
  children: readonly OoxmlNode[],
  op: Pick<SetNotePropertiesOp, 'footnote' | 'endnote'>,
  nextId: () => string
): OoxmlNode[] {
  return withNoteProperties(children, op, nextId, followsInSection);
}

type NoteFields = NonNullable<SetNotePropertiesOp['footnote']>;

const isW = (node: OoxmlNode, localName: string): boolean =>
  node.kind !== 'textValue' && node.namespaceUri === W && node.localName === localName;

/**
 * `CT_Settings` members that follow `w:footnotePr` and `w:endnotePr`. Word reads
 * `settings.xml` in schema order and reports a file whose members are out of sequence as
 * damaged, so the note properties go before the first of these, never at the end.
 */
const SETTINGS_AFTER_NOTE_PROPERTIES: ReadonlySet<string> = new Set([
  'compat',
  'docVars',
  'rsids',
  'attachedSchema',
  'themeFontLang',
  'clrSchemeMapping',
  'doNotIncludeSubdocsInStats',
  'doNotAutoCompressPictures',
  'forceUpgrade',
  'captions',
  'readModeInkLockDown',
  'smartTagType',
  'shapeDefaults',
  'doNotEmbedSmartTags',
  'decimalSymbol',
  'listSeparator',
]);

/**
 * A settings member after the note properties: one of the list, or a member of another
 * namespace (`m:mathPr`, `sl:schemaLibrary`, and Word's `w14:` and `w15:` members, which it
 * writes last).
 */
const followsInSettings = (node: OoxmlNode): boolean =>
  node.kind !== 'textValue' &&
  (node.namespaceUri !== W || SETTINGS_AFTER_NOTE_PROPERTIES.has(node.localName));

/**
 * A section member after the note properties: `CT_SectPr` opens with the header and footer
 * references, and `w:footnotePr` and `w:endnotePr` come first after them.
 */
const followsInSection = (node: OoxmlNode): boolean =>
  node.kind !== 'textValue' && !isW(node, 'headerReference') && !isW(node, 'footerReference');

/**
 * `children` with the op's note properties: each field the op gives replaces the authored one,
 * every other authored child stays (the separator references of a settings-level element
 * too), and `w:footnotePr` then `w:endnotePr` sit before the first child that follows them in
 * the schema.
 */
function withNoteProperties(
  children: readonly OoxmlNode[],
  op: Pick<SetNotePropertiesOp, 'footnote' | 'endnote'>,
  nextId: () => string,
  follows: (node: OoxmlNode) => boolean
): OoxmlNode[] {
  const footnote = children.find((child) => isW(child, 'footnotePr'));
  const endnote = children.find((child) => isW(child, 'endnotePr'));
  return placed(
    children,
    [footnote, endnote],
    [
      mergedProps(nextId, 'footnotePr', footnote, op.footnote),
      mergedProps(nextId, 'endnotePr', endnote, op.endnote),
    ],
    follows
  );
}

/**
 * `children` without the `old` note property elements, and the `written` ones (`w:footnotePr`
 * then `w:endnotePr`) before the first child that follows them in the schema.
 */
function placed(
  children: readonly OoxmlNode[],
  old: readonly (OoxmlNode | undefined)[],
  written: readonly (OoxmlNode | null)[],
  follows: (node: OoxmlNode) => boolean
): OoxmlNode[] {
  const rest = children.filter((child) => !old.includes(child));
  const found = rest.findIndex(follows);
  const at = found < 0 ? rest.length : found;
  const nodes = written.filter((node): node is OoxmlNode => node !== null);
  return [...rest.slice(0, at), ...nodes, ...rest.slice(at)];
}

/** The ids of the two separator notes a new notes part holds (`ensureNotesPart`). */
const SEPARATOR_NOTE_IDS = ['-1', '0'] as const;

/**
 * The package with the settings' `w:footnotePr` or `w:endnotePr` listing a notes part's separator
 * notes (`<w:footnote w:id="-1"/><w:footnote w:id="0"/>`), as Word writes them for its notes
 * parts: ECMA-376 17.11.9 and 17.11.3 list the special notes a document uses there. Authored
 * fields and references stay; a missing reference goes after them (`CT_FtnDocProps`).
 *
 * A package with no settings part keeps none: Word reads its notes without the list, and two
 * people who add a first note at once in a Yjs room would each create the part, which shared
 * state cannot merge.
 */
export function writeNoteSeparatorReferences(
  pkg: OoxmlPackage,
  noteKind: 'footnote' | 'endnote'
): OoxmlPackage | null {
  const settings = settingsPartOf(pkg);
  if (!settings) return pkg;
  const nextId = createNodeIdAllocator(settings);
  const children = settings.root.children;
  const footnote = children.find((child) => isW(child, 'footnotePr'));
  const endnote = children.find((child) => isW(child, 'endnotePr'));
  const own = noteKind === 'footnote' ? footnote : endnote;
  const authored = own && own.kind !== 'textValue' ? own.children : [];
  const missing = SEPARATOR_NOTE_IDS.filter(
    (id) => !authored.some((child) => isW(child, noteKind) && wmlAttribute(child, 'id') === id)
  );
  if (own && missing.length === 0) return pkg;
  const listed = {
    id: own?.id ?? nextId(),
    kind: 'generic',
    namespaceUri: W,
    localName: noteKind === 'footnote' ? 'footnotePr' : 'endnotePr',
    prefix: 'w',
    namespaceBindings: [],
    attributes: own && own.kind !== 'textValue' ? own.attributes : [],
    children: [...authored, ...missing.map((id) => separatorReference(nextId, noteKind, id))],
  } as unknown as OoxmlNode;
  const written = noteKind === 'footnote' ? [listed, endnote ?? null] : [footnote ?? null, listed];
  const rebuilt = replaceChildren(
    settings,
    settings.root.id,
    placed(children, [footnote, endnote], written, followsInSettings)
  );
  return rebuilt.ok ? withPart(pkg, rebuilt.part) : null;
}

function wmlAttribute(node: OoxmlNode, localName: string): string | undefined {
  if (node.kind === 'textValue') return undefined;
  return node.attributes.find((a) => a.namespaceUri === W && a.localName === localName)?.value;
}

function separatorReference(
  nextId: () => string,
  noteKind: 'footnote' | 'endnote',
  id: string
): OoxmlNode {
  return {
    id: nextId(),
    kind: 'generic',
    namespaceUri: W,
    localName: noteKind,
    prefix: 'w',
    namespaceBindings: [],
    attributes: [
      { kind: 'genericExtension', namespaceUri: W, localName: 'id', prefix: 'w', value: id },
    ],
    children: [],
  } as unknown as OoxmlNode;
}

/** The `CT_FtnProps` fields in schema order: the op's name and the element's. */
const NOTE_FIELDS = [
  ['position', 'pos'],
  ['numFmt', 'numFmt'],
  ['numStart', 'numStart'],
  ['numRestart', 'numRestart'],
] as const;

/** The authored element with the given fields in place, or the element as it is. */
function mergedProps(
  nextId: () => string,
  localName: 'footnotePr' | 'endnotePr',
  existing: OoxmlNode | undefined,
  fields: NoteFields | undefined
): OoxmlNode | null {
  if (!fields) return existing ?? null;
  const authored = existing && existing.kind !== 'textValue' ? existing.children : [];
  const own: OoxmlNode[] = [];
  for (const [field, element] of NOTE_FIELDS) {
    const value = fields[field];
    const kept = authored.find((child) => isW(child, element));
    if (value !== undefined) own.push(valChild(nextId, element, String(value)));
    else if (kept) own.push(kept);
  }
  const others = authored.filter((child) => !NOTE_FIELDS.some(([, name]) => isW(child, name)));
  return {
    id: existing?.id ?? nextId(),
    kind: 'generic',
    namespaceUri: W,
    localName,
    prefix: 'w',
    namespaceBindings: [],
    attributes: existing && existing.kind !== 'textValue' ? existing.attributes : [],
    children: [...own, ...others],
  } as unknown as OoxmlNode;
}

function valChild(nextId: () => string, localName: string, value: string): OoxmlNode {
  return {
    id: nextId(),
    kind: 'generic',
    namespaceUri: W,
    localName,
    prefix: 'w',
    namespaceBindings: [],
    attributes: [
      {
        kind: 'wmlVal',
        namespaceUri: W,
        localName: 'val',
        prefix: 'w',
        value,
      },
    ],
    children: [],
  } as unknown as OoxmlNode;
}
