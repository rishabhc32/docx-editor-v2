import {
  readOoxmlPart,
  WML_NAMESPACE_URI,
  withPart,
  type OoxmlElement,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { readStyleDefinition } from '../layout/style-definition-reader.ts';
import { stylesPartOf } from '../store/package/ooxml-indexes.ts';
import { createNodeIdAllocator, insertChildren } from '../store/package/ooxml-edit.ts';
import {
  withNewPart,
  withRelationship,
  withRelationshipsPartFor,
} from '../store/package/package-edit.ts';
import { propertyElement } from '../store/store/tree-op-properties.ts';

/** Word's built-in hyperlink character style, by name: Word localizes the id, not the name. */
const HYPERLINK_NAME = 'hyperlink';
/** The styleId Word gives the built-in. */
const HYPERLINK_STYLE_ID = 'Hyperlink';

/** The character style a new link's text takes, and the edit that declares it if needed. */
export interface HyperlinkStyleWrite {
  readonly styleId: string;
  readonly packageEdits: readonly ((pkg: OoxmlPackage) => OoxmlPackage)[];
}

/**
 * The `Hyperlink` style for a new link: the document's own, resolved by name before id as
 * `listParagraphStyleId` does, or Word's definition under a free id, added in the link's own
 * transaction as Word does on a first link.
 */
export function hyperlinkStyleWrite(
  styles: readonly { type: string; name: string; styleId: string }[]
): HyperlinkStyleWrite {
  let byId: string | null = null;
  for (const style of styles) {
    if (style.type !== 'character') continue;
    if (style.name.trim().toLowerCase() === HYPERLINK_NAME) {
      return { styleId: style.styleId, packageEdits: [] };
    }
    if (byId === null && style.styleId === HYPERLINK_STYLE_ID) byId = style.styleId;
  }
  if (byId !== null) return { styleId: byId, packageEdits: [] };
  const usedStyleIds = new Set(styles.map((style) => style.styleId));
  let styleId = HYPERLINK_STYLE_ID;
  for (let suffix = 1; usedStyleIds.has(styleId); suffix++) {
    styleId = `${HYPERLINK_STYLE_ID}${suffix}`;
  }
  return { styleId, packageEdits: [ensureHyperlinkStyle(styleId)] };
}

/**
 * Declare Word's `Hyperlink` style under `styleId`. A document with no styles part gets one; a
 * style already declared under `styleId` is left as it is.
 *
 * The package writes repeat `ensureSeparatorStyle` in `list-separator-style.ts`, which declares
 * only List Paragraph. They are repeated, not shared, so that upstream file stays unchanged.
 */
function ensureHyperlinkStyle(styleId: string): (pkg: OoxmlPackage) => OoxmlPackage {
  return (pkg) => {
    const existing = stylesPartOf(pkg);
    const name = existing?.name ?? '/word/styles.xml';
    const contentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
    const parsed = existing
      ? null
      : readOoxmlPart(
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
          { name, contentType }
        );
    const part = existing ?? (parsed?.ok ? parsed.part : null);
    if (!part) throw new Error(`Cannot create style ${styleId}`);
    if (
      part.root.children.some(
        (node) =>
          node.kind !== 'textValue' &&
          node.attributes.some((a) => a.localName === 'styleId' && a.value === styleId)
      )
    )
      return pkg;
    const style = hyperlinkStyleElement(part, styleId);
    if (existing) {
      const inserted = insertChildren(part, part.root.id, part.root.children.length, [style]);
      if (!inserted.ok) throw new Error(`Cannot insert style ${styleId}`);
      return withPart(pkg, inserted.part);
    }
    let next = withNewPart(
      pkg,
      name,
      { ...part.root, children: [style] } as OoxmlElement,
      contentType
    );
    next = withRelationshipsPartFor(next, pkg.mainDocumentPart);
    const related = withRelationship(
      next,
      pkg.mainDocumentPart,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
      name
    );
    if (!related.ok) throw new Error(`Cannot link style ${styleId}`);
    return related.pkg;
  };
}

/**
 * Word's own definition: the theme's hyperlink colour, underlined, based on the default
 * character style of `styles`, with node ids that are fresh in that part.
 */
function hyperlinkStyleElement(styles: OoxmlPart, styleId: string): OoxmlElement {
  const nextId = createNodeIdAllocator(styles);
  const element = (
    localName: string,
    attributes: Record<string, string> = {},
    children: readonly OoxmlElement[] = []
  ): OoxmlElement =>
    ({
      ...(propertyElement({ localName, attributes }, nextId()) as OoxmlElement),
      children,
      ...(localName === 'rPr' ? { kind: 'runProperties' } : {}),
    }) as OoxmlElement;
  const basedOn = defaultCharacterStyleId(styles);
  const style = element('style', { type: 'character', styleId }, [
    element('name', { val: 'Hyperlink' }),
    ...(basedOn ? [element('basedOn', { val: basedOn })] : []),
    element('uiPriority', { val: '99' }),
    element('unhideWhenUsed'),
    element('rPr', {}, [
      element('color', { val: '0563C1', themeColor: 'hyperlink' }),
      element('u', { val: 'single' }),
    ]),
  ]);
  // Keep the inserted subtree valid even when the source uses another XML prefix.
  return { ...style, namespaceBindings: [{ prefix: 'w', namespaceUri: WML_NAMESPACE_URI }] };
}

/** The default character style (Word's Default Paragraph Font); the last one wins. */
function defaultCharacterStyleId(styles: OoxmlPart): string | null {
  let found: string | null = null;
  for (const child of styles.root.children) {
    if (child.kind === 'textValue' || child.localName !== 'style') continue;
    const style = readStyleDefinition(child);
    if (style?.type !== 'character') continue;
    if (style.isDefault) found = style.styleId;
    else if (found === style.styleId) found = null;
  }
  return found;
}
