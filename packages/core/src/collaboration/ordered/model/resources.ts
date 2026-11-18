// What edits add outside the story text (design D14): media named by content hash, new story
// parts, a pasted chart's or diagram's own parts, relationships, list definitions, styles, and
// settings, each by key with the suggestion that added it.

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import type {
  OrderedFragmentObjectUse,
  OrderedFragmentPartType,
} from '../change/fragment-names-objects.ts';
import type { OoxmlProperty } from '../../../store/store/tree-op-types.ts';
import { sameMaker, type OrderedMaker } from './person.ts';
import { OrderedRefusal } from './refusal.ts';
import type { OrderedDocument, OrderedItem, OrderedNotePropertiesOp } from './types.ts';
import { recordUndo } from './undo-record.ts';

/**
 * Media content types a change may add: the images the engine reads in a Word file. Raster
 * images, an SVG version (painted as an image, never run), and TIFF, EMF and WMF (kept as they
 * are; a host may convert them to paint). And bytes kept as they are, never opened or run: a
 * pasted chart's embedded workbook, a pasted embedded object's part (an OLE object, or an
 * embedded workbook, document or presentation), and a pasted ink part (InkML).
 *
 * @public
 */
export type OrderedMediaType =
  | 'image/png'
  | 'image/jpeg'
  | 'image/gif'
  | 'image/bmp'
  | 'image/webp'
  | 'image/svg+xml'
  | 'image/tiff'
  | 'image/x-emf'
  | 'image/x-wmf'
  | 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  | 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  | 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  | 'application/vnd.openxmlformats-officedocument.oleObject'
  | 'application/inkml+xml';

/** The file name extension of each media content type. @public */
export const ORDERED_MEDIA_EXTENSIONS: Readonly<Record<OrderedMediaType, string>> = Object.freeze({
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/webp': 'webp',
  'image/svg+xml': 'svg',
  'image/tiff': 'tiff',
  'image/x-emf': 'emf',
  'image/x-wmf': 'wmf',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
  'application/vnd.openxmlformats-officedocument.oleObject': 'bin',
  'application/inkml+xml': 'xml',
});

/** A chart's embedded workbook (ECMA-376 Part 1, §15.2.11): media a chart's part names. */
export const ORDERED_WORKBOOK_TYPE: OrderedMediaType =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** An embedded object's OLE part (ECMA-376 Part 1, §15.2.10): a compound file, never opened. */
export const ORDERED_OLE_OBJECT_TYPE: OrderedMediaType =
  'application/vnd.openxmlformats-officedocument.oleObject';

/**
 * The parts of an embedded object (`o:OLEObject/@r:id`): an OLE part, or an embedded Office
 * file (ECMA-376 Part 1, §15.2.11), as Word embeds a workbook, a document or a presentation.
 */
export const ORDERED_EMBEDDED_OBJECT_TYPES: readonly OrderedMediaType[] = Object.freeze([
  ORDERED_OLE_OBJECT_TYPE,
  ORDERED_WORKBOOK_TYPE,
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
]);

/** An ink part ([MS-DOCX] `w14:contentPart`): InkML, XML text that only Word reads. */
export const ORDERED_INK_TYPE: OrderedMediaType = 'application/inkml+xml';

/** What media serves in a paste: each use takes only its own content types. */
export type OrderedMediaUse = 'picture' | 'workbook' | OrderedFragmentObjectUse;

/** Whether media of `type` may serve `use`: the room checks it, whatever a client sends. */
export function isOrderedMediaFor(type: string, use: OrderedMediaUse): boolean {
  switch (use) {
    case 'picture':
      return type.startsWith('image/');
    case 'workbook':
      return type === ORDERED_WORKBOOK_TYPE;
    case 'object':
      return ORDERED_EMBEDDED_OBJECT_TYPES.some((own) => own === type);
    case 'ink':
      return type === ORDERED_INK_TYPE;
  }
}

/** Every media content type, once: the readers' and the stores' one list. */
export const ORDERED_MEDIA_TYPES: readonly OrderedMediaType[] = Object.freeze(
  Object.keys(ORDERED_MEDIA_EXTENSIONS) as OrderedMediaType[]
);

/** A content type the room stores, or null. */
export function orderedMediaType(value: string): OrderedMediaType | null {
  return ORDERED_MEDIA_TYPES.find((type) => type === value) ?? null;
}

/**
 * The media of a new picture or a new image (Insert Picture, Change Picture, a script's
 * picture): the raster images the engine's picture commands write. A paste carries any type.
 */
export const ORDERED_PICTURE_MEDIA_TYPES: readonly OrderedMediaType[] = Object.freeze([
  'image/png',
  'image/jpeg',
  'image/gif',
  'image/bmp',
  'image/webp',
]);

/** What every resource has: its key, and the suggestion that added it. @public */
export interface OrderedResourceBase {
  /** `p:<change>.<k>`: the key the change's minter made. */
  readonly key: string;
  /** The suggestion that added it: a party's pending one, or a direct edit's accepted one. */
  readonly suggestion: string;
}

/**
 * A media part named by its bytes' SHA-256: the bytes are in the room's storage, not in the
 * package (design D14).
 *
 * @public
 */
export interface OrderedMediaReference {
  /** The part name, such as `/word/media/image1.png`. */
  readonly part: string;
  readonly contentType: OrderedMediaType;
  /** 64 lower-case hex digits. */
  readonly sha256: string;
  /** The size in bytes. */
  readonly size: number;
}

/**
 * A media part of the loaded file that a room checkpoint names by reference, and its place
 * among the file's parts: a save writes the part there, as the Word file had it.
 *
 * @public
 */
export interface OrderedFileMedia extends OrderedMediaReference {
  /** The part's place among the entries of the file, from 0. */
  readonly entry: number;
}

/**
 * A media part a change adds: bytes the change names by their SHA-256, which the room server
 * stores and serves (design D14). The bytes never travel inside a change.
 *
 * @public
 */
export interface OrderedMediaResource extends OrderedResourceBase, OrderedMediaReference {
  readonly kind: 'media';
}

/** What a new story part is. @public */
export type OrderedStoryRole = 'header' | 'footer' | 'footnotes' | 'endnotes';

/**
 * A new story part: its story is in `doc.stories` under `part` (`addOrderedStory`); this says
 * what it is, so compose writes its content type, and only while a relationship to it is used.
 *
 * @public
 */
export interface OrderedStoryResource extends OrderedResourceBase {
  readonly kind: 'story';
  readonly part: string;
  readonly role: OrderedStoryRole;
  /**
   * A loaded part whose relationships the new part takes, under the same ids: a header or
   * footer a section stops sharing with the one before it is a copy of it (group 9.8).
   */
  readonly relationshipsOf?: string;
}

/**
 * A part a pasted chart or diagram needs (`OrderedFragmentPart`): written while a relationship
 * to it is written, with its own relationships read in turn.
 *
 * @public
 */
export interface OrderedObjectPartResource extends OrderedResourceBase {
  readonly kind: 'part';
  readonly part: string;
  readonly type: OrderedFragmentPartType;
  /** The part's root, read by the engine's bounded part reader. */
  readonly root: OoxmlElement;
}

/** The relationship types a change may add. @public */
export type OrderedRelationshipType =
  | 'image'
  | 'hyperlink'
  | 'header'
  | 'footer'
  | 'footnotes'
  | 'endnotes'
  | 'package'
  | 'oleObject'
  | 'ink'
  | OrderedFragmentPartType;

/**
 * A relationship from a part to a resource the session added, or to an address outside the
 * package (a link's, through `sanitizeHref`).
 *
 * @public
 */
export interface OrderedRelationshipResource extends OrderedResourceBase {
  readonly kind: 'relationship';
  /** The part that holds it: a story, loaded or added. */
  readonly owner: string;
  /** The relationship id, minted from the key. */
  readonly id: string;
  readonly type: OrderedRelationshipType;
  /** The key of the resource it points at, or an external address. */
  readonly target: { readonly resource: string } | { readonly external: string };
}

/** One level of a list definition. @public */
export interface OrderedListLevel {
  /** 0 to 8. */
  readonly level: number;
  readonly start: number;
  /** `w:numFmt`: decimal, bullet, lowerLetter, ... (an allowlist). */
  readonly format: string;
  /** `w:lvlText`, such as `%1.` or a bullet character. */
  readonly text: string;
  readonly justification: 'left' | 'center' | 'right';
  /** Left indent and hanging indent, in twips. */
  readonly indent: { readonly left: number; readonly hanging: number };
  /** The bullet's font, when the level shows a symbol. */
  readonly font?: string;
}

/**
 * A new list definition: a `w:abstractNum` and the `w:num` that names it.
 *
 * @public
 */
export interface OrderedListResource extends OrderedResourceBase {
  readonly kind: 'list';
  /** `w:num/@w:numId`, minted. Paragraphs name it in `w:numPr`. */
  readonly numId: string;
  /** `w:abstractNum/@w:abstractNumId`, minted. */
  readonly abstractNumId: string;
  readonly levels: readonly OrderedListLevel[];
  /**
   * A pasted list's whole definition, written in place of `levels`: its `w:abstractNum` (the
   * apply sets its number) and the `w:lvlOverride`s of its `w:num`.
   */
  readonly definition?: {
    readonly abstract: OoxmlElement;
    readonly overrides: readonly OoxmlElement[];
  };
  /** A pasted list's picture bullets, which its levels name (`w:lvlPicBulletId`). */
  readonly bullets?: readonly OrderedListBullet[];
}

/**
 * A picture bullet of a pasted list (`w:numPicBullet`, ECMA-376 Part 1, §17.9.20): written in
 * the numbering part while its list is, with its images as relationships of the numbering part.
 *
 * @public
 */
export interface OrderedListBullet {
  /** `w:numPicBulletId`, minted from the list's key. */
  readonly id: string;
  /** The `w:numPicBullet`, with `id` and its images' minted relationship ids. */
  readonly element: OoxmlElement;
  /** Its images: the minted relationship id, and the media resource it points at. */
  readonly images: readonly { readonly id: string; readonly media: string }[];
}

/**
 * A list level a list's definition does not declare, which Increase Indent declares before it
 * moves a paragraph there: the engine's own declaration (`ensureNumberingLevel`, Word's default
 * format for that depth), replayed on the numbering part at compose. Word records no change in
 * the numbering part (`CT_Numbering` has no revision children), so it goes with its suggestion.
 *
 * @public
 */
export interface OrderedListLevelResource extends OrderedResourceBase {
  readonly kind: 'list-level';
  /** `w:num/@w:numId`: the list whose definition takes the level. */
  readonly numId: string;
  /** 0 to 8. */
  readonly level: number;
  /** Which of Word's default formats the level takes: a bullet, or a number. */
  readonly format: 'bullet' | 'ordered';
}

/** What a style styles. @public */
export type OrderedStyleType = 'paragraph' | 'character' | 'table';

/**
 * A new style: a link's Hyperlink style, a list's List Paragraph, a paste's style. Written
 * only when the file has no style with its id.
 *
 * @public
 */
export interface OrderedStyleResource extends OrderedResourceBase {
  readonly kind: 'style';
  readonly styleId: string;
  readonly type: OrderedStyleType;
  readonly name: string;
  readonly basedOn?: string;
  readonly next?: string;
  readonly link?: string;
  /** `w:uiPriority`: where Word's style gallery sorts it. */
  readonly uiPriority?: number;
  /** `w:unhideWhenUsed`: Word shows it in its lists once something uses it. */
  readonly unhideWhenUsed?: boolean;
  /** Paragraph properties, on the paragraph allowlist. */
  readonly paragraph: readonly OoxmlProperty[];
  /** Run properties, on the run allowlist. */
  readonly run: readonly OoxmlProperty[];
  /** A pasted style's whole `w:style`, written in place of the typed values. */
  readonly element?: OoxmlElement;
}

/**
 * A settings write: the engine's own op, replayed on the settings part at compose, in the order
 * changes made them. Word records no settings change as a suggestion, so only a direct edit makes
 * one (a party's is refused as `policy-direct`). The document's note properties are one (group
 * 9.9).
 *
 * @public
 */
export interface OrderedSettingResource extends OrderedResourceBase {
  readonly kind: 'setting';
  readonly op:
    | { readonly op: 'setDocumentProtection'; readonly enforce: boolean }
    | { readonly op: 'setSectionFurnitureOptions'; readonly evenAndOddHeaders: boolean }
    | { readonly op: 'setDocumentProperties'; readonly values: Readonly<Record<string, string>> }
    | OrderedNotePropertiesOp;
}

/**
 * A custom node's payload (a data chip's data): a node in the custom XML data store of its
 * namespace, which the control's `w:dataBinding` names. Written while a written control binds
 * it; the store is the file's, or a new one under the `ds:itemID` the binding quotes.
 *
 * @public
 */
export interface OrderedCustomXmlResource extends OrderedResourceBase {
  readonly kind: 'custom-xml';
  /** The store's namespace, which names the store among the package's stores. */
  readonly namespaceUri: string;
  /** The store's root element name. */
  readonly root: string;
  /** The store's `ds:itemID`, a GUID in braces, as the binding quotes it. */
  readonly storeItemId: string;
  /** The node's own id, which the binding's xpath quotes. */
  readonly nodeId: string;
  /** The text Word paints the control from. */
  readonly label: string;
  /** The payload, as its host serialized it; never parsed here. */
  readonly data: string;
  /** The open item of the control that binds it. */
  readonly control: OrderedItem;
}

/** Anything a session adds outside the story text. @public */
export type OrderedResource =
  | OrderedMediaResource
  | OrderedStoryResource
  | OrderedObjectPartResource
  | OrderedRelationshipResource
  | OrderedListResource
  | OrderedListLevelResource
  | OrderedStyleResource
  | OrderedSettingResource
  | OrderedCustomXmlResource;

/** A custom XML data store the loaded main document part relates. @public */
export interface OrderedFileStore {
  /** `ds:itemID`, braced and upper-case. */
  readonly storeItemId: string;
  /** The store's root element name. */
  readonly root: string;
  /** The nodes the loaded stories' controls bind: compose takes out those no control binds. */
  readonly boundNodes: ReadonlySet<string>;
}

/** The resources of a document, and the loaded file's names they must not take. @public */
export interface OrderedResources {
  /** Part names the loaded package holds, lower case (OPC names ignore case). Load writes it. */
  readonly fileParts: Set<string>;
  /** Relationship ids the loaded package holds, by owner part. Load writes it. */
  readonly fileRelationships: Map<string, Set<string>>;
  /**
   * List numbers, abstract list numbers and picture bullet ids the loaded numbering part holds.
   * Load writes it.
   */
  readonly fileLists: Set<string>;
  /** The custom XML data stores the loaded main document part relates, by namespace. */
  readonly fileStores: Map<string, OrderedFileStore>;
  /**
   * The file's media that a room checkpoint names by reference, by part name key
   * (`partNameKey`): the package holds no bytes for them. Load writes it.
   */
  readonly fileMedia: Map<string, OrderedFileMedia>;
  /** Resources by key, in the order changes added them. */
  readonly added: Map<string, OrderedResource>;
}

/** No resources: for a new document, filled by load. */
export function createOrderedResources(): OrderedResources {
  return {
    fileParts: new Set(),
    fileRelationships: new Map(),
    fileLists: new Set(),
    fileStores: new Map(),
    fileMedia: new Map(),
    added: new Map(),
  };
}

/** The Word names a resource takes. A name of another resource or of the file is taken. */
function namesOf(resource: OrderedResource): string[] {
  switch (resource.kind) {
    case 'media':
    case 'story':
    case 'part':
      return [`part ${resource.part.toLowerCase()}`];
    case 'relationship':
      return [`rel ${resource.owner} ${resource.id}`];
    case 'list':
      return [
        `num ${resource.numId}`,
        `abstract ${resource.abstractNumId}`,
        ...(resource.bullets ?? []).flatMap((bullet) => [
          `bullet ${bullet.id}`,
          ...bullet.images.map((image) => `rel numbering ${image.id}`),
        ]),
      ];
    case 'list-level':
    case 'style':
    case 'setting':
    case 'custom-xml':
      return [];
  }
}

/** Whether a relationship id is one the loaded file holds, in any part. */
function fileRelationship(doc: OrderedDocument, id: string): boolean {
  for (const ids of doc.resources.fileRelationships.values()) if (ids.has(id)) return true;
  return false;
}

function takenByFile(doc: OrderedDocument, resource: OrderedResource): boolean {
  const { fileParts, fileRelationships, fileLists } = doc.resources;
  switch (resource.kind) {
    case 'media':
    case 'story':
    case 'part':
      return fileParts.has(resource.part.toLowerCase());
    case 'relationship':
      return fileRelationships.get(resource.owner)?.has(resource.id) ?? false;
    case 'list':
      return (
        fileLists.has(`num ${resource.numId}`) ||
        fileLists.has(`abstract ${resource.abstractNumId}`) ||
        (resource.bullets ?? []).some(
          (bullet) =>
            fileLists.has(`bullet ${bullet.id}`) ||
            bullet.images.some((image) => fileRelationship(doc, image.id))
        )
      );
    case 'list-level':
    case 'style':
    case 'setting':
    case 'custom-xml':
      return false;
  }
}

/**
 * Add a resource, inside the caller's undo record. The list never removes one during a room
 * session. Refused by name when its key is in use, or when one of its names is the file's or
 * another resource's (`name-taken`): every copy refuses alike, as every copy holds the same file
 * and the same resources. Every key and Word name (part name, relationship id, list number)
 * comes from the change id through the change's id minter, so every copy adds the same resource
 * under the same name in any arrival order.
 */
export function addOrderedResource(doc: OrderedDocument, resource: OrderedResource): void {
  if (doc.undo === null) throw new Error('A resource is added inside an undo record.');
  const { added } = doc.resources;
  if (added.has(resource.key)) throw new OrderedRefusal('name-taken', resource.key);
  if (takenByFile(doc, resource)) throw new OrderedRefusal('name-taken', resource.key);
  const names = new Set(namesOf(resource));
  for (const other of added.values())
    for (const name of namesOf(other))
      if (names.has(name)) throw new OrderedRefusal('name-taken', resource.key);
  added.set(resource.key, resource);
  recordUndo(doc, () => {
    added.delete(resource.key);
  });
}

/**
 * Whether a resource can still be written: its suggestion is pending or accepted. Compose writes
 * those that show and that something uses (`compose/resources.ts`). A rejected suggestion's
 * resource goes; a taken-back one's goes when nothing uses it, as its content does not show.
 */
export function orderedResourceShows(doc: OrderedDocument, resource: OrderedResource): boolean {
  const status = doc.suggestions.get(resource.suggestion)?.status;
  return status === 'pending' || status === 'accepted';
}

/**
 * A resource key's Word token: `<namespace>_<counter>_<k>` for `p:<namespace>:<counter>.<k>`.
 * Counter and k are digits, so the token names one key.
 */
export function orderedResourceToken(key: string): string {
  const match = /^p:([A-Za-z0-9_-]{1,64}):([0-9]{1,15})\.([0-9]{1,9})$/.exec(key);
  if (!match) throw new RangeError('Not a resource key.');
  return `${match[1]}_${match[2]}_${match[3]}`;
}

/** A relationship id for a resource key: an XML name, as `Id` must be. */
export const orderedRelationshipId = (key: string): string => `rO${orderedResourceToken(key)}`;

/**
 * A media part name for a resource key, in the folder Word uses: an embedded object's part and
 * a chart's workbook in `embeddings`, ink in `ink`, an image in `media`.
 */
export const orderedMediaPart = (key: string, type: OrderedMediaType): string => {
  let folder = 'media';
  if (ORDERED_EMBEDDED_OBJECT_TYPES.includes(type)) {
    folder = 'embeddings';
  } else if (type === ORDERED_INK_TYPE) {
    folder = 'ink';
  }

  return `/word/${folder}/ordered_${orderedResourceToken(key)}.${ORDERED_MEDIA_EXTENSIONS[type]}`;
};

/** A chart's or a diagram's part name for a resource key, in the folder Word uses. */
export const orderedObjectPart = (key: string, type: OrderedFragmentPartType): string => {
  let folder = 'diagrams';
  if (type === 'chartUserShapes') {
    folder = 'drawings';
  } else if (type === 'themeOverride') {
    folder = 'theme';
  } else if (type.startsWith('chart')) {
    folder = 'charts';
  }

  return `/word/${folder}/ordered_${type}_${orderedResourceToken(key)}.xml`;
};

/** A story part name for a resource key. */
export const orderedStoryPart = (key: string, role: OrderedStoryRole): string =>
  `/word/${role}_${orderedResourceToken(key)}.xml`;

/**
 * A list number for a resource key: a decimal from 100000 to 899999999, a function of the key
 * alone (FNV-1a), so the same on every copy. Nine digits at most, the engine's own bound for a
 * list number (`setListNumbering`). Two keys meet with a chance of about one in 900 million;
 * the second then refuses as `name-taken` on every copy alike.
 */
export function orderedListNumber(
  key: string,
  role: 'num' | 'abstract' | 'bullet',
  index?: number
): string {
  let hash = 0x811c9dc5;
  const text = index === undefined ? `ordered ${role} ${key}` : `ordered ${role} ${index} ${key}`;
  for (let k = 0; k < text.length; k += 1) {
    hash ^= text.charCodeAt(k);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }

  return String(100_000 + (hash % 899_900_000));
}

/**
 * The list definition a change added with this number, when an edit by `maker` may use it: its
 * suggestion is accepted, or pending and the maker's own. Another person's pending list is not
 * theirs to use (design D14).
 */
export function orderedListOf(
  doc: OrderedDocument,
  numId: string,
  maker: OrderedMaker
): OrderedListResource | null {
  for (const resource of doc.resources.added.values()) {
    if (resource.kind !== 'list' || resource.numId !== numId) continue;
    const suggestion = doc.suggestions.get(resource.suggestion);
    if (suggestion?.status === 'accepted') return resource;
    if (suggestion?.status === 'pending' && sameMaker(suggestion.attribution, maker)) {
      return resource;
    }
  }

  return null;
}
