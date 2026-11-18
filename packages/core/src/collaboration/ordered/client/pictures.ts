// What the client reads of the editor's picture ops (group 9.4): the atom of the list that an
// op's drawing id names, and a new drawing as the typed values a change carries.

import { buildInlinePictureDrawing } from '../../../store/package/drawing-package-edit.ts';
import { canonicalTreeDifference } from '../../../store/package/ooxml-difference.ts';
import { RELATIONSHIPS_NAMESPACE_URI } from '../../../store/package/ooxml-shared.ts';
import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { EMU_PER_POINT } from '../../../store/package/drawing-package-edit.ts';
import { buildTextboxDrawing } from '../../../store/package/textbox-drawing-builder.ts';
import { inheritedNamespaceBindings } from '../../../store/package/wml-namespace.ts';
import type { OrderedNewObject } from '../edits/pictures.ts';
import { isHidden, itemAt } from '../model/document.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from '../model/types.ts';
import type { OrderedTranslateBatch } from './ops/types.ts';
import { paragraphAtIndex, type ListParagraph } from './paragraphs.ts';
import { OrderedClientRefusal } from './refusal.ts';

/** An object's atom, its story, and the shown paragraph and offset it is at. */
export interface ListObject {
  readonly story: OrderedStory;
  readonly atom: OrderedItem;
  readonly paragraph: ListParagraph;
  readonly offset: number;
}

/** The drawing ids an atom answers to: its element's, and every `w:drawing` inside it. */
function drawingIds(node: OoxmlNode, into: string[], depth = 0): void {
  if (node.kind === 'textValue' || depth > 8) return;
  if (depth === 0 || node.kind === 'drawing') into.push(node.id);
  if (node.kind === 'drawing') return;
  for (const child of node.children) drawingIds(child, into, depth + 1);
}

/** A drawing's atom and its story. */
interface DrawingAtom {
  readonly story: OrderedStory;
  readonly atom: OrderedItem;
}

/** The atoms of every story by drawing id, rebuilt when a lookup misses. */
const DRAWING_INDEX = new WeakMap<OrderedDocument, Map<string, DrawingAtom>>();

/** Build the drawing index of a document again, and keep it. */
function rebuildDrawingIndex(doc: OrderedDocument): Map<string, DrawingAtom> {
  const index = new Map<string, DrawingAtom>();
  for (const story of doc.stories.values()) {
    for (const atom of story.tree.items()) {
      if (atom.kind !== 'atom' || !atom.node || atom.pieces) {
        continue;
      }

      const ids: string[] = [];
      drawingIds(atom.node, ids);
      for (const id of ids) {
        index.set(id, { story, atom });
      }
    }
  }

  DRAWING_INDEX.set(doc, index);
  return index;
}

/** The atom a drawing id names, still in its story, or null. */
function atomOfDrawing(doc: OrderedDocument, drawingNodeId: string): DrawingAtom | null {
  const live = (found: DrawingAtom | undefined) =>
    found !== undefined && itemAt(found.story, found.atom) === found.atom ? found : null;
  return (
    live(DRAWING_INDEX.get(doc)?.get(drawingNodeId)) ??
    live(rebuildDrawingIndex(doc).get(drawingNodeId))
  );
}

/**
 * The shown object a drawing id of the editor names, and where it is. The paragraph's text is
 * checked against the editor's through the batch, as for every op. Refused by name when the list
 * has no such object, or does not show it.
 */
export function objectOf(batch: OrderedTranslateBatch, drawingNodeId: string): ListObject {
  const found = atomOfDrawing(batch.doc, drawingNodeId);
  if (!found || isHidden(batch.doc, found.atom))
    throw new OrderedClientRefusal(
      'ordered-refused',
      'out-of-step',
      `no picture ${drawingNodeId.slice(0, 160)}`
    );
  const { story, atom } = found;
  const index = story.tree.before(atom).marks;
  const shown = paragraphAtIndex(batch.doc, story, index);
  const nodeId = shown.mark.node?.id;
  const paragraph = nodeId === undefined ? shown : batch.paragraph(nodeId);
  const offset = paragraph.view.chars.findIndex((char) => char.item === atom);
  if (offset < 0)
    throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'a picture off its paragraph');
  return { story, atom, paragraph, offset };
}

// ---- a new drawing, read back ----

const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

function find(
  node: OoxmlNode,
  test: (element: OoxmlElement) => boolean,
  depth = 0
): OoxmlElement | null {
  if (node.kind === 'textValue' || depth > 16) return null;
  if (test(node)) return node;
  for (const child of node.children) {
    const found = find(child, test, depth + 1);
    if (found) return found;
  }

  return null;
}

const named = (ns: string, localName: string) => (element: OoxmlElement) =>
  element.namespaceUri === ns && element.localName === localName;

const plain = (element: OoxmlElement | null, localName: string): string | undefined =>
  element?.attributes.find((a) => a.namespaceUri === '' && a.localName === localName)?.value;

const relationshipOf = (element: OoxmlElement | null, localName: string): string | undefined =>
  element?.attributes.find(
    (a) => a.namespaceUri === RELATIONSHIPS_NAMESPACE_URI && a.localName === localName
  )?.value;

function whole(value: string | undefined): number | null {
  if (value === undefined || !/^[1-9][0-9]{0,15}$/.test(value)) return null;
  const n = Number(value);
  return Number.isSafeInteger(n) ? n : null;
}

const unsaid = (what: string): OrderedClientRefusal =>
  new OrderedClientRefusal('ordered-unsupported-op', `insertDrawing.${what}`);

/** Whether `rebuilt` is the editor's drawing, ids aside, under the scope it goes into. */
function same(
  drawing: OoxmlNode,
  rebuilt: OoxmlNode,
  bindings: ReadonlyMap<string, string> | undefined
): boolean {
  return canonicalTreeDifference(drawing, rebuilt, { bindings }).equal;
}

/**
 * A new drawing of the editor's in `part` as the typed values a change carries: an inline
 * picture whose image is in the package the write's ops apply to (`batch.media`), or a text box.
 * A change carries no XML: the drawing is read back into the values the engine's builders take
 * (the inline picture of Insert Picture, the text box of Insert Text Box), built again from them,
 * and compared with the editor's drawing. A drawing those values do not say exactly is refused
 * by name, never sent as an approximation.
 */
export function newObjectOf(
  batch: OrderedTranslateBatch,
  drawing: OoxmlNode,
  part: string
): OrderedNewObject {
  const docPr = find(drawing, named(WP, 'docPr'));
  const docPrId = whole(plain(docPr, 'id'));
  const extent = find(drawing, named(WP, 'extent'));
  const cx = whole(plain(extent, 'cx'));
  const cy = whole(plain(extent, 'cy'));
  if (docPrId === null || cx === null || cy === null) throw unsaid('frame');
  const holder = (batch.gated() ?? batch.shown())?.parts.get(part);
  const scope = holder ? inheritedNamespaceBindings(holder, drawing) : undefined;
  if (find(drawing, named(WPS, 'txbx'))) {
    if (!same(drawing, buildTextboxDrawing(docPrId, cx / EMU_PER_POINT), scope))
      throw unsaid('shape');
    return { type: 'textbox', extent: { cx } };
  }

  const embed = relationshipOf(find(drawing, named(A, 'blip')), 'embed');
  const media = embed === undefined ? null : batch.media(part, embed);
  if (embed === undefined || media === null) throw unsaid('media');
  const hlink = relationshipOf(find(drawing, named(A, 'hlinkClick')), 'id');
  const link = hlink === undefined ? undefined : batch.link(part, hlink);
  if (link === null) throw unsaid('link');
  const title = plain(docPr, 'title');
  const description = plain(docPr, 'descr');
  const lockAspect = plain(find(drawing, named(A, 'graphicFrameLocks')), 'noChangeAspect') === '1';
  const rebuilt = buildInlinePictureDrawing({
    docPrId,
    relationshipId: embed,
    extentEmu: { cx, cy },
    ...(title === undefined ? {} : { title }),
    ...(description === undefined ? {} : { description }),
    ...(hlink === undefined ? {} : { hyperlinkRelationshipId: hlink }),
    ...(lockAspect ? { lockAspectRatio: true } : {}),
  });
  if (!same(drawing, rebuilt, scope)) throw unsaid('drawing');
  return {
    type: 'picture',
    media,
    extent: { cx, cy },
    ...(title === undefined ? {} : { title }),
    ...(description === undefined ? {} : { description }),
    ...(link === undefined ? {} : { link }),
    ...(lockAspect ? { lockAspect: true } : {}),
  };
}
