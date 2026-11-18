// Text box stories (group 9.9, `model/text-boxes.ts`), at load and when an edit adds an object
// that holds a text box.

import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNamespaceBinding,
  type OoxmlNode,
  type OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import { usedParaIds, w14RootPrefix, withFreshParaIds } from '../../../store/package/para-id.ts';
import { textboxStoriesInPart } from '../../../store/package/textbox-stories.ts';
import { MAX_XML_DEPTH } from '../../../store/package/xml-reader.ts';
import { addOrderedStory, labelUnindex } from '../model/document.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { registerTextBox, textBoxStoryName } from '../model/text-boxes.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from '../model/types.ts';
import { recordUndo } from '../model/undo-record.ts';
import { withChildren } from '../ooxml-nodes.ts';
import { loadBlocks } from './blocks.ts';
import {
  itemsWithPrompts,
  loadInEdit,
  newLoadBudget,
  newLoadState,
  RevisionReader,
  type LoadBudget,
} from './state.ts';

/**
 * The story parts whose text boxes the editor edits: the main document (the editor's text box
 * editing), headers and footers (the editing API's shape bodies). A note's text box is not one.
 */
const TEXT_BOX_HOSTS: ReadonlySet<string> = new Set(['document', 'hdr', 'ftr']);

export const holdsTextBoxStories = (story: OrderedStory): boolean =>
  TEXT_BOX_HOSTS.has(story.part.root.localName) &&
  story.part.root.namespaceUri === WML_NAMESPACE_URI;

/**
 * The content of a text box an edit adds (a new one, a pasted or a copied one) with a fresh
 * `w14:paraId` on each paragraph, as every paragraph an edit makes has one: a save then holds
 * them, and a reopen keeps them instead of minting its own. The ids come from the object's node
 * id, which the change minted, so every copy makes the same, past every id the host part and
 * this session hold. Inside the caller's undo record.
 */
function withSessionParagraphIds(
  doc: OrderedDocument,
  host: OrderedStory,
  atom: OrderedItem,
  root: OoxmlElement
): OoxmlElement {
  const prefix = w14RootPrefix(host.part.root);
  if (prefix === null || !atom.node) return root;
  const used = new Set([...usedParaIds(host.part.root), ...doc.paraIds]);
  const before = new Set(used);
  const fresh = withFreshParaIds(
    root,
    used,
    `ordered text box ${atom.node.id}`,
    { value: 0 },
    prefix
  );
  const minted = [...used].filter((id) => !before.has(id));
  for (const id of minted) doc.paraIds.add(id);
  recordUndo(doc, () => {
    for (const id of minted) doc.paraIds.delete(id);
  });
  return fresh.kind === 'textValue' ? root : fresh;
}

/** One text box as its own story, of the object `atom` in `host`. */
function addTextBoxStory(
  doc: OrderedDocument,
  host: OrderedStory,
  atom: OrderedItem,
  loaded: OoxmlElement,
  revisions: RevisionReader,
  budget: LoadBudget
): void {
  // An edit's text box takes paragraph ids; a file's has the ones its open gave
  const root = doc.undo === null ? loaded : withSessionParagraphIds(doc, host, atom, loaded);
  const name = textBoxStoryName(host.name, root.id);
  if (doc.stories.has(name)) throw new OrderedRefusal('name-taken', name.slice(0, 160));
  const state = newLoadState(doc, name, revisions, budget);
  loadInEdit(doc, () => loadBlocks(state, root.children, 1));
  const owner = atom.by;

  // A text box in an object a suggestion inserted is that suggestion's: its plain content takes
  // the insertion's label under its marker, as a new note's content does (`load/notes.ts`), so a
  // person types inside another person's new text box only as inside their suggestion
  const items = itemsWithPrompts(state).map(
    (item): OrderedItem =>
      owner !== null && item.by === null ? { ...item, by: owner, quiet: true } : item
  );

  // The text box's paragraphs read the host's namespaces (`w14`, for new paragraph ids); a
  // prefix the content declares itself is bound once, as the content binds it
  const bindings = new Map<string, OoxmlNamespaceBinding>();
  for (const binding of [...host.part.root.namespaceBindings, ...root.namespaceBindings])
    bindings.set(binding.prefix, binding);
  const part: OoxmlPart = {
    id: `${host.part.id}${name.slice(host.name.length)}`,
    name,
    contentType: host.part.contentType,
    root: { ...root, namespaceBindings: [...bindings.values()] },
  };
  const story = addOrderedStory(doc, name, part, items);
  recordUndo(doc, () => {
    for (const item of story.tree.items()) labelUnindex(doc, item);
    doc.stories.delete(name);
  });
  registerTextBox(doc, name, {
    host: host.name,
    atom: { space: atom.space, offset: atom.offset },
    root: root.id,
  });
}

/**
 * At load: every text box of a host's part (`holdsTextBoxStories`) as its own story, with the
 * load's revision reader (a revision's triple is one suggestion across the stories) and budget.
 * The text boxes are the ones the editor itself reaches (`textboxStoriesInPart`): the live copy
 * of each (the DrawingML one, or the VML one when the file has no other), in a drawing that
 * shows. The package writer keeps the VML fallback copy in step on save, as for the editor alone.
 */
export function loadTextBoxStories(
  doc: OrderedDocument,
  host: OrderedStory,
  revisions: RevisionReader,
  budget: LoadBudget
): void {
  if (!holdsTextBoxStories(host)) return;
  const roots = textboxStoriesInPart(host.part);
  if (roots.length === 0) return;
  const wanted = new Set(roots.map((root) => root.drawingNodeId));
  const atoms = new Map<string, OrderedItem>();
  for (const item of host.tree.items())
    if (item.kind === 'atom' && item.node && wanted.has(item.node.id))
      atoms.set(item.node.id, item);
  for (const root of roots) {
    const atom = atoms.get(root.drawingNodeId);
    if (atom) addTextBoxStory(doc, host, atom, root.root, revisions, budget);
  }
}

/** Whether a node has a `w:txbxContent` below it: a cheap, bounded test before a full read. */
function holdsTextBoxContent(node: OoxmlNode): boolean {
  const stack: { node: OoxmlNode; depth: number }[] = [{ node, depth: 0 }];
  for (let next = stack.pop(); next !== undefined; next = stack.pop()) {
    const at = next.node;
    if (at.kind === 'textValue' || next.depth > MAX_XML_DEPTH) continue;
    if (at.localName === 'txbxContent' && at.namespaceUri === WML_NAMESPACE_URI) return true;
    for (const child of at.children) stack.push({ node: child, depth: next.depth + 1 });
  }

  return false;
}

const SCRATCH_KINDS: ReadonlyMap<string, OoxmlNode['kind']> = new Map([
  ['p', 'paragraph'],
  ['r', 'run'],
]);

/** A scratch main part that holds one object in one run, for the editor's own text box read. */
function scratchPart(host: OrderedStory, atom: OoxmlElement): OoxmlPart {
  const element = (id: string, localName: string, children: OoxmlNode[]): OoxmlElement =>
    ({
      id: `${atom.id}~${id}`,
      kind: SCRATCH_KINDS.get(localName) ?? 'generic',
      namespaceUri: WML_NAMESPACE_URI,
      localName,
      prefix: 'w',
      namespaceBindings: [],
      attributes: [],
      children,
    }) as unknown as OoxmlElement;
  const body = element('body', 'body', [element('p', 'p', [element('r', 'r', [atom])])]);
  return { ...host.part, root: withChildren(host.part.root, [body]) };
}

/**
 * Objects that go into a host story during the session: each text box they hold becomes
 * its own story, inside the caller's undo record. Its revisions are the object's suggestion.
 */
export function addTextBoxStories(
  doc: OrderedDocument,
  host: OrderedStory,
  items: readonly OrderedItem[]
): void {
  if (!holdsTextBoxStories(host)) return;
  for (const atom of items) {
    const node = atom.node;
    if (atom.kind !== 'atom' || !node || node.kind === 'textValue') continue;
    if (!holdsTextBoxContent(node)) continue;
    const budget = newLoadBudget();
    const revisions = new RevisionReader(doc, atom.by ?? undefined);
    for (const root of textboxStoriesInPart(scratchPart(host, node)))
      if (root.drawingNodeId === node.id)
        addTextBoxStory(doc, host, atom, root.root, revisions, budget);
  }
}
