// Text box stories: their names, and the index from each one to the object that holds it.

import type { OrderedDocument, OrderedItem, OrderedItemAddress } from './types.ts';
import { recordUndo } from './undo-record.ts';

/** Where a text box's story sits: its host story, its object's address, its content's node. */
export interface OrderedTextBoxHost {
  /** The story that holds the object: the main document, a header or a footer. */
  readonly host: string;
  /** The object's address in the host story. */
  readonly atom: OrderedItemAddress;
  /** The node id of the `w:txbxContent` the story was loaded from, in the object's node. */
  readonly root: string;
}

const MARK = '!txbx:';

/**
 * The name of a text box's story: the host part, and the content's node id. A text box's
 * paragraphs live inside an object (a drawing, a VML shape), so the object would be one unit no
 * edit reaches. Each text box of the main document, a header or a footer is its own story
 * instead: its `w:txbxContent` loaded as items. The object keeps its node; compose writes the
 * story back into it (`compose/text-boxes.ts`).
 */
export const textBoxStoryName = (host: string, root: string): string => `${host}${MARK}${root}`;

/** Whether a story name is a text box's. */
export const isTextBoxStoryName = (name: string): boolean => name.includes(MARK);

interface TextBoxIndex {
  readonly byName: Map<string, OrderedTextBoxHost>;
  /** Object address (`space|offset`) to its text box stories' names, in node order. */
  readonly byAtom: Map<string, string[]>;
}

const INDEXES = new WeakMap<OrderedDocument, TextBoxIndex>();

function indexOf(doc: OrderedDocument): TextBoxIndex {
  let index = INDEXES.get(doc);
  if (!index) {
    index = { byName: new Map(), byAtom: new Map() };
    INDEXES.set(doc, index);
  }

  return index;
}

const atomKey = (address: OrderedItemAddress): string => `${address.space}|${address.offset}`;

/**
 * Record a text box story's host, inside the caller's undo record when one is open (a change),
 * or for good (a load). A take-back removes what a change added.
 */
export function registerTextBox(
  doc: OrderedDocument,
  name: string,
  host: OrderedTextBoxHost
): void {
  const index = indexOf(doc);
  index.byName.set(name, host);
  const key = atomKey(host.atom);
  const names = index.byAtom.get(key) ?? [];
  index.byAtom.set(key, [...names, name]);
  recordUndo(doc, () => {
    index.byName.delete(name);
    const left = (index.byAtom.get(key) ?? []).filter((other) => other !== name);
    if (left.length > 0) {
      index.byAtom.set(key, left);
    } else {
      index.byAtom.delete(key);
    }
  });
}

/** The host of a text box story, or null for any other story. */
export function textBoxHostOf(doc: OrderedDocument, name: string): OrderedTextBoxHost | null {
  return INDEXES.get(doc)?.byName.get(name) ?? null;
}

/** The text box stories an object holds, by name, in node order. */
export function textBoxesOfAtom(doc: OrderedDocument, atom: OrderedItem): readonly string[] {
  if (atom.kind !== 'atom') return [];
  return INDEXES.get(doc)?.byAtom.get(atomKey(atom)) ?? [];
}

/** Whether a document has text box stories at all (a cheap test for the common case). */
export function hasTextBoxes(doc: OrderedDocument): boolean {
  return (INDEXES.get(doc)?.byName.size ?? 0) > 0;
}
