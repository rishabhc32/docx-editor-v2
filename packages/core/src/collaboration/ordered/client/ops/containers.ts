// Containers the editor names by node id (a table, a row, a cell, a control, a link), found in
// the list as their open items, as paragraphs are found by their marks.

import type { OrderedChangeAddress } from '../../change/types.ts';
import { isKeptLink } from '../../compose/kept-links.ts';
import { isHidden } from '../../model/document.ts';
import type {
  OrderedContainerToken,
  OrderedDocument,
  OrderedItem,
  OrderedStory,
} from '../../model/types.ts';
import { addressOf, openOfNode } from '../paragraphs.ts';
import { OrderedClientRefusal } from '../refusal.ts';

/** A container the editor names, as the list holds it. */
export interface ListContainer {
  readonly story: OrderedStory;
  readonly open: OrderedItem;
}

/** The shown container with this node id, of one of these kinds, in any story. */
export function containerOf(
  doc: OrderedDocument,
  nodeId: string,
  tokens: readonly OrderedContainerToken[]
): ListContainer {
  for (const story of doc.stories.values()) {
    const open = openOfNode(doc, story, nodeId);
    if (!open || open.token === undefined || !tokens.includes(open.token)) continue;
    if (isHidden(doc, open) && !isKeptLink(doc, story, open)) break;
    return { story, open };
  }

  throw new OrderedClientRefusal(
    'ordered-refused',
    'out-of-step',
    `no container ${nodeId.slice(0, 160)}`
  );
}

/** The addresses of containers' open items. */
export const addresses = (containers: readonly ListContainer[]): OrderedChangeAddress[] =>
  containers.map((container) => addressOf(container.open));
