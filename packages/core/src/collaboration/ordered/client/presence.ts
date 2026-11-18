// Presence (design D7, task 6.4): a person's selection as two points, each a paragraph's node id
// and a UTF-16 offset in its text. The list gives every copy the same node ids, so a point names
// the same paragraph everywhere.

import type {
  CollaborationIdentity,
  CollaborationParticipant,
  CollaborationRemoteSelection,
  CollaborationRemoteSelectionAddress,
} from '../../index.ts';
import { safeParticipantColor } from '../../participant-color.ts';
import { orderedAuthorName, type OrderedPerson } from '../model/person.ts';
import {
  ORDERED_PROTOCOL_LIMITS,
  type OrderedPresencePoint,
  type OrderedPresenceState,
  type OrderedServerPresenceMessage,
} from '../protocol/messages.ts';

/** Most people kept, as the editor's participant list caps them. */
const MAX_PARTICIPANTS = 256;

/** A point of this editor: the paragraph's stable id, its node id and its text now. */
export interface OrderedPresenceParagraph {
  readonly paragraphId: string;
  readonly nodeId: string;
  readonly text: string;
}

const field = (
  value: string,
  max: number = ORDERED_PROTOCOL_LIMITS.presenceValue
): string | null => (value.length > 0 && value.length <= max ? value : null);

/** The state a person publishes: their selection and who they are. */
export function orderedPresenceState(
  identity: CollaborationIdentity,
  selection: { readonly anchor: OrderedPresencePoint; readonly head: OrderedPresencePoint } | null
): OrderedPresenceState {
  const data: Record<string, string> = {};
  const actorId = field(identity.actorId);
  if (actorId !== null) data.actorId = actorId;
  const color = identity.color === undefined ? null : field(identity.color, 64);
  if (color !== null) data.color = color;
  if (identity.role === 'agent') data.role = 'agent';
  return { selection, data };
}

interface Remote {
  readonly person: OrderedPerson;
  readonly state: OrderedPresenceState;
}

/** Other people's presence, by namespace. */
export class OrderedPresence {
  readonly #remote = new Map<string, Remote>();

  /**
   * Take another person's presence; false when it changed nothing. Presence needs no order: the
   * last message from a person wins, and a point this copy cannot place is left out until the
   * next one.
   */
  take(message: OrderedServerPresenceMessage, ownNamespace: string): boolean {
    if (message.namespace === ownNamespace) return false;
    if (message.state === null) return this.#remote.delete(message.namespace);
    if (!this.#remote.has(message.namespace) && this.#remote.size >= MAX_PARTICIPANTS) return false;
    this.#remote.set(message.namespace, { person: message.person, state: message.state });
    return true;
  }

  /** Everyone gone: the connection dropped. False when nobody was there. */
  clear(): boolean {
    if (this.#remote.size === 0) return false;
    this.#remote.clear();
    return true;
  }

  participants(local: CollaborationIdentity): readonly CollaborationParticipant[] {
    const out: CollaborationParticipant[] = [{ ...local, isLocal: true }];
    for (const remote of this.#remote.values()) {
      const identity = identityOf(remote);
      out.push({ ...identity, isLocal: false });
      if (out.length >= MAX_PARTICIPANTS) break;
    }

    return Object.freeze(out);
  }

  /** Other people's selections in this editor's paragraphs; points it cannot place are left out. */
  selections(
    resolve: (nodeId: string) => OrderedPresenceParagraph | null
  ): readonly CollaborationRemoteSelection[] {
    const out: CollaborationRemoteSelection[] = [];
    for (const remote of this.#remote.values()) {
      const selection = remote.state.selection;
      if (!selection) continue;
      const anchor = place(selection.anchor, resolve);
      const head = place(selection.head, resolve);
      if (!anchor || !head) continue;
      const identity = identityOf(remote);
      out.push({
        actorId: identity.actorId,
        name: identity.name,
        ...(identity.color === undefined ? {} : { color: identity.color }),
        anchor,
        head,
      });
    }

    return Object.freeze(out);
  }
}

function place(
  point: OrderedPresencePoint,
  resolve: (nodeId: string) => OrderedPresenceParagraph | null
): CollaborationRemoteSelectionAddress | null {
  const paragraph = resolve(point.paragraph);
  if (!paragraph) return null;
  return {
    paragraphId: paragraph.paragraphId,
    nodeId: paragraph.nodeId,
    offset: Math.min(point.offset, paragraph.text.length),
  };
}

/**
 * Who a remote person is: the room names them (`person`, shown by their Word name); the rest is
 * what they sent.
 */
function identityOf(remote: Remote): CollaborationIdentity {
  const data = remote.state.data ?? {};
  const actorId = typeof data.actorId === 'string' ? data.actorId : remote.person.id;
  const color = typeof data.color === 'string' ? safeParticipantColor(data.color) : undefined;
  return {
    actorId,
    name: orderedAuthorName(remote.person),
    ...(color ? { color } : {}),
    role: data.role === 'agent' ? 'agent' : 'human',
  };
}
