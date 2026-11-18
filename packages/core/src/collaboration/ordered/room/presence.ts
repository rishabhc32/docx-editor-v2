// Presence rules, as the host's server keeps them today: a client publishes only its own
// presence, under a size cap, and the room names it from the login, whatever it said.
// Presence needs no order and changes no document.

import {
  ORDERED_PROTOCOL_LIMITS,
  ORDERED_PROTOCOL_VERSION,
  type OrderedClientPresenceMessage,
  type OrderedServerPresenceMessage,
} from '../protocol/messages.ts';
import { checkedPerson } from '../model/person.ts';
import type { OrderedActor } from './policy.ts';

/**
 * Why the room dropped a presence message. Stable.
 *
 * - `foreign-presence`: the message names another client's namespace, or the namespace belongs
 *   to another person.
 * - `presence-too-large`: the state is over the presence cap.
 *
 * @public
 */
export type OrderedPresenceRefusalCode = 'foreign-presence' | 'presence-too-large';

/** The presence to relay to everyone else, or why it was dropped. @public */
export type OrderedPresenceVerdict =
  | { readonly ok: true; readonly message: OrderedServerPresenceMessage }
  | { readonly ok: false; readonly code: OrderedPresenceRefusalCode };

/** What presence admission checks against. @public */
export interface OrderedPresenceOptions {
  /** The ID of the person who owns a namespace in the room (`OrderedRoom.namespaceOwner`). */
  readonly ownerOf?: (namespace: string) => string | undefined;
  /** Longest state as JSON. Default {@link ORDERED_PROTOCOL_LIMITS}.presence. */
  readonly maxLength?: number;
}

/**
 * Admit a client's presence: its own namespace only, under the size cap, named by the login's
 * person.
 * Read-only people may send presence too.
 *
 * @throws RangeError when the actor's person is not valid.
 * @public
 */
export function admitOrderedPresence(
  actor: OrderedActor,
  message: OrderedClientPresenceMessage,
  options: OrderedPresenceOptions = {}
): OrderedPresenceVerdict {
  if (message.namespace !== actor.namespace) {
    return { ok: false, code: 'foreign-presence' };
  }

  // The room keeps the checked person's ID as a namespace's owner
  const person = checkedPerson(actor.person);
  const owner = options.ownerOf?.(actor.namespace);
  if (owner !== undefined && owner !== person.id) {
    return { ok: false, code: 'foreign-presence' };
  }

  const cap = options.maxLength ?? ORDERED_PROTOCOL_LIMITS.presence;
  if (JSON.stringify(message.state).length > cap) {
    return { ok: false, code: 'presence-too-large' };
  }

  return {
    ok: true,
    message: {
      v: ORDERED_PROTOCOL_VERSION,
      type: 'presence',
      namespace: actor.namespace,
      person,
      state: message.state,
    },
  };
}
