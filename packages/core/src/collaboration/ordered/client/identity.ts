import type { CollaborationIdentity } from '../../index.ts';
import { safeParticipantColor } from '../../participant-color.ts';
import { ORDERED_NAMESPACE_PATTERN } from '../change/ids.ts';
import { stylesPartOf } from '../../../store/package/ooxml-indexes.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { attributeOf, isWml } from '../ooxml-nodes.ts';

const MAX_IDENTITY = 256;

/**
 * The identity, trimmed and checked: an actor id and a name of 1 to 256 characters, a colour of
 * at most 64 (an unsafe one is dropped).
 *
 * @throws RangeError when the identity is not valid.
 */
export function checkedIdentity(identity: CollaborationIdentity): CollaborationIdentity {
  const actorId = identity.actorId.trim();
  const name = identity.name.trim();
  if (actorId.length === 0 || actorId.length > MAX_IDENTITY)
    throw new RangeError('An identity needs an actor id of 1 to 256 characters.');
  if (name.length === 0 || name.length > MAX_IDENTITY)
    throw new RangeError('An identity needs a name of 1 to 256 characters.');
  if (identity.color !== undefined && identity.color.length > 64)
    throw new RangeError('A participant colour has at most 64 characters.');
  const color = safeParticipantColor(identity.color);
  return Object.freeze({
    actorId,
    name,
    ...(color ? { color } : {}),
    role: identity.role ?? 'human',
  });
}

/** Random letters and digits, from the host's crypto source where it has one. */
function randomSuffix(length: number): string {
  const alphabet = 'abcdefghijklmnopqrstuvwxyz0123456789';
  const bytes = new Uint8Array(length);
  const scope: { readonly crypto?: unknown } = globalThis;
  const crypto = scope.crypto;
  const fill =
    typeof crypto === 'object' && crypto !== null && 'getRandomValues' in crypto
      ? crypto.getRandomValues
      : undefined;
  if (typeof fill === 'function') Reflect.apply(fill, crypto, [bytes]);
  else for (let k = 0; k < length; k += 1) bytes[k] = Math.floor(Math.random() * 256);
  let out = '';
  for (const byte of bytes) out += alphabet.charAt(byte % alphabet.length);
  return out;
}

/**
 * The session's namespace: the host's when it gives one, else a new one made from the actor id
 * and a random suffix. A namespace is new for every session: a counter that starts at 1 again
 * under an old namespace would repeat change ids, and the room answers a repeated id with its
 * first order number and applies nothing.
 *
 * @throws RangeError when the host's namespace is not valid.
 */
export function sessionNamespace(actorId: string, given?: string): string {
  if (given !== undefined) {
    if (!ORDERED_NAMESPACE_PATTERN.test(given))
      throw new RangeError('A namespace is 1 to 64 letters, digits, "_" or "-".');
    return given;
  }

  const slug = actorId.replace(/[^A-Za-z0-9_-]/g, '').slice(0, 24) || 'p';
  return `${slug}-${randomSuffix(16)}`;
}

/**
 * The style id of the document's default paragraph style, or null without one. The layout's
 * rule: the main document's styles part, `w:default` as any on value, the last default wins,
 * and a later style of that id without the flag clears it.
 */
export function defaultParagraphStyle(pkg: OoxmlPackage): string | null {
  let found: string | null = null;
  for (const style of stylesPartOf(pkg)?.root.children ?? []) {
    if (style.kind === 'textValue' || !isWml(style, 'style')) {
      continue;
    }

    const id = attributeOf(style, 'styleId');
    if (attributeOf(style, 'type') !== 'paragraph' || id === undefined) {
      continue;
    }

    const flag = attributeOf(style, 'default');
    if (flag === '1' || flag === 'true' || flag === 'on') {
      found = id;
    } else if (found === id) {
      found = null;
    }
  }

  return found;
}
