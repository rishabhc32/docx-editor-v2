// The properties a change may write, by target. A property is a `w:` element name on the
// target's allowlist; a few take one level of children, each on its own list. Run and
// paragraph properties are the engine's own op vocabulary (`ACCEPTED_*_PROPERTIES`), so a
// change never writes what the engine's ops cannot. A table, row or cell takes the engine's
// own table op instead (`table-op.ts`).

import {
  ACCEPTED_PARAGRAPH_PROPERTIES,
  ACCEPTED_RUN_PROPERTIES,
} from '../../../store/store/tree-op-types.ts';

/** Where a property list goes. @public */
export type OrderedPropertyScope = 'run' | 'paragraph';

/** Property name to the names its children may take; an empty list takes no children. */
type Allowlist = ReadonlyMap<string, ReadonlySet<string>>;

function allowlist(
  names: readonly string[],
  children: Readonly<Record<string, readonly string[]>> = {}
): Allowlist {
  const out = new Map<string, ReadonlySet<string>>();
  for (const name of names)
    out.set(name, new Set(Object.hasOwn(children, name) ? children[name] : []));
  return out;
}

const ALLOWLISTS: Readonly<Record<OrderedPropertyScope, Allowlist>> = {
  run: allowlist(ACCEPTED_RUN_PROPERTIES),
  paragraph: allowlist(ACCEPTED_PARAGRAPH_PROPERTIES, {
    tabs: ['tab'],
    numPr: ['ilvl', 'numId'],
  }),
};

/** Whether a property name is on a scope's allowlist. @public */
export function isOrderedPropertyAllowed(scope: OrderedPropertyScope, localName: string): boolean {
  return ALLOWLISTS[scope].has(localName);
}

/** The child names a property may take in a scope, or null when the property is not allowed. */
export function allowedChildren(
  scope: OrderedPropertyScope,
  localName: string
): ReadonlySet<string> | null {
  return ALLOWLISTS[scope].get(localName) ?? null;
}
