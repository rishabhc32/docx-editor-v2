// The parts that hold a file's tracked changes outside the stories (a style's change records in
// the styles part), as the decisions leave them.

import { withPart, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { resolveRevisions } from '../../../store/store/tree-op-revisions.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument, OrderedPartRevision } from '../model/types.ts';

/** The last part each file's part gave, by the decisions it was given: a publish reads it again. */
const decidedParts = new WeakMap<OoxmlPart, { readonly key: string; readonly part: OoxmlPart }>();

/** A decided change record: its suggestion, its record and the decision. */
interface Decided {
  readonly suggestion: string;
  readonly revision: OrderedPartRevision;
  readonly action: 'accept' | 'reject';
}

/** `part` with the decided records resolved, in the order load made their suggestions. */
function decidedPart(part: OoxmlPart, decided: readonly Decided[]): OoxmlPart {
  let next = part;
  for (const { revision, action } of decided) {
    const { element, revision: address } = revision;

    // The resolver moves its node index to the root it builds: give it a root of its own
    const own = { ...next, root: { ...next.root } };
    const resolved = resolveRevisions(own, action, address, { localName: element });
    if (!resolved.ok || !resolved.part)
      throw new OrderedRefusal(
        'engine-refused',
        `${action} of ${element} ${address.id} in ${part.name}: ${resolved.reason ?? 'no part'}`
      );
    next = resolved.part;
  }

  return next;
}

/**
 * `pkg` with each part that holds a decided change record outside the stories written as the
 * decisions leave it. `pkg` holds the file's part there. The engine's own resolver decides each
 * record, as the editor alone and Word decide it: an accept drops the record; a reject puts back
 * the properties it recorded and drops it. Pending records stay as the file wrote them. The part
 * is a function of the file's part and the statuses alone, so every copy, every publish and the
 * save write the same, in any order of decisions, and a take-back of a decision gives the file's
 * records back.
 */
export function withDecidedPartRevisions(pkg: OoxmlPackage, doc: OrderedDocument): OoxmlPackage {
  const byPart = new Map<string, Decided[]>();
  for (const [suggestion, revision] of doc.partRevisions) {
    const status = doc.suggestions.get(suggestion)?.status;
    if (status !== 'accepted' && status !== 'rejected') continue;
    const decided = byPart.get(revision.part) ?? [];
    decided.push({ suggestion, revision, action: status === 'accepted' ? 'accept' : 'reject' });
    byPart.set(revision.part, decided);
  }

  let next = pkg;
  for (const [name, decided] of byPart) {
    const part = pkg.parts.get(name);
    if (!part) continue;
    const key = decided.map(({ suggestion, action }) => `${suggestion}:${action}`).join(' ');
    let cached = decidedParts.get(part);
    if (cached?.key !== key) {
      cached = { key, part: decidedPart(part, decided) };
      decidedParts.set(part, cached);
    }

    next = withPart(next, cached.part);
  }

  return next;
}
