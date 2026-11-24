// The parts that hold a file's tracked changes outside the stories (a style's change records in
// the styles part), as the decisions and the attributions leave them.

import { replaceNode } from '../../../store/package/ooxml-edit.ts';
import { withPart, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import {
  collectRevisionSitesTransient,
  resolveRevisions,
  revisionGroupKey,
} from '../../../store/store/tree-op-revisions.ts';
import { recordRevision } from '../load/part-revisions.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { fileMark } from '../model/revision-ids.ts';
import type { OrderedDocument, OrderedPartRevision } from '../model/types.ts';

/** The last part each file's part gave, by the writes it was given: a publish reads it again. */
const writtenParts = new WeakMap<OoxmlPart, { readonly key: string; readonly part: OoxmlPart }>();

/**
 * A change record group the part writes otherwise than the file: decided, or pending with
 * another attribution (`attribute`).
 */
interface Written {
  readonly suggestion: string;
  readonly revision: OrderedPartRevision;
  readonly action: 'accept' | 'reject' | 'attribute';
}

/** `part` with each record of a group given its suggestion's attribution (`fileMark`). */
function attributedPart(doc: OrderedDocument, part: OoxmlPart, written: Written): OoxmlPart {
  const { revision, element } = written.revision;
  const group = revisionGroupKey(revision, element);
  let next = part;
  for (const { node } of collectRevisionSitesTransient(part)) {
    const own = recordRevision(node);
    if (own === undefined || revisionGroupKey(own, node.localName) !== group) {
      continue;
    }

    // The edit moves its node index to the root it builds: give it a root of its own
    const replaced = replaceNode(
      { ...next, root: { ...next.root } },
      node.id,
      fileMark(doc, node, written.suggestion)
    );
    if (replaced.ok) {
      next = replaced.part;
    }
  }

  return next;
}

/** `part` with the written records, in the order load made their suggestions. */
function writtenPart(doc: OrderedDocument, part: OoxmlPart, writes: readonly Written[]): OoxmlPart {
  let next = part;
  for (const written of writes) {
    const { action, revision } = written;
    if (action === 'attribute') {
      next = attributedPart(doc, next, written);
      continue;
    }

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

/** What a part writes for a suggestion's record group otherwise than the file, if anything. */
function actionOf(doc: OrderedDocument, suggestion: string): Written['action'] | null {
  const status = doc.suggestions.get(suggestion)?.status;
  if (status === 'accepted') {
    return 'accept';
  }

  if (status === 'rejected') {
    return 'reject';
  }

  return doc.revisionIds.reattributed.has(suggestion) ? 'attribute' : null;
}

/** A write as the part's cache key names it: an attribution by what it writes. */
function writeKey(doc: OrderedDocument, { suggestion, action }: Written): string {
  if (action !== 'attribute') {
    return `${suggestion}:${action}`;
  }

  return `${suggestion}:${JSON.stringify(doc.suggestions.get(suggestion)?.attribution)}`;
}

/**
 * `pkg` with each part that holds a decided change record outside the stories written as the
 * decisions leave it. `pkg` holds the file's part there. The engine's own resolver decides each
 * record, as the editor alone and Word decide it: an accept drops the record; a reject puts back
 * the properties it recorded and drops it. Pending records stay as the file wrote them, but for
 * those whose suggestion a change gave another attribution, which take it (`fileMark`). The part
 * is a function of the file's part, the statuses and the attributions alone, so every copy, every
 * publish and the save write the same, in any order of decisions, and a take-back of a decision
 * gives the file's records back.
 */
export function withDecidedPartRevisions(pkg: OoxmlPackage, doc: OrderedDocument): OoxmlPackage {
  const byPart = new Map<string, Written[]>();
  for (const [suggestion, revision] of doc.partRevisions) {
    const action = actionOf(doc, suggestion);
    if (action === null) {
      continue;
    }

    const writes = byPart.get(revision.part) ?? [];
    writes.push({ suggestion, revision, action });
    byPart.set(revision.part, writes);
  }

  let next = pkg;
  for (const [name, writes] of byPart) {
    const part = pkg.parts.get(name);
    if (!part) continue;
    const key = writes.map((written) => writeKey(doc, written)).join(' ');
    let cached = writtenParts.get(part);
    if (cached?.key !== key) {
      cached = { key, part: writtenPart(doc, part, writes) };
      writtenParts.set(part, cached);
    }

    next = withPart(next, cached.part);
  }

  return next;
}
