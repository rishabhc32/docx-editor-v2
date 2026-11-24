// Everything an edit can write, as plain data: what an undo record must take back exactly.

import type { OrderedDocument } from '../index.ts';

const identities = new WeakMap<object, number>();
let nextIdentity = 0;
const identityOf = (item: object): number => {
  let id = identities.get(item);
  if (id === undefined) {
    id = nextIdentity += 1;
    identities.set(item, id);
  }
  return id;
};

const keyOf = (item: { space: string; offset: number; length: number }) =>
  `${item.space}|${item.offset}+${item.length}`;

/** Everything an edit can write, as plain data; `withIdentity` also names each item object. */
export function snapshot(doc: OrderedDocument, withIdentity = true): unknown {
  const stories = [...doc.stories].map(([name, story]) => {
    const problems = story.tree.checkInvariants(
      (a, b) => a.units === b.units && a.shown === b.shown && a.marks === b.marks
    );
    return {
      name,
      problems,
      summary: { ...story.tree.summary },
      items: story.tree.toArray().map((item) => ({
        ...(withIdentity ? { identity: identityOf(item) } : {}),
        key: keyOf(item),
        kind: item.kind,
        text: item.text,
        node: item.node?.id ?? null,
        by: item.by,
        dels: [...item.dels],
        fmts: item.fmts.map((f) => `${f.id}:${f.suggestion}:${f.seq}`),
        props: item.props.map((r) => `${r.suggestion}:${r.target}:${r.seq}`),
        move: item.move,
        retracted: item.retracted,
        gone: item.gone,
        quiet: item.quiet,
        within: { by: [...item.within.by], dels: [...item.within.dels] },
        face: item.face?.id ?? null,
      })),
      spaces: [...story.spaces]
        .map(([space, list]) => [space, list.map(keyOf)] as const)
        .sort(([a], [b]) => (a < b ? -1 : 1)),
    };
  });
  return {
    stories,
    suggestions: [...doc.suggestions.values()].map((s) => ({ ...s })),
    labelled: [...doc.labelled]
      .map(([id, set]) => [id, [...set].map(keyOf).sort()] as const)
      .sort(([a], [b]) => (a < b ? -1 : 1)),
    counters: { ...doc.counters },
    resources: [...doc.resources.added.keys()],
    runShells: [...doc.runShells.keys()].sort(),
    // The tables past the items that an edit may write: every copy holds the same, whatever
    // take-backs it made, or the ids it draws from them differ
    revisionIds: {
      byRole: [...doc.revisionIds.byRole].sort(([a], [b]) => (a < b ? -1 : 1)),
      attributesByRole: [...doc.revisionIds.attributesByRole.keys()].sort(),
      fileIds: [...doc.revisionIds.fileIds].sort(),
      reattributed: [...doc.revisionIds.reattributed].sort(([a], [b]) => (a < b ? -1 : 1)),
    },
    movesWithRanges: [...doc.movesWithRanges].sort(),
    paraIds: [...doc.paraIds].sort(),
    commentIds: [...doc.commentIds].sort(),
    comments: [...doc.comments.values()],
    commentRanges: [...doc.commentRanges],
    originals: [...doc.originals.keys()].sort(),
  };
}
