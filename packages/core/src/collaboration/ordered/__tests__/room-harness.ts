// A room harness for tests: the room's Word file, the real `applyChange` (`applyOrderedChange`),
// a small fake one built only from the model's mutations (for tests that need an apply that
// fails halfway), actors, and a plain-data snapshot of a document.

import { canonicalOoxmlFingerprint } from '../../../store/package/ooxml-tree.ts';
import {
  createOrderedItem as createItem,
  isHidden,
  itemAt,
  storyOfItem,
} from '../model/document.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { withAttributes, withId } from '../ooxml-nodes.ts';
import { docx, p, packageOf, t } from './word-fixtures.ts';
import {
  insertItems,
  newSuggestion,
  nextSeq,
  relabel,
  setStatus,
  splitItem,
} from '../model/mutations.ts';
import { personAttribution } from '../model/person.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type {
  OrderedDocument,
  OrderedItem,
  OrderedStory,
  OrderedSuggestion,
} from '../model/types.ts';
import type { OrderedChangePlace, OrderedEditChange } from '../change/types.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { applyOrderedChange } from '../change/apply.ts';
import { settleDecidedTables } from '../change/settle-columns.ts';
import type { OrderedApplyChange, OrderedApplyContext } from '../room/room.ts';
import type { OrderedActor } from '../room/policy.ts';
import { testPerson } from './test-persons.ts';

export const BODY = '/word/document.xml';
export const HEADER = '/word/header1.xml';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The room's Word file: three paragraphs, and a header. */
export const ROOM_FILE = docx(
  p(t('The fee is 100 USD.')) + p(t('Due in 30 days.')) + p(t('Signed here.')),
  {
    parts: {
      'word/header1.xml': `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>`,
    },
    relationships: `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes:
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
  }
);

/** A fresh copy of the room's document. */
export const loadRoom = (): OrderedDocument => loadOrderedPackage(packageOf(ROOM_FILE));

/** The n-th paragraph of a story as loaded: its text space and its mark's address. */
export function paragraphOf(doc: OrderedDocument, n: number, part = BODY) {
  const story = doc.stories.get(part);
  const marks =
    story?.tree.toArray().filter((item) => item.kind === 'mark' && item.space.startsWith('m:')) ??
    [];
  const mark = marks[n];
  if (!mark) throw new Error(`no paragraph ${n} in ${part}`);
  return { text: `b:${mark.space.slice(2)}`, mark: { space: mark.space, offset: 0 } };
}

// ---- the applies ----

/** The real apply, as the room and the client run it. */
export const realApply: OrderedApplyChange = applyOrderedChange;

function storyOf(doc: OrderedDocument, part: string): OrderedStory {
  const story = doc.stories.get(part);
  if (!story) throw new OrderedRefusal('unknown-story', part);
  return story;
}

/** The item that starts at a place, split out when the place is inside a text item. */
function cut(doc: OrderedDocument, story: OrderedStory, place: OrderedChangePlace): OrderedItem {
  const address = 'mark' in place ? place.mark : place;
  const item = itemAt(story, address);
  if (!item) throw new OrderedRefusal('unknown-address', `${address.space}|${address.offset}`);
  if (address.offset > item.offset) return splitItem(doc, story, item, address.offset)?.[1] ?? item;
  return item;
}

function suggestionFor(
  doc: OrderedDocument,
  change: OrderedEditChange,
  ctx: OrderedApplyContext
): OrderedSuggestion | null {
  if (change.mode === 'edit') return null;
  if (change.into !== undefined) {
    const into = doc.suggestions.get(change.into);
    if (!into || into.status !== 'pending')
      throw new OrderedRefusal('unknown-suggestion', change.into);
    return into;
  }
  return newSuggestion(doc, personAttribution(ctx.person, ctx.date), ctx.ids.suggestion());
}

/** The items from the one starting at `from` to the one before `to`. */
function range(
  doc: OrderedDocument,
  story: OrderedStory,
  from: OrderedChangePlace,
  to: OrderedChangePlace
) {
  const end = cut(doc, story, to);
  const start = cut(doc, story, from);
  const out: OrderedItem[] = [];
  for (const item of story.tree.items(start)) {
    if (item === end) break;
    out.push(item);
  }
  return out;
}

function retract(doc: OrderedDocument, id: string): void {
  for (const item of [...(doc.labelled.get(id) ?? [])]) {
    const story = storyOfItem(doc, item);
    if (!story) continue;
    if (item.by === id && !item.retracted) relabel(doc, story, item, { retracted: true });
    if (item.dels.includes(id))
      relabel(doc, story, item, { dels: item.dels.filter((d) => d !== id) });
    if (item.fmts.some((f) => f.suggestion === id))
      relabel(doc, story, item, { fmts: item.fmts.filter((f) => f.suggestion !== id) });
  }
}

/** Apply a change with the model's mutations only: insert, delete, format, split, decide, retract. */
export const testApply: OrderedApplyChange = (doc, change, ctx) => {
  switch (change.kind) {
    case 'decide': {
      const decided = change.cards.flatMap((card) => card.suggestions.map((s) => s.suggestion));
      for (const id of decided) {
        setStatus(doc, id, change.status);
      }

      settleDecidedTables(doc, decided, ctx.ids);
      return;
    }
    case 'retract':
      for (const id of change.suggestions) retract(doc, id);
      return;
    case 'comment':
    case 'comment-reply':
    case 'comment-resolve':
      return;
    case 'insert': {
      const story = storyOf(doc, change.part);
      const at = cut(doc, story, change.place);
      const suggestion = suggestionFor(doc, change, ctx);
      const space = ctx.ids.units();
      const item = {
        space,
        kind: 'text' as const,
        length: change.text.length,
        text: change.text,
        run: space,
        by: suggestion?.id ?? null,
      };
      insertItems(doc, story, story.tree.prev(at), [createItem(item)]);
      return;
    }
    case 'delete':
    case 'format': {
      const story = storyOf(doc, change.part);
      const items = range(doc, story, change.from, change.to);
      if (change.kind === 'format' && change.mode === 'edit')
        throw new OrderedRefusal('unknown-address', 'the test apply has no direct format');
      const suggestion = suggestionFor(doc, change, ctx);
      const seq = nextSeq(doc);
      const record = ctx.ids.record();
      for (const item of items) {
        if (isHidden(doc, item) || (item.kind !== 'text' && item.kind !== 'mark')) continue;
        if (change.kind === 'delete') {
          if (suggestion === null) relabel(doc, story, item, { gone: true });
          else if (!item.dels.includes(suggestion.id))
            relabel(doc, story, item, { dels: [...item.dels, suggestion.id] });
        } else if (suggestion !== null && item.kind === 'text') {
          const properties = change.properties.map((p) => ({ localName: p.localName }));
          relabel(doc, story, item, {
            fmts: [
              ...item.fmts,
              { kind: 'properties', id: record, suggestion: suggestion.id, seq, properties },
            ],
          });
        }
      }
      return;
    }
    case 'split': {
      const story = storyOf(doc, change.part);
      const at = cut(doc, story, change.place);
      let own: OrderedItem | null = null;
      for (const item of story.tree.items(at))
        if (item.kind === 'mark') {
          own = item;
          break;
        }
      if (!own || !own.node || own.node.kind === 'textValue')
        throw new OrderedRefusal('unknown-address', 'no paragraph after the place');
      const suggestion = suggestionFor(doc, change, ctx);
      const paragraph = ctx.ids.paragraph();
      const shell = withAttributes(
        withId(own.node, paragraph.space),
        own.node.attributes.filter((a) => a.localName !== 'paraId' && a.localName !== 'textId')
      );
      const mark = createItem({
        space: paragraph.space,
        kind: 'mark',
        node: shell,
        face: own.face,
        by: suggestion?.id ?? null,
      });
      insertItems(doc, story, story.tree.prev(at), [mark]);
      return;
    }
    default:
      throw new OrderedRefusal('unknown-address', `the test apply has no ${change.kind}`);
  }
};

// ---- snapshots ----

const keyOf = (item: OrderedItem) => `${item.space}|${item.offset}+${item.length}`;

/** Everything an edit can write, as plain data, without object identity. */
export function snapshot(doc: OrderedDocument): unknown {
  return {
    stories: [...doc.stories].map(([name, story]) => ({
      name,
      summary: { ...story.tree.summary },
      items: story.tree.toArray().map((item) => ({
        key: keyOf(item),
        kind: item.kind,
        text: item.text,
        node: item.node?.id ?? null,
        by: item.by,
        dels: [...item.dels],
        fmts: item.fmts.map((f) => `${f.id}:${f.suggestion}:${f.seq}`),
        retracted: item.retracted,
        gone: item.gone,
      })),
      spaces: [...story.spaces]
        .map(([space, list]) => [space, list.map(keyOf)] as const)
        .sort(([a], [b]) => (a < b ? -1 : 1)),
    })),
    suggestions: [...doc.suggestions.values()].map((s) => ({
      ...s,
      attribution: { ...s.attribution },
    })),
    labelled: [...doc.labelled]
      .map(([id, set]) => [id, [...set].map(keyOf).sort()] as const)
      .sort(([a], [b]) => (a < b ? -1 : 1)),
    counters: { ...doc.counters },
    resources: [...doc.resources.added.keys()],
  };
}

/** The body's composed part, as the fidelity oracle sees it. */
export const bodyFingerprint = (doc: OrderedDocument): string =>
  canonicalOoxmlFingerprint(composeOrderedPart(doc, BODY).part);

// ---- actors ----

export function actor(
  name: string,
  namespace: string,
  role: 'party' | 'organiser' | 'reader'
): OrderedActor {
  return {
    person: testPerson(name),
    namespace,
    readOnly: role === 'reader',
    suggests: role === 'party',
    decides: role === 'organiser',
    comments: role !== 'reader',
  };
}
