// The side state of a saved item list (checkpoint version 6): state modules keep for a document
// beside its fields, each saved and given back through the table below.

import {
  emptiedControlsOf,
  restoreEmptiedControls,
  restoreSplitParts,
  splitPartsOf,
  type PromptedByEdit,
  type Untracked,
} from '../change/settle-emptied.ts';
import {
  namedPromptStateOf,
  restoreNamedPromptState,
  type NamedPromptState,
  type Swap,
} from '../change/settle-named-prompts.ts';
import { at, fieldsOf, required } from '../change/read.ts';
import {
  droppedMarksOf,
  filledPromptsOf,
  firstFillsOf,
  noteFirstFill,
  recordDroppedMark,
  recordFilledPrompt,
} from './prompt-marks.ts';
import {
  bool,
  count,
  id,
  item,
  list,
  map,
  maybe,
  node,
  nullable,
  record,
  set,
  stored,
  type SnapshotCodec,
  type SnapshotFields,
  type SnapshotReadContext,
  type SnapshotWriteContext,
} from './snapshot-codec.ts';
import { ADDRESS, NAMESPACE } from './snapshot-fields.ts';
import { registerTextBox, textBoxHostsOf, type OrderedTextBoxHost } from './text-boxes.ts';
import { keepLoadedTypers, loadedTypersOf, type TypedRun } from './typers.ts';
import type { OrderedDocument } from './types.ts';

/**
 * State a module keeps for a document beside its fields (a side map by document): how to take it
 * from a document, and how to give it back to an opened one. Caches and indexes that `indexItem`
 * makes again are no side state.
 */
export interface SnapshotSide<V> {
  readonly codec: SnapshotCodec<V>;
  of(doc: OrderedDocument): V;
  restore(doc: OrderedDocument, value: V): void;
}

const side = <V>(
  codec: SnapshotCodec<V>,
  of: (doc: OrderedDocument) => V,
  restore: (doc: OrderedDocument, value: V) => void
): SnapshotSide<V> => ({ codec, of, restore });

export const TEXT_BOX_HOST: SnapshotFields<OrderedTextBoxHost> = {
  host: stored(id),
  atom: stored(record(ADDRESS)),
  root: stored(id),
};

export const TYPED_RUN: SnapshotFields<TypedRun> = {
  from: stored(count),
  to: stored(count),
  typer: stored(NAMESPACE),
};

const items = list(item);

export const UNTRACKED: SnapshotFields<Untracked> = { by: stored(id), units: stored(items) };

export const PROMPTED_BY_EDIT: SnapshotFields<PromptedByEdit> = {
  node: stored(nullable(node)),
  original: maybe(node),
  dropped: stored(items),
  went: stored(items),
  untracked: stored(nullable(record(UNTRACKED))),
  mark: stored(item),
};

export const SWAP: SnapshotFields<Swap> = {
  replaced: stored(items),
  units: stored(items),
  went: stored(items),
};

export const NAMED_PROMPT_STATE: SnapshotFields<NamedPromptState> = {
  swaps: stored(map(item, record(SWAP))),
  settled: stored(items),
  replaced: stored(items),
  kept: stored(set(id)),
};

/** Every side state, by the name the checkpoint saves it under. */
export const SIDE: Readonly<Record<string, SnapshotSide<unknown>>> = {
  textBoxes: side(map(id, record(TEXT_BOX_HOST)), textBoxHostsOf, (doc, hosts) => {
    for (const [name, host] of hosts) {
      registerTextBox(doc, name, host);
    }
  }),
  typers: side(map(id, list(record(TYPED_RUN))), loadedTypersOf, keepLoadedTypers),
  droppedMarks: side(items, droppedMarksOf, (doc, marks) => {
    for (const mark of marks) {
      recordDroppedMark(doc, mark, true);
    }
  }),
  filledPrompts: side(set(id), filledPromptsOf, (doc, spaces) => {
    for (const filled of spaces) {
      recordFilledPrompt(doc, filled);
    }
  }),
  firstFills: side(map(id, bool), firstFillsOf, (doc, fills) => {
    for (const [control, own] of fills) {
      noteFirstFill(doc, control, own);
    }
  }),
  emptiedControls: side(
    map(item, record(PROMPTED_BY_EDIT)),
    emptiedControlsOf,
    restoreEmptiedControls
  ),
  namedPrompts: side(record(NAMED_PROMPT_STATE), namedPromptStateOf, restoreNamedPromptState),
  splitParts: side(map(item, list(record(ADDRESS))), splitPartsOf, restoreSplitParts),
};

const NAMES = new Set(Object.keys(SIDE));

/** Every side state of `doc`, as the checkpoint writes it. */
export function sideJson(doc: OrderedDocument, w: SnapshotWriteContext): Record<string, unknown> {
  return Object.fromEntries(
    Object.entries(SIDE).map(([name, state]) => [name, state.codec.write(state.of(doc), w)])
  );
}

/**
 * Give an opened document every side state the checkpoint saved.
 *
 * @throws Invalid when one does not read.
 */
export function restoreSide(doc: OrderedDocument, json: unknown, r: SnapshotReadContext): void {
  const f = fieldsOf(json, 'side', NAMES);
  for (const [name, state] of Object.entries(SIDE)) {
    state.restore(doc, state.codec.read(required(f, name), at('side', name), r));
  }
}

/** The first side state where `b` differs from `a`, or null. */
export function sideDifference(a: OrderedDocument, b: OrderedDocument): string | null {
  for (const [name, state] of Object.entries(SIDE)) {
    const found = state.codec.diff(state.of(a), state.of(b), at('side', name));
    if (found !== null) {
      return found;
    }
  }

  return null;
}
