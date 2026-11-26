// The undo record. An edit runs in place; every write to the document records how to take it
// back. A copy can then take an edit back without a copy of the document: when the edit throws
// midway, or to put the server's change before its own pending ones and apply them again.

import type { OrderedCounters, OrderedDocument, OrderedUndoStep } from './types.ts';

/** How to take back one edit's writes, last first. @public */
export interface OrderedUndoRecord {
  readonly steps: readonly OrderedUndoStep[];
}

/** An edit's result, and the record that takes the edit back. @public */
export interface OrderedUndoRecorded<T> {
  readonly value: T;
  readonly record: OrderedUndoRecord;
}

/** Each document's state: a new object after each write. */
const STATES = new WeakMap<OrderedDocument, object>();

/**
 * The document's state: the same object until the document's next write (each write records how
 * to take it back, and a take-back writes). What is kept by a state goes with it.
 */
export function stateOf(doc: OrderedDocument): object {
  let state = STATES.get(doc);
  if (state === undefined) {
    state = {};
    STATES.set(doc, state);
  }

  return state;
}

/** Record how to take back one write, when an undo record is open. The document's state moves. */
export function recordUndo(doc: OrderedDocument, step: OrderedUndoStep): void {
  STATES.delete(doc);
  doc.undo?.push(step);
}

/**
 * Run an edit inside an undo record. When the edit throws, its writes are taken back and the
 * error goes on. Records do not nest: one edit is one record.
 *
 * @public
 */
export function withUndoRecord<T>(doc: OrderedDocument, edit: () => T): OrderedUndoRecorded<T> {
  if (doc.undo !== null) throw new Error('An undo record is already open.');
  const counters: OrderedCounters = { ...doc.counters };
  const steps: OrderedUndoStep[] = [() => Object.assign(doc.counters, counters)];
  doc.undo = steps;
  let value: T;
  try {
    value = edit();
  } catch (error) {
    doc.undo = null;
    takeBack(doc, { steps });
    throw error;
  }

  doc.undo = null;
  return { value, record: { steps } };
}

/**
 * Take back an edit's writes, last first. The document must be in the state the edit left,
 * or in a state that later records were already taken back from.
 *
 * @public
 */
export function takeBack(doc: OrderedDocument, record: OrderedUndoRecord): void {
  if (doc.undo !== null) throw new Error('An undo record is open.');
  STATES.delete(doc);
  for (let at = record.steps.length - 1; at >= 0; at -= 1) record.steps[at]?.();
}
