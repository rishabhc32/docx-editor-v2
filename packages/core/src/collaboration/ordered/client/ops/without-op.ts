// Changes the client makes that no editor op makes (design D13): undo and redo, the host's own
// requests, and a kind the format keeps for hosts. The guard checks that every change kind is
// made by an op or listed here.

import type { OrderedChangeKind } from '../../change/types.ts';
import { DECIDE_RULE, type OrderedPartyRule } from './types.ts';

/** A change kind no editor op makes, who makes it, and the room's rule for a party. */
export interface OrderedChangeWithoutOp {
  readonly kind: OrderedChangeKind;
  readonly madeBy: string;
  /** For a party: the room's rule that refuses it, or null when a party may make it. */
  readonly party: OrderedPartyRule | null;
}

export const ORDERED_CHANGES_WITHOUT_OP: readonly OrderedChangeWithoutOp[] = [
  {
    kind: 'retract',
    madeBy: "a person's undo of their own suggestion (task 6.3)",
    party: null,
  },
  {
    kind: 'revert',
    madeBy:
      "the organiser's undo and redo of every own direct edit, also one inside another person's suggestion (tasks 6.7 and 9.12)",
    party: { code: 'policy-direct', reason: 'text', why: 'a direct edit: a party never makes one' },
  },
  {
    kind: 'decide',
    madeBy: "the host's `requestDecision`, beside the editor's accept and reject (D15)",
    party: DECIDE_RULE,
  },
  {
    kind: 'comment-reply',
    madeBy: "the host's `replyToComment` (D15): the editor's reply names no comment of the room",
    party: null,
  },
  {
    kind: 'comment-resolve',
    madeBy: "the host's `resolveComment` (D15): the editor's resolve names no comment of the room",
    party: { code: 'policy-comment', reason: 'resolve', why: 'a party never resolves a thread' },
  },
  {
    kind: 'move',
    madeBy:
      'no editor op today: the editor moves text by a deletion and an insertion; the kind stays for hosts and the matrix',
    party: null,
  },
];
