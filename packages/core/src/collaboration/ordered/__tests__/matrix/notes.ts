// The matrix's note and equation scenarios (group 9.9): a note deleted, directly (Olive, the
// organiser) and as a suggestion (Ana: its reference struck); a note converted to the other
// kind, alone and with Ana's typing beside its reference in either order; the document's note
// properties; and an equation given a new linear form or removed, alone and after Ana's typing.
// Text boxes are stories of their own: the `boxtext` kinds (`kinds.ts`) meet every edit in the
// single, pairs, presses and trio sets.

import {
  AFTER,
  ANA,
  BEFORE,
  OLIVE,
  present,
  scenario,
  type Scenario,
  type Who,
} from './actions.ts';
import { KINDS, paragraphOf, type MatrixKind } from './catalogue.ts';

/** The kinds whose paragraph cites a note, and the note's kind. */
const NOTES: Readonly<Record<string, 'footnote' | 'endnote'>> = {
  fnref: 'footnote',
  enref: 'endnote',
};

/** The kinds whose paragraph holds an equation. */
const EQUATIONS = new Set(['math', 'mathfrac', 'mathpara']);

const type = (who: Who, p: number, after = false) => ({
  ...who,
  p,
  kind: 'insert' as const,
  at: after ? AFTER : BEFORE,
  text: 'x ',
});

function forNote(kind: MatrixKind, note: 'footnote' | 'endnote'): (Scenario | null)[] {
  const p = paragraphOf(kind);
  const remove = (who: Who) => ({ ...who, p, kind: 'deleteNote' as const, note });
  const convert = { ...OLIVE, p, kind: 'convertNote' as const, note };
  return [
    scenario('notes', kind, 'delete the note directly', [remove(OLIVE)]),
    scenario('notes', kind, 'Ana deletes the note', [remove(ANA)]),
    scenario('notes', kind, 'convert the note', [convert]),
    scenario('notes', kind, 'Ana types before it, Olive converts the note', [
      type(ANA, p),
      convert,
    ]),
    scenario('notes', kind, 'Olive converts the note, Ana types after it', [
      convert,
      type(ANA, p, true),
    ]),
    note === 'footnote'
      ? scenario('notes', kind, 'note properties', [
          { ...OLIVE, p, kind: 'noteProperties', numFmt: 'lowerRoman', numStart: 3 },
        ])
      : null,
  ];
}

function forEquation(kind: MatrixKind): (Scenario | null)[] {
  const p = paragraphOf(kind);
  const form = (linear: string | null) => ({ ...OLIVE, p, kind: 'equation' as const, linear });
  return [
    scenario('notes', kind, 'a new form, directly', [form('y=2')]),
    scenario('notes', kind, 'remove it directly', [form(null)]),
    scenario('notes', kind, 'Ana types before it, Olive gives it a new form', [
      type(ANA, p),
      form('y=2'),
    ]),
  ];
}

export const NOTE_SET: readonly Scenario[] = KINDS.flatMap((kind) => {
  const note = NOTES[kind.key];
  if (note) return forNote(kind, note);
  return EQUATIONS.has(kind.key) ? forEquation(kind) : [];
}).filter(present);
