// Note, equation and text box ops: a new note (task 6.1); deleted and converted notes, note
// properties, equations and typing inside text boxes (group 9.9).

import { ORDERED_CHANGE_LIMITS } from '../../change/types.ts';
import { addressOf, placeAt } from '../paragraphs.ts';
import { OrderedClientRefusal } from '../refusal.ts';
import { equationOf, noteReferences, sectionHolders, shownNotes } from './note-references.ts';
import { translator, type OrderedOpOption, type OrderedOpTable } from './types.ts';
import { untrackedRule } from './untracked.ts';

export const NOTE_OPS: OrderedOpTable<
  | 'insertNote'
  | 'deleteNote'
  | 'convertNote'
  | 'convertAllNotes'
  | 'setNoteProperties'
  | 'setMathEquation'
  | 'removeMathEquation'
> = {
  insertNote: {
    status: 'translated',
    group: '6.1',
    changes: ['insert-note'],
    party: null,
    note: 'A footnote or an endnote at a place, its paragraph empty.',
    translate: translator('insertNote', (op, batch) => () => {
      const p = batch.paragraph(op.paragraphId);
      const place = placeAt(p, op.offset);
      return [{ kind: 'insert-note', part: p.story.name, place, note: op.noteKind, text: '' }];
    }),
  },
  deleteNote: {
    status: 'translated',
    group: '9.9',
    changes: ['delete'],
    party: untrackedRule('text', 'the reference and the note go; suggesting strikes the reference'),
    note: 'A note and its references removed: a direct deletion of each reference, which takes the note with it.',
    translate: translator('deleteNote', (op, batch) => () => {
      const references = noteReferences(batch, op.noteKind, op.noteId);
      if (references.length === 0)
        throw new OrderedClientRefusal(
          'ordered-refused',
          'out-of-step',
          'a note with no reference'
        );
      return references.map(({ paragraph, offset }) => ({
        kind: 'delete',
        part: paragraph.story.name,
        from: placeAt(paragraph, offset),
        to: placeAt(paragraph, offset + 1),
      }));
    }),
  },
  convertNote: {
    status: 'translated',
    group: '9.9',
    changes: ['convert-note'],
    party: untrackedRule('package', 'a footnote becomes an endnote or back'),
    tracking: 'direct',
    note: 'One note to the other kind, as Word converts it directly: a copy of the note and of its references.',
    translate: translator('convertNote', (op, batch) => () => {
      const found = shownNotes(batch.doc, op.fromKind, op.noteId);
      const open = found?.opens[0];
      if (!found || !open)
        throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'no such note');
      return [{ kind: 'convert-note', part: found.story.name, notes: [addressOf(open)] }];
    }),
  },
  convertAllNotes: {
    status: 'translated',
    group: '9.9',
    changes: ['convert-note'],
    party: untrackedRule('package', 'every note of a kind becomes the other kind'),
    tracking: 'direct',
    note: 'Every note of one kind to the other, in the order of their story, a change for each 1000.',
    translate: translator('convertAllNotes', (op, batch) => () => {
      const found = shownNotes(batch.doc, op.fromKind);
      if (!found) return [];
      const notes = found.opens.map(addressOf);
      const size = ORDERED_CHANGE_LIMITS.targets;
      const out = [];
      for (let k = 0; k < notes.length; k += size)
        out.push({
          kind: 'convert-note' as const,
          part: found.story.name,
          notes: notes.slice(k, k + size),
        });
      return out;
    }),
  },
  setNoteProperties: {
    status: 'translated',
    group: '9.9',
    changes: ['note-properties'],
    party: untrackedRule('package', "notes' number format, start and place"),
    tracking: 'direct',
    note: "Notes' number format, restart, start and position, for the document's settings or for a section, by the item that holds its `w:sectPr`.",
    translate: translator('setNoteProperties', (op, batch) => () => {
      const found = sectionHolders(batch.doc);
      if (!found)
        throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'no main document');
      const fields = {
        ...(op.footnote ? { footnote: op.footnote } : {}),
        ...(op.endnote ? { endnote: op.endnote } : {}),
      };
      const part = found.story.name;
      if (op.scope === 'document') return [{ kind: 'note-properties', part, ...fields }];
      const holder = op.sectionIndex === undefined ? undefined : found.holders[op.sectionIndex];
      if (!holder)
        throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'no such section');
      const section =
        holder.kind === 'zero' ? { body: addressOf(holder) } : { mark: addressOf(holder) };
      return [{ kind: 'note-properties', part, section, ...fields }];
    }),
  },
  setMathEquation: {
    status: 'translated',
    group: '9.9',
    changes: ['equation'],
    party: untrackedRule('text', 'an equation is rewritten'),
    note: 'An equation from its linear form: the old object struck and the new one inserted, one suggestion.',
    translate: translator('setMathEquation', (op, batch) => () => {
      const found = equationOf(batch, op.equationId);
      return [
        {
          kind: 'equation',
          part: found.paragraph.story.name,
          object: addressOf(found.item),
          equation: found.equation,
          linear: op.linear,
        },
      ];
    }),
  },
  removeMathEquation: {
    status: 'translated',
    group: '9.9',
    changes: ['delete', 'equation'],
    party: untrackedRule('text', 'an equation goes'),
    note: "An equation removed: a deletion of its object's unit, or the display without it when the display holds others.",
    translate: translator('removeMathEquation', (op, batch) => () => {
      const found = equationOf(batch, op.equationId);
      const part = found.paragraph.story.name;
      if (found.count > 1)
        return [
          {
            kind: 'equation',
            part,
            object: addressOf(found.item),
            equation: found.equation,
            linear: null,
          },
        ];
      return [
        {
          kind: 'delete',
          part,
          from: placeAt(found.paragraph, found.offset),
          to: placeAt(found.paragraph, found.offset + 1),
        },
      ];
    }),
  },
};

/** Options of the note ops, and typing inside a text box. */
export const NOTE_OPTIONS: readonly OrderedOpOption[] = [
  {
    op: 'insertText',
    option: 'text-box',
    status: 'translated',
    group: '9.9',
    note: 'Every text edit inside a text box: each text box of the main document is a story of its own, written back into its object.',
  },
  {
    op: 'acceptRevision',
    option: 'equation-part',
    status: 'planned',
    group: '9.9',
    refusal: 'ordered-unsupported-op:acceptRevision.equation-part',
    note: "A decision (accept or reject, one or all) that names Word's mark inside some math runs of an equation only (a `w:ins` or `w:del` in `m:r`): the list has no unit inside an equation, so a room keeps such marks as file content and refuses their decision by this name. The editor alone decides them run by run. A whole equation marked in one direction by one author at one date is a suggestion on the equation (D20).",
  },
];
