// Every matrix cell that does not pass every check, by cause. A cell on this list must give
// exactly the codes listed; every other cell must pass every check. An entry that matches no
// cell fails too (`known.test.ts`), so the list only shrinks as the engine or the model is fixed.
//
// Owners:
//   engine   the engine does something else than Word or the decided policy (a defect, or a rule
//            the model does not follow on purpose);
//   harness  the oracle cannot make the same edit (the engine has no move; a command that
//            flips sides in a right-to-left table);
//   word     Word's file form cannot hold the state, so a save and reopen loses it;
//   model    a limit of the model, refused by name or listed as an open item;
//   both     the engine and the model both refuse the edit by name.

import type { SetName } from './actions.ts';
import {
  BOOKMARK_TYPING,
  PROMPT_DELETE_BOTH,
  PROMPT_DELETE_MEET,
  PROMPT_IN_CONTAINER,
} from './known-cells.ts';

export interface KnownDifference {
  readonly cause: string;
  readonly owner: 'engine' | 'harness' | 'word' | 'model' | 'both';
  /** The codes each matching cell gives (`VODWLISERN`, see `check.ts`). */
  readonly codes: string;
  readonly sets: readonly SetName[];
  /** The kinds (the part of a name before ` | `), or a test of the kind. */
  readonly kinds: readonly string[] | ((kind: string) => boolean);
  /** The edits (the part after ` | `), or a test of the edit. */
  readonly edits: readonly string[] | ((edit: string) => boolean);
}

const PROMPTS = ['placeholder', 'dropprompt', 'dateprompt', 'richprompt'];
const FIELDS = [
  'pagefield',
  'ref',
  'numpages',
  'datefield',
  'hyperfield',
  'noteref',
  'seq',
  'merge',
  'autonum',
  'styleref',
  'citation',
  'formtext',
  'formcheck',
];
const ONE_COLUMN = [
  'cantsplit',
  'floattable',
  'headrow',
  'listcell',
  'nested',
  'pictable',
  'rtltable',
  'tblstyle',
  'textdir',
  'cellrev',
  'span',
];

const inside = (containers: readonly string[], items: readonly string[]) => (kind: string) =>
  containers.some((c) => items.some((i) => kind === `${i} in ${c}`));
const all = () => true;

export const KNOWN: readonly KnownDifference[] = [
  // ---- the engine ----
  {
    cause:
      'The engine has no move: its reference is a cut and a plain paste, which loses formats and objects',
    owner: 'harness',
    codes: 'VD',
    sets: ['single'],
    kinds: [
      'bold',
      'br',
      'caps',
      'charstyle',
      'checkbox',
      'column',
      'complex',
      'cr',
      'dstrike',
      'eastasia',
      'effects',
      'face',
      'fmt',
      'font',
      'glow',
      'hidden',
      'italic',
      'link',
      'nbh',
      'page',
      'rtl',
      'runshade',
      'size',
      'smallcaps',
      'soft',
      'strike',
      'styled',
      'sub',
      'sup',
      'sym',
      'tab',
      'theme',
      'udouble',
      'uwave',
      'watermark',
    ],
    edits: ['move it'],
  },
  {
    cause:
      'The engine: a decision on moved-from text in a combo box leaves its value record (a deletion updates it)',
    owner: 'engine',
    codes: 'W',
    sets: ['single'],
    kinds: ['combo'],
    edits: ['move it'],
  },
  {
    cause:
      "The engine: a range deletion strikes a control's prompt, and accepting it leaves the control empty (the model leaves a prompt alone, rule 7)",
    owner: 'engine',
    codes: 'VD',
    sets: ['single', 'pairs'],
    kinds: PROMPTS,
    edits: [
      'delete across',
      'replace across',
      'delete from the paragraph before',
      'delete into the paragraph after',
      'Ana deletes across, Ben deletes it',
      'Ana deletes it, Ben replaces across',
    ],
  },
  {
    cause:
      "The engine: a range deletion strikes a control's prompt, and accepting it leaves the control empty (the model leaves a prompt alone, rule 7)",
    owner: 'engine',
    codes: 'VD',
    sets: ['mix'],
    kinds: PROMPT_DELETE_BOTH,
    edits: ['delete both'],
  },
  {
    cause:
      "The engine: a range deletion strikes a control's prompt, and accepting it leaves the control empty (the model leaves a prompt alone, rule 7)",
    owner: 'engine',
    codes: 'VD',
    sets: ['mix'],
    kinds: PROMPT_DELETE_MEET,
    edits: ['delete where they meet'],
  },
  {
    cause:
      "The engine: a range deletion strikes a control's prompt, and accepting it leaves the control empty (the model leaves a prompt alone, rule 7)",
    owner: 'engine',
    codes: 'VD',
    sets: ['nested'],
    kinds: PROMPT_IN_CONTAINER,
    edits: ['delete the container', 'delete across its start'],
  },
  {
    cause: 'The engine: typing or Enter moves a bookmark start in front of the markers before it',
    owner: 'engine',
    codes: 'VD',
    sets: ['mix'],
    kinds: BOOKMARK_TYPING,
    edits: ['type between'],
  },
  {
    cause: 'The engine: typing or Enter moves a bookmark start in front of the markers before it',
    owner: 'engine',
    codes: 'VD',
    sets: ['mix'],
    kinds: ['comment+bookmark', 'comment+emptymark', 'ptab+bookmark', 'ptab+emptymark'],
    edits: ['Enter between'],
  },
  {
    cause:
      'The engine: a field inside a text control refuses typing beside it and drops deletions of it',
    owner: 'engine',
    codes: 'E',
    sets: ['nested'],
    kinds: inside(['plain', 'rich', 'group', 'gallery'], ['simple']),
    edits: ['type before it', 'type after it', 'Ana types in it, Ben deletes the container'],
  },
  {
    cause:
      'The engine: a field inside a text control refuses typing beside it and drops deletions of it',
    owner: 'engine',
    codes: 'VD',
    sets: ['nested'],
    kinds: inside(['plain', 'rich', 'group', 'gallery'], ['simple']),
    edits: ['delete it', 'delete the container', 'delete across its start'],
  },
  {
    cause:
      'The engine: a direct insertion beside struck text that ends in an object lands before the struck text',
    owner: 'engine',
    codes: 'VD',
    sets: ['pairs', 'trio'],
    kinds: [
      'anchor',
      'behind',
      'dateprompt',
      'dropprompt',
      'enref',
      'fnref',
      'formcheck',
      'placeholder',
      'richprompt',
      'simple',
      'tight',
      'wraptb',
      'listcell+picture+footnote',
    ],
    edits: [
      'Ana deletes across, Olive types after',
      'Ana types after, Cy deletes across, Olive types before',
    ],
  },
  {
    cause:
      "On purpose: a direct deletion of a comment's whole range keeps the comment at the deletion's place (the words of a discussion deleted stay discussed); the engine takes the comment away with its words",
    owner: 'engine',
    codes: 'VD',
    sets: ['pairs'],
    kinds: ['comment'],
    edits: ['Ana bolds, Olive deletes it', 'Ana types after, Olive deletes across'],
  },
  {
    cause:
      'The engine: a direct deletion across a control leaves it empty, or keeps its value record',
    owner: 'engine',
    codes: 'VD',
    sets: ['pairs'],
    kinds: [
      'combo',
      'date',
      'dateprompt',
      'dropdown',
      'dropprompt',
      'gallery',
      'placeholder',
      'plain',
      // A control with no type element is rich text, so its prompt comes back when emptied
      'rich',
      'richprompt',
    ],
    edits: ['Ana types after, Olive deletes across'],
  },
  {
    cause:
      'The engine: a direct deletion across a control leaves it empty, or keeps its value record',
    owner: 'engine',
    codes: 'VD',
    sets: ['pairs'],
    kinds: ['date', 'dropdown'],
    edits: ['Ana bolds, Olive deletes it'],
  },
  {
    cause:
      'The engine refuses a second deletion of a struck simple field (unsupported); policy 1 makes it do nothing',
    owner: 'engine',
    codes: 'E',
    sets: ['pairs'],
    kinds: ['simple'],
    edits: ['Ana deletes across, Ben deletes it', 'Ana deletes it, Ben replaces across'],
  },
  {
    cause: 'The engine refuses a second deletion of a struck row; policy 1 makes it do nothing',
    owner: 'engine',
    codes: 'E',
    sets: ['structures'],
    kinds: ['headrow', 'mergeboth', 'span', 'table', 'tblstyle', 'vmerge'],
    edits: ['Ana deletes the row, Ben deletes it too'],
  },
  {
    cause:
      'The engine gives a struck paragraph mark to the second deleter; policy 1 keeps the first',
    owner: 'engine',
    codes: 'VD',
    sets: ['single'],
    kinds: ['markdel'],
    edits: ['Delete at end', 'delete into the paragraph after'],
  },
  {
    cause:
      'The engine strikes moved-from text a second time (a w:del inside w:moveFrom); policy 1 leaves the move alone',
    owner: 'engine',
    codes: 'VD',
    sets: ['single'],
    kinds: ['movefile'],
    edits: ['delete from the paragraph before'],
  },
  {
    cause:
      "The engine formats a check box form field's separator run (its result is empty); a format reaches a field's result only (rule 9)",
    owner: 'engine',
    codes: 'VD',
    sets: ['single', 'pairs'],
    kinds: ['formcheck'],
    edits: ['bold across', 'italic it', 'Ana bolds, Ben italics it', 'Ana bolds, Olive italics it'],
  },
  {
    cause: 'The engine: a decision in a table with cells merged down rows rewrites the table grid',
    owner: 'engine',
    codes: 'DW',
    sets: ['structures'],
    kinds: ['vmerge', 'mergeboth'],
    edits: ['delete row', 'Ana types in the row, Ben deletes the row'],
  },
  {
    cause: 'The engine: a decision in a table with cells merged down rows rewrites the table grid',
    owner: 'engine',
    codes: 'DW',
    sets: ['structures'],
    kinds: ['mergeboth'],
    edits: [
      'insert row above',
      'insert row below',
      'row height',
      'table centered',
      'table style',
      'delete column',
      'column widths',
      'cell fill',
      'cell vertical alignment',
      'cell borders',
      'Ana fills the cell, Ben types inside',
      'Ana fills the cell, Ben gives it the same fill',
      'Ana sets the row height, then fills the cell',
    ],
  },
  // ---- the oracle ----
  {
    cause:
      "The editor's command flips left and right in a right-to-left table; the model takes the side the client sends",
    owner: 'harness',
    codes: 'VD',
    sets: ['structures'],
    kinds: ['rtltable'],
    edits: ['insert column left', 'insert column right', 'Olive inserts a column'],
  },
  {
    cause:
      "The editor's command flips left and right in a right-to-left table; the model takes the side the client sends",
    owner: 'harness',
    codes: 'V',
    sets: ['structures'],
    kinds: ['rtltable'],
    edits: ['Ana inserts a column, Ben inserts one too'],
  },
  // ---- Word's form, and the model's limits ----
  {
    cause:
      "Both refuse another person's change of a cell's or a table's properties over a pending one: Word keeps one change record for each (`unsupported-revision`)",
    owner: 'both',
    codes: 'R',
    sets: ['structures'],
    kinds: (kind) => kind !== 'cellrev',
    edits: ['Ana fills the cell, Ben: cell borders', 'Ana: table centered, Ben: table style'],
  },
  {
    cause:
      "A field's result is one unit of the model, so a format over it records on each of its runs, also on a run that already has the value, where the engine writes no record (backlog 25); a direct format over that run then writes Word's one run record for it, so a rejection takes the direct format there too",
    owner: 'model',
    codes: 'VD',
    sets: ['pairs'],
    kinds: ['ref'],
    edits: ['Ana bolds, Olive italics it'],
  },
  {
    cause:
      'Both refuse Enter in a control that holds one line (a plain text control without `w:multiLine`, a list, a date) and inside an inline rich text control, as Word does (`single-line-control`)',
    owner: 'both',
    codes: 'R',
    sets: ['single'],
    kinds: ['plain', 'placeholder', 'dropdown', 'combo', 'date', 'rich'],
    edits: ['Enter inside'],
  },
  {
    cause:
      'Both refuse Enter in a plain text control that holds one line, or in an inline rich text control, before a kind inside it (`single-line-control`)',
    owner: 'both',
    codes: 'R',
    sets: ['nested'],
    kinds: (kind) => kind.endsWith(' in plain') || kind.endsWith(' in rich'),
    edits: ['Enter before it'],
  },
  {
    cause:
      'Both refuse Enter in the prompt of a list, date or inline rich text control (the model: `break-in-container`, policy 4; the engine: `single-line-control`)',
    owner: 'both',
    codes: 'R',
    sets: ['single'],
    kinds: ['dropprompt', 'dateprompt', 'richprompt'],
    edits: ['Enter inside'],
  },
  {
    cause:
      "A tab or a break typed in a block control's prompt: the engine puts it inside the prompt text and keeps the prompt; the model replaces the prompt, as typed text does (rule 7)",
    owner: 'engine',
    codes: 'VD',
    sets: ['single'],
    kinds: ['blockprompt'],
    edits: ['tab before', 'line break after', 'page break before'],
  },
  {
    cause: 'The model refuses a move of text another suggestion holds, or of a prompt, by name',
    owner: 'model',
    codes: 'S',
    sets: ['single'],
    kinds: [
      'blockprompt',
      'dateprompt',
      'del',
      'dropprompt',
      'ins',
      'placeholder',
      'richprompt',
      'rowins',
    ],
    edits: ['move it'],
  },
  {
    cause:
      'The model refuses a move of words out of a control bound to custom XML, by name; the engine cuts them',
    owner: 'model',
    codes: 'S',
    sets: ['single'],
    kinds: ['bound'],
    edits: ['move it'],
  },
  {
    cause:
      "The model refuses deleting a one-column table's column by the engine's own refusal (block-required); the editor's command removes the table",
    owner: 'model',
    codes: 'S',
    sets: ['structures'],
    kinds: ONE_COLUMN.filter((kind) => kind !== 'span'),
    edits: [
      'delete column',
      'Olive deletes a column',
      'Ana deletes the column, Ben inserts one beside it',
    ],
  },
  {
    cause:
      "The caret's cell covers every column, so the editor's command deletes the table (`deleteBlock`); the model's column change deletes the column as Word does (its cell, and the other row's cell under it)",
    owner: 'harness',
    codes: 'VD',
    sets: ['structures'],
    kinds: ['span'],
    edits: [
      'delete column',
      'Olive deletes a column',
      'Ana deletes the column, Ben inserts one beside it',
    ],
  },
  {
    cause:
      "The step's words are gone after the first edit (typing replaced the prompt), so the step does not apply",
    owner: 'harness',
    codes: 'N',
    sets: ['pairs'],
    kinds: ['blockprompt'],
    edits: ['Ana types after, Ben deletes across', 'Ana types after, Olive deletes across'],
  },
  // ---- refused by both, by name ----
  {
    cause:
      'Both refuse a place between the two UTF-16 units of one character (splits-surrogate-pair)',
    owner: 'both',
    codes: 'R',
    sets: ['single', 'pairs'],
    kinds: ['emoji'],
    edits: ['type inside', 'Enter inside', 'Ana types inside, Ben deletes it'],
  },
  {
    cause: 'Both refuse edits of a control bound to custom XML (not-covered; the engine: bound)',
    owner: 'both',
    codes: 'R',
    sets: ['single', 'pairs'],
    kinds: ['bound'],
    edits: [
      'type inside',
      'Enter inside',
      'delete it',
      'delete across',
      'replace across',
      'bold across',
      'italic it',
      'delete from the paragraph before',
      'delete into the paragraph after',
      'Ana types inside, Ben deletes it',
      'Ana types after, Ben deletes across',
      'Ana deletes it, Ben types before',
      'Ana deletes it, Ben types inside',
      'Ana bolds, Ben italics it',
      'Ana deletes across, Ben deletes it',
      'Ana deletes it, Ben replaces across',
      'Ana deletes across, Olive types after',
      'Ana types after, Olive deletes across',
      'Ana bolds, Olive deletes it',
      'Ana bolds, Olive italics it',
    ],
  },
  {
    cause:
      'The model refuses deleting a prompt in a control bound to custom XML, by name; the engine leaves a selected prompt alone, as Word does, so the editor sends nothing',
    owner: 'model',
    codes: 'S',
    sets: ['nested'],
    kinds: ['dropprompt in bound'],
    edits: ['delete it'],
  },
  {
    cause: 'Both refuse edits of a control bound to custom XML (not-covered; the engine: bound)',
    owner: 'both',
    codes: 'R',
    sets: ['nested'],
    kinds: (k) => k.endsWith(' in bound'),
    edits: all,
  },
  {
    cause:
      'Both refuse edits of a table of contents result (not-covered; the engine: generated and read-only)',
    owner: 'both',
    codes: 'R',
    sets: ['single', 'pairs'],
    kinds: ['toc'],
    edits: (e) => e !== 'Backspace at start' && e !== 'center',
  },
  {
    cause: 'Both refuse an outline level press: the engine does not author `w:outlineLvl`',
    owner: 'both',
    codes: 'R',
    sets: ['presses'],
    kinds: all,
    edits: ['outline level'],
  },
  {
    cause: 'Both refuse a new column in a table with a pending cell change (unsupported-revision)',
    owner: 'both',
    codes: 'R',
    sets: ['structures'],
    kinds: ['cellrev'],
    edits: [
      'insert column left',
      'insert column right',
      'Ana inserts a column, Ben inserts one too',
    ],
  },
  {
    cause: 'Both refuse a table press or a drag on a table with a pending cell change',
    owner: 'both',
    codes: 'R',
    sets: ['structures'],
    kinds: ['cellrev'],
    edits: ['table style', 'right edge'],
  },
  {
    cause: 'Both refuse a cell press on a cell with a pending change (invalidArgs)',
    owner: 'both',
    codes: 'R',
    sets: ['structures'],
    kinds: ['cellrev'],
    edits: [
      'cell fill',
      'cell vertical alignment',
      'cell borders',
      'Ana fills the cell, Ben types inside',
      'Ana fills the cell, Ben: cell borders',
      'Ana fills the cell, Ben gives it the same fill',
    ],
  },
  {
    cause:
      "The model refuses a picture in a table of contents result (not-covered, group 9.5); the engine's picture insert does not check the read-only result its typing checks",
    owner: 'model',
    codes: 'S',
    sets: ['pictures'],
    kinds: ['toc'],
    edits: all,
  },
  {
    cause:
      "The engine's picture insert keeps a block control's placeholder text and puts the picture beside it; the list replaces the prompt, as Word replaces a control's placeholder with what is inserted",
    owner: 'engine',
    codes: 'VD',
    sets: ['pictures'],
    kinds: ['blockprompt'],
    edits: ['a picture before', 'a picture after, directly'],
  },
  {
    cause:
      "The engine's picture insert keeps a block control's placeholder text (above); a later edit of the picture changes nothing a decision changes",
    owner: 'engine',
    codes: 'V',
    sets: ['pictures'],
    kinds: ['blockprompt'],
    edits: ['Ana adds a picture, Olive resizes it', 'Ana adds a picture, Ben deletes it'],
  },
  {
    cause:
      "The engine's untracked picture insert goes past a zero-width marker at the caret (a comment's end and reference, a permission end, a proofing mark, a positional tab, an object it does not count); its typed text, and the list, stop before it",
    owner: 'engine',
    codes: 'VD',
    sets: ['pictures'],
    kinds: ['ole', 'ink', 'barebr', 'ptab', 'ruby', 'comment', 'perm', 'proof', 'fmt'],
    edits: ['a picture after, directly'],
  },
  {
    cause: 'Both refuse a row or table press on a table with a pending cell change',
    owner: 'both',
    codes: 'R',
    sets: ['structures'],
    kinds: ['cellrev'],
    edits: [
      'row height',
      'table centered',
      'Ana: table centered, Ben: table style',
      'Ana sets the row height, then fills the cell',
    ],
  },
];

const matches = <T>(test: readonly T[] | ((value: T) => boolean), value: T) =>
  typeof test === 'function' ? test(value) : test.includes(value);

/** The known differences a cell of a set matches. */
export function knownFor(set: SetName, name: string): KnownDifference[] {
  const at = name.indexOf(' | ');
  const kind = name.slice(0, at);
  const edit = name.slice(at + 3);
  return KNOWN.filter(
    (entry) => entry.sets.includes(set) && matches(entry.kinds, kind) && matches(entry.edits, edit)
  );
}
