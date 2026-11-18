// Random changes of every kind, at places read from a copy's view now: typing, tabs and breaks,
// deletions and formats over ranges that cross paragraphs, Enter and joins, paragraph presses,
// table presses, rows and columns, new tables, control values, link targets, moves, notes,
// comments and replies, take-backs and, for an organiser, decisions and resolves. Places go anywhere in any
// story, so some changes are refused by name: a session counts them.

import { orderedCardSeen } from '../../change/apply-decide.ts';
import type {
  OrderedChange,
  OrderedChangeAddress,
  OrderedChangePlace,
} from '../../change/types.ts';
import { isHidden } from '../../model/document.ts';
import type { OrderedDocument, OrderedItem } from '../../model/types.ts';
import { randomPasteChange } from './random-paste.ts';

/** A change without its version and id: what a client fills in. */
export type ChangeBody = OrderedChange extends infer C
  ? C extends OrderedChange
    ? Omit<C, 'v' | 'id'>
    : never
  : never;

/** Who makes a change: a party suggests; an organiser also edits directly and decides. */
export interface Who {
  readonly author: string;
  readonly organiser: boolean;
}

interface Paragraph {
  readonly units: OrderedChangeAddress[];
  readonly mark: OrderedChangeAddress;
}

interface View {
  readonly paragraphs: Paragraph[];
  readonly units: OrderedChangeAddress[];
  readonly opens: Map<string, OrderedItem[]>;
}

const BODY = '/word/document.xml';
const WORDS = ['a', 'b', 'x', 'fee ', ' net', ' due', '€', '😀', 'Ltd'];
const PATHS = ['terms', 'annex', 'fees', 'a?b=c'];

const address = (item: OrderedItem, offset = item.offset): OrderedChangeAddress => ({
  space: item.space,
  offset,
});

/** The shown paragraphs of a story, their units, and the shown containers by token. */
function viewOf(doc: OrderedDocument, name: string): View {
  const paragraphs: Paragraph[] = [];
  const opens = new Map<string, OrderedItem[]>();
  let units: OrderedChangeAddress[] = [];
  for (const item of doc.stories.get(name)?.tree.items() ?? []) {
    if (isHidden(doc, item)) continue;
    if (item.kind === 'text')
      for (let k = 0; k < item.length; k += 1) units.push(address(item, item.offset + k));
    else if (item.kind === 'atom') units.push(address(item));
    else if (item.kind === 'mark') {
      paragraphs.push({ units, mark: address(item) });
      units = [];
    } else if (item.kind === 'open' && item.token !== undefined)
      opens.set(item.token, [...(opens.get(item.token) ?? []), item]);
  }
  return { paragraphs, units: paragraphs.flatMap((p) => p.units), opens };
}

/** A new comment thread, a reply, or (an organiser's) a resolve with or without a note. */
function randomComment(
  doc: OrderedDocument,
  who: Who,
  rand: () => number,
  part: string,
  range: () => { from: OrderedChangePlace; to: OrderedChangePlace } | null,
  word: () => string
): ChangeBody | null {
  const threads = [...doc.comments.keys()];
  const comment = threads[Math.floor(rand() * threads.length)];
  const roll = rand();
  if (comment !== undefined && roll < 0.3) {
    return { kind: 'comment-reply', comment, text: `Yes ${word()}` };
  }

  if (comment !== undefined && roll < 0.45 && who.organiser) {
    const note = rand() < 0.5 ? { note: `Done ${word()}` } : {};
    return { kind: 'comment-resolve', comment, ...note };
  }

  const r = range();
  return r ? { kind: 'comment', part, ...r, text: `Why ${word()}?\nSee ${word()}` } : null;
}

/** A random change by `who` in their own copy, or null when the copy offers no target. */
export function randomChange(
  doc: OrderedDocument,
  who: Who,
  rand: () => number,
  /** Group 9.10's changes too (paste, new lists, paragraph removal, protection). */
  paste = false
): ChangeBody | null {
  const pick = <T>(list: readonly T[]): T | undefined => list[Math.floor(rand() * list.length)];
  const stories = [...doc.stories.keys()];
  const part = rand() < 0.75 ? BODY : (pick(stories) ?? BODY);
  const view = viewOf(doc, part);
  const own = [...doc.suggestions.values()].filter(
    (s) => s.status === 'pending' && s.attribution.author === who.author
  );
  const mode = who.organiser && rand() < 0.3 ? 'edit' : 'suggest';
  const into = mode === 'suggest' && rand() < 0.1 ? pick(own)?.id : undefined;
  const edit = { part, mode, ...(into === undefined ? {} : { into }) } as const;
  const paragraph = () => pick(view.paragraphs);
  const place = (): OrderedChangePlace | undefined => {
    const p = paragraph();
    if (!p) return undefined;
    const unit = rand() < 0.2 ? undefined : pick(p.units);
    return unit ?? { mark: p.mark };
  };
  const range = (): { from: OrderedChangePlace; to: OrderedChangePlace } | null => {
    const i = Math.floor(rand() * view.units.length);
    const j = i + 1 + Math.floor(rand() * 12);
    const from = view.units[i];
    const to = view.units[j];
    return from && to ? { from, to } : null;
  };
  const word = () => pick(WORDS) ?? 'a';
  const open = (token: string) => pick(view.opens.get(token) ?? []);
  const roll = rand();
  // Group 9.10 (paste, lists, paragraph removal, protection) takes a slice of typing's share
  if (paste && roll < 0.03) {
    const mark = () => paragraph()?.mark;
    return randomPasteChange(who, rand, edit, place, mark);
  }
  if (roll < 0.24) {
    const at = place();
    return at ? { kind: 'insert', ...edit, place: at, text: word() } : null;
  }
  if (roll < 0.28) {
    const at = place();
    if (!at) return null;
    const kind = pick(['tab', 'line', 'page', 'column'] as const) ?? 'tab';
    return kind === 'tab'
      ? { kind: 'insert-tab', ...edit, place: at }
      : { kind: 'insert-break', ...edit, place: at, break: kind };
  }
  if (roll < 0.4) {
    const r = range();
    return r ? { kind: 'delete', ...edit, ...r } : null;
  }
  if (roll < 0.47) {
    const r = range();
    const properties = pick([
      [{ localName: 'b' }],
      [{ localName: 'i' }],
      [{ localName: 'u', attributes: { val: 'single' } }],
      [{ localName: 'color', attributes: { val: 'C00000' } }],
      [{ localName: 'b', attributes: { val: '0' } }],
    ]);
    return r && properties ? { kind: 'format', ...edit, ...r, properties } : null;
  }
  if (roll < 0.53) {
    const at = place();
    return at ? { kind: 'split', ...edit, place: at } : null;
  }
  if (roll < 0.57) {
    const p = pick(view.paragraphs.slice(0, -1));
    return p ? { kind: 'join', ...edit, mark: p.mark } : null;
  }
  if (roll < 0.62) {
    const marks = [paragraph(), paragraph()].flatMap((p) => (p ? [p.mark] : []));
    const unique = [...new Map(marks.map((m) => [`${m.space}|${m.offset}`, m])).values()];
    const properties = pick([
      [{ localName: 'jc', attributes: { val: 'center' } }],
      [{ localName: 'jc', attributes: { val: 'right' } }],
      [{ localName: 'spacing', attributes: { after: '240' } }],
      [{ localName: 'ind', attributes: { left: '720' } }],
      [{ localName: 'keepNext' }],
      [
        {
          localName: 'numPr',
          children: [
            { localName: 'ilvl', attributes: { val: '0' } },
            { localName: 'numId', attributes: { val: '1' } },
          ],
        },
      ],
      [
        {
          localName: 'tabs',
          children: [{ localName: 'tab', attributes: { val: 'right', pos: '9000' } }],
        },
      ],
    ]);
    return unique.length > 0 && properties
      ? { kind: 'paragraph-property', ...edit, marks: unique, properties }
      : null;
  }
  if (roll < 0.65) {
    const cell = open('tc');
    const row = open('tr');
    const table = open('tbl');
    const press = pick([
      cell && {
        on: [address(cell)],
        op: { op: 'setTableCellFill', color: { kind: 'hex', value: 'FFCC00' } },
      },
      cell && {
        on: [address(cell)],
        op: { op: 'setTableCellVerticalAlignment', alignment: 'center' },
      },
      row && { on: [address(row)], op: { op: 'setTableRowHeight', heightTwips: 600 } },
      table && { on: [address(table)], op: { op: 'setTableProperties', justification: 'center' } },
      table && {
        on: [address(table)],
        op: { op: 'setTableProperties', styleId: 'TableGrid', headerRowCount: 1 },
      },
      // Drags: the edge on the cell's right (it fits two columns of 3000), the right edge
      cell && {
        on: [address(cell)],
        op: { op: 'setTableColumnWidths', leftWidthTwips: 3600, rightWidthTwips: 2400 },
      },
      cell && {
        on: [address(cell)],
        op: { op: 'setTableRightEdgeWidth', columnWidthTwips: 2600, tableWidthTwips: 5600 },
      },
    ] as const);
    return press ? ({ kind: 'table-property', ...edit, ...press } as ChangeBody) : null;
  }
  if (roll < 0.69) {
    const row = open('tr');
    const cell = open('tc');
    const which = rand();
    if (which < 0.25)
      return row ? { kind: 'insert-row', ...edit, row: address(row), where: 'below' } : null;
    if (which < 0.5) return row ? { kind: 'delete-row', ...edit, row: address(row) } : null;
    if (which < 0.75)
      return cell ? { kind: 'insert-column', ...edit, cell: address(cell), where: 'right' } : null;
    return cell ? { kind: 'delete-column', ...edit, cell: address(cell) } : null;
  }
  if (roll < 0.7) {
    const p = paragraph();
    const at = p ? (p.units[0] ?? { mark: p.mark }) : undefined;
    return at
      ? { kind: 'insert-table', ...edit, place: at, rows: 2, columns: 2, columnWidth: 2400 }
      : null;
  }
  if (roll < 0.73) {
    const control = open('sdt') ?? open('sdtBlock');
    return control
      ? {
          kind: 'control-value',
          ...edit,
          control: address(control),
          value: { kind: 'text', text: `Party ${word()}` },
        }
      : null;
  }
  if (roll < 0.75) {
    const link = open('hyperlink');
    if (!link) return null;
    return rand() < 0.7
      ? {
          kind: 'link-target',
          ...edit,
          link: address(link),
          href: `https://example.com/${pick(PATHS) ?? 'x'}`,
        }
      : { kind: 'link-target', ...edit, link: address(link), anchor: 'Terms' };
  }
  if (roll < 0.78) {
    const p = pick(view.paragraphs.filter((q) => q.units.length > 3));
    const at = place();
    if (!p || !at) return null;
    const i = Math.floor(rand() * (p.units.length - 2));
    const from = p.units[i];
    const to = p.units[i + 2];
    return from && to ? { kind: 'move', ...edit, from, to, place: at } : null;
  }
  if (roll < 0.8) {
    const at = place();
    return at && part === BODY
      ? {
          kind: 'insert-note',
          ...edit,
          place: at,
          note: rand() < 0.5 ? 'footnote' : 'endnote',
          text: ` Note ${word()}`,
        }
      : null;
  }
  if (roll < 0.83) {
    return randomComment(doc, who, rand, part, range, word);
  }
  if (roll < 0.9) {
    const suggestion = pick(own);
    return suggestion ? { kind: 'retract', suggestions: [suggestion.id] } : null;
  }
  if (!who.organiser) return null;
  const pending = [...doc.suggestions.values()].filter((s) => s.status === 'pending');
  const chosen = pick(pending);
  const seen = chosen ? orderedCardSeen(doc, chosen.id) : null;
  return seen
    ? {
        kind: 'decide',
        cards: [{ suggestions: seen }],
        status: rand() < 0.5 ? 'accepted' : 'rejected',
      }
    : null;
}
