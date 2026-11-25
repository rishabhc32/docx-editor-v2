// Random changes that reach the shapes a fold once lost or changed: typing, formats and values
// over controls' prompts beside a stored value, a control wrapped round words and its value given
// while it is pending, moves of words that hold a note reference, section breaks and page setups
// beside another person's section records, row heights and columns in a table with notes, links
// one person makes and another types in, and decisions. Half the changes are `randomChange`'s.

import type {
  OrderedChangeAddress,
  OrderedChangePlace,
  OrderedControlValue,
} from '../../change/types.ts';
import { isHidden } from '../../model/document.ts';
import type { OrderedDocument, OrderedItem } from '../../model/types.ts';
import { asWml, childNamed, isWml } from '../../ooxml-nodes.ts';
import { randomChange, viewOf, type ChangeBody, type Who } from './random-change.ts';

const BODY = '/word/document.xml';
const WORDS = ['a', 'net ', ' due', 'Ltd', '€'];
const NOTE_REFERENCES = ['footnoteReference', 'endnoteReference'];
const CONTROL_TYPES = ['date', 'comboBox', 'dropDownList', 'plainText', 'richText'] as const;

/** The places one shape needs, read from a copy's body now. */
interface Shapes {
  readonly doc: OrderedDocument;
  readonly rand: () => number;
  readonly edit: { readonly part: string; readonly mode: 'suggest' | 'edit' };
  readonly view: ReturnType<typeof viewOf>;
}

const address = (item: OrderedItem, offset = item.offset): OrderedChangeAddress => ({
  space: item.space,
  offset,
});

const pickOf =
  (rand: () => number) =>
  <T>(list: readonly T[]): T | undefined =>
    list[Math.floor(rand() * list.length)];

/** The shown units between a container's open item and its close item. */
function unitsInside(doc: OrderedDocument, open: OrderedItem): OrderedChangeAddress[] {
  const units: OrderedChangeAddress[] = [];
  let inside = false;
  for (const item of doc.stories.get(BODY)?.tree.items() ?? []) {
    if (item === open) {
      inside = true;
    } else if (inside && item.kind === 'close' && item.pair === open.pair) {
      return units;
    } else if (inside && !isHidden(doc, item) && (item.kind === 'text' || item.kind === 'atom')) {
      for (let k = 0; k < item.length; k += 1) {
        units.push(address(item, item.offset + k));
      }
    }
  }

  return units;
}

/** A control's type, from its properties (`w:sdtPr`): a date, a list, words, or rich content. */
function controlType(open: OrderedItem): string {
  const shell = open.node === null ? null : asWml(open.node, 'sdt');
  const props = shell ? childNamed(shell, 'sdtPr') : undefined;
  const type = ['date', 'comboBox', 'dropDownList', 'text'].find(
    (name) => props && childNamed(props, name)
  );
  return type ?? 'richText';
}

/** Typing into a control's prompt, or into a link's words at their start or end. */
function typingInside(s: Shapes, token: 'sdt' | 'hyperlink'): ChangeBody | null {
  const pick = pickOf(s.rand);
  const open = pick(s.view.opens.get(token) ?? []);
  const units = open ? unitsInside(s.doc, open) : [];
  const place = s.rand() < 0.5 ? units[0] : units.at(-1);
  return place ? { kind: 'insert', ...s.edit, place, text: pick(WORDS) ?? 'a' } : null;
}

/** A format over the whole of a control's prompt or value. */
function formatControl(s: Shapes): ChangeBody | null {
  const open = pickOf(s.rand)(s.view.opens.get('sdt') ?? []);
  const units = open ? unitsInside(s.doc, open) : [];
  const from = units[0];
  const to = units.at(-1);
  const properties = [{ localName: s.rand() < 0.5 ? 'b' : 'i' }];
  return from && to ? { kind: 'format', ...s.edit, from, to, properties } : null;
}

/** A value for a control, set directly or suggested. */
function controlValue(s: Shapes): ChangeBody | null {
  const pick = pickOf(s.rand);
  const open = pick(s.view.opens.get('sdt') ?? []);
  if (!open) {
    return null;
  }

  return { kind: 'control-value', ...s.edit, control: address(open), value: valueOf(s, open) };
}

/** A value of a control's own type: a date, an item, or words. */
function valueOf(s: Shapes, open: OrderedItem): OrderedControlValue {
  const pick = pickOf(s.rand);
  const type = controlType(open);
  if (type === 'date') {
    const day = String(1 + Math.floor(s.rand() * 28)).padStart(2, '0');
    return { kind: 'date', iso: `2026-09-${day}` };
  }

  if (type === 'dropDownList' || (type === 'comboBox' && s.rand() < 0.6)) {
    return { kind: 'listItem', value: pick(['30', '60', 'EUR', 'USD']) ?? '30' };
  }

  return { kind: 'text', text: `Acme ${pick(WORDS) ?? 'a'}` };
}

/** Three units of one shown paragraph, the first of them at `from`. */
function threeUnits(s: Shapes): { from: OrderedChangePlace; to: OrderedChangePlace } | null {
  const paragraph = pickOf(s.rand)(s.view.paragraphs.filter((q) => q.units.length > 3));
  if (!paragraph) {
    return null;
  }

  const i = Math.floor(s.rand() * (paragraph.units.length - 2));
  const from = paragraph.units[i];
  const to = paragraph.units[i + 2];
  return from && to ? { from, to } : null;
}

/** A new control round words of one paragraph. */
function wrapControl(s: Shapes): ChangeBody | null {
  const range = threeUnits(s);
  const type = pickOf(s.rand)(CONTROL_TYPES) ?? 'date';
  return range ? { kind: 'wrap-control', ...s.edit, ...range, control: { type } } : null;
}

/** A new link round words of one paragraph. */
function insertLink(s: Shapes): ChangeBody | null {
  const range = threeUnits(s);
  return range
    ? { kind: 'insert-link', ...s.edit, ...range, href: 'https://example.com/fees' }
    : null;
}

/** A move of the words round a note reference to a random place. */
function moveNote(s: Shapes): ChangeBody | null {
  const pick = pickOf(s.rand);
  const units = s.view.units;
  const references = units.flatMap((unit, k) => (isNoteReference(s.doc, unit) ? [k] : []));
  const k = pick(references) ?? -1;
  const from = units[k - 1];
  const to = units[k + 1];
  const place = pick(units);
  return from && to && place ? { kind: 'move', ...s.edit, from, to, place } : null;
}

/** Whether a unit is a footnote or an endnote reference. */
function isNoteReference(doc: OrderedDocument, unit: OrderedChangeAddress): boolean {
  const items = doc.stories.get(BODY)?.spaces.get(unit.space) ?? [];
  const node = items.find((item) => item.offset === unit.offset)?.node;
  return node != null && NOTE_REFERENCES.some((name) => isWml(node, name));
}

/** A section break at a paragraph's end, or a page setup of the section a paragraph is in. */
function section(s: Shapes): ChangeBody | null {
  const pick = pickOf(s.rand);
  const mark = pick(s.view.paragraphs)?.mark;
  if (!mark) {
    return null;
  }

  if (s.rand() < 0.4) {
    const type = pick(['nextPage', 'continuous'] as const) ?? 'nextPage';
    return { kind: 'section-break', ...s.edit, mark, type };
  }

  const page = pick([
    { orientation: 'landscape' },
    { orientation: 'portrait' },
    { marginLeftTwips: 1800 },
    { marginTopTwips: 720, marginBottomTwips: 720 },
  ] as const);
  return { kind: 'section-property', ...s.edit, sections: [], page, pageAt: mark };
}

/** A row height, a column width, or a column or row added or taken out, in a table. */
function tablePart(s: Shapes): ChangeBody | null {
  const pick = pickOf(s.rand);
  const cell = pick(s.view.opens.get('tc') ?? []);
  const row = pick(s.view.opens.get('tr') ?? []);
  const roll = s.rand();
  if (!cell || !row) {
    return null;
  }

  if (roll < 0.3) {
    const op = { op: 'setTableRowHeight', heightTwips: 400 + Math.floor(s.rand() * 4) * 100 };
    return { kind: 'table-property', ...s.edit, on: [address(row)], op } as ChangeBody;
  }

  if (roll < 0.5) {
    const op = { op: 'setTableColumnWidths', leftWidthTwips: 3600, rightWidthTwips: 2400 };
    return { kind: 'table-property', ...s.edit, on: [address(cell)], op } as ChangeBody;
  }

  if (roll < 0.7) {
    return { kind: 'insert-column', ...s.edit, cell: address(cell), where: 'right' };
  }

  if (roll < 0.85) {
    return { kind: 'delete-column', ...s.edit, cell: address(cell) };
  }

  return { kind: 'insert-row', ...s.edit, row: address(row), where: 'below' };
}

/** The shapes, by name. */
const SHAPES: readonly (readonly [string, (s: Shapes) => ChangeBody | null])[] = [
  ['typing in a control', (s) => typingInside(s, 'sdt')],
  ['typing in a link', (s) => typingInside(s, 'hyperlink')],
  ['a format on a control', formatControl],
  ['a control value', controlValue],
  ['a new control round words', wrapControl],
  ['a new link', insertLink],
  ['a move round a note reference', moveNote],
  ['a section break or page setup', section],
  ['a table row or column', tablePart],
];

/** The names of the shapes, and of the other changes (`any`). */
export const SHAPE_NAMES = [...SHAPES.map(([name]) => name), 'any'];

/** A random change by `who` in their copy, and its shape's name; null when none fits. */
export function randomShapeChange(
  doc: OrderedDocument,
  who: Who,
  rand: () => number
): { readonly shape: string; readonly body: ChangeBody } | null {
  if (rand() < 0.5) {
    const body = randomChange(doc, who, rand);
    return body === null ? null : { shape: 'any', body };
  }

  const mode = who.organiser && rand() < 0.4 ? 'edit' : 'suggest';
  const [shape, make] = pickOf(rand)(SHAPES) ?? SHAPES[0]!;
  const body = make({ doc, rand, edit: { part: BODY, mode }, view: viewOf(doc, BODY) });
  return body === null ? null : { shape, body };
}
