// The reader for a table property change's op: the engine's own table op without its table,
// targets and revision (`OrderedTableOp`). It checks the op's shape and caps; the engine checks
// what the values mean when the change applies, and a value it refuses refuses the change.

import { TABLE_BORDER_STYLES } from '../../../store/table-border-style.ts';
import type {
  TableBorderEdgeTarget,
  TableBorderSpecInput,
  TreeDocColorValue,
} from '../../../store/store/tree-op-types.ts';
import type { OrderedTableOp } from '../edits/tables.ts';
import {
  at,
  fieldsOf,
  int,
  Invalid,
  list,
  oneOf,
  plainText,
  required,
  str,
  variantFields,
  type Fields,
} from './read.ts';
import { ORDERED_CHANGE_LIMITS as LIMITS } from './types.ts';

const HEX = /^[0-9A-Fa-f]{6}$/;
const THEME_SLOT = /^[A-Za-z][A-Za-z0-9]{0,31}$/;
const EDGES: readonly TableBorderEdgeTarget[] = [
  'all',
  'outside',
  'inside',
  'top',
  'bottom',
  'left',
  'right',
];

/** The widest border Word draws: 12 points, in eighths of a point. */
const MAX_BORDER_SIZE = 96;

/** The engine's cap on a length in twips (22 inches): a width, a row height. */
const MAX_TWIPS = LIMITS.columnWidth;

function hex(value: unknown, path: string): string {
  const text = str(value, path, 6);
  if (!HEX.test(text)) throw new Invalid('invalid-value', path, 'six hex digits');
  return text;
}

/** Each kind of color, and its own fields. */
const COLOR_FIELDS = {
  hex: ['value'],
  theme: ['slot', 'resolvedHex', 'tint', 'shade'],
  auto: ['resolvedHex'],
} as const;

function color(value: unknown, path: string): TreeDocColorValue {
  const { name: kind, fields: f } = variantFields(value, path, 'kind', COLOR_FIELDS);
  if (kind === 'hex') {
    return { kind, value: hex(required(f, 'value'), at(path, 'value')) };
  }

  if (kind === 'auto') {
    return { kind, resolvedHex: hex(required(f, 'resolvedHex'), at(path, 'resolvedHex')) };
  }

  const slot = str(required(f, 'slot'), at(path, 'slot'), 32);
  if (!THEME_SLOT.test(slot)) throw new Invalid('invalid-value', at(path, 'slot'));
  return {
    kind,
    slot,
    resolvedHex: hex(required(f, 'resolvedHex'), at(path, 'resolvedHex')),
    ...(f.has('tint') ? { tint: int(f.get('tint'), at(path, 'tint'), 0, 255) } : {}),
    ...(f.has('shade') ? { shade: int(f.get('shade'), at(path, 'shade'), 0, 255) } : {}),
  };
}

function border(value: unknown, path: string): TableBorderSpecInput {
  const f = fieldsOf(value, path, new Set(['style', 'size', 'color']));
  return {
    style: oneOf(required(f, 'style'), at(path, 'style'), TABLE_BORDER_STYLES),
    size: int(required(f, 'size'), at(path, 'size'), 1, MAX_BORDER_SIZE),
    color: color(required(f, 'color'), at(path, 'color')),
  };
}

const TABLE_FIELDS = [
  'styleId',
  'headerRowCount',
  'columnWidthsTwips',
  'justification',
  'preferredWidth',
] as const;

function tableProperties(f: Fields, path: string): OrderedTableOp {
  if (!TABLE_FIELDS.some((key) => f.has(key))) {
    throw new Invalid('empty', path, 'no property to set');
  }

  const op: {
    op: 'setTableProperties';
    styleId?: string;
    headerRowCount?: number;
    columnWidthsTwips?: number[];
    justification?: 'left' | 'center' | 'right' | null;
    preferredWidth?: { type: 'auto' | 'dxa' | 'pct'; value: number } | null;
  } = { op: 'setTableProperties' };
  if (f.has('styleId')) {
    op.styleId = plainText(f.get('styleId'), at(path, 'styleId'), 1, LIMITS.attributeValue);
  }

  if (f.has('headerRowCount')) {
    op.headerRowCount = int(
      f.get('headerRowCount'),
      at(path, 'headerRowCount'),
      0,
      LIMITS.tableRows
    );
  }

  if (f.has('columnWidthsTwips')) {
    op.columnWidthsTwips = list(
      f.get('columnWidthsTwips'),
      at(path, 'columnWidthsTwips'),
      1,
      LIMITS.tableColumns
    ).map((width, k) => int(width, at(at(path, 'columnWidthsTwips'), k), 1, LIMITS.columnWidth));
  }

  if (f.has('justification')) {
    const raw = f.get('justification');
    op.justification =
      raw === null
        ? null
        : oneOf(raw, at(path, 'justification'), ['left', 'center', 'right'] as const);
  }

  if (f.has('preferredWidth')) {
    const raw = f.get('preferredWidth');
    if (raw === null) {
      op.preferredWidth = null;
    } else {
      const where = at(path, 'preferredWidth');
      const w = fieldsOf(raw, where, new Set(['type', 'value']));
      op.preferredWidth = {
        type: oneOf(required(w, 'type'), at(where, 'type'), ['auto', 'dxa', 'pct'] as const),
        // Every type is capped at 22 inches in twips; the engine checks what a `pct` value means
        value: int(required(w, 'value'), at(where, 'value'), 0, MAX_TWIPS),
      };
    }
  }

  return op;
}

/** Each table op, and its own fields. */
const OP_FIELDS = {
  setTableProperties: TABLE_FIELDS,
  setTableRowHeight: ['heightTwips'],
  setTableCellFill: ['color'],
  setTableCellVerticalAlignment: ['alignment'],
  setTableCellBorders: ['scope', 'target', 'spec'],
  setTableColumnWidths: ['leftWidthTwips', 'rightWidthTwips'],
  setTableRightEdgeWidth: ['columnWidthTwips', 'tableWidthTwips'],
} as const;

/** A width of a column or a table in twips: a whole number up to 22 inches. */
const width = (f: Fields, key: string, path: string): number =>
  int(required(f, key), at(path, key), 1, MAX_TWIPS);

/**
 * Check a table op: its name, then exactly its fields.
 *
 * @throws Invalid on a shape the format does not allow.
 */
export function tableOp(value: unknown, path: string): OrderedTableOp {
  const { name, fields: f } = variantFields(value, path, 'op', OP_FIELDS);
  switch (name) {
    case 'setTableProperties':
      return tableProperties(f, path);
    case 'setTableRowHeight':
      return {
        op: name,
        heightTwips: int(required(f, 'heightTwips'), at(path, 'heightTwips'), 1, MAX_TWIPS),
      };
    case 'setTableCellFill': {
      const raw = required(f, 'color');
      return { op: name, color: raw === null ? null : color(raw, at(path, 'color')) };
    }
    case 'setTableCellVerticalAlignment':
      return {
        op: name,
        alignment: oneOf(required(f, 'alignment'), at(path, 'alignment'), [
          'top',
          'center',
          'bottom',
        ] as const),
      };
    case 'setTableCellBorders': {
      const scope = oneOf(required(f, 'scope'), at(path, 'scope'), ['none', ...EDGES] as const);
      if (scope === 'none') {
        if (f.has('spec')) throw new Invalid('unknown-field', at(path, 'spec'));
        return {
          op: name,
          scope,
          target: oneOf(required(f, 'target'), at(path, 'target'), EDGES),
        };
      }

      if (f.has('target')) throw new Invalid('unknown-field', at(path, 'target'));
      return { op: name, scope, spec: border(required(f, 'spec'), at(path, 'spec')) };
    }
    case 'setTableColumnWidths':
      return {
        op: name,
        leftWidthTwips: width(f, 'leftWidthTwips', path),
        rightWidthTwips: width(f, 'rightWidthTwips', path),
      };
    case 'setTableRightEdgeWidth':
      return {
        op: name,
        columnWidthTwips: width(f, 'columnWidthTwips', path),
        tableWidthTwips: width(f, 'tableWidthTwips', path),
      };
  }
}

/** The container token a table op targets. */
export function tableOpToken(op: OrderedTableOp): 'tbl' | 'tr' | 'tc' {
  if (op.op === 'setTableProperties') {
    return 'tbl';
  }

  if (op.op === 'setTableRowHeight') {
    return 'tr';
  }

  return 'tc';
}
