// The records a checkpoint's manifest keeps that were made over a pending record with no change
// record of their own (`over`), by the change record they were made over (design D14).

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import type { OoxmlProperty } from '../../../store/store/tree-op-types.ts';
import { ALIGNMENTS, NUM_ID, TAB_LEADERS } from '../change/apply-properties.ts';
import { readFurniture, readPage, noteWrites } from '../change/kinds/sections/read.ts';
import { isOrderedPropertyAllowed, type OrderedPropertyScope } from '../change/properties.ts';
import {
  at,
  bool,
  int,
  Invalid,
  oneOf,
  plainFields,
  required,
  type Fields,
} from '../change/read.ts';
import { ORDERED_CHANGE_LIMITS as LIMITS } from '../change/types.ts';
import { properties } from '../change/values.ts';
import { manifestList, manifestString, readNamedOnce, xmlText } from './checkpoint-suggestions.ts';
import { propertiesXml, readPropertiesXml } from './checkpoint-takeovers.ts';
import type { OrderedSectionOp } from './sections.ts';
import type { OrderedFormatWrite, OrderedParagraphOp, OrderedStyleWrite } from './types.ts';

/** A format's write as its record gives it: the properties it sets and takes off, and its style. */
export interface CheckpointFormatWrite extends OrderedFormatWrite {
  readonly style?: OrderedStyleWrite;
}

/**
 * A record's write, by the record's kind: a format's (under `w:rPr`), a paragraph press (under
 * `w:pPr`) or a section's write (under `w:sectPr`).
 */
export type CheckpointOverWrite =
  | { readonly format: CheckpointFormatWrite }
  | { readonly paragraphOp: OrderedParagraphOp }
  | { readonly sectionOp: OrderedSectionOp };

/** A record made over the change record's pending records, with its write. */
export interface CheckpointOverRecord {
  /** The suggestion, by its place in the manifest's list. */
  readonly suggestion: number;
  /** The pending records it was made over: their suggestions, by their places. */
  readonly over: readonly number[];
  /**
   * A format on its author's own pending words, or a press on their own new paragraph, shown with
   * no change record.
   */
  readonly untracked?: boolean;
  readonly write: CheckpointOverWrite;
}

/**
 * The records made over the pending records of one change record, by the label on it: the
 * properties those pending records give without the records made over them (the properties
 * element the change record names), and the records, in the order they apply.
 */
export interface CheckpointMadeOver {
  readonly label: number;
  readonly properties: OoxmlElement;
  readonly records: readonly CheckpointOverRecord[];
}

/** The write of a record under each properties element, and how it reads. */
const WRITES = {
  rPr: ['format', (value: unknown, path: string) => ({ format: readFormatWrite(value, path) })],
  pPr: [
    'paragraphOp',
    (value: unknown, path: string) => ({ paragraphOp: readParagraphOp(value, path) }),
  ],
  sectPr: [
    'sectionOp',
    (value: unknown, path: string) => ({ sectionOp: readSectionOp(value, path) }),
  ],
} as const;

/** The made-over records as the manifest lists them: JSON leaves out a field with no value. */
export function madeOverJson(madeOver: readonly CheckpointMadeOver[]): unknown[] {
  return madeOver.map(({ properties, records, ...entry }) => ({
    ...entry,
    properties: propertiesXml(properties),
    records: records.map(({ write, ...record }) => ({ ...record, ...writeJson(write) })),
  }));
}

/** A record's write as the manifest lists it: a section mark's properties as XML. */
function writeJson(write: CheckpointOverWrite): CheckpointOverWrite | object {
  if (!('sectionOp' in write) || write.sectionOp.op !== 'mark') {
    return write;
  }

  return { sectionOp: { ...write.sectionOp, sectPr: propertiesXml(write.sectionOp.sectPr) } };
}

/**
 * The listed made-over records, each checked: every value comes from storage. Each names a
 * suggestion of the manifest's list of `suggestions`, and a change record once; each record's
 * write is one of the properties the entry gives.
 *
 * @throws Invalid when one does not read.
 */
export function readMadeOver(value: unknown, suggestions: number): CheckpointMadeOver[] {
  return readNamedOnce(
    value,
    'madeOver',
    (entry, path) => readEntry(entry, path, suggestions),
    ({ label }) => label
  );
}

function readEntry(value: unknown, path: string, suggestions: number): CheckpointMadeOver {
  const f = plainFields(value, path);
  const properties = readPropertiesXml(
    required(f, 'properties'),
    at(path, 'properties'),
    new Set(Object.keys(WRITES))
  );
  const write = WRITES[properties.localName as keyof typeof WRITES];
  const records = manifestList(required(f, 'records'), at(path, 'records'), 1);
  return {
    label: int(required(f, 'label'), at(path, 'label')),
    properties,
    records: records.map((record, k) =>
      readRecord(record, at(at(path, 'records'), k), suggestions, write)
    ),
  };
}

function readRecord(
  value: unknown,
  path: string,
  suggestions: number,
  [key, readWrite]: (typeof WRITES)[keyof typeof WRITES]
): CheckpointOverRecord {
  const f = plainFields(value, path);
  const over = manifestList(required(f, 'over'), at(path, 'over'), 1);
  return {
    suggestion: int(required(f, 'suggestion'), at(path, 'suggestion'), 0, suggestions - 1),
    over: over.map((k, j) => int(k, at(at(path, 'over'), j), 0, suggestions - 1)),
    untracked: f.has('untracked') ? bool(f.get('untracked'), at(path, 'untracked')) : undefined,
    write: readWrite(required(f, key), at(path, key)),
  };
}

/** Properties a scope allows, flat as a record holds them; none for an empty list. */
function recordProperties(
  value: unknown,
  path: string,
  scope: OrderedPropertyScope
): OoxmlProperty[] {
  if (manifestList(value, path).length === 0) {
    return [];
  }

  return properties(value, path, scope).map(({ localName, attributes, children }, k) => {
    if (children !== undefined) {
      throw new Invalid('unknown-field', at(at(path, k), 'children'));
    }

    return attributes === undefined ? { localName } : { localName, attributes };
  });
}

/** The names of the properties a write takes off, each allowed in the scope and named once. */
function removed(f: Fields, scope: OrderedPropertyScope): string[] | undefined {
  if (!f.has('remove')) {
    return undefined;
  }

  const path = at(f.path, 'remove');
  const names = manifestList(f.get('remove'), path, 1).map((name, k) => {
    const read = manifestString(name, at(path, k));
    if (!isOrderedPropertyAllowed(scope, read)) {
      throw new Invalid('unknown-property', at(path, k));
    }

    return read;
  });
  if (new Set(names).size !== names.length) {
    throw new Invalid('repeated-property', path);
  }

  return names;
}

function readFormatWrite(value: unknown, path: string): CheckpointFormatWrite {
  const f = plainFields(value, path);
  const remove = removed(f, 'run');
  return {
    properties: f.has('properties')
      ? recordProperties(f.get('properties'), at(path, 'properties'), 'run')
      : [],
    ...(remove === undefined ? {} : { remove }),
    ...(f.has('style') ? { style: readStyle(f.get('style'), at(path, 'style')) } : {}),
  };
}

/** A character style a format sets or takes off, by its id. */
function readStyle(value: unknown, path: string): OrderedStyleWrite {
  const f = plainFields(value, path);
  if (f.has('set') === f.has('remove')) {
    throw new Invalid('invalid-value', path, 'one of set and remove');
  }

  const key = f.has('set') ? 'set' : 'remove';
  const id = xmlText(manifestString(f.get(key), at(path, key)), at(path, key));
  return key === 'set' ? { set: id } : { remove: id };
}

/** A list level, 0 to 8. */
const level = (value: unknown, path: string): number => int(value, path, 0, 8);

/** A position in twips, as far as a column's width goes. */
const twips = (value: unknown, path: string): number =>
  int(value, path, -LIMITS.columnWidth, LIMITS.columnWidth);

function readParagraphOp(value: unknown, path: string): OrderedParagraphOp {
  const f = plainFields(value, path);
  const op = oneOf(required(f, 'op'), at(path, 'op'), [
    'setParagraphProperties',
    'setListLevel',
    'setListNumbering',
    'setParagraphTabStops',
  ] as const);
  switch (op) {
    case 'setParagraphProperties': {
      const remove = removed(f, 'paragraph');
      const set = recordProperties(required(f, 'properties'), at(path, 'properties'), 'paragraph');
      return { op, properties: set, ...(remove === undefined ? {} : { remove }) };
    }
    case 'setListLevel':
      return { op, level: level(required(f, 'level'), at(path, 'level')) };
    case 'setListNumbering': {
      const numId = required(f, 'numId');
      if (numId !== null && (typeof numId !== 'string' || !NUM_ID.test(numId))) {
        throw new Invalid('invalid-value', at(path, 'numId'), 'not a list number');
      }

      return {
        op,
        numId,
        ...(f.has('level') ? { level: level(f.get('level'), at(path, 'level')) } : {}),
      };
    }
    default:
      return readTabStops(f);
  }
}

function readTabStops(f: Fields): OrderedParagraphOp {
  const { path } = f;
  const alignments = [...new Set(Object.values(ALIGNMENTS))];
  const stops = manifestList(required(f, 'stops'), at(path, 'stops')).map((stop, k) => {
    const where = at(at(path, 'stops'), k);
    const s = plainFields(stop, where);
    const leader = s.has('leader')
      ? oneOf(s.get('leader'), at(where, 'leader'), TAB_LEADERS)
      : undefined;
    return {
      positionTwips: twips(required(s, 'positionTwips'), at(where, 'positionTwips')),
      alignment: oneOf(required(s, 'alignment'), at(where, 'alignment'), alignments),
      ...(leader === undefined ? {} : { leader }),
    };
  });
  const inForce = f.has('inForcePositionsTwips')
    ? manifestList(f.get('inForcePositionsTwips'), at(path, 'inForcePositionsTwips'))
    : undefined;
  return {
    op: 'setParagraphTabStops',
    stops,
    ...(inForce === undefined
      ? {}
      : {
          inForcePositionsTwips: inForce.map((n, k) =>
            twips(n, at(at(path, 'inForcePositionsTwips'), k))
          ),
        }),
  };
}

/** A section's write; a reference write names no pending record, so none is made over one. */
function readSectionOp(value: unknown, path: string): OrderedSectionOp {
  const f = plainFields(value, path);
  const op = oneOf(required(f, 'op'), at(path, 'op'), [
    'page',
    'furniture',
    'type',
    'mark',
    'setNoteProperties',
  ] as const);

  // The op's own fields, read as the change that makes it reads them
  const { op: _op, ...fields } = value as Record<string, unknown>;
  switch (op) {
    case 'page':
      return { op, ...readPage(fields, path) };
    case 'furniture':
      return { op, ...readFurniture(fields, path) };
    case 'type':
      return {
        op,
        type: oneOf(required(f, 'type'), at(path, 'type'), ['nextPage', 'continuous'] as const),
      };
    case 'mark':
      return {
        op,
        sectPr: readPropertiesXml(required(f, 'sectPr'), at(path, 'sectPr'), new Set(['sectPr'])),
      };
    default:
      return { op, ...noteWrites(f) };
  }
}
