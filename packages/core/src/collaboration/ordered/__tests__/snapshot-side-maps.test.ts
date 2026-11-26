// Every map a module keeps by document, and what a checkpoint does with it.
//
// The field table (`model/snapshot-fields.ts`) covers each field of `OrderedDocument` through
// the type check. State a module keeps beside the document, in a map by document, has no field,
// so the type check does not see it. This guard lists each such map with its class, and a new
// map fails until it has one. It reads each file whole, so a declaration over several lines counts,
// and every other place that names such a map type is listed too:
//
//   saved:     the checkpoint saves it and gives it back (`SIDE` in `model/snapshot-side.ts`).
//   base:      the open makes it again from the base package (`recordLoadedPackage`).
//   index:     `indexOrderedStory` makes it again from the items.
//   cache:     made again on the first read after an open.
//   transient: empty again before a close can see it.

import { describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { SIDE } from '../model/snapshot-side.ts';

type DocumentStateClass = 'saved' | 'base' | 'index' | 'cache' | 'transient';

interface DocumentState {
  readonly kind: DocumentStateClass;
  /** For a saved map: its name in `SIDE`. */
  readonly side?: string;
}

const saved = (side: string): DocumentState => ({ kind: 'saved', side });

/** Each map by document, as `<file under ordered/> <name>`. */
const DOCUMENT_STATE: Readonly<Record<string, DocumentState>> = {
  'change/settle-emptied.ts promptedByEdit': saved('emptiedControls'),
  'change/settle-named-prompts.ts states': saved('namedPrompts'),
  'change/settle-emptied.ts splitParts': saved('splitParts'),
  'model/prompt-marks.ts droppedOf': saved('droppedMarks'),
  'model/prompt-marks.ts filledOf': saved('filledPrompts'),
  'model/prompt-marks.ts firstFillOf': saved('firstFills'),
  'model/text-boxes.ts INDEXES': saved('textBoxes'),
  'model/typers.ts LOADED': saved('typers'),
  'load/load.ts LOADED': { kind: 'base' },
  'load/sections.ts FURNITURE': { kind: 'base' },
  'model/prompt-marks.ts marksOf': { kind: 'index' },
  'model/prompt-marks.ts controlsOf': { kind: 'index' },
  'model/prompt-marks.ts promptlessOf': { kind: 'index' },
  'model/prompt-marks.ts promptedOf': { kind: 'index' },
  'client/pictures.ts DRAWING_INDEX': { kind: 'cache' },
  'compose/objects.ts COMPOSED': { kind: 'cache' },
  'model/undo-record.ts STATES': { kind: 'cache' },
  'model/prompt-marks.ts changedOf': { kind: 'transient' },
};

/** Each other place that names a map or set type keyed by document, by file, and why. */
const OTHER_USES: Readonly<Record<string, number>> = {
  // `addFor`'s parameter: one of the index maps above
  'model/prompt-marks.ts': 1,
};

const ORDERED = join(import.meta.dir, '..');

/** A map or set type whose key is a document. */
const DOCUMENT_KEYED = /\b(?:Weak)?(?:Map|Set)<\s*OrderedDocument\b/g;

/** A module-level declaration, over as many lines as it takes, whose type is such a map or set. */
const DOCUMENT_MAP =
  /^(?:export )?(?:const|let) (\w+)\b[^;]*?\b(?:Weak)?(?:Map|Set)<\s*OrderedDocument\b/gm;

function sourceFiles(dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) {
      return entry.name === '__tests__' ? [] : sourceFiles(path);
    }

    return entry.name.endsWith('.ts') ? [path] : [];
  });
}

/** Each map by document, and the number of other places that name such a type, by file. */
function documentMaps(): { readonly maps: string[]; readonly others: Record<string, number> } {
  const maps: string[] = [];
  const others: Record<string, number> = {};
  for (const path of sourceFiles(ORDERED)) {
    const file = relative(ORDERED, path);
    const source = readFileSync(path, 'utf8');
    const declared = [...source.matchAll(DOCUMENT_MAP)];
    for (const found of declared) {
      maps.push(`${file} ${found[1]}`);
    }

    const other = [...source.matchAll(DOCUMENT_KEYED)].length - declared.length;
    if (other !== 0) {
      others[file] = other;
    }
  }

  return { maps, others };
}

describe('the maps a module keeps by document', () => {
  test('each has a class: a new one says whether a checkpoint saves it', () => {
    const { maps, others } = documentMaps();
    expect(maps.sort()).toEqual(Object.keys(DOCUMENT_STATE).sort());
    expect(others).toEqual(OTHER_USES);
  });

  test('each saved map is a side state of the checkpoint, and each side state is one map', () => {
    const sides = Object.values(DOCUMENT_STATE).flatMap((state) =>
      state.kind === 'saved' && state.side !== undefined ? [state.side] : []
    );

    expect(sides.sort()).toEqual(Object.keys(SIDE).sort());
  });
});
