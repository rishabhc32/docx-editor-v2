// Paste with formatting, story block replacement, block deletion, new list definitions,
// document protection and revision attribution (group 9.10).

import { mainStory } from '../../model/document.ts';
import type { OrderedDocument, OrderedStory } from '../../model/types.ts';
import { OrderedClientRefusal } from '../refusal.ts';
import { attributionChanges } from './paste/attribution.ts';
import { deleteBlockChanges } from './paste/delete-block.ts';
import { insertFragmentChanges } from './paste/fragment.ts';
import { translator, type OrderedOpOption, type OrderedOpTable } from './types.ts';
import { untrackedRule } from './untracked.ts';

/** The story whose part root (or body) has this node id. */
function storyOfRoot(doc: OrderedDocument, rootId: string): string {
  for (const story of doc.stories.values()) {
    const root = story.part.root;
    if (root.id === rootId) return story.name;
    for (const child of root.children)
      if (child.kind !== 'textValue' && child.localName === 'body' && child.id === rootId)
        return story.name;
  }

  throw new OrderedClientRefusal(
    'ordered-refused',
    'out-of-step',
    `no story ${rootId.slice(0, 160)}`
  );
}

/**
 * The main document's story, or a refusal by name when the list has none. A change that names
 * no place there still names its part.
 */
export function mainStoryOf(doc: OrderedDocument): OrderedStory {
  const story = mainStory(doc);
  if (!story) {
    throw new OrderedClientRefusal('ordered-refused', 'unknown-story', 'no main document');
  }

  return story;
}

export const PASTE_BLOCK_OPS: OrderedOpTable<
  | 'insertFragment'
  | 'replaceStoryBlocks'
  | 'deleteBlock'
  | 'setDocumentProtection'
  | 'setRevisionAttribution'
> = {
  insertFragment: {
    status: 'translated',
    group: '9.10',
    changes: ['insert-fragment'],
    party: null,
    note: 'A paste with formatting: blocks with their runs, styles, lists, links and notes (D14), tracked as Word tracks a paste.',
    translate: translator('insertFragment', (op, batch) => () => insertFragmentChanges(op, batch)),
  },
  replaceStoryBlocks: {
    status: 'translated',
    group: '9.10',
    changes: ['replace-story'],
    party: untrackedRule('text', "a story's blocks replaced with plain paragraphs"),
    tracking: 'direct',
    note: "A story's blocks replaced with plain paragraphs (automation): every item struck with one settled deletion, the new paragraphs after.",
    translate: translator('replaceStoryBlocks', (op, batch) => () => [
      {
        kind: 'replace-story',
        part: storyOfRoot(batch.doc, op.storyRootId),
        paragraphs: op.paragraphs,
      },
    ]),
  },
  deleteBlock: {
    status: 'translated',
    group: '9.10',
    changes: ['delete-row', 'delete-paragraph'],
    party: null,
    note: "A table (tracked: every row struck, as the engine's tracked deletion), a row, or a paragraph (direct).",
    translate: translator('deleteBlock', (op, batch) => () => deleteBlockChanges(op, batch)),
  },
  setDocumentProtection: {
    status: 'translated',
    group: '9.10',
    changes: ['setting'],
    party: untrackedRule('package', 'protection is a setting'),
    tracking: 'direct',
    note: 'Editing restrictions on or off: a setting resource (D14), never a suggestion.',
    translate: translator('setDocumentProtection', (op, batch) => () => [
      {
        kind: 'setting',
        part: mainStoryOf(batch.doc).name,
        setting: { op: 'setDocumentProtection', enforce: op.enforce },
      },
    ]),
  },
  setRevisionAttribution: {
    status: 'translated',
    group: '9.10',
    changes: ['attribution'],
    party: {
      code: 'policy-not-own',
      reason: 'others_suggestions',
      why: "it rewrites suggestions' date",
    },
    tracking: 'review',
    note: "The date of suggestions: the review write `revision-attribution`; a party may change only the date of their own. Another author is refused for everyone (`fixed-author`): a room item's author comes only from its person.",
    translate: translator(
      'setRevisionAttribution',
      (op, batch) => () => attributionChanges(op, batch)
    ),
  },
};

/** Options of these ops that change their meaning. */
export const PASTE_BLOCK_OPTIONS: readonly OrderedOpOption[] = [
  {
    op: 'insertFragment',
    option: 'object',
    status: 'planned',
    group: '9.10',
    refusal: 'ordered-unsupported-op:insertFragment.object',
    note: "A paste that holds an element of a drawing namespace the format does not hold outside an extension and a choice, or an object the change carries no part for: the change format carries pictures (a VML picture's image by `r:id`, `o:relid` or `r:pict` too), shapes, groups and drawing canvases, VML pictures and shapes, SVG versions, charts (also Office 2016 charts, `cx`) and diagrams (SmartArt) with their own parts, embedded objects and ink with their parts, and a list's picture bullets, their media and parts by hash, never opened (the editor's copy leaves out an ActiveX control, a linked object and a chart linked to a workbook file; a paste that still names a linked object is refused by name, `not-covered`). An extension (`a:ext`) or a choice (`mc:Choice`) of a namespace the format does not hold is left out, as Word older than the namespace reads it. A picture's address never enters a room (`external-picture-link`, D18).",
  },
  {
    op: 'insertText',
    option: 'mixed-tracking',
    status: 'translated',
    group: '9.10',
    note: "A batch of tracked and direct ops: each change takes its own op's mode; an armed format of typed text follows the typing.",
  },
];
