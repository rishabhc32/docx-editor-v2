// Header, footer and section ops (group 9.8): create, delete, link and unlink, page furniture
// options, section properties and section breaks, watermark.

import type { PackageWatermarkWrite } from '../../../../store/package/watermark-package.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import type {
  OrderedChangeWatermark,
  OrderedChangeWatermarkBox,
} from '../../change/kinds/sections.ts';
import type { OrderedChangeMedia } from '../../change/resources.ts';
import type { OrderedSectionView } from '../../compose/sections.ts';
import { sectionsOfStory } from '../../edits/sections.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from '../../model/types.ts';
import { addressOf, paragraphAtIndex } from '../paragraphs.ts';
import { OrderedClientRefusal } from '../refusal.ts';
import { mainStoryOf } from './paste-blocks.ts';
import {
  translator,
  type OrderedOpOption,
  type OrderedOpTable,
  type OrderedPartyRule,
  type OrderedTranslateBatch,
} from './types.ts';

/** The room's rule: a new, changed or removed header or footer part is the organiser's alone. */
const FURNITURE: OrderedPartyRule = {
  code: 'policy-furniture',
  reason: 'furniture',
  why: "headers and footers are the organiser's alone: a new, changed or removed header or footer part",
};

/**
 * The section at the editor's index, as the engine counts them. The client names it by the item
 * that holds its properties (`orderedSectionsOf`), an address every copy shares.
 */
function sectionAtIndex(
  doc: OrderedDocument,
  story: OrderedStory,
  index: number
): OrderedSectionView {
  const section = sectionsOfStory(doc, story)[index];
  if (!section) {
    throw new OrderedClientRefusal('ordered-refused', 'engine-refused', `sectionIndex ${index}`);
  }

  return section;
}

/** The fields of an object that are defined. */
function defined<T extends object>(fields: T): Partial<T> {
  const out: Partial<T> = {};
  for (const key of Object.keys(fields) as (keyof T)[])
    if (fields[key] !== undefined) out[key] = fields[key];
  return out;
}

/**
 * The mark a section break ends its section at: Enter's new mark when the batch splits the
 * paragraph first (the head keeps the editor's id, and the list's new mark ends it), else the
 * paragraph's own.
 */
function breakMark(
  batch: OrderedTranslateBatch,
  op: Extract<TreeDocOp, { op: 'setSectionMark' }>
): OrderedItem {
  const index = batch.ops.indexOf(op);
  const split = batch.ops
    .slice(0, index < 0 ? 0 : index)
    .some((other) => other.op === 'splitParagraph' && other.paragraphId === op.paragraphId);
  const p = batch.paragraph(op.paragraphId);
  if (!split) return p.mark;

  // The list keeps the paragraph's own mark on the tail: the head is the paragraph before it
  return paragraphAtIndex(batch.doc, p.story, p.index - 1).mark;
}

/**
 * The shape size of the editor's watermark, and its size for each section, as the change gives
 * them.
 */
function watermarkBox(write: PackageWatermarkWrite): OrderedChangeWatermarkBox {
  const sizes = write.sectionSizes;
  return {
    widthPt: write.widthPt,
    heightPt: write.heightPt,
    ...(sizes === undefined
      ? {}
      : {
          sectionSizes: sizes.map((size) =>
            size ? { widthPt: size.widthPt, heightPt: size.heightPt } : null
          ),
        }),
  };
}

/** The editor's text watermark as the change gives it. */
function textWatermark(
  write: Extract<PackageWatermarkWrite, { kind: 'text' }>
): OrderedChangeWatermark {
  return {
    kind: 'text',
    text: write.text,
    fontFamily: write.fontFamily,
    fontSizePt: write.fontSizePt,
    bold: write.bold,
    italic: write.italic,
    colorHex: write.colorHex,
    opacity: write.opacity,
    diagonal: write.diagonal,
    ...watermarkBox(write),
  };
}

/** The editor's picture watermark as the change gives it: its image by hash. */
function pictureWatermark(
  write: Extract<PackageWatermarkWrite, { kind: 'picture' }>,
  media: OrderedChangeMedia
): OrderedChangeWatermark {
  return { kind: 'picture', washout: write.washout, media, ...watermarkBox(write) };
}

/**
 * The editor's watermark as the change gives it, or null for none. A picture's bytes are held
 * under their hash: the session sends them before the change.
 */
function watermarkOf(
  write: PackageWatermarkWrite | null,
  batch: OrderedTranslateBatch
): OrderedChangeWatermark | null {
  if (write === null) {
    return null;
  }

  if (write.kind === 'picture') {
    return pictureWatermark(write, batch.heldMedia(write.image.bytes, write.image.mime));
  }

  return textWatermark(write);
}

type Lifecycle = Extract<
  TreeDocOp,
  { op: 'createHeaderFooter' | 'deleteHeaderFooter' | 'linkToPrevious' | 'unlinkFromPrevious' }
>;

const ACTIONS = {
  createHeaderFooter: 'create',
  deleteHeaderFooter: 'delete',
  linkToPrevious: 'link',
  unlinkFromPrevious: 'unlink',
} as const;

/**
 * A header or footer of a section made, removed, linked or unlinked: one `header-footer` change.
 * A new header or footer part is a story the change adds, with its part and relationship as
 * resources (design D14).
 */
function furnitureStep(op: Lifecycle, batch: OrderedTranslateBatch) {
  return () => {
    const main = mainStoryOf(batch.doc);
    const section = sectionAtIndex(batch.doc, main, op.sectionIndex);
    const create = op.op === 'createHeaderFooter';
    return [
      {
        kind: 'header-footer' as const,
        part: main.name,
        action: ACTIONS[op.op],
        section: addressOf(section.holder),
        story: op.kind,
        variant: op.variant,
        ...(create && op.titlePage === true ? { titlePage: true } : {}),
        ...(create && op.evenAndOddHeaders === true ? { evenAndOddHeaders: true } : {}),
      },
    ];
  };
}

function lifecycle<K extends Lifecycle['op']>(op: K, note: string) {
  return {
    status: 'translated',
    group: '9.8',
    changes: ['header-footer'],
    party: FURNITURE,
    tracking: 'direct',
    note,
    translate: translator(op, furnitureStep),
  } as const;
}

export const SECTION_OPS: OrderedOpTable<
  | 'createHeaderFooter'
  | 'deleteHeaderFooter'
  | 'linkToPrevious'
  | 'unlinkFromPrevious'
  | 'setSectionFurnitureOptions'
  | 'setWatermark'
  | 'setSectionProperties'
  | 'setSectionMark'
> = {
  createHeaderFooter: lifecycle(
    'createHeaderFooter',
    "A new header or footer part for a section's variant, and its reference."
  ),
  deleteHeaderFooter: lifecycle(
    'deleteHeaderFooter',
    "A section's header or footer reference removed; the part goes when unused."
  ),
  linkToPrevious: lifecycle(
    'linkToPrevious',
    "A section takes the previous section's header or footer."
  ),
  unlinkFromPrevious: lifecycle(
    'unlinkFromPrevious',
    "A section gets its own copy of the previous section's header or footer."
  ),
  setSectionFurnitureOptions: {
    status: 'translated',
    group: '9.8',
    changes: ['section-property', 'setting'],
    party: null,
    note: "A different first page and header and footer distances on a section; suggested, Word's `w:sectPrChange`. Different odd and even pages, a setting.",
    translate: translator('setSectionFurnitureOptions', (op, batch) => () => {
      const main = mainStoryOf(batch.doc);
      const furniture = defined({
        titlePage: op.titlePage,
        headerDistanceTwips: op.headerDistanceTwips,
        footerDistanceTwips: op.footerDistanceTwips,
      });
      const section =
        Object.keys(furniture).length === 0
          ? []
          : [
              {
                kind: 'section-property' as const,
                part: main.name,
                sections: [addressOf(sectionAtIndex(batch.doc, main, op.sectionIndex ?? 0).holder)],
                furniture,
              },
            ];

      // Odd and even pages is the document's setting, written as every setting is
      const setting =
        op.evenAndOddHeaders === undefined
          ? []
          : [
              {
                kind: 'setting' as const,
                part: main.name,
                setting: {
                  op: 'setSectionFurnitureOptions' as const,
                  evenAndOddHeaders: op.evenAndOddHeaders,
                },
              },
            ];
      return [...section, ...setting];
    }),
  },
  setWatermark: {
    status: 'translated',
    group: '9.8',
    changes: ['watermark'],
    party: FURNITURE,
    tracking: 'direct',
    note: 'A watermark, which Word keeps in every header a section shows.',
    translate: translator('setWatermark', (op, batch) => {
      const watermark = watermarkOf(op.watermark, batch);
      return () => [{ kind: 'watermark', part: mainStoryOf(batch.doc).name, watermark }];
    }),
  },
  setSectionProperties: {
    status: 'translated',
    group: '9.8',
    changes: ['section-property'],
    party: null,
    note: "Page size, orientation and margins of every section, or of the one a paragraph is in; suggested, Word's `w:sectPrChange`.",
    translate: translator('setSectionProperties', (op, batch) => () => {
      // The op's revision is its mode: a suggestion, which the change's attribution names
      const { op: _name, anchorParagraphId, revision: _revision, ...fields } = op;

      // The room finds the sections when it applies the change: every one, or the anchor's
      const pageAt =
        anchorParagraphId === undefined
          ? {}
          : { pageAt: addressOf(batch.paragraph(anchorParagraphId).mark) };
      return [
        {
          kind: 'section-property',
          part: mainStoryOf(batch.doc).name,
          sections: [],
          page: defined(fields),
          ...pageAt,
        },
      ];
    }),
  },
  setSectionMark: {
    status: 'translated',
    group: '9.8',
    changes: ['section-break'],
    party: null,
    tracking: 'either',
    note: "A section break on a paragraph's mark; tracked with the mark the batch proposes, its type for the next section untracked, as Word writes it.",
    translate: translator('setSectionMark', (op, batch) => () => {
      const mark = breakMark(batch, op);
      const main = mainStoryOf(batch.doc);
      return [
        {
          kind: 'section-break',
          part: main.name,
          mark: addressOf(mark),
          ...(op.breakType === undefined ? {} : { type: op.breakType }),
        },
      ];
    }),
  },
};

/** Options of the section ops, and decisions on a file's own change records. */
export const SECTION_OPTIONS: readonly OrderedOpOption[] = [
  {
    op: 'setSectionFurnitureOptions',
    option: 'odd-and-even-pages',
    status: 'translated',
    group: '9.8',
    note: "Different odd and even pages: a setting, which has no tracked form, so the editor writes it untracked in suggesting mode too and the room makes it a direct edit. A party's is refused (`policy-direct`, `package`): the room's rule hashes `word/settings.xml`.",
  },
  {
    op: 'acceptRevision',
    option: 'section-record',
    status: 'translated',
    group: '9.8',
    note: "A file's own `w:sectPrChange`: a suggestion since load reads it, decided as a status.",
  },
  {
    op: 'setSectionMark',
    option: 'retype-suggested',
    status: 'translated',
    group: '9.8',
    note: "A suggested break that gives the next section another type: the type goes on untracked, as Word writes it; a reject leaves it, a take-back removes it. A party's is refused (`policy-direct`, `package`): the room's rule sees the type stay.",
  },
  {
    op: 'acceptRevision',
    option: 'file-record',
    status: 'translated',
    group: '9.8',
    note: "A file's own `w:tblPrExChange`, `w:numberingChange` or `w:numPr/w:ins`: a suggestion since load reads it, decided as a status, as Word decides it (Word for Mac, 2026-10-29). A rejected row record resets the row's cells with no record of their own, as the engine does; both decisions of a numbering record remove it and keep the list. A `w:tblGridChange` is a live grid record of its table's suggestion (group 9.7, design D16).",
  },
];
