// The `header-footer` change kind (task 9.8): create, delete, link and unlink a section's header
// or footer, as the engine's lifecycle ops (`createHeaderFooter`, `deleteHeaderFooter`,
// `linkToPrevious`, `unlinkFromPrevious`).

import { suggestionFor } from '../../../edits/context.ts';
import { addSetting } from '../../../edits/resources.ts';
import {
  createFurniture,
  furnitureOf,
  removeFurniture,
  unlinkFurniture,
} from '../../../edits/sections-furniture.ts';
import { sectionsOfStory } from '../../../edits/sections.ts';
import { itemAt } from '../../../model/document.ts';
import { OrderedRefusal } from '../../../model/refusal.ts';
import type { OrderedFurnitureKind, OrderedFurnitureVariant } from '../../../model/sections.ts';
import type { OrderedDocument } from '../../../model/types.ts';
import { at, bool, Invalid, oneOf, required, type Fields } from '../../read.ts';
import type { OrderedChangeAddress, OrderedEditBase, OrderedEditChange } from '../../types.ts';
import { address } from '../../values.ts';
import { runKindEdit } from '../run.ts';
import { orderedKindSpec } from '../spec.ts';
import { touchSection } from './read.ts';

/**
 * Create, delete, link or unlink the header or footer of one variant of a section, named by the
 * item that holds the section's properties. A create may also turn on a different first page
 * (`titlePage`) and different odd and even pages (`evenAndOddHeaders`), as the editor does when
 * it creates a first-page or an even-page header.
 *
 * @public
 */
export interface OrderedHeaderFooterChange extends OrderedEditBase {
  readonly kind: 'header-footer';
  readonly action: 'create' | 'delete' | 'link' | 'unlink';
  readonly section: OrderedChangeAddress;
  readonly story: OrderedFurnitureKind;
  readonly variant: OrderedFurnitureVariant;
  readonly titlePage?: boolean;
  readonly evenAndOddHeaders?: boolean;
}

function read(fields: Fields): Record<string, unknown> {
  const path = fields.path;
  const action = oneOf(required(fields, 'action'), at(path, 'action'), [
    'create',
    'delete',
    'link',
    'unlink',
  ] as const);
  const out: Record<string, unknown> = {
    action,
    section: address(required(fields, 'section'), at(path, 'section')),
    story: oneOf(required(fields, 'story'), at(path, 'story'), ['header', 'footer'] as const),
    variant: oneOf(required(fields, 'variant'), at(path, 'variant'), [
      'default',
      'first',
      'even',
    ] as const),
  };
  for (const name of ['titlePage', 'evenAndOddHeaders'] as const)
    if (fields.has(name)) {
      if (action !== 'create') throw new Invalid('unknown-field', at(path, name));
      out[name] = bool(fields.get(name), at(path, name));
    }

  return out;
}

const isMine = (change: OrderedEditChange): change is OrderedHeaderFooterChange =>
  change.kind === 'header-footer';

/**
 * Whether a party's change would change a header or footer part, as the host's check sees it: a
 * create and an unlink add one; a delete and a link remove one unless another section still
 * names it. The check refuses a party's change that does (`furniture`); a reference that goes
 * while another section still names the part changes the main part only (`package`).
 */
function changesFurniture(change: OrderedEditChange, doc: OrderedDocument): boolean {
  if (!isMine(change)) return false;
  if (change.action === 'create' || change.action === 'unlink') return true;
  const story = doc.stories.get(change.part);
  const holder = story ? itemAt(story, change.section) : null;
  if (!story || holder === null) return true;
  const sections = sectionsOfStory(doc, story);
  const index = sections.findIndex((s) => s.holder === holder);
  if (index < 0) return true;
  const furniture = furnitureOf(doc, story, sections);
  const slots = (s: (typeof furniture)[number]) =>
    change.story === 'header' ? s.headers : s.footers;

  // `furnitureOf` gives one entry for each section
  const slot = slots(furniture[index]!).get(change.variant);
  if (!slot) return true;

  // The part stays when another section names it in its own reference
  return !furniture.some(
    (other, k) =>
      k !== index && [...slots(other).values()].some((s) => !s.inherited && s.story === slot.story)
  );
}

export const HEADER_FOOTER_SPEC = orderedKindSpec<OrderedHeaderFooterChange>({
  kind: 'header-footer',
  fields: ['action', 'section', 'story', 'variant', 'titlePage', 'evenAndOddHeaders'],
  read,
  touches: (change, reader) => touchSection(reader, change.section),
  apply(change, input) {
    // Word has no change record for a section's header or footer reference (`CT_SectPrChange`
    // holds `CT_SectPrBase`, ECMA-376 Part 1, §17.13.5.32): the organiser's direct edit only
    if (input.edit.direct !== true) {
      throw new OrderedRefusal(
        'invalid-edit',
        "Word records no change of a section's header or footer as a suggestion"
      );
    }

    return runKindEdit(
      input,
      (ctx) => {
        const suggestion = suggestionFor(ctx);
        const target = { section: change.section, story: change.story, variant: change.variant };
        switch (change.action) {
          case 'create':
            createFurniture(ctx, target, suggestion, change.titlePage === true);
            if (change.evenAndOddHeaders === true) {
              addSetting(ctx, suggestion.id, {
                op: 'setSectionFurnitureOptions',
                evenAndOddHeaders: true,
              });
            }

            break;
          case 'delete':
          case 'link':
            removeFurniture(ctx, target, suggestion, change.action === 'link');
            break;
          case 'unlink':
            unlinkFurniture(ctx, target, suggestion);
            break;
        }

        return suggestion.id;
      },
      { settleDirect: true }
    );
  },
  policy: { text: false, properties: false, furniture: changesFurniture },
  publish: 'story',
});
