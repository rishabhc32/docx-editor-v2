// The `attribution` change kind (group 9.10): suggestions' date, without a decision. Their author
// stays: a room item's comes only from its person.

import { setAttribution } from '../../../model/attribution.ts';
import { madeBy } from '../../../model/document.ts';
import { OrderedRefusal } from '../../../model/refusal.ts';
import type { OrderedAttribution } from '../../../model/types.ts';
import { fieldPath, Invalid, plainText, required, str } from '../../read.ts';
import type { OrderedEditBase } from '../../types.ts';
import { suggestionIds } from '../../values.ts';
import { orderedKindSpec } from '../spec.ts';

/** Word's cap on an author's name ([MS-OI29500] §17.13.4.2 note c). */
const AUTHOR = 255;
const INSTANT = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d{1,3})?Z$/;

/**
 * Give pending suggestions another date, without deciding them (`setRevisionAttribution`). Word
 * has no command for it: these are the `w:author` and `w:date` of each revision mark. A
 * suggestion's marks carry its one attribution, so the change names whole suggestions; their
 * `w:id`s stay. `author` must be each suggestion's own: another one is refused for every sender
 * (`fixed-author`).
 *
 * @public
 */
export interface OrderedAttributionChange extends OrderedEditBase {
  readonly kind: 'attribution';
  readonly suggestions: readonly string[];
  readonly author: string;
  /** `YYYY-MM-DDTHH:MM:SSZ`, written as `w:date`. */
  readonly date?: string;
}

export const ATTRIBUTION_SPEC = orderedKindSpec<OrderedAttributionChange>({
  kind: 'attribution',
  fields: ['suggestions', 'author', 'date'],
  read(fields) {
    // Every path starts at the change, so a refusal names its place in the message
    const author = plainText(required(fields, 'author'), fieldPath(fields, 'author'), 1, AUTHOR);
    if (author.trim().length === 0)
      throw new Invalid('invalid-text', fieldPath(fields, 'author'), 'blank');
    const out: { suggestions: string[]; author: string; date?: string } = {
      suggestions: suggestionIds(required(fields, 'suggestions'), fieldPath(fields, 'suggestions')),
      author,
    };
    if (fields.has('date')) {
      const date = str(fields.get('date'), fieldPath(fields, 'date'), 24);
      if (!INSTANT.test(date)) throw new Invalid('invalid-date', fieldPath(fields, 'date'));
      out.date = date;
    }

    return out;
  },
  // A change of attribution names suggestions, not places
  touches: () => {},
  apply(change, { doc }) {
    for (const id of change.suggestions) {
      const suggestion = doc.suggestions.get(id);
      if (!suggestion) throw new OrderedRefusal('unknown-suggestion', id);
      if (suggestion.status !== 'pending') throw new OrderedRefusal('not-pending', id);
      const before = suggestion.attribution;
      if (change.author !== before.author) {
        throw new OrderedRefusal('fixed-author', id);
      }

      const date = change.date ?? before.date;

      // A new date drops what the file said beside the old one (`w16du:dateUtc`)
      const next: OrderedAttribution = {
        author: before.author,
        person: before.person,
        ...(date === undefined ? {} : { date }),
        ...(change.date === undefined && before.attributes !== undefined
          ? { attributes: before.attributes }
          : {}),
      };
      setAttribution(doc, id, next);
    }

    // It makes and extends no suggestion: its step in the history has no inverse, as a
    // decision's has none (the room writes it for good)
    return { suggestions: [], replaced: [] };
  },
  policy: {
    text: false,
    properties: false,
    // The host's check skips a party's own suggestions before it compares: their date may
    // change; another person's suggestion is not theirs to write (another author's name is
    // refused for everyone as it applies: `fixed-author`)
    party(change, doc, maker) {
      if (change.kind !== 'attribution') return null;
      const notOwn = (detail: string) =>
        ({ ok: false, code: 'policy-not-own', reason: 'others_suggestions', detail }) as const;
      for (const id of change.suggestions) {
        const pending = doc.suggestions.get(id)?.status === 'pending';
        if (!pending || !madeBy(doc, id, maker)) {
          return notOwn(id);
        }
      }

      return { ok: true };
    },
  },
  // Every mark of the suggestions changes: their paragraphs are drawn again, in every story
  publish: 'paragraphs',
  rewrites: (change) => change.suggestions,
});
