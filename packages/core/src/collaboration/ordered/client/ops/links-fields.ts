// Link and field ops (group 9.5): new and removed links, a link's new target, address, tooltip
// or part, page fields, field codes and results, a table of contents and bookmarks.

import { stylesPartOf } from '../../../../store/package/ooxml-indexes.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import type { OrderedChangeStyle } from '../../change/resources.ts';
import { isHidden } from '../../model/document.ts';
import { attributeOf } from '../../ooxml-nodes.ts';
import type { OrderedItem } from '../../model/types.ts';
import { OrderedClientRefusal } from '../refusal.ts';
import { addressOf, placeAt, type ListParagraph } from '../paragraphs.ts';
import { containerOf, type ListContainer } from './containers.ts';
import {
  translateDeleteBookmark,
  translateFieldCode,
  translateFieldResults,
  translateInsertBookmark,
  translateInsertToc,
  translatePageField,
  translateTocPages,
  translateTocResult,
} from './fields.ts';
import {
  translator,
  type OrderedChangeBody,
  type OrderedOpOption,
  type OrderedOpTable,
  type OrderedTranslateBatch,
} from './types.ts';
import { untrackedRule } from './untracked.ts';
import { addedStyleId } from './paste/definitions.ts';

/**
 * Where an op's link points: the address of the relationship the editor minted for it (read on
 * the editor's package, or on the package it shows), or the bookmark it names.
 */
function targetOf(
  batch: OrderedTranslateBatch,
  part: string,
  op: { readonly relationshipId?: string; readonly anchor?: string }
): { readonly href: string } | { readonly anchor: string } {
  if (op.anchor !== undefined && op.relationshipId === undefined) return { anchor: op.anchor };
  if (op.relationshipId === undefined || op.anchor !== undefined)
    throw new OrderedClientRefusal('ordered-refused', 'invalid-change', 'one link target');
  const id = op.relationshipId;
  const href =
    batch.hyperlinkTarget(part, id) ??
    batch.shown()?.externalTargets.find((t) => t.ownerPart === part && t.id === id)?.rawTarget;
  if (href === undefined)
    throw new OrderedClientRefusal(
      'ordered-refused',
      'out-of-step',
      `no address ${id.slice(0, 64)}`
    );
  return { href };
}

/**
 * Word's Hyperlink character style under `styleId`, when the document holds no style of its own
 * under that id: the definition the editor adds with a first link (`hyperlinkStyleWrite`), based
 * on the default character style. A style a change added (pending, or another person's) is no
 * style of the file's, so a link declares its own, which a reject of the other leaves in place.
 */
function hyperlinkStyle(
  batch: OrderedTranslateBatch,
  styleId: string
): OrderedChangeStyle | undefined {
  const shown = batch.shown();
  const styles = shown ? stylesPartOf(shown) : undefined;
  const declared = styles?.root.children.some(
    (c) =>
      c.kind !== 'textValue' && c.localName === 'style' && attributeOf(c, 'styleId') === styleId
  );
  if (declared === true && !addedStyleId(batch.doc, styleId)) {
    return undefined;
  }

  let basedOn: string | null = null;
  for (const c of styles?.root.children ?? []) {
    if (c.kind === 'textValue' || c.localName !== 'style' || attributeOf(c, 'type') !== 'character')
      continue;
    const id = attributeOf(c, 'styleId') ?? null;
    const isDefault = ['1', 'true', 'on'].includes(attributeOf(c, 'default') ?? '');

    // The last default character style wins, as the editor reads it
    if (isDefault) basedOn = id;
    else if (basedOn === id) basedOn = null;
  }

  return {
    styleId,
    type: 'character',
    name: 'Hyperlink',
    ...(basedOn === null ? {} : { basedOn }),
    uiPriority: 99,
    unhideWhenUsed: true,
    run: [
      { localName: 'color', attributes: { val: '0563C1', themeColor: 'hyperlink' } },
      { localName: 'u', attributes: { val: 'single' } },
    ],
  };
}

/** The shown paragraph a container's open item is in, as the batch reads it. */
function paragraphOfLink(batch: OrderedTranslateBatch, link: ListContainer): ListParagraph {
  for (let at: OrderedItem | null = link.open; at; at = link.story.tree.next(at)) {
    if (at.kind !== 'mark' || isHidden(batch.doc, at)) continue;
    const node = at.node;
    if (!node || node.kind === 'textValue') break;
    return batch.paragraph(node.id);
  }

  throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'a link in no paragraph');
}

/** A part of a link's words, by paragraph offsets, as places. */
function partOf(
  batch: OrderedTranslateBatch,
  link: ListContainer,
  range: { readonly start: number; readonly end: number } | undefined
): { from?: ReturnType<typeof placeAt>; to?: ReturnType<typeof placeAt> } {
  if (range === undefined) return {};
  const p = paragraphOfLink(batch, link);
  return { from: placeAt(p, range.start), to: placeAt(p, range.end) };
}

type LinkOp = Extract<TreeDocOp, { op: 'insertHyperlink' | 'setHyperlinkTarget' }>;

const tooltipOf = (op: LinkOp) => (op.tooltip === undefined ? {} : { tooltip: op.tooltip });

export const LINK_FIELD_OPS: OrderedOpTable<
  | 'setHyperlinkTarget'
  | 'insertHyperlink'
  | 'removeHyperlink'
  | 'insertPageField'
  | 'setFieldCode'
  | 'refreshFieldResults'
  | 'insertToc'
  | 'replaceTocResult'
  | 'rewriteTocPageNumbers'
  | 'insertBookmark'
  | 'deleteBookmark'
> = {
  setHyperlinkTarget: {
    status: 'translated',
    group: '6.1',
    changes: ['link-target'],
    party: null,
    tracking: 'either',
    note: "A link's new address or bookmark, tooltip, or a part of its words; in place with tracking off.",
    translate: translator('setHyperlinkTarget', (op, batch) => () => {
      const link = containerOf(batch.doc, op.linkId, ['hyperlink']);
      const part = link.story.name;
      return [
        {
          kind: 'link-target',
          part,
          link: addressOf(link.open),
          ...targetOf(batch, part, op),
          ...tooltipOf(op),
          ...partOf(batch, link, op.range),
        },
      ];
    }),
  },
  insertHyperlink: {
    status: 'translated',
    group: '9.5',
    changes: ['insert-link'],
    party: null,
    tracking: 'either',
    note: "A link around a range; suggesting wraps the batch's own tracked copy of the words.",
    translate: translator('insertHyperlink', (op, batch) => () => {
      const p = batch.paragraph(op.paragraphId);
      const part = p.story.name;
      const newStyle = op.styleId === undefined ? undefined : hyperlinkStyle(batch, op.styleId);
      const body: OrderedChangeBody = {
        kind: 'insert-link',
        part,
        from: placeAt(p, op.start),
        to: placeAt(p, op.end),
        ...targetOf(batch, part, op),
        ...tooltipOf(op),
        ...(op.styleId === undefined ? {} : { style: op.styleId }),
        ...(newStyle === undefined ? {} : { newStyle }),
      };
      return [body];
    }),
  },
  removeHyperlink: {
    status: 'translated',
    group: '9.5',
    changes: ['remove-link'],
    party: null,
    tracking: 'either',
    note: "A link's words without the link; suggesting unwraps the batch's own tracked copy.",
    translate: translator('removeHyperlink', (op, batch) => () => {
      const link = containerOf(batch.doc, op.linkId, ['hyperlink']);
      return [
        {
          kind: 'remove-link',
          part: link.story.name,
          link: addressOf(link.open),
          ...partOf(batch, link, op.range),
          ...(op.styleId === undefined ? {} : { style: op.styleId }),
        },
      ];
    }),
  },
  insertPageField: {
    status: 'translated',
    group: '9.5',
    changes: ['insert-field'],
    party: null,
    note: 'A PAGE, NUMPAGES, SECTIONPAGES or "Page X of Y" field: a whole field, one insertion when tracked.',
    translate: translator('insertPageField', (op, batch) => () => translatePageField(op, batch)),
  },
  setFieldCode: {
    status: 'translated',
    group: '9.5',
    changes: ['field-code'],
    party: untrackedRule('package', "a field's instruction"),
    tracking: 'direct',
    note: "A field's instruction (PAGE and NUMPAGES only), the engine's own rewrite; never run.",
    translate: translator('setFieldCode', (op, batch) => () => translateFieldCode(op, batch)),
  },
  refreshFieldResults: {
    status: 'translated',
    group: '9.5',
    changes: ['field-result'],
    party: untrackedRule('text', "a field's shown result"),
    tracking: 'direct',
    note: "Fields' shown results, the text in the change; the instruction stays.",
    translate: translator(
      'refreshFieldResults',
      (op, batch) => () => translateFieldResults(op, batch)
    ),
  },
  insertToc: {
    status: 'translated',
    group: '9.5',
    changes: ['insert-toc'],
    party: untrackedRule('text', 'a table of contents with its result'),
    tracking: 'direct',
    note: 'A table of contents control with its first rows (text and pages in the change) and heading bookmarks.',
    translate: translator('insertToc', (op, batch) => () => translateInsertToc(op, batch)),
  },
  replaceTocResult: {
    status: 'translated',
    group: '9.5',
    changes: ['toc-result'],
    party: untrackedRule('text', "a table of contents' result"),
    tracking: 'direct',
    note: "A table of contents' rows made again, the rows in the change.",
    translate: translator('replaceTocResult', (op, batch) => () => translateTocResult(op, batch)),
  },
  rewriteTocPageNumbers: {
    status: 'translated',
    group: '9.5',
    changes: ['field-result', 'toc-pages'],
    party: untrackedRule('text', "a table of contents' page numbers"),
    tracking: 'direct',
    note: "A table of contents' page numbers only: a PAGEREF field's last text, or a row's plain text.",
    translate: translator(
      'rewriteTocPageNumbers',
      (op, batch) => () => translateTocPages(op, batch)
    ),
  },
  insertBookmark: {
    status: 'translated',
    group: '9.5',
    changes: ['bookmark'],
    party: untrackedRule('package', 'a bookmark is no revision'),
    tracking: 'direct',
    note: 'A named bookmark around a range or at a point; the same name elsewhere goes.',
    translate: translator(
      'insertBookmark',
      (op, batch) => () => translateInsertBookmark(op, batch)
    ),
  },
  deleteBookmark: {
    status: 'translated',
    group: '9.5',
    changes: ['bookmark'],
    party: untrackedRule('package', 'a bookmark is no revision'),
    tracking: 'direct',
    note: "A bookmark's markers removed, only its copies, or the pair a deletion encloses.",
    translate: translator(
      'deleteBookmark',
      (op, batch) => () => translateDeleteBookmark(op, batch)
    ),
  },
};

/** Options of the link and field ops that change their meaning. */
export const LINK_FIELD_OPTIONS: readonly OrderedOpOption[] = [
  {
    op: 'setHyperlinkTarget',
    option: 'relationshipId',
    status: 'translated',
    group: '9.5',
    note: "A new address: the address of the relationship the editor minted, as the link's own relationship resource (D14).",
  },
  {
    op: 'setHyperlinkTarget',
    option: 'range-or-tooltip',
    status: 'translated',
    group: '9.5',
    note: "A new target for part of a link's words (a link of its own), or a new tooltip.",
  },
];
