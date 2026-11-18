// What checks each entry of the editor's feature list (`docs/site/data/word-features.ts`) in the
// ordered model: the matrix kinds that put it in a story and meet every edit, the by-name
// refusals of what the first release does not cover (`refusals.ts`), or, for an entry with no
// content in a story (an export, a part of the editor's screen), the reason it has none.
// `coverage.test.ts` fails when an entry of the list has none of the three.

export interface FeatureCoverage {
  readonly kinds?: readonly string[];
  readonly refusals?: readonly string[];
  /** Why the entry has no content in a story for an edit to reach. */
  readonly outside?: string;
}

const NO_CONTENT = 'not content of a story: the model edits no part of it';

export const COVERAGE: Readonly<Record<string, FeatureCoverage>> = {
  'export.markdown': { outside: `${NO_CONTENT} (an export of the composed document)` },
  'export.pdf': { outside: `${NO_CONTENT} (an export of the composed document)` },
  'export.print': { outside: `${NO_CONTENT} (an export of the composed document)` },
  'text.basic-formatting': { kinds: ['bold', 'italic', 'udouble', 'uwave', 'strike', 'dstrike'] },
  'text.input-method': {
    outside: `${NO_CONTENT}: composition ends in the editor before its text reaches the model`,
  },
  'text.format-painter': { kinds: ['bold', 'face', 'fmt'] },
  'text.sub-superscript': { kinds: ['sup', 'sub'] },
  'text.fonts': { kinds: ['font', 'size'] },
  'text.embedded-fonts': { kinds: ['font'] },
  'text.color': { kinds: ['face', 'theme'] },
  'text.highlight': { kinds: ['face', 'runshade'] },
  'text.rtl': { kinds: ['rtl', 'complex', 'bidi', 'dir'] },
  'text.effects': { kinds: ['glow', 'effects', 'caps', 'smallcaps'] },
  'text.hidden': { kinds: ['hidden'] },
  'text.math': { kinds: ['math', 'mathpara', 'mathfrac'] },
  'text.symbols': { kinds: ['sym', 'emoji', 'emojiedge'] },
  'paragraphs.alignment': { kinds: ['left', 'centered', 'right', 'justify'] },
  'paragraphs.east-asian-typography': { kinds: ['eastasia', 'eastpara'] },
  'paragraphs.spacing': { kinds: ['lineauto', 'exact', 'atleast', 'beforeafter', 'contextual'] },
  'paragraphs.pagination': { kinds: ['keeps', 'widow'] },
  'paragraphs.indentation': { kinds: ['indents', 'hanging', 'firstline'] },
  'paragraphs.styles': { kinds: ['heading', 'quote', 'styled', 'charstyle'] },
  'paragraphs.borders': { kinds: ['borders', 'formatted'] },
  'paragraphs.tabs': { kinds: ['tabstops', 'tab', 'ptab'] },
  'paragraphs.frames': { kinds: ['frame', 'dropcap'] },
  'paragraphs.hyphenation': { kinds: ['nohyphen', 'soft', 'nbh'] },
  'lists.bullets': { kinds: ['bullet', 'multilevel'] },
  'lists.numbered': { kinds: ['list', 'roman', 'listrun', 'listcell'] },
  'lists.custom-numbering': { kinds: ['custom'], refusals: ['new-list-definition'] },
  'lists.continuation': { kinds: ['restart', 'listrun'] },
  'lists.picture-bullets': { kinds: ['picbullet'] },
  'tables.editing': { kinds: ['table', 'listcell', 'pictable'] },
  'tables.rtl': { kinds: ['rtltable'] },
  'tables.rows-columns': { kinds: ['table', 'widths', 'headrow'] },
  'tables.borders-shading': { kinds: ['tblborders', 'tblstyle'] },
  'tables.merge': { kinds: ['span', 'vmerge', 'mergeboth'] },
  'tables.page-break': { kinds: ['cantsplit', 'headrow'] },
  'tables.nested': { kinds: ['nested'] },
  'tables.conditional-formatting': { kinds: ['tblstyle', 'headrow'] },
  'tables.floating': { kinds: ['floattable'] },
  'tables.text-direction': { kinds: ['textdir'] },
  // New pictures, their properties, a new image and a deletion: the `pictures` set (group 9.4)
  'images.inline': { kinds: ['picture', 'pictable', 'picheader'] },
  'images.anchored': { kinds: ['anchor', 'wraptb', 'behind', 'tight'] },
  // A picture is one item whatever its media's format: the format never reaches the model
  'images.bmp-webp': { kinds: ['picture'] },
  'images.svg': { kinds: ['picture'] },
  'images.wmf': { kinds: ['picture'] },
  'images.tiff': { kinds: ['picture'] },
  'images.tracked': { kinds: ['picins'] },
  // A text box is a story of its own: every edit inside it is ordinary (group 9.9, `boxtext`)
  'images.textboxes': { kinds: ['textbox', 'vml', 'boxtext'] },
  'images.shapes': { kinds: ['shape'] },
  'images.legacy-vml': { kinds: ['vml', 'watermark'] },
  'images.crop': { kinds: ['crop'] },
  'images.adjustments': { kinds: ['picfx'] },
  'images.effects': { kinds: ['picfx'] },
  'images.charts': { kinds: ['chart'] },
  'images.smartart': { kinds: ['smartart'] },
  'images.ink': { kinds: ['ink'] },
  'layout.pagination': { outside: `${NO_CONTENT}: pages are the layout of the composed stories` },
  'layout.sections': { kinds: ['section', 'nextpage'] },
  'layout.headers-footers': { kinds: ['header', 'footer', 'firstheader', 'picheader'] },
  'layout.watermarks': { kinds: ['watermark'] },
  'layout.footnotes': { kinds: ['fnref', 'enref', 'footnote', 'endnote'] },
  'layout.columns': { kinds: ['cols', 'column'] },
  'layout.page-borders': { kinds: ['sectfancy'] },
  'layout.line-numbers': { kinds: ['sectfancy'] },
  'layout.even-odd-headers': { kinds: ['evenheader'] },
  'layout.vertical-align': { kinds: ['sectfancy'] },
  'layout.background': { kinds: ['background'] },
  'layout.page-num-format': { kinds: ['sectfancy'] },
  'review.tracked-changes': { kinds: ['ins', 'del', 'fmt', 'picins'] },
  'review.accept-reject': { kinds: ['ins', 'del', 'fmt', 'pprchange', 'markdel', 'rowins'] },
  'review.comments': { kinds: ['comment'] },
  'review.ai-redlining': { kinds: ['bold', 'ins'] },
  'review.moves': { kinds: ['movefile'] },
  'review.cell-revisions': { kinds: ['cellrev', 'rowins'] },
  'fields.hyperlinks': { kinds: ['link', 'hyperfield'] },
  'fields.bookmarks': { kinds: ['bookmark', 'emptymark', 'anchorlink'] },
  'fields.page-numbers': { kinds: ['pagefield', 'numpages'] },
  'fields.toc': { kinds: ['toc'], refusals: ['table-of-contents'] },
  'fields.cross-references': { kinds: ['ref', 'noteref'] },
  'fields.autonum': { kinds: ['autonum'] },
  'fields.styleref': { kinds: ['styleref'] },
  'fields.other-codes': { kinds: ['simple', 'datefield', 'seq', 'merge'] },
  'fields.citations': { kinds: ['citation'] },
  'fields.legacy-forms': { kinds: ['formtext', 'formcheck'] },
  'structure.content-controls': {
    kinds: ['plain', 'rich', 'sdtblock', 'placeholder', 'blockprompt', 'richprompt'],
  },
  'structure.repeating-sections': { kinds: ['repeating'] },
  'structure.typed-controls': {
    kinds: [
      'dropdown',
      'combo',
      'date',
      'checkbox',
      'piccontrol',
      'gallery',
      'dropprompt',
      'dateprompt',
    ],
  },
  'structure.custom-xml': { kinds: ['xml', 'xmlblock', 'bound'], refusals: ['bound-control'] },
  'structure.misplaced-breaks': { kinds: ['barebr'] },
  'structure.macros': { outside: `${NO_CONTENT}: a macro project is a part of its own` },
  'structure.ole': { kinds: ['ole'] },
  'structure.protection': { kinds: ['perm'] },
  'collab.realtime': {
    outside: `${NO_CONTENT}: the room and the client (tasks 5 to 7) carry the matrix's edits`,
  },
  'collab.ordered-rooms': {
    outside: `${NO_CONTENT}: the ordered room itself, which carries every edit the matrix makes`,
  },
  'collab.find-replace': { kinds: ['bold', 'face'] },
  'collab.anchor-navigation': { outside: `${NO_CONTENT} (a part of the editor's screen)` },
  'collab.clipboard': { kinds: ['picture', 'crop'], refusals: ['paste-with-formatting'] },
  'collab.undo-redo': {
    outside: `${NO_CONTENT}: undo is the client's (task 6.3), over the retraction that edits.test.ts checks`,
  },
  'collab.mount-modes': { outside: `${NO_CONTENT} (a part of the editor's screen)` },
  'collab.i18n': { outside: `${NO_CONTENT} (a part of the editor's screen)` },
  'collab.zoom-fit': { outside: `${NO_CONTENT} (a part of the editor's screen)` },
  'collab.document-refresh': {
    outside: `${NO_CONTENT}: the client publishes the composed paragraphs (task 6.2)`,
  },
  'collab.agent-tools': {
    outside: `${NO_CONTENT}: the automation API drives the editor, whose edits reach the model as the matrix's do`,
  },
};
