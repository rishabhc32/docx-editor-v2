// The exact mix and nested cells of the families in `known.ts`: a family holds most pairs of
// its kinds, not all, so its cells are listed by name.

/** Mix cells where typing or Enter moves a bookmark start in front of other markers. */
export const BOOKMARK_TYPING: readonly string[] = (
  'comment+bookmark comment+emptymark dropprompt+bookmark dropprompt+emptymark ' +
  'ptab+bookmark ptab+emptymark '
)
  .trim()
  .split(' ');

/** Mix cells where a deletion of both kinds strikes a control's prompt. */
export const PROMPT_DELETE_BOTH: readonly string[] = (
  'anchor+dropprompt bold+dropprompt bookmark+dropprompt br+dropprompt checkbox+dropprompt ' +
  'comment+dropprompt complex+dropprompt crop+dropprompt datefield+dropprompt ' +
  'del+dropprompt dropdown+dropprompt dropprompt+anchor dropprompt+bold dropprompt+bookmark ' +
  'dropprompt+br dropprompt+checkbox dropprompt+comment dropprompt+complex dropprompt+crop ' +
  'dropprompt+datefield dropprompt+del dropprompt+dropdown dropprompt+fnref dropprompt+font ' +
  'dropprompt+formtext dropprompt+hidden dropprompt+ins dropprompt+link dropprompt+math ' +
  'dropprompt+pagefield dropprompt+piccontrol dropprompt+picture dropprompt+plain ' +
  'dropprompt+rich dropprompt+shape dropprompt+simple dropprompt+sup dropprompt+tab ' +
  'dropprompt+textbox dropprompt+udouble fnref+dropprompt font+dropprompt ' +
  'formtext+dropprompt hidden+dropprompt ins+dropprompt link+dropprompt math+dropprompt ' +
  'pagefield+dropprompt piccontrol+dropprompt picture+dropprompt plain+dropprompt ' +
  'rich+dropprompt shape+dropprompt simple+dropprompt sup+dropprompt tab+dropprompt ' +
  'textbox+dropprompt udouble+dropprompt '
)
  .trim()
  .split(' ');

/** Mix cells where a deletion where the kinds meet strikes a control's prompt. */
export const PROMPT_DELETE_MEET: readonly string[] = (
  'anchor+dropprompt bold+dropprompt bookmark+dropprompt br+dropprompt checkbox+dropprompt ' +
  'comment+dropprompt complex+dropprompt crop+dropprompt datefield+dropprompt ' +
  'del+dropprompt dropdown+dropprompt dropprompt+anchor dropprompt+bold dropprompt+bookmark ' +
  'dropprompt+br dropprompt+checkbox dropprompt+comment dropprompt+complex dropprompt+crop ' +
  'dropprompt+datefield dropprompt+del dropprompt+dropdown dropprompt+fnref dropprompt+font ' +
  'dropprompt+formtext dropprompt+hidden dropprompt+ins dropprompt+link dropprompt+math ' +
  'dropprompt+pagefield dropprompt+piccontrol dropprompt+picture dropprompt+plain ' +
  'dropprompt+rich dropprompt+shape dropprompt+simple dropprompt+sup dropprompt+tab ' +
  'dropprompt+textbox dropprompt+udouble fnref+dropprompt font+dropprompt hidden+dropprompt ' +
  'ins+dropprompt link+dropprompt math+dropprompt pagefield+dropprompt ' +
  'piccontrol+dropprompt picture+dropprompt plain+dropprompt rich+dropprompt ' +
  'shape+dropprompt simple+dropprompt sup+dropprompt tab+dropprompt textbox+dropprompt ' +
  'udouble+dropprompt '
)
  .trim()
  .split(' ');

/** Nested cells where deleting the container strikes the prompt inside it. */
export const PROMPT_IN_CONTAINER: readonly string[] = [
  'dropprompt in bookmark',
  'dropprompt in comment',
  'dropprompt in gallery',
  'dropprompt in link',
  'dropprompt in plain',
  'dropprompt in rich',
  'dropprompt in smart',
  'dropprompt in xml',
];
