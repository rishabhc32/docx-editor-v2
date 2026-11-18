// What a write in a text box changes on the screen (group 9.9): the editor shows a text box
// inside the paragraph that holds its object.

import { itemAt } from '../model/document.ts';
import { hasTextBoxes, isTextBoxStoryName, textBoxHostOf } from '../model/text-boxes.ts';
import type { OrderedDocument } from '../model/types.ts';
import { nextShownMark } from './paragraphs.ts';
import { OrderedDirty } from './publish.ts';

/**
 * Name the host paragraph of a text box story in `into`: a write that names a text box's story
 * names that paragraph of the host story. A text box story a take-back removed names nothing:
 * the take-back of its object names the host paragraph itself.
 */
function nameHost(doc: OrderedDocument, name: string, into: OrderedDirty): void {
  const host = textBoxHostOf(doc, name);
  const story = host ? doc.stories.get(host.host) : undefined;
  if (!host || !story) return;
  const atom = itemAt(story, host.atom);
  const mark = atom ? nextShownMark(doc, story, atom) : null;
  if (mark) {
    into.mark(story.name, { space: mark.space, offset: mark.offset });
  } else {
    into.story(story.name);
  }
}

/** `dirty` with every text box story named by its host paragraph instead. */
export function withTextBoxHosts(doc: OrderedDocument, dirty: OrderedDirty): OrderedDirty {
  if (!hasTextBoxes(doc)) return dirty;
  const boxes = [...dirty.stories, ...dirty.marks.keys()].filter(isTextBoxStoryName);
  if (boxes.length === 0) return dirty;
  const out = new OrderedDirty();
  for (const name of dirty.stories) if (!isTextBoxStoryName(name)) out.story(name);
  for (const [name, marks] of dirty.marks)
    if (!isTextBoxStoryName(name)) for (const address of marks.values()) out.mark(name, address);
  for (const name of boxes) nameHost(doc, name, out);
  out.stale = dirty.stale;
  return out;
}
