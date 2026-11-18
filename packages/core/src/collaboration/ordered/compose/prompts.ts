// Content control prompts and value records in the composed part.

import {
  contentControlPropertiesOf as summaryOf,
  contentControlTextOf,
  orderedContentControlProperties,
  type ContentControlProperties,
} from '../../../store/package/content-control-nodes.ts';
import { createNodeIdAllocator, findNode, replaceNode } from '../../../store/package/ooxml-edit.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { promptToRetain } from '../../../store/store/content-control-prompt-glossary.ts';
import {
  withDecisionPrompt,
  withRestoredPrompt,
} from '../../../store/store/content-control-prompt-restore.ts';
import {
  parseContentControlDate,
  parseIsoDate,
} from '../../../store/store/content-control-dates.ts';
import { editedProperties, wmlElement } from '../../../store/store/tree-op-content-controls.ts';
import {
  contentControlPropertiesOf,
  findContentControl,
} from '../../../store/store/tree-op-nodes.ts';
import type { OrderedDocument } from '../model/types.ts';
import { attributeOf, childNamed, isWml, withChildren } from '../ooxml-nodes.ts';
import type { ComposeContext } from './context.ts';
import { withStoryBlocks } from './paragraph-ops.ts';

/** The `w:docPart/@w:val` a control's `w:placeholder` names, if any. */
export function placeholderNameOf(control: OoxmlElement): string | undefined {
  const sdtPr = childNamed(control, 'sdtPr');
  const placeholder = sdtPr ? childNamed(sdtPr, 'placeholder') : undefined;
  const docPart = placeholder ? childNamed(placeholder, 'docPart') : undefined;
  const name = docPart ? attributeOf(docPart, 'val')?.trim() : undefined;
  return name && name.length <= 512 ? name : undefined;
}

/** How `hidePrompt` treats a control's prompt. */
interface HidePromptOptions {
  /** False when the prompt shows again after a suggestion that filled it went. Default true. */
  readonly hide?: boolean;
  /** False for a prompt Enter removed for good, as Word does: no block keeps it. Default true. */
  readonly retain?: boolean;
  /** The control sits in a paragraph: the engine names the block it keeps by that too. */
  readonly inline?: boolean;
}

/**
 * The control's properties without `w:showingPlcHdr`, naming the placeholder block that keeps
 * its prompt, and the block reported for the caller to write when the glossary lacks it: the
 * engine's own choice of block and name (`promptToRetain`), so one prompt gets one block
 * whichever side keeps it. `original` is the control as it was loaded, showing its prompt.
 *
 * A prompt shows only while its control holds nothing else. While a suggestion's text fills the
 * control, the prompt cannot be the control's content, so Word's other place for it is used: a
 * placeholder block in the glossary, named by the control's `w:placeholder` (ECMA-376
 * §17.5.2.25). Compose names the block on the control and reports the block for the caller to
 * write; it never writes another part itself.
 */
export function hidePrompt(
  ctx: ComposeContext,
  props: readonly OoxmlNode[],
  original: OoxmlNode | undefined,
  { hide = true, retain = true, inline = false }: HidePromptOptions
): readonly OoxmlNode[] {
  let named: string | undefined;
  if (retain && original && original.kind !== 'textValue') {
    named = placeholderNameOf(original);
    const retained = promptToRetain(partHolding(ctx, original, inline), original, {
      retainPrompt: () => undefined,
      placeholderBlocks: (name) => ctx.doc.placeholders.get(name),
    });
    if (retained) {
      named = retained.name;
      if (!ctx.prompts.some((prompt) => prompt.name === retained.name)) ctx.prompts.push(retained);
    }
  }

  return props.map((child) => {
    if (child.kind === 'textValue' || !isWml(child, 'sdtPr')) return child;
    const kept = hide ? child.children.filter((c) => !isWml(c, 'showingPlcHdr')) : child.children;
    if (named === undefined || childNamed(child, 'placeholder')) return withChildren(child, kept);
    const reference = wmlElement(ctx.mint, 'placeholder', {
      children: [wmlElement(ctx.mint, 'docPart', { attributes: [['val', named]] })],
    });
    return withChildren(child, orderedContentControlProperties([...kept, reference]));
  });
}

/**
 * The story's part, or, for a control it lacks (a copy that Enter made when it split a control),
 * a part where the control sits as compose draws it: the engine names the block it keeps by the
 * prompt and by whether the control is inline.
 */
function partHolding(ctx: ComposeContext, control: OoxmlElement, inline: boolean): OoxmlPart {
  if (findNode(ctx.story.part, control.id)) return ctx.story.part;
  const block = inline ? build(ctx.mint(), 'paragraph', 'p', [], [control]) : control;
  return withStoryBlocks(ctx.story.part, [block]);
}

/**
 * `part` with each emptied control's prompt shown again, the glossary's own when it has one. A
 * control a decision on a suggestion emptied (`decided`) shows what the engine's decision leaves:
 * Word's blank prompt when it names no prompt (`withDecisionPrompt`). One a direct edit emptied
 * shows its type's prompt, as the engine's emptying edit does.
 */
export function withRestoredPrompts(
  doc: OrderedDocument,
  part: OoxmlPart,
  controls: readonly string[],
  decided: ReadonlySet<string>
): OoxmlPart {
  const options = { placeholderBlocks: (name: string) => doc.placeholders.get(name) };
  for (const id of controls)
    part = decided.has(id)
      ? withDecisionPrompt(part, id, options)
      : withRestoredPrompt(part, id, options);
  return part;
}

/** `part` with each control's value record (`w:lastValue`, `w:fullDate`) following its text. */
export function withControlRecords(part: OoxmlPart, controls: readonly string[]): OoxmlPart {
  for (const id of controls) {
    const control = findContentControl(part, id);
    if (!control) continue;
    const edits = valueRecordEdits(control);
    if (
      edits.lastValue === undefined &&
      edits.fullDate === undefined &&
      edits.checked === undefined
    )
      continue;
    const properties = contentControlPropertiesOf(control);
    if (!properties) continue;
    const edited = editedProperties(properties, edits, createNodeIdAllocator(part));
    const written = replaceNode(part, properties.id, edited);
    if (written.ok) part = written.part;
  }

  return part;
}

/** The value record writes that bring a control's record in step with its text. */
interface ValueRecordEdits {
  lastValue?: string | null;
  fullDate?: string | null;
  checked?: boolean;
}

function valueRecordEdits(control: OoxmlElement): ValueRecordEdits {
  const summary = summaryOf(control);
  const text = contentControlTextOf(control);
  if (summary.type === 'dropDownList' || summary.type === 'comboBox')
    return listValueEdits(summary, text);
  if (summary.type === 'checkbox') return checkboxEdits(summary, text);
  if (summary.type === 'date') return dateEdits(summary, text);
  return {};
}

/** A list's last value: the value of the item it shows (a combo box's own text). */
function listValueEdits(summary: ContentControlProperties, text: string): ValueRecordEdits {
  if (summary.showingPlaceholder) return summary.lastValue === undefined ? {} : { lastValue: null };
  const item = summary.listItems.find((candidate) => candidate.displayText === text);
  const value = item?.value ?? (summary.type === 'comboBox' ? text : undefined);
  return value === undefined || value === summary.lastValue ? {} : { lastValue: value };
}

/** A check box's state: it follows the glyph the control shows. */
function checkboxEdits(summary: ContentControlProperties, text: string): ValueRecordEdits {
  const box = summary.checkbox;
  const glyph = (state: { readonly value: string } | undefined) =>
    state === undefined ? undefined : String.fromCodePoint(Number.parseInt(state.value, 16) || 0);
  let checked: boolean | undefined;
  if (text === glyph(box?.checkedState)) {
    checked = true;
  } else if (text === glyph(box?.uncheckedState)) {
    checked = false;
  }

  return checked === undefined || checked === box?.checked ? {} : { checked };
}

/** A date's full date: the date it shows, in the century of the one before. */
function dateEdits(summary: ContentControlProperties, text: string): ValueRecordEdits {
  const current = summary.date?.fullDate;
  if (summary.showingPlaceholder) return current === undefined ? {} : { fullDate: null };
  const previous = current === undefined ? null : parseIsoDate(current);
  const century = previous ? Math.floor(previous.year / 100) * 100 : undefined;
  const shown = parseContentControlDate(text, summary.date?.dateFormat, century);
  if (!shown) return {};
  const pad = (value: number, width: number) => String(value).padStart(width, '0');
  const fullDate = `${pad(shown.year, 4)}-${pad(shown.month, 2)}-${pad(shown.day, 2)}T00:00:00Z`;
  return fullDate === current ? {} : { fullDate };
}
