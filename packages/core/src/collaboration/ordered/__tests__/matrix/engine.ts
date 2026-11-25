// The oracle: today's editor in suggesting mode, driven through the surface a browser uses (or,
// where the editor's command is a store op, that op), and the engine's own accept and reject.

import { wmlAttributeValue, wmlChildNamed } from '../../../../store/store/tree-op-table-shared.ts';
import { planTableCommand } from '../../../../editor/table-command-plan.ts';
import {
  mountPaginatedSurface,
  type PaginatedSurface,
} from '../../../../editor/paginated-surface.ts';
import type { OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { findNode, parentNodeOf } from '../../../../store/package/ooxml-edit.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlDrawingNode,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../../../../store/package/ooxml-tree.ts';
import { placeholderBlocksIn } from '../../../../store/store/content-control-prompt-glossary.ts';
import {
  directParagraphProperties,
  mergedProperties,
} from '../../../../store/store/direct-properties.ts';
import { paragraphTextOf } from '../../../../store/store/paragraph-model-text.ts';
import { applyTreeOp } from '../../../../store/store/tree-op-apply.ts';
import { wholeMergedCellIds } from '../../../../store/store/tree-op-table-merged-cells.ts';
import { cloneWithNewIds } from '../../../../store/store/tree-op-nodes.ts';
import type { StoryScope } from '../../../../store/store/tree-package-store.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import { buildInlinePictureDrawing } from '../../../../store/package/drawing-package-edit.ts';
import { storyParagraphs } from '../../../../store/package/story-blocks.ts';
import {
  textboxStoriesInPart,
  type TextboxStoryRoot,
} from '../../../../store/package/textbox-stories.ts';
import type { OrderedObjectChangeOp, OrderedParagraphOp, OrderedTableOp } from '../../index.ts';
import { NEW_PICTURE_EMU } from './model.ts';
import type { Action } from './actions.ts';
import type { Selector } from './catalogue.ts';
import type { StoryTarget } from './kinds.ts';

export const DATE = '2026-10-18T00:00:00Z';

/** The editor on a fresh container; `close` destroys it and takes the container away. */
export function mount(
  bytes: Uint8Array,
  author: string
): { readonly surface: PaginatedSurface; readonly close: () => void } {
  const container = document.createElement('div');
  document.body.append(container);
  const opened = mountPaginatedSurface(container, bytes, {
    scale: 1,
    editingMode: 'suggest',
    author,
    revisionDisplayMode: 'all-markup',
  });
  if (!opened.ok) {
    container.remove();
    throw new Error(opened.reason);
  }
  const surface = opened.surface;
  return {
    surface,
    close: () => {
      surface.destroy();
      container.remove();
    },
  };
}

function find(text: string, selector: Selector): number {
  let index = -1;
  for (let n = 0; n <= (selector.nth ?? 0); n += 1) {
    index = text.indexOf(selector.find, index + 1);
    if (index < 0) throw new Error(`selector "${selector.find}" not found in "${text}"`);
  }
  return index;
}

export function caretOf(text: string, selector: Selector): number {
  const index = find(text, selector);
  return index + (selector.edge === 'end' ? selector.find.length : 0) + (selector.delta ?? 0);
}

export function rangeOf(text: string, selector: Selector): [number, number] {
  const index = find(text, selector);
  return [index, index + selector.find.length];
}

export interface Triple {
  readonly id: string;
  readonly author: string;
  readonly date?: string;
}

const TRIPLE_TAGS = new Set([
  'ins',
  'del',
  'rPrChange',
  'pPrChange',
  'cellIns',
  'cellDel',
  'cellMerge',
  'tblPrChange',
  'trPrChange',
  'tcPrChange',
  'tblGridChange',
  'moveFrom',
  'moveTo',
]);

/** Every revision triple in a part (content wrappers and change records). */
export function triplesOf(part: OoxmlPart): Triple[] {
  const found = new Map<string, Triple>();
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.namespaceUri === WML_NAMESPACE_URI && TRIPLE_TAGS.has(node.localName)) {
      const read = (name: string) =>
        node.attributes.find((a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === name)
          ?.value;
      const id = read('id');
      // A `w:tblGridChange` is `CT_Markup`: the engine names it by its `w:id` and no author
      const author = read('author') ?? (node.localName === 'tblGridChange' ? '' : undefined);
      if (id !== undefined && author !== undefined) {
        const date = read('date');
        found.set(`${id}|${author}|${date ?? ''}`, {
          id,
          author,
          ...(date === undefined ? {} : { date }),
        });
      }
    }
    for (const child of node.children) visit(child);
  };
  visit(part.root);
  return [...found.values()];
}

/**
 * Decide revision triples in the engine's part, one after another, reading the package's
 * glossary as the store's transaction does (a prompt comes back from its block). A grid record
 * (`w:tblGridChange`, named by its `w:id` alone) decided with its resize's other records goes
 * with them, so its own triple then finds nothing.
 */
export function engineDecide(
  part: OoxmlPart,
  triples: readonly Triple[],
  action: 'accept' | 'reject',
  pkg: OoxmlPackage
): OoxmlPart {
  let current = part;
  for (const revision of triples) {
    const result = applyTreeOp(
      current,
      { op: action === 'accept' ? 'acceptRevision' : 'rejectRevision', revision },
      { placeholderBlocks: (name) => placeholderBlocksIn(pkg, name) }
    );
    if (result.ok) current = result.part;
  }
  return current;
}

// ---- stories ----

/** The store scope of a story; null for the body (and a text box, which is in the body's part). */
export function scopeOf(story: StoryTarget | undefined): StoryScope | null {
  if (!story || story.scope === 'textbox') return null;
  if (story.scope !== 'note') return { kind: 'headerFooter', rId: story.id };
  return { kind: 'notesPart', noteKind: story.id.startsWith('end') ? 'endnote' : 'footnote' };
}

/** The body's one text box, as the editor reaches it. */
function textBoxOf(surface: PaginatedSurface): TextboxStoryRoot {
  const [box] = textboxStoriesInPart(surface.session.part());
  if (!box) throw new Error('no text box in the body');
  return box;
}

/** Open the story for editing, as a person double-clicks a header, a note or a text box. */
export function enterStory(surface: PaginatedSurface, story: StoryTarget | undefined): void {
  if (!story) return;
  const box = story.scope === 'textbox' ? textBoxOf(surface) : null;
  const opened = box
    ? // The text box's frame, as a click in it or Find's match opens it
      surface.setActiveScope({
        kind: 'frame',
        id: box.root.id,
        drawingNodeId: box.drawingNodeId,
        hostParagraphId: box.hostParagraphId,
      })
    : story.scope === 'note'
      ? surface.enterNote(story.id)
      : surface.enterHeaderFooter({
          rId: story.id,
          kind: story.scope === 'footer' ? 'footer' : 'header',
          ...(story.page === undefined ? {} : { pageIndex: story.page }),
        });
  if (!opened) throw new Error(`the editor did not open ${story.part}`);
}

export function storyPart(surface: PaginatedSurface, story: StoryTarget | undefined): OoxmlPart {
  const scope = scopeOf(story);
  const part = scope ? surface.session.partFor(scope) : surface.session.part();
  if (!part) throw new Error(`no part for ${story?.part ?? 'the body'}`);
  return part;
}

/** The story's paragraph ids, by the model's index (the scope's own leaves some out). */
function paragraphIds(surface: PaginatedSurface, story: StoryTarget | undefined): string[] {
  if (story?.scope === 'textbox')
    return storyParagraphs(textBoxOf(surface).root).map((paragraph) => paragraph.id);
  const scope = scopeOf(story);
  if (!scope) return surface.session.paragraphIds();
  const ids = surface.session.paragraphIdsIn(scope);
  return [...Array.from({ length: story?.skip ?? 0 }, () => ''), ...ids];
}

const TOGGLES: Readonly<Record<string, 'bold' | 'italic' | 'underline' | 'strikethrough'>> = {
  b: 'bold',
  i: 'italic',
  u: 'underline',
  strike: 'strikethrough',
};

/**
 * Whether the editor's toggle turns a property off on this range (the selection already shows
 * it, from the text or its style): the model's edit then writes it off too.
 */
export function togglesOff(
  surface: PaginatedSurface,
  story: StoryTarget | undefined,
  action: Action
): boolean {
  if (action.kind !== 'formatRange') return false;
  const flag = TOGGLES[action.prop];
  if (!flag) return false;
  const point = (index: number, selector: Selector) => ({
    paragraphId: paragraphId(surface, story, index),
    offset: caretOf(engineText(surface, story, index), selector),
  });
  surface.setSelection({
    anchor: point(action.p, action.from),
    head: point(action.to.p, action.to),
  });
  return surface.formatting()[flag];
}

function paragraphId(
  surface: PaginatedSurface,
  story: StoryTarget | undefined,
  index: number
): string {
  const id = paragraphIds(surface, story)[index];
  if (id === undefined) throw new Error(`no engine paragraph ${index}`);
  return id;
}

/** The engine's model text of the index-th paragraph of the story. */
export function engineText(
  surface: PaginatedSurface,
  story: StoryTarget | undefined,
  index: number
): string {
  const id = paragraphIds(surface, story)[index];
  return id === undefined ? '<none>' : (paragraphTextOf(storyPart(surface, story), id) ?? '');
}

// ---- actions ----

/** The paragraph's nearest table, row and cell, by node id. */
function tableOf(
  part: OoxmlPart,
  paragraphId: string
): { table: string; row: string; cell: string } {
  let cell: string | null = null;
  let row: string | null = null;
  for (let node = parentNodeOf(part, paragraphId); node; node = parentNodeOf(part, node.id)) {
    if (cell === null && node.kind === 'tableCell') cell = node.id;
    if (row === null && node.kind === 'tableRow') row = node.id;
    if (node.kind === 'table') return { table: node.id, row: row ?? '', cell: cell ?? '' };
  }
  throw new Error('the paragraph is not in a table');
}

function first(
  node: OoxmlNode,
  test: (n: OoxmlElement) => boolean,
  depth = 0
): OoxmlElement | null {
  if (node.kind === 'textValue' || depth > 40) return null;
  if (test(node)) return node;
  for (const child of node.children) {
    const found = first(child, test, depth + 1);
    if (found) return found;
  }
  return null;
}

function paragraphOpOf(part: OoxmlPart, paragraphId: string, op: OrderedParagraphOp): TreeDocOp {
  if (op.op === 'setParagraphProperties')
    return {
      op: 'setParagraphProperties',
      paragraphId,
      properties: mergedProperties(directParagraphProperties(part, paragraphId), op.properties),
    };
  if (op.op === 'setListLevel') return { op: 'setListLevel', paragraphId, level: op.level };
  if (op.op === 'setListNumbering')
    return {
      op: 'setListNumbering',
      paragraphId,
      numId: op.numId,
      ...(op.level === undefined ? {} : { level: op.level }),
    };
  return { op: 'setParagraphTabStops', paragraphId, stops: op.stops };
}

function tableOpOf(
  part: OoxmlPart,
  ids: { table: string; row: string; cell: string },
  op: OrderedTableOp
): TreeDocOp {
  const tableId = ids.table;
  // As the editor's cell commands name a caret's cell: a merged cell whole
  const cellIds = wholeMergedCellIds(part, tableId, [ids.cell]);
  switch (op.op) {
    case 'setTableProperties':
      return { ...op, tableId };
    case 'setTableRowHeight':
      return { op: 'setTableRowHeight', tableId, rowId: ids.row, heightTwips: op.heightTwips };
    case 'setTableCellFill':
      return { op: 'setTableCellFill', tableId, cellIds, color: op.color };
    case 'setTableCellVerticalAlignment':
      return { op: 'setTableCellVerticalAlignment', tableId, cellIds, alignment: op.alignment };
    case 'setTableCellBorders':
      return 'spec' in op
        ? { op: 'setTableCellBorders', tableId, cellIds, scope: op.scope, spec: op.spec }
        : { op: 'setTableCellBorders', tableId, cellIds, scope: 'none', target: op.target };
    case 'setTableColumnWidths':
    case 'setTableRightEdgeWidth': {
      // As the editor's drag names them: the caret's cell's right edge, or the last column
      const table = findNode(part, tableId);
      const grid =
        table && table.kind !== 'textValue'
          ? (table.children.find((c) => c.kind !== 'textValue' && c.localName === 'tblGrid') ??
            null)
          : null;
      const columns =
        grid && grid.kind !== 'textValue'
          ? grid.children.filter((c) => c.kind !== 'textValue' && c.localName === 'gridCol')
          : [];
      // The caret cell's right edge: the last grid column it covers (a merged cell covers more)
      const row = findNode(part, ids.row);
      let index = -1;
      let at = 0;
      for (const c of row && row.kind !== 'textValue' ? row.children : []) {
        if (c.kind !== 'tableCell') continue;
        const props = wmlChildNamed(c, 'tcPr');
        const span = props && wmlChildNamed(props, 'gridSpan');
        const value = span && wmlAttributeValue(span, 'val');
        at += Number(value ?? 1);
        if (c.id === ids.cell) index = at - 1;
      }
      const left = columns[index]?.id ?? '';
      if (op.op === 'setTableRightEdgeWidth')
        return {
          op: op.op,
          tableId,
          gridColumnId: left,
          columnWidthTwips: op.columnWidthTwips,
          tableWidthTwips: op.tableWidthTwips,
        };
      return {
        op: op.op,
        tableId,
        leftGridColumnId: left,
        rightGridColumnId: columns[index + 1]?.id ?? '',
        leftWidthTwips: op.leftWidthTwips,
        rightWidthTwips: op.rightWidthTwips,
      };
    }
  }
}

/** One action, through the surface as the editor's commands make it. */
export function performEngine(
  surface: PaginatedSurface,
  story: StoryTarget | undefined,
  action: Action
): void {
  const id = (index: number) => paragraphId(surface, story, index);
  const text = (index: number) => engineText(surface, story, index);
  const select = (index: number, start: number, end = start) =>
    surface.setSelection({
      anchor: { paragraphId: id(index), offset: start },
      head: { paragraphId: id(index), offset: end },
    });
  const point = (index: number, selector: Selector) => ({
    paragraphId: id(index),
    offset: caretOf(text(index), selector),
  });
  const revision =
    action.mode === 'edit' ? {} : { revision: { author: action.author, date: DATE } };
  const commit = (ops: readonly TreeDocOp[]) => {
    const scope = scopeOf(story);
    const result = surface.session.applyTreeOps(
      // Every op this harness makes takes the optional revision attribution
      ops.map((op) => ({ ...op, ...revision }) as TreeDocOp),
      undefined,
      undefined,
      scope ?? undefined
    );
    if (!result.committed) throw new Error(`refused: ${result.reason}`);
  };
  const tableCommand = (command: Record<string, string>) => {
    select(action.p, 0);
    const plan = planTableCommand({
      // The command shapes are the editor's; this harness names them by their type
      command: command as never,
      part: storyPart(surface, story),
      layout: surface.layout(),
      storeRevision: surface.session.packageRevision(),
      selection: surface.state().selection,
      cellSelection: surface.state().cellSelection,
      themeColors: surface.session.documentThemeColors(),
      editable: surface.session.editable,
      viewing: false,
    });
    const result = surface.applyTableCommandPlan(plan);
    if (!result.ok) throw new Error(`table command refused: ${result.reason ?? result.code}`);
  };
  switch (action.kind) {
    case 'insert':
      select(action.p, caretOf(text(action.p), action.at));
      surface.insertPlainText(action.text);
      return;
    case 'insertBreak':
      select(action.p, caretOf(text(action.p), action.at));
      if (action.break === 'tab') surface.insertTab();
      else if (action.break === 'line') surface.insertLineBreak();
      else surface.insertPageBreak();
      return;
    case 'deleteRange':
    case 'replaceRange':
    case 'formatRange':
      surface.setSelection({
        anchor: point(action.p, action.from),
        head: point(action.to.p, action.to),
      });
      if (action.kind === 'formatRange') surface.toggleRunProperty(action.prop, action.attributes);
      else surface.deleteSelection();
      if (action.kind === 'replaceRange') surface.insertPlainText(action.text);
      return;
    case 'enter':
      select(action.p, caretOf(text(action.p), action.at));
      surface.splitParagraph();
      return;
    case 'join':
      select(action.p, 0);
      surface.deleteBackward();
      return;
    case 'move':
    case 'copy': {
      const [start, end] = rangeOf(text(action.p), action.range);
      const words = text(action.p).slice(start, end);
      const at = () => caretOf(text(action.to.p), action.to.at);
      if (action.kind === 'move') {
        // The engine has no move: a cut, then a paste of the words at the destination. A refused
        // cut refuses the move: the paste's commit would clear the refusal.
        const before = surface.state().lastRejection;
        select(action.p, start, end);
        surface.deleteSelection();
        const refusal = surface.state().lastRejection;
        if (refusal && refusal !== before) {
          throw new Error(`refused: ${refusal}`);
        }

        select(action.to.p, at());
        surface.insertPlainText(words);
        return;
      }
      // A copy of a picture: the store op the editor's picture paste commits
      const paragraph = findNode(storyPart(surface, story), id(action.p));
      const drawing = paragraph ? first(paragraph, (n) => n.kind === 'drawing') : null;
      if (!drawing) throw new Error('no picture to copy');
      let next = 0;
      const copy = cloneWithNewIds(drawing, () => `matrix-picture-${(next += 1)}`);
      commit([
        {
          op: 'insertDrawing',
          paragraphId: id(action.to.p),
          offset: at(),
          // A copy of a node the reader made as a drawing
          drawing: copy as OoxmlDrawingNode,
        },
      ]);
      return;
    }
    case 'paragraph':
      commit(action.ops.map((op) => paragraphOpOf(storyPart(surface, story), id(action.p), op)));
      return;
    case 'insertRow':
      tableCommand({ type: 'insertRow', where: action.where });
      return;
    case 'deleteRow':
      tableCommand({ type: 'deleteRow' });
      return;
    case 'insertColumn':
      tableCommand({ type: 'insertColumn', where: action.where });
      return;
    case 'deleteColumn':
      tableCommand({ type: 'deleteColumn' });
      return;
    case 'tableProperty':
      commit([
        tableOpOf(
          storyPart(surface, story),
          tableOf(storyPart(surface, story), id(action.p)),
          action.op
        ),
      ]);
      return;
    case 'insertTable':
      commit([
        {
          op: 'insertTable',
          beforeParagraphId: id(action.p),
          rows: action.rows,
          cols: action.cols,
          columnWidthTwips: 2000,
        },
      ]);
      return;
    case 'link':
      select(action.p, caretOf(text(action.p), action.at));
      surface.hyperlinks.applyHyperlink({ url: action.url });
      return;
    case 'control': {
      const paragraph = findNode(storyPart(surface, story), id(action.p));
      const control = paragraph
        ? first(paragraph, (n) => n.localName === 'sdt' && n.id !== paragraph.id)
        : null;
      if (!control) throw new Error('no control in the paragraph');
      if (!surface.contentControls.setValue(control.id, action.value))
        throw new Error('value refused');
      return;
    }
    case 'insertNote':
      commit([
        {
          op: 'insertNote',
          noteKind: action.note,
          paragraphId: id(action.p),
          offset: caretOf(text(action.p), action.at),
        },
      ]);
      return;
    case 'pageField':
      commit([
        {
          op: 'insertPageField',
          paragraphId: id(action.p),
          offset: caretOf(text(action.p), action.at),
          field: action.field,
        },
      ]);
      return;
    case 'bookmark': {
      // A bookmark is no revision: the op takes none, in either mode
      const op: TreeDocOp = {
        op: 'insertBookmark',
        name: action.name,
        start: point(action.p, action.at),
        end: point(action.p, action.to),
      };
      const result = surface.session.applyTreeOps(
        [op],
        undefined,
        undefined,
        scopeOf(story) ?? undefined
      );
      if (!result.committed) throw new Error(`refused: ${result.reason}`);
      return;
    }
    case 'insertPicture':
      // The drawing the editor's Insert Picture builds, under a relationship the package does
      // not hold yet (the editor's package edit adds it)
      commit([
        {
          op: 'insertDrawing',
          paragraphId: id(action.p),
          offset: caretOf(text(action.p), action.at),
          drawing: buildInlinePictureDrawing({
            docPrId: 900,
            relationshipId: 'rIdNew1',
            extentEmu: { cx: NEW_PICTURE_EMU, cy: NEW_PICTURE_EMU },
          }) as OoxmlDrawingNode,
        },
      ]);
      return;
    case 'pictureProperty': {
      const drawing = drawingOf(surface, story, id(action.p));
      commit(action.ops.map((op) => drawingOpOf(drawing, op)));
      return;
    }
    case 'replacePicture':
      commit([
        {
          op: 'replaceDrawingResource',
          drawingNodeId: drawingOf(surface, story, id(action.p)),
          relationshipId: 'rIdNew2',
        },
      ]);
      return;
    case 'deletePicture': {
      if (action.mode === 'edit') {
        commit([{ op: 'deleteDrawing', drawingNodeId: drawingOf(surface, story, id(action.p)) }]);
        return;
      }
      // Suggesting: the editor strikes the picture's unit (`deleteImageTracked`)
      const at = text(action.p).indexOf('\uFFFC');
      select(action.p, at, at + 1);
      surface.deleteSelection();
      return;
    }
    case 'deleteNote':
      if (!surface.deleteNote(action.note, noteOf(surface, story, id(action.p), action.note)))
        throw new Error(`refused: ${surface.state().lastRejection}`);
      return;
    case 'convertNote':
      if (!surface.convertNote(action.note, noteOf(surface, story, id(action.p), action.note)))
        throw new Error(`refused: ${surface.state().lastRejection}`);
      return;
    case 'noteProperties':
      if (
        !surface.setNoteProperties({
          scope: 'document',
          footnote: { numFmt: action.numFmt, numStart: action.numStart },
        })
      )
        throw new Error(`refused: ${surface.state().lastRejection}`);
      return;
    case 'equation': {
      // The paragraph's first equation, as the editor's equation editor names it
      const paragraph = findNode(storyPart(surface, story), id(action.p));
      const math = paragraph ? first(paragraph, (n) => n.localName === 'oMath') : null;
      if (!math) throw new Error('no equation in the paragraph');
      const done =
        action.linear === null
          ? surface.equations.removeEquation(math.id)
          : surface.equations.applyEquation(math.id, action.linear);
      if (!done) throw new Error(`refused: ${surface.state().lastRejection}`);
      return;
    }
  }
}

/** The `w:id` of a paragraph's first note reference of a kind. */
function noteOf(
  surface: PaginatedSurface,
  story: StoryTarget | undefined,
  paragraphId: string,
  kind: 'footnote' | 'endnote'
): number {
  const paragraph = findNode(storyPart(surface, story), paragraphId);
  const reference = paragraph ? first(paragraph, (n) => n.localName === `${kind}Reference`) : null;
  const id = reference?.attributes.find((a) => a.localName === 'id')?.value;
  if (id === undefined) throw new Error(`no ${kind} in the paragraph`);
  return Number(id);
}

/** The first object of a paragraph, by the node id the editor names it with. */
function drawingOf(
  surface: PaginatedSurface,
  story: StoryTarget | undefined,
  paragraphId: string
): string {
  const paragraph = findNode(storyPart(surface, story), paragraphId);
  // The run's child that holds the drawing: the drawing, or the `mc:AlternateContent` around
  // it, as the editor names the object it selects
  const isDrawing = (n: OoxmlElement) => n.kind === 'drawing' || n.localName === 'drawing';
  const run = paragraph
    ? first(
        paragraph,
        (n) => n.kind === 'run' && n.children.some((c: OoxmlNode) => first(c, isDrawing) !== null)
      )
    : null;
  const held = run?.children.find((c: OoxmlNode) => first(c, isDrawing) !== null);
  if (!held) throw new Error('no picture in the paragraph');
  return held.id;
}

/** The engine's op for one of a property change's ops. */
function drawingOpOf(drawingNodeId: string, op: OrderedObjectChangeOp): TreeDocOp {
  switch (op.op) {
    case 'resize':
      return { op: 'resizeDrawing', drawingNodeId, extentEmu: { cx: op.cx, cy: op.cy } };
    case 'crop':
      return { op: 'cropDrawing', drawingNodeId, crop: op.crop };
    case 'position':
      return { op: 'positionDrawing', drawingNodeId, position: op.position };
    case 'wrap':
      return { op: 'setDrawingWrap', drawingNodeId, wrap: op.wrap };
    case 'metadata':
      return {
        op: 'setDrawingMetadata',
        drawingNodeId,
        title: op.title,
        description: op.description,
      };
    case 'locks':
      return { op: 'setDrawingLocks', drawingNodeId, locks: op.locks };
    case 'transform':
      return { op: 'transformDrawing', drawingNodeId, action: op.action };
    case 'image':
      throw new Error('a new image is its own action');
  }
}
