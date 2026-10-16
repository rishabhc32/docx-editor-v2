// Suggesting mode proposes a content control's value instead of writing it.
//
// A dropdown pick, a checkbox toggle and a date pick wrote the control outright in suggesting
// mode: the value changed with no card and nothing to reject. They now strike the old display
// text and insert the new one inside the control, the form Word records. The control's own
// record of the value has no tracked form; a checkbox's follows its glyph back on Reject, and the
// others keep the value the suggestion wrote.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { resolveRevisionMarkup } from '../../contracts/revision-markup.ts';
import { revisionItemsOf } from '../../store/store/review-reads.ts';
import { planRevisionBatch } from '../../store/store/revision-batch.ts';
import { findContentControl } from '../../store/store/tree-op-nodes.ts';
import {
  contentControlWidgetDate,
  contentControlWidgetItems,
  contentControlWidgetValue,
} from '../content-control-widget-session.ts';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';
import { stubReviewModule } from './review-test-module.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const AUTHOR = 'Bea';

const sdt = (properties: string, content: string) =>
  `<w:p><w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt></w:p>`;
const DROPDOWN = sdt(
  '<w:dropDownList w:lastValue="1"><w:listItem w:displayText="One" w:value="1"/>' +
    '<w:listItem w:displayText="Two" w:value="2"/></w:dropDownList>',
  '<w:r><w:t>One</w:t></w:r>'
);
const CHECKBOX = sdt(
  '<w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/>' +
    '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>',
  '<w:r><w:sym w:font="MS Gothic" w:char="2610"/></w:r>'
);
const DATE = sdt(
  '<w:date w:fullDate="2026-01-02T00:00:00Z"><w:dateFormat w:val="M/d/yyyy"/></w:date>',
  '<w:r><w:t>1/2/2026</w:t></w:r>'
);

function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

const mounted: { destroy(): void; container: HTMLElement }[] = [];

afterEach(() => {
  for (const entry of mounted.splice(0)) {
    entry.destroy();
    entry.container.remove();
  }
});

function suggesting(body: string): { surface: PaginatedSurface; container: HTMLElement } {
  const container = document.createElement('div');
  document.body.append(container);
  const opened = mountPaginatedSurface(container, docx(body), {
    scale: 1,
    editingMode: 'suggest',
    author: AUTHOR,
    revisionDisplayMode: 'all-markup',
  });
  if (!opened.ok) throw new Error(opened.reason);
  mounted.push({ destroy: () => opened.surface.destroy(), container });
  return { surface: opened.surface, container };
}

/** The editor facade in suggesting mode, for what `can` and `exec` answer. */
function suggestingEditor(body: string): { editor: DocxEditorInstance; surface: PaginatedSurface } {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docx(body),
    author: AUTHOR,
    mode: 'suggesting',
    modules: [stubReviewModule()],
  });
  mounted.push({ destroy: () => editor.destroy(), container });
  if (!editor.surface) throw new Error('surface failed to mount');
  return { editor, surface: editor.surface };
}

function caret(surface: PaginatedSurface, index: number, offset: number): void {
  const paragraphId = surface.session.paragraphIds()[index]!;
  surface.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
}

const controlId = (surface: PaginatedSurface) => surface.layout().contentControls![0]!.id;
const control = (surface: PaginatedSurface) =>
  findContentControl(surface.session.part(), controlId(surface));
const xml = (surface: PaginatedSurface) => serializeOoxmlPart(surface.session.part());
/** What the control's `w:sdtContent` holds, ids and dates left out. */
const content = (surface: PaginatedSurface) =>
  (/<w:sdtContent>(.*?)<\/w:sdtContent>/s.exec(xml(surface))?.[1] ?? '')
    .replace(/ w:date="[^"]*"/g, '')
    .replace(/ w:id="[^"]*"/g, '');

function resolveAll(surface: PaginatedSurface, action: 'accept' | 'reject'): void {
  const plan = planRevisionBatch(surface.session.part(), action);
  const result = surface.applyAutomationOps(() => [...plan.ops]);
  if (result.rejected) throw new Error(String(result.reason));
}

function pressCheckbox(container: HTMLElement): HTMLElement {
  const widget = container.querySelector<HTMLElement>('[data-docx-cc-widget="checkbox"]')!;
  widget.dispatchEvent(
    new PointerEvent('pointerdown', {
      bubbles: true,
      cancelable: true,
      button: 0,
      pointerId: 1,
      pointerType: 'mouse',
    })
  );
  return container.querySelector<HTMLElement>('[data-docx-cc-widget="checkbox"]')!;
}

describe('content-control values in suggesting mode', () => {
  test('a dropdown pick is a deletion and an insertion inside the control', () => {
    const { surface } = suggesting(DROPDOWN);
    expect(surface.contentControls.setValue(controlId(surface), '2')).toBe(true);
    expect(content(surface)).toBe(
      `<w:del w:author="${AUTHOR}"><w:r><w:delText>One</w:delText></w:r></w:del>` +
        `<w:ins w:author="${AUTHOR}"><w:r><w:t>Two</w:t></w:r></w:ins>`
    );
    expect(revisionItemsOf(surface.session.part())).toMatchObject([
      { revisionKind: 'delete', author: AUTHOR, text: 'One' },
      { revisionKind: 'insert', author: AUTHOR, text: 'Two' },
    ]);
    // The list reads the item the suggestion shows, not the struck one run into it.
    const pending = control(surface);
    expect(contentControlWidgetValue(pending, contentControlWidgetItems(pending))).toBe('2');

    resolveAll(surface, 'reject');
    expect(content(surface)).toBe('<w:r><w:t>One</w:t></w:r>');
  });

  test('a checkbox toggle reports its proposed state, and Reject puts the old one back', () => {
    const { surface, container } = suggesting(CHECKBOX);
    surface.contentControls.setShowAll(true);
    expect(pressCheckbox(container).getAttribute('aria-checked')).toBe('true');
    expect(content(surface)).toContain(`<w:del w:author="${AUTHOR}">`);
    expect(xml(surface)).toContain('<w14:checked w14:val="1"/>');

    resolveAll(surface, 'reject');
    surface.contentControls.setShowAll(true);
    const widget = container.querySelector('[data-docx-cc-widget="checkbox"]')!;
    expect(widget.getAttribute('aria-checked')).toBe('false');
    expect(xml(surface)).toContain('<w14:checked w14:val="0"/>');
    expect(content(surface)).toBe('<w:r><w:sym w:char="2610" w:font="MS Gothic"/></w:r>');

    // The next press proposes checking it again, not the glyph it already shows.
    expect(pressCheckbox(container).getAttribute('aria-checked')).toBe('true');
    expect(content(surface)).toContain('w:char="2612"');
  });

  test('Reject restores a date’s text; w:fullDate keeps the picked date', () => {
    const { surface } = suggesting(DATE);
    surface.contentControls.setValue(controlId(surface), '2026-03-04');
    expect(content(surface)).toBe(
      `<w:del w:author="${AUTHOR}"><w:r><w:delText>1/2/2026</w:delText></w:r></w:del>` +
        `<w:ins w:author="${AUTHOR}"><w:r><w:t>3/4/2026</w:t></w:r></w:ins>`
    );

    resolveAll(surface, 'reject');
    expect(content(surface)).toBe('<w:r><w:t>1/2/2026</w:t></w:r>');
    // `w:sdtPr` has no tracked form, so nothing records the date the pick replaced.
    expect(contentControlWidgetDate(control(surface))).toBe('2026-03-04T00:00:00Z');
  });

  test('a value is content, so Track formatting off still proposes it', () => {
    const { surface } = suggesting(DROPDOWN);
    surface.setRevisionMarkup(resolveRevisionMarkup({ trackFormatting: false }));
    surface.contentControls.setValue(controlId(surface), '2');
    expect(content(surface)).toContain(`<w:ins w:author="${AUTHOR}">`);
  });

  test('one undo takes the proposal back, and redo proposes it again', () => {
    const { surface } = suggesting(DROPDOWN);
    surface.contentControls.setValue(controlId(surface), '2');
    surface.undo();
    expect(content(surface)).toBe('<w:r><w:t>One</w:t></w:r>');
    expect(xml(surface)).toContain('w:lastValue="1"');
    surface.redo();
    expect(content(surface)).toContain(
      `<w:ins w:author="${AUTHOR}"><w:r><w:t>Two</w:t></w:r></w:ins>`
    );
  });

  test('a temporary control, which a value write removes, is refused', () => {
    const { surface } = suggesting(sdt('<w:temporary/><w:text/>', '<w:r><w:t>Acme</w:t></w:r>'));
    const before = xml(surface);
    expect(surface.contentControls.setValue(controlId(surface), 'Beta')).toBe(false);
    expect(xml(surface)).toBe(before);
  });
});

// `can` validated the value without the revision suggesting mode adds, so it approved a value
// the tracked write then refused as "this control type is not supported".
describe('can and exec agree on a value suggesting mode cannot propose', () => {
  const command = { type: 'setContentControlValue' as const, value: 'Beta' };

  test('a temporary control is refused with the reason, unless the author proposed it', () => {
    const temporary = sdt('<w:temporary/><w:text/>', '<w:r><w:t>Acme</w:t></w:r>');
    const own = temporary
      .replace('<w:sdt>', `<w:ins w:id="9" w:author="${AUTHOR}"><w:sdt>`)
      .replace('</w:sdt>', '</w:sdt></w:ins>');
    const { editor, surface } = suggestingEditor(temporary + own);
    const before = xml(surface);

    caret(surface, 0, 1);
    const can = editor.can(command);
    expect(can).toEqual({
      ok: false,
      code: 'unsupported',
      reason: 'suggesting mode cannot propose a value for a temporary content control',
    });
    expect(editor.exec(command)).toEqual(can);
    expect(xml(surface)).toBe(before);

    // Rejecting the author's own insertion takes the control and its value together.
    caret(surface, 1, 1);
    expect(editor.can(command)).toEqual({ ok: true });
    expect(editor.exec(command)).toEqual({ ok: true, changed: true });
  });

  test('a block control over two paragraphs is refused with the reason', () => {
    const { editor, surface } = suggestingEditor(
      '<w:sdt><w:sdtPr><w:text w:multiLine="1"/></w:sdtPr><w:sdtContent>' +
        '<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>' +
        '</w:sdtContent></w:sdt>'
    );
    const before = xml(surface);

    caret(surface, 0, 0);
    const can = editor.can(command);
    expect(can).toEqual({
      ok: false,
      code: 'unsupported',
      reason: 'suggesting mode can propose a value only in a content control with one paragraph',
    });
    expect(editor.exec(command)).toEqual(can);
    expect(xml(surface)).toBe(before);
  });
});
