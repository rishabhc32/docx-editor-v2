// Every op an everyday suggesting gesture commits is one suggesting mode can record.
//
// The rule is the one the opt-in strict mode applies: after `attributeTrackedOps`, an op must
// answer `isTrackedEdit`, or the batch changes the document with no card for a reviewer. Lists,
// list indents, tab stops and content-control values each committed ops that did not, so the
// strict mode refused contract edits a suggester makes every day. This pins the batches the
// session is actually handed, through every lane those gestures take.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import type { TreeDocOp } from '@docx-editor.dev/core/store';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { isTrackedEdit } from '../surface-tracked-ops.ts';
import { runToolbarCommand } from '../toolbar-commands.ts';
import { stubReviewModule } from './review-test-module.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

const STYLES =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>' +
  '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>';
const NUMBERING =
  `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0">` +
  '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/></w:lvl>' +
  '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';

const PLAIN = '<w:p><w:r><w:t>alpha</w:t></w:r></w:p>';
const item = (level: number, text: string) =>
  '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>' +
  `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
  (text ? `<w:r><w:t>${text}</w:t></w:r>` : '') +
  '</w:p>';
const sdt = (properties: string, content: string) =>
  `<w:p><w:r><w:t xml:space="preserve">Field: </w:t></w:r><w:sdt><w:sdtPr>${properties}</w:sdtPr>` +
  `<w:sdtContent>${content}</w:sdtContent></w:sdt></w:p>`;
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
const TEXT = sdt('<w:text/>', '<w:r><w:t>Acme</w:t></w:r>');

function docx(body: string): Uint8Array {
  const type = (name: string) =>
    `application/vnd.openxmlformats-officedocument.wordprocessingml.${name}+xml`;
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${type('document.main')}"/>` +
        `<Override PartName="/word/styles.xml" ContentType="${type('styles')}"/>` +
        `<Override PartName="/word/numbering.xml" ContentType="${type('numbering')}"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}">` +
        `<Relationship Id="rIdS" Type="${R}/styles" Target="styles.xml"/>` +
        `<Relationship Id="rIdN" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`
    ),
    'word/styles.xml': strToU8(`<w:styles xmlns:w="${W}">${STYLES}</w:styles>`),
    'word/numbering.xml': strToU8(NUMBERING),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

interface Suggesting {
  readonly editor: DocxEditorInstance;
  readonly surface: PaginatedSurface;
  readonly container: HTMLElement;
  /** Every batch the surface handed the session since the mount. */
  readonly batches: (readonly TreeDocOp[])[];
}

const mounted: { editor: DocxEditorInstance; container: HTMLElement }[] = [];

afterEach(() => {
  for (const entry of mounted.splice(0)) {
    entry.editor.destroy();
    entry.container.remove();
  }
});

function suggesting(body: string): Suggesting {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docx(body),
    author: 'Bea',
    mode: 'suggesting',
    modules: [stubReviewModule()],
  });
  mounted.push({ editor, container });
  const surface = editor.surface;
  if (!surface) throw new Error('surface failed to mount');
  // The surface commits through this session object, so the batches recorded here are the
  // ones after attribution: exactly what the store receives.
  const batches: (readonly TreeDocOp[])[] = [];
  const session = surface.session as { applyTreeOps: PaginatedSurface['session']['applyTreeOps'] };
  const apply = session.applyTreeOps.bind(session);
  session.applyTreeOps = (ops, ...rest) => {
    batches.push(ops);
    return apply(ops, ...rest);
  };
  return { editor, surface, container, batches };
}

function caret(surface: PaginatedSurface, index: number, offset: number): void {
  const paragraphId = surface.session.paragraphIds()[index]!;
  surface.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
}

function controlId(surface: PaginatedSurface): string {
  return surface.layout().contentControls![0]!.id;
}

/** The committed ops suggesting mode could not show as a suggestion. */
function unrecorded(batches: readonly (readonly TreeDocOp[])[]): string[] {
  return batches.flat().flatMap((op) => (isTrackedEdit(op) ? [] : [op.op]));
}

/** Run a gesture; it must change the document, and every op it commits must be recorded. */
function recorded(target: Suggesting, gesture: () => unknown): void {
  const before = target.surface.session.packageRevision();
  gesture();
  expect(target.surface.session.packageRevision()).not.toBe(before);
  expect(target.batches.length).toBeGreaterThan(0);
  expect(unrecorded(target.batches)).toEqual([]);
}

describe('list gestures in suggesting mode', () => {
  test('Bullets on a plain paragraph', () => {
    const target = suggesting(PLAIN);
    caret(target.surface, 0, 2);
    recorded(target, () => runToolbarCommand(target.editor, 'list.bullet'));
  });

  for (const [slot, gesture] of [
    ['list.numbered', 'Numbering converts a bulleted item'],
    ['list.bullet', 'Bullets turns a bulleted item off'],
    ['list.indent', 'Increase Indent demotes a list item'],
  ] as const) {
    test(gesture, () => {
      const target = suggesting(item(0, 'alpha'));
      caret(target.surface, 0, 2);
      recorded(target, () => runToolbarCommand(target.editor, slot));
    });
  }

  test('Decrease Indent promotes a list item', () => {
    const target = suggesting(item(1, 'alpha'));
    caret(target.surface, 0, 2);
    recorded(target, () => runToolbarCommand(target.editor, 'list.outdent'));
  });

  test('Tab and Shift+Tab over a list item change its level', () => {
    const target = suggesting(item(0, 'alpha') + item(0, 'beta'));
    const [first, second] = target.surface.session.paragraphIds();
    target.surface.setSelection({
      anchor: { paragraphId: first!, offset: 0 },
      head: { paragraphId: second!, offset: 2 },
    });
    recorded(target, () => target.surface.indentWithTab('increase'));
    recorded(target, () => target.surface.indentWithTab('decrease'));
  });

  test('Enter on an empty item leaves the list', () => {
    const target = suggesting(item(0, 'alpha') + item(0, ''));
    caret(target.surface, 1, 0);
    recorded(target, () => target.surface.exitListOnEmptyItem());
  });
});

describe('tab stops in suggesting mode', () => {
  test('the paragraph dialog with tab stops and spacing', () => {
    const target = suggesting(PLAIN);
    caret(target.surface, 0, 2);
    recorded(target, () =>
      target.editor.exec({
        type: 'setParagraphFormat',
        tabStops: [{ positionTwips: 1440, alignment: 'right' }],
        spaceAfterPt: 6,
      })
    );
  });
});

describe('content-control values in suggesting mode', () => {
  test('a dropdown pick', () => {
    const target = suggesting(DROPDOWN);
    recorded(target, () => target.surface.contentControls.setValue(controlId(target.surface), '2'));
  });

  test('a checkbox toggle through its widget', () => {
    const target = suggesting(CHECKBOX);
    target.surface.contentControls.setShowAll(true);
    const widget = target.container.querySelector('[data-docx-cc-widget="checkbox"]')!;
    recorded(target, () =>
      widget.dispatchEvent(
        new PointerEvent('pointerdown', {
          bubbles: true,
          cancelable: true,
          button: 0,
          pointerId: 1,
          pointerType: 'mouse',
        })
      )
    );
  });

  test('a date pick', () => {
    const target = suggesting(DATE);
    recorded(target, () =>
      target.surface.contentControls.setValue(controlId(target.surface), '2026-03-04')
    );
  });

  test('a text value written through the command', () => {
    const target = suggesting(TEXT);
    caret(target.surface, 0, 9);
    recorded(target, () =>
      expect(target.editor.exec({ type: 'setContentControlValue', value: 'Beta' })).toMatchObject({
        ok: true,
      })
    );
  });
});
