/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `reviewModule({ untrackedEdits: 'refuse' })`: suggesting refuses a local edit it cannot record
// as a suggestion, whole, and keeps every edit it can record. Page setup is the example of an edit
// with no tracked form; the default keeps applying it directly.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { diagnosticCodeOfRefusal } from '@docx-editor.dev/core';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { contentControlsIn, serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { reviewModule, type ReviewModuleOptions } from '../index';
import { REL, W, zipDocument } from '../collaboration/__tests__/document-peer-support';

const SECTION =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
const BODY =
  '<w:p><w:r><w:t>First paragraph</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p>' +
  SECTION;

/** Two list items separated by Word's blank paragraph, which Enter continues. */
const item = (text: string) =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const SEPARATED_LIST = `${item('Item one')}<w:p/>${item('Item two')}${SECTION}`;
const NUMBERING_PART =
  `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">` +
  '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
const withNumbering = (body: string) =>
  zipDocument(body, {
    overrides:
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
    documentRels:
      `<Relationships xmlns="${REL}"><Relationship Id="rIdN" ` +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>',
    extraXml: { 'word/numbering.xml': NUMBERING_PART },
  });

/** BODY with a styles part, so a paragraph style has something to apply. */
const withHeadingStyle = () =>
  zipDocument(BODY, {
    overrides:
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    documentRels:
      `<Relationships xmlns="${REL}"><Relationship Id="rIdS" ` +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    extraXml: {
      'word/styles.xml':
        `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Heading1">` +
        '<w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>',
    },
  });

/** A drop-down list control, the shape a contract's field picker takes. */
const CHOICE =
  '<w:p><w:sdt><w:sdtPr><w:alias w:val="Choice"/><w:tag w:val="choice"/><w:dropDownList>' +
  '<w:listItem w:displayText="Yes" w:value="Yes"/><w:listItem w:displayText="No" w:value="No"/>' +
  '</w:dropDownList></w:sdtPr><w:sdtContent><w:r><w:t>Yes</w:t></w:r></w:sdtContent></w:sdt></w:p>' +
  SECTION;

type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
});

function mount(
  options: {
    review?: ReviewModuleOptions;
    mode?: 'edit' | 'suggesting';
    trackFormatting?: boolean;
    document?: Uint8Array;
  } = {}
): Editor {
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  const editor = createDocxEditor({
    container,
    document: options.document ?? zipDocument(BODY),
    author: 'Reviewer',
    mode: options.mode ?? 'suggesting',
    ...(options.trackFormatting === undefined
      ? {}
      : { revisionMarkup: { trackFormatting: options.trackFormatting } }),
    modules: [reviewModule(options.review ?? { untrackedEdits: 'refuse' })],
  });
  editors.push(editor);
  return editor;
}

function select(editor: Editor, paragraph: number, start: number, end = start) {
  const paragraphId = editor.surface!.session.paragraphIds()[paragraph]!;
  editor.surface!.setSelection({
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  });
}

const body = (editor: Editor) => serializeOoxmlPart(editor.surface!.session.part());
const leftMargin = (editor: Editor) => /<w:pgMar[^>]*w:left="(\d+)"/.exec(body(editor))?.[1];
const accepted = (editor: Editor) => editor.snapshot().lastRejection ?? null;
const controlId = (editor: Editor) =>
  contentControlsIn(editor.surface!.session.part().root)[0]!.node.id;
const REASON = 'suggesting cannot record this change as a suggestion';
/** Ask for a left margin through the toolbar's command; true when it was applied. */
const setMargin = (editor: Editor, twips: number) =>
  editor.exec({ type: 'setPageSetup', marginLeft: twips }).ok;

describe("untrackedEdits: 'refuse' while suggesting", () => {
  test('a page-setup change is refused, with the reason and its diagnostic code', () => {
    const editor = mount();
    const before = body(editor);
    const result = editor.exec({ type: 'setPageSetup', marginLeft: 720 });
    expect(result.ok).toBe(false);
    expect(accepted(editor)).toBe(REASON);
    expect(!result.ok && result.reason).toBe(REASON);
    expect(diagnosticCodeOfRefusal(REASON)).toBe('edit.tracked-structure-unsupported');
    expect(editor.exec({ type: 'setPageSetup', orientation: 'landscape' }).ok).toBe(false);
    expect(body(editor)).toBe(before);
    expect(editor.getTrackedChanges()).toEqual([]);
  });

  test('tab stops pressed with spacing land as one suggestion', () => {
    const editor = mount();
    select(editor, 1, 2);
    const result = editor.exec({
      type: 'setParagraphFormat',
      spaceBeforePt: 12,
      tabStops: [{ positionTwips: 720, alignment: 'left' }],
    });
    expect(result.ok).toBe(true);
    expect(body(editor)).toMatch(/<w:pPrChange[^>]*>.*<\/w:pPrChange>/);
    expect(editor.getTrackedChanges()).toHaveLength(1);
  });

  test('edits it can record still land as suggestions beside a refused one', () => {
    const editor = mount();
    expect(setMargin(editor, 720)).toBe(false);
    select(editor, 0, 5);
    editor.surface!.type(' new');
    select(editor, 1, 0, 6);
    expect(editor.exec({ type: 'toggleMark', mark: 'bold' }).ok).toBe(true);
    expect(editor.exec({ type: 'setAlignment', align: 'center' }).ok).toBe(true);
    expect(editor.addComment('Check this').ok).toBe(true);
    // Each rides on a record the same batch writes: the inserted mark, the tracked link text.
    select(editor, 1, 6);
    editor.surface!.splitParagraph();
    select(editor, 0, 3);
    expect(editor.exec({ type: 'insertBreak', kind: 'section' }).ok).toBe(true);
    select(editor, 0, 1);
    const link = { type: 'insertHyperlink', href: 'https://example.com', text: 'see' } as const;
    expect(editor.exec(link as never).ok).toBe(true);
    expect(accepted(editor)).toBeNull();
    expect(body(editor)).toContain('w:hyperlink');
    expect(editor.surface!.session.bodyText()).toContain('new');
    expect(editor.getTrackedChanges().length).toBeGreaterThanOrEqual(5);
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
    expect(editor.getTrackedChanges()).toEqual([]);
    expect(leftMargin(editor)).toBe('1440');
  });

  test('Enter continues a blank-separated list as a suggestion', () => {
    const editor = mount({ document: withNumbering(SEPARATED_LIST) });
    expect(setMargin(editor, 720)).toBe(false);
    select(editor, 2, 'Item two'.length);
    editor.surface!.splitParagraph();
    expect(accepted(editor)).toBeNull();
    // A blank separator and a new item: two proposed paragraph marks.
    expect(editor.surface!.session.paragraphIds()).toHaveLength(5);
    expect(body(editor).match(/<w:ins /g)?.length).toBe(2);
  });

  test('a content-control pick is recorded, so the widget write lands as a suggestion', () => {
    const editor = mount({ document: zipDocument(CHOICE) });
    const id = controlId(editor);
    expect(editor.surface!.contentControls.disabledReason(id, 'edit')).toBeNull();
    expect(editor.surface!.contentControls.setValue(id, 'No')).toBe(true);
    expect(body(editor)).toMatch(/<w:del .*Yes.*<w:ins .*No/);
  });

  test('bookmarks are inserted, removed and deduplicated: they are not revisions', () => {
    const copies =
      '<w:p><w:bookmarkStart w:id="1" w:name="Same"/><w:r><w:t>one</w:t></w:r><w:bookmarkEnd w:id="1"/>' +
      '<w:bookmarkStart w:id="2" w:name="SAME"/><w:r><w:t>two</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>' +
      SECTION;
    const editor = mount({ document: zipDocument(copies) });
    const names = () => editor.query({ type: 'bookmarks' }).map(({ name }) => name);
    const run = (command: Parameters<Editor['exec']>[0]) => {
      expect(editor.can(command).ok).toBe(true);
      const result = editor.exec(command);
      expect(result.ok ? null : result.reason).toBeNull();
    };
    select(editor, 0, 0, 3);
    run({ type: 'insertBookmark', name: 'Intro' });
    expect(names()).toEqual(['Intro', 'Same']);
    run({ type: 'deleteBookmark', name: 'Intro' });
    run({ type: 'keepFirstBookmark', name: 'Same' });
    expect(editor.query({ type: 'bookmarks' })).toMatchObject([{ name: 'Same', duplicates: 0 }]);
    expect(editor.getTrackedChanges()).toEqual([]);
  });

  test("a link's new text is a suggestion; its own target is no edit, a new one is refused", () => {
    const linked = zipDocument(
      '<w:p><w:r><w:t xml:space="preserve">see </w:t></w:r>' +
        `<w:hyperlink r:id="rIdL"><w:r><w:t>site</w:t></w:r></w:hyperlink></w:p>${SECTION}`,
      {
        documentRels:
          `<Relationships xmlns="${REL}"><Relationship Id="rIdL" ` +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ' +
          'Target="https://example.com/" TargetMode="External"/></Relationships>',
      }
    );
    const editor = mount({ document: linked });
    select(editor, 0, 6);
    // The link lane names the link's own target beside the tracked strike and insertion.
    const rename = { type: 'insertHyperlink', href: 'https://example.com/', text: 'page' } as const;
    expect(editor.exec(rename as never).ok).toBe(true);
    expect(accepted(editor)).toBeNull();
    expect(editor.getTrackedChanges()).not.toEqual([]);
    const renamed = body(editor);
    // A renamed link is struck text and a copy now, so its target cannot change in place.
    expect(editor.exec({ ...rename, href: 'https://other.example/' } as never).ok).toBe(false);
    expect(accepted(editor)).not.toBeNull();
    expect(body(editor)).toBe(renamed);
  });

  test('a link proposed a new target or removed lands as one suggestion each', () => {
    const linked = zipDocument(
      '<w:p><w:r><w:t xml:space="preserve">see </w:t></w:r>' +
        `<w:hyperlink r:id="rIdL"><w:r><w:t>site</w:t></w:r></w:hyperlink></w:p>${SECTION}`,
      {
        documentRels:
          `<Relationships xmlns="${REL}"><Relationship Id="rIdL" ` +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ' +
          'Target="https://example.com/" TargetMode="External"/></Relationships>',
      }
    );
    for (const remove of [false, true]) {
      const editor = mount({ document: linked });
      const before = body(editor);
      select(editor, 0, 6);
      const result = remove
        ? editor.exec({ type: 'removeHyperlink' })
        : editor.exec({ type: 'insertHyperlink', href: 'https://other.example/' } as never);
      expect(result.ok ? null : result.reason).toBeNull();
      expect(editor.getTrackedChanges().length).toBeGreaterThan(0);
      expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
      expect(body(editor)).toBe(before);
    }
  });

  test("removing the author's own pending link is part of that suggestion", () => {
    const editor = mount();
    select(editor, 0, 0, 'First'.length);
    expect(editor.exec({ type: 'insertHyperlink', href: 'https://a.example/' } as never).ok).toBe(
      true
    );
    // In the linked copy, after the struck words.
    select(editor, 0, 'FirstFi'.length);
    const unlink = { type: 'removeHyperlink' } as const;
    expect(editor.can(unlink).ok).toBe(true);
    const result = editor.exec(unlink);
    expect(result.ok ? null : result.reason).toBeNull();
    expect(body(editor)).not.toMatch(/w:hyperlink|w:rStyle/);
    expect(editor.getTrackedChanges().every((change) => change.author === 'Reviewer')).toBe(true);
    // Rejecting gives the words back as they were.
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
    expect(editor.surface!.session.bodyText()).toContain('First paragraph');
    expect(editor.getTrackedChanges()).toEqual([]);
  });

  test("a new target for the author's own pending link is part of that suggestion", async () => {
    const editor = mount();
    select(editor, 0, 5);
    const link = { type: 'insertHyperlink', href: 'https://a.example/', text: 'link' } as const;
    expect(editor.exec(link as never).ok).toBe(true);
    select(editor, 0, 7);
    const reaim = { type: 'insertHyperlink', href: 'https://b.example/' } as const;
    expect(editor.exec(reaim as never).ok).toBe(true);
    expect(accepted(editor)).toBeNull();
    expect(editor.getTrackedChanges()).toMatchObject([{ kind: 'insert', author: 'Reviewer' }]);
    const files = unzipSync(new Uint8Array(await editor.save()));
    const linked = /<w:hyperlink [^>]*r:id="([^"]+)"/.exec(strFromU8(files['word/document.xml']!));
    // Hyperlink relationships only: a first link can also add one for the styles part.
    const relationships = strFromU8(files['word/_rels/document.xml.rels']!).match(
      /<Relationship [^>]*>/g
    );
    const targets = new Map(
      (relationships ?? [])
        .filter((element) => /Type="[^"]*\/hyperlink"/.test(element))
        .map((element) => [
          /Id="([^"]+)"/.exec(element)?.[1],
          /Target="([^"]+)"/.exec(element)?.[1],
        ])
    );
    expect(targets.get(linked?.[1] ?? '')).toBe('https://b.example/');
    // The first target stays behind unreferenced, as every re-aim leaves it; nothing else does.
    expect([...targets.values()].sort()).toEqual(['https://a.example/', 'https://b.example/']);
    // Rejecting the insertion takes the link, and its new target, with it.
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
    expect(body(editor)).not.toContain('w:hyperlink');
  });

  test('a save keeps the cached REF results instead of writing new ones', async () => {
    // Fresh on load, so the field goes live; typing at the bookmark's start makes it stale.
    const fields = zipDocument(
      '<w:p><w:bookmarkStart w:id="1" w:name="term"/><w:r><w:t>Closing Date</w:t></w:r>' +
        '<w:bookmarkEnd w:id="1"/></w:p>' +
        '<w:p><w:fldSimple w:instr=" REF term "><w:r><w:t>Closing Date</w:t></w:r></w:fldSimple></w:p>' +
        SECTION
    );
    const savedResult = async (editor: Editor) => {
      const saved = strFromU8(unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!);
      return /<w:fldSimple[^>]*>(.*?)<\/w:fldSimple>/.exec(saved)?.[1]!.replace(/<[^>]+>/g, '');
    };
    for (const [review, refreshed] of [
      [{ untrackedEdits: 'refuse' }, false],
      [{}, true],
    ] as const) {
      const editor = mount({ review, mode: 'edit', document: fields });
      select(editor, 0, 0);
      editor.surface!.type('The ');
      expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
      const revision = editor.getDocumentHandle().revision;
      expect(await savedResult(editor)).toBe(refreshed ? 'The Closing Date' : 'Closing Date');
      expect(editor.getDocumentHandle().revision === revision).toBe(!refreshed);
    }
  });

  test('formatting left untracked on purpose is applied while page setup is still refused', () => {
    const editor = mount({ trackFormatting: false });
    expect(setMargin(editor, 720)).toBe(false);
    select(editor, 1, 0, 6);
    expect(editor.exec({ type: 'toggleMark', mark: 'bold' }).ok).toBe(true);
    expect(accepted(editor)).toBeNull();
    expect(body(editor)).toContain('<w:b/>');
    expect(editor.getTrackedChanges()).toEqual([]);
  });

  test('undo and redo of an earlier direct edit are not refused', () => {
    const editor = mount({ mode: 'edit' });
    expect(setMargin(editor, 720)).toBe(true);
    expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
    expect(setMargin(editor, 900)).toBe(false);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(leftMargin(editor)).toBe('1440');
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(leftMargin(editor)).toBe('720');
  });

  test('the editing API lane keeps its own tracked-change policy', () => {
    // Its host refuses operations with no tracked form while suggesting, before this lane.
    const editor = mount();
    expect(setMargin(editor, 900)).toBe(false);
    const result = editor.surface!.applyAutomationOps(() => [
      { op: 'setSectionProperties', marginLeftTwips: 720 },
    ]);
    expect(result.committed).toBe(true);
    expect(leftMargin(editor)).toBe('720');
  });
});

describe("everyday edits under untrackedEdits: 'refuse'", () => {
  test('paste, deletes across paragraphs, styles, tables and rows land as suggestions', () => {
    const editor = mount({ document: withHeadingStyle() });
    const before = editor.surface!.session.bodyText();
    const run = (command: Parameters<Editor['exec']>[0]) => {
      const result = editor.exec(command);
      // The reason, not a bare false, so a refusal names itself in the failure.
      expect(result.ok ? null : result.reason).toBeNull();
    };
    select(editor, 0, 5);
    run({ type: 'paste', text: 'one\ntwo' });
    run({ type: 'paste', text: 'x', html: '<p><b>bold</b></p><p>next</p>' });
    const ids = editor.surface!.session.paragraphIds();
    editor.surface!.setSelection({
      anchor: { paragraphId: ids[0]!, offset: 2 },
      head: { paragraphId: ids[1]!, offset: 2 },
    });
    run({ type: 'deleteText' });
    select(editor, 1, 0);
    editor.surface!.deleteBackward();
    expect(accepted(editor)).toBeNull();
    run({ type: 'setParagraphStyle', styleId: 'Heading1' });
    run({ type: 'insertTable', rows: 2, cols: 2 });
    editor.surface!.type('cell');
    expect(accepted(editor)).toBeNull();
    run({ type: 'insertRow', where: 'below' });
    expect(editor.getTrackedChanges().length).toBeGreaterThan(0);
    run({ type: 'resolveAllReviewChanges', action: 'reject' });
    expect(editor.surface!.session.bodyText()).toBe(before);
  });
});

describe('only refuse, and only while suggesting', () => {
  test('the default and apply keep applying the change directly', () => {
    const policies: [ReviewModuleOptions, boolean][] = [
      [{ untrackedEdits: 'refuse' }, false],
      [{}, true],
      [{ untrackedEdits: 'apply' }, true],
    ];
    for (const [review, applies] of policies) {
      const editor = mount({ review });
      expect(setMargin(editor, 720)).toBe(applies);
      expect(leftMargin(editor)).toBe(applies ? '720' : '1440');
    }
  });

  test('refuse applies the change once the editor leaves suggesting', () => {
    const editor = mount();
    expect(setMargin(editor, 720)).toBe(false);
    expect(editor.exec({ type: 'setEditingMode', mode: 'editing' }).ok).toBe(true);
    expect(setMargin(editor, 720)).toBe(true);
    expect(leftMargin(editor)).toBe('720');
  });

  test('an unknown value throws where the host wrote it', () => {
    expect(() => reviewModule({ untrackedEdits: 'warn' as never })).toThrow(TypeError);
  });
});
