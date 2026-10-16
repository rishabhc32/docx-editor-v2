// Suggesting mode records tab stops as a `w:pPrChange`, like every other paragraph property.
//
// `w:tabs` is a `CT_PPrBase` child, so Word keeps the paragraph's old tab stops in the
// `w:pPrChange` a paragraph property change writes. The editor wrote them with a dedicated op that
// had no tracked form, so a suggester's tab stops applied outright: no card, nothing to reject.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { resolveRevisionMarkup } from '../../contracts/revision-markup.ts';
import { revisionItemsOf } from '../../store/store/review-reads.ts';
import { planRevisionBatch } from '../../store/store/revision-batch.ts';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const AUTHOR = 'Bea';

const TABBED =
  '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>' +
  '<w:r><w:t>alpha</w:t></w:r></w:p>';
const STOPS = [{ positionTwips: 2880, alignment: 'right' as const }];

function docx(body: string, settings = ''): Uint8Array {
  const type = (name: string) =>
    `application/vnd.openxmlformats-officedocument.wordprocessingml.${name}+xml`;
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${type('document.main')}"/>` +
        `<Override PartName="/word/settings.xml" ContentType="${type('settings')}"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rIdT" Type="${R}/settings" Target="settings.xml"/></Relationships>`
    ),
    'word/settings.xml': strToU8(`<w:settings xmlns:w="${W}">${settings}</w:settings>`),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

const mounted: { surface: PaginatedSurface; container: HTMLElement }[] = [];

afterEach(() => {
  for (const entry of mounted.splice(0)) {
    entry.surface.destroy();
    entry.container.remove();
  }
});

function suggesting(bytes: Uint8Array): PaginatedSurface {
  const container = document.createElement('div');
  document.body.append(container);
  const opened = mountPaginatedSurface(container, bytes, {
    editingMode: 'suggest',
    author: AUTHOR,
    revisionDisplayMode: 'all-markup',
  });
  if (!opened.ok) throw new Error(opened.reason);
  mounted.push({ surface: opened.surface, container });
  const paragraphId = opened.surface.session.paragraphIds()[0]!;
  opened.surface.setSelection({
    anchor: { paragraphId, offset: 2 },
    head: { paragraphId, offset: 2 },
  });
  return opened.surface;
}

/** A paragraph's `w:pPr`, with the attributes the assertions do not care about removed. */
function pPr(surface: PaginatedSurface, index = 0): string {
  const paragraph = serializeOoxmlPart(surface.session.part()).split('</w:p>')[index] ?? '';
  return (/<w:pPr>.*?<\/w:pPr>(?=<w:r>)/s.exec(paragraph)?.[0] ?? '')
    .replace(/ w:date="[^"]*"/g, '')
    .replace(/ w:id="[^"]*"/g, '');
}

/** Select from the first paragraph into the last, as a dialog over several paragraphs is. */
function selectAll(surface: PaginatedSurface): void {
  const ids = surface.session.paragraphIds();
  surface.setSelection({
    anchor: { paragraphId: ids[0]!, offset: 0 },
    head: { paragraphId: ids[ids.length - 1]!, offset: 2 },
  });
}

function resolveAll(surface: PaginatedSurface, action: 'accept' | 'reject'): void {
  const plan = planRevisionBatch(surface.session.part(), action);
  const result = surface.applyAutomationOps(() => [...plan.ops]);
  if (result.rejected) throw new Error(String(result.reason));
}

describe('tab stops in suggesting mode', () => {
  test('are recorded with the stops they replaced, and Reject puts those back', () => {
    const surface = suggesting(docx(TABBED));
    expect(surface.setParagraphFormat({ tabStops: STOPS })).toBe(true);

    expect(pPr(surface)).toContain('<w:tab w:pos="2880" w:val="right"/>');
    expect(pPr(surface)).toContain(
      `<w:pPrChange w:author="${AUTHOR}"><w:pPr><w:tabs><w:tab w:pos="720" w:val="left"/>` +
        '</w:tabs></w:pPr></w:pPrChange>'
    );
    expect(revisionItemsOf(surface.session.part())).toMatchObject([
      { revisionKind: 'format', author: AUTHOR },
    ]);

    resolveAll(surface, 'reject');
    expect(pPr(surface)).toBe('<w:pPr><w:tabs><w:tab w:pos="720" w:val="left"/></w:tabs></w:pPr>');
    expect(revisionItemsOf(surface.session.part())).toEqual([]);
  });

  test('Accept keeps the new stops and drops the record', () => {
    const surface = suggesting(docx(TABBED));
    surface.setParagraphFormat({ tabStops: STOPS });
    expect(pPr(surface)).toContain('pPrChange');
    resolveAll(surface, 'accept');
    expect(pPr(surface)).toContain('<w:tab w:pos="2880" w:val="right"/>');
    expect(pPr(surface)).not.toContain('pPrChange');
  });

  test('one dialog OK over two paragraphs, with spacing, is one suggestion', () => {
    const surface = suggesting(docx(TABBED + TABBED));
    selectAll(surface);
    surface.setParagraphFormat({ tabStops: STOPS, spaceAfterPt: 6 });
    expect(revisionItemsOf(surface.session.part())).toHaveLength(1);

    resolveAll(surface, 'reject');
    for (const index of [0, 1]) {
      expect(pPr(surface, index)).toBe(
        '<w:pPr><w:tabs><w:tab w:pos="720" w:val="left"/></w:tabs></w:pPr>'
      );
    }
  });

  test('Track formatting off and w:doNotTrackFormatting apply them untracked', () => {
    const surface = suggesting(docx(TABBED + TABBED));
    surface.setParagraphFormat({ tabStops: STOPS });
    expect(pPr(surface, 0)).toContain('pPrChange');

    surface.setRevisionMarkup(resolveRevisionMarkup({ trackFormatting: false }));
    const second = surface.session.paragraphIds()[1]!;
    surface.setSelection({
      anchor: { paragraphId: second, offset: 2 },
      head: { paragraphId: second, offset: 2 },
    });
    surface.setParagraphFormat({ tabStops: STOPS });
    expect(pPr(surface, 1)).toContain('<w:tab w:pos="2880" w:val="right"/>');
    expect(pPr(surface, 1)).not.toContain('pPrChange');

    const untracked = suggesting(docx(TABBED, '<w:doNotTrackFormatting/>'));
    untracked.setParagraphFormat({ tabStops: STOPS });
    expect(pPr(untracked)).toContain('<w:tab w:pos="2880" w:val="right"/>');
    expect(pPr(untracked)).not.toContain('pPrChange');
  });

  test('a paragraph the author proposed adding records nothing more', () => {
    const own =
      `<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="${AUTHOR}"/></w:rPr></w:pPr>` +
      '<w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p>';
    const surface = suggesting(docx(own));
    selectAll(surface);
    surface.setParagraphFormat({ tabStops: STOPS });
    expect(pPr(surface, 0)).toContain('<w:tab w:pos="2880" w:val="right"/>');
    expect(pPr(surface, 0)).not.toContain('pPrChange');
    // The paragraph after it was there before the suggestion, so its stops are a proposal.
    expect(pPr(surface, 1)).toContain(`<w:pPrChange w:author="${AUTHOR}"><w:pPr/></w:pPrChange>`);
  });
});
