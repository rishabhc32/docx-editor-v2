// The change kinds of group 9.10 applied as the room applies them: a story replacement, a
// paragraph removal and a setting are direct edits whose revert shows the document exactly as
// before; a refused change writes nothing.

import { describe, expect, test } from 'bun:test';
import { orderedIdMinter } from '../change/ids.ts';
import { applyOrderedChange } from '../change/apply.ts';
import type { OrderedChange } from '../change/types.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import {
  apply,
  at,
  BODY,
  DATE,
  DIRECT,
  EDIT,
  FOOTNOTES,
  load,
  markOf,
  partsOf,
  refusal,
  textOf,
} from './change-apply-fixture.ts';
import { testAttribution, testPerson } from './test-persons.ts';
import { docx, MATRIX_EXTRAS, p, t } from './word-fixtures.ts';

/** Take back or make again a direct edit's suggestions, as the organiser's undo and redo do. */
function revert(
  doc: ReturnType<typeof load>,
  id: string,
  suggestions: readonly string[],
  status: 'accepted' | 'rejected'
): void {
  const change = { v: 1, id, kind: 'revert', suggestions, status } as OrderedChange;
  withUndoRecord(doc, () =>
    applyOrderedChange(doc, change, {
      person: testPerson('Olive'),
      date: DATE,
      ids: orderedIdMinter(id),
    })
  );
}

const texts = (doc: ReturnType<typeof load>): string[] => {
  const out: string[] = [];
  for (let k = 0; k < 20; k += 1) {
    const text = textOf(doc, k);
    if (text === '<none>') break;
    out.push(text);
  }
  return out;
};

describe('replace-story', () => {
  test('puts plain paragraphs in place of the body; revert shows the body again exactly', () => {
    const doc = load();
    const before = partsOf(doc);
    const outcome = apply(
      doc,
      'olive-1:1',
      { ...DIRECT, kind: 'replace-story', paragraphs: ['New one.', '', 'New three.'] },
      'Olive'
    );
    expect(texts(doc)).toEqual(['New one.', '', 'New three.']);
    const body = partsOf(doc)[0] ?? '';
    expect(body).not.toContain('<w:tbl>');
    expect(body).not.toContain('<w:sdt>');
    expect(body).not.toContain('<w:ins ');
    revert(doc, 'olive-1:2', outcome.suggestions, 'rejected');
    expect(partsOf(doc)).toEqual(before);
    revert(doc, 'olive-1:3', outcome.suggestions, 'accepted');
    expect(texts(doc)).toEqual(['New one.', '', 'New three.']);
  });

  test('is refused as a suggestion, and writes nothing', () => {
    const doc = load();
    expect(refusal(doc, 'ana-1:1', { ...EDIT, kind: 'replace-story', paragraphs: ['x'] })).toBe(
      'invalid-edit'
    );
  });
});

describe('delete-paragraph', () => {
  test('removes a paragraph beside another; revert shows it again', () => {
    const doc = load();
    const before = partsOf(doc);
    const all = texts(doc);
    const index = all.indexOf('Due in 30 days.');
    const outcome = apply(
      doc,
      'olive-1:1',
      { ...DIRECT, kind: 'delete-paragraph', mark: markOf(doc, index) },
      'Olive'
    );
    expect(texts(doc)).toEqual(all.filter((text) => text !== 'Due in 30 days.'));
    revert(doc, 'olive-1:2', outcome.suggestions, 'rejected');
    expect(partsOf(doc)).toEqual(before);
  });

  test('removes the paragraph right before a table, and the table stays', () => {
    const doc = load();
    const all = texts(doc);
    const index = all.indexOf('First item');
    apply(
      doc,
      'olive-1:1',
      { ...DIRECT, kind: 'delete-paragraph', mark: markOf(doc, index) },
      'Olive'
    );
    expect(texts(doc)).toEqual(all.filter((text) => text !== 'First item'));
    expect(partsOf(doc)[0]).toContain('<w:tbl>');
  });

  // The engine's `deleteBlock` keeps the paragraph a note, a cell or a section's end needs
  test('is refused where the engine keeps the paragraph', () => {
    const cell = '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>';
    const file = docx(
      p(t('Ends a section.'), '<w:sectPr/>') + `${cell}${p(t('Cell.'))}</w:tc></w:tr></w:tbl>`,
      MATRIX_EXTRAS
    );
    const refused = (
      [
        [BODY, 0],
        [BODY, 1],
        [FOOTNOTES, 1],
        [FOOTNOTES, 0],
      ] as const
    ).map(([part, last]) => {
      const doc = openOrderedWordFile(file, { file: 1 }).document;
      apply(
        doc,
        'pia-1:1',
        { ...EDIT, kind: 'insert-note', place: markOf(doc, 0), note: 'footnote', text: ' Note' },
        'Pia'
      );

      // A note's only paragraph: the file's, and Pia's new one at the story's end
      const notes = doc.stories.get(FOOTNOTES)!.tree.summary.marks;
      const index = part === BODY ? last : notes - 1 - last;
      const change = {
        part,
        mode: 'edit',
        kind: 'delete-paragraph',
        mark: markOf(doc, index, part),
      };
      return refusal(doc, 'olive-1:1', change as never, 'Olive');
    });
    expect(refused).toEqual(Array(4).fill('engine-refused'));
  });

  // The engine removes the whole paragraph, its markers and empty runs too: at a cell's end
  // nothing is left after the cell's last paragraph
  test("removes a cell's last paragraph with the markers at its end", () => {
    const end = '<w:bookmarkStart w:id="1" w:name="b"/><w:bookmarkEnd w:id="1"/><w:r><w:t/></w:r>';
    const doc = openOrderedWordFile(
      docx(
        '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>' +
          `${p(t('First.'))}${p(t('Second.') + end)}</w:tc></w:tr></w:tbl>${p(t('After.'))}`
      ),
      { file: 1 }
    ).document;
    apply(doc, 'olive-1:1', { ...DIRECT, kind: 'delete-paragraph', mark: markOf(doc, 1) }, 'Olive');
    const body = partsOf(doc)[0] ?? '';
    expect(body).toContain('<w:t>First.</w:t></w:r></w:p></w:tc>');
    expect(body).not.toContain('<w:t/>');
    expect(body).not.toContain('w:bookmark');
  });

  test('is refused as a suggestion: Word tracks it as struck text and mark', () => {
    const doc = load();
    expect(
      refusal(doc, 'ana-1:1', { ...EDIT, kind: 'delete-paragraph', mark: markOf(doc, 0) })
    ).toBe('invalid-edit');
  });
});

describe('setting', () => {
  test('is a direct edit; as a suggestion it is refused', () => {
    const doc = load();
    const setting = { op: 'setDocumentProtection', enforce: true } as const;
    expect(refusal(doc, 'ana-1:1', { ...EDIT, kind: 'setting', setting })).toBe('invalid-edit');
    const outcome = apply(doc, 'olive-1:1', { ...DIRECT, kind: 'setting', setting }, 'Olive');
    expect(outcome.suggestions).toEqual(['S:olive-1:1.1']);
    expect(doc.suggestions.get('S:olive-1:1.1')?.status).toBe('accepted');
    expect(doc.suggestions.get('S:olive-1:1.1')?.direct).toBe(true);
    expect([...doc.resources.added.values()].map((r) => r.kind)).toEqual(['setting']);
    void BODY;
  });
});

describe('attribution', () => {
  test('changes the date on every mark of a suggestion, and keeps its author and ids', () => {
    const doc = load();
    const typed = apply(doc, 'ana-1:1', {
      ...EDIT,
      kind: 'insert',
      place: markOf(doc, 0),
      text: ' now',
    });
    const sid = typed.suggestions[0] ?? '';
    const marks = (xml: string) =>
      [...xml.matchAll(/<w:ins ([^>]*)>/g)].map((m) => {
        const read = (name: string) => new RegExp(`w:${name}="([^"]*)"`).exec(m[1] ?? '')?.[1];
        return { id: read('id'), author: read('author'), date: read('date') };
      });
    const before = marks(partsOf(doc)[0] ?? '');
    expect(before).toHaveLength(1);
    apply(
      doc,
      'olive-1:1',
      {
        ...DIRECT,
        kind: 'attribution',
        suggestions: [sid],
        author: 'Ana',
        date: '2026-10-29T10:00:00Z',
      },
      'Olive'
    );
    expect(marks(partsOf(doc)[0] ?? '')).toEqual([
      { id: before[0]?.id, author: 'Ana', date: '2026-10-29T10:00:00Z' },
    ]);
    expect(doc.suggestions.get(sid)?.attribution.person).toEqual(testPerson('Ana'));
    expect(doc.suggestions.get(sid)?.status).toBe('pending');
  });

  // A room item's author comes only from its person: no sender renames it, the organiser too
  test('is refused by name when it names another author, and writes nothing', () => {
    const doc = load();
    const typed = apply(doc, 'ana-1:1', {
      ...EDIT,
      kind: 'insert',
      place: markOf(doc, 0),
      text: ' now',
    });
    const sid = typed.suggestions[0] ?? '';
    const before = partsOf(doc)[0];
    const renamed = {
      ...DIRECT,
      kind: 'attribution' as const,
      suggestions: [sid],
      author: 'Ana Legal',
    };
    expect(refusal(doc, 'olive-1:1', renamed, 'Olive')).toBe('fixed-author');
    expect(refusal(doc, 'ana-1:2', { ...renamed, mode: 'suggest' }, 'Ana')).toBe('fixed-author');
    expect(partsOf(doc)[0]).toBe(before);
    expect(doc.suggestions.get(sid)?.attribution).toEqual(testAttribution('Ana', DATE));
  });

  test('is refused on a decided suggestion, and writes nothing', () => {
    const doc = load();
    const typed = apply(doc, 'ana-1:1', {
      ...EDIT,
      kind: 'insert',
      place: markOf(doc, 0),
      text: ' now',
    });
    const sid = typed.suggestions[0] ?? '';
    const suggestion = doc.suggestions.get(sid);
    if (!suggestion) throw new Error('no suggestion');
    apply(
      doc,
      'olive-1:1',
      {
        kind: 'decide',
        cards: [{ suggestions: [{ suggestion: sid, version: suggestion.version }] }],
        status: 'accepted',
      },
      'Olive'
    );
    expect(
      refusal(
        doc,
        'olive-1:2',
        { ...DIRECT, kind: 'attribution', suggestions: [sid], author: 'X' },
        'Olive'
      )
    ).toBe('not-pending');
  });
});

describe('insert-fragment', () => {
  const run = (text: string) => ({ e: 'w:r', c: [{ e: 'w:t', c: [text] }] });
  const CARA = { 'w:id': '1', 'w:author': 'Cara', 'w:date': '2026-01-02T10:00:00Z' };
  const STRUCK = { e: 'w:r', c: [{ e: 'w:delText', c: ['old'] }] };

  // Design D18: the client sends a paste's accepted final view; the server refuses a pasted
  // tracked change, which would bring another person's name and date into the room
  test.each([
    ['an insertion', { e: 'w:p', c: [{ e: 'w:ins', a: CARA, c: [run('x')] }] }],
    ['a deletion', { e: 'w:p', c: [run('x'), { e: 'w:del', a: CARA, c: [STRUCK] }] }],
    [
      'a format change',
      {
        e: 'w:p',
        c: [
          {
            e: 'w:r',
            c: [
              { e: 'w:rPr', c: [{ e: 'w:b' }, { e: 'w:rPrChange', a: CARA, c: [{ e: 'w:rPr' }] }] },
              { e: 'w:t', c: ['x'] },
            ],
          },
        ],
      },
    ],
    [
      'an inserted paragraph mark',
      {
        e: 'w:p',
        c: [{ e: 'w:pPr', c: [{ e: 'w:rPr', c: [{ e: 'w:ins', a: CARA }] }] }, run('x')],
      },
    ],
  ])(
    'holding %s is refused, by the organiser and by a party, and writes nothing',
    (_name, block) => {
      for (const [who, edit] of [
        ['Olive', DIRECT],
        ['Ana', EDIT],
      ] as const) {
        const doc = load();
        const before = partsOf(doc);
        const body = {
          ...edit,
          kind: 'insert-fragment' as const,
          place: at(doc, 0, 2),
          blocks: [block],
        };
        expect(refusal(doc, `${who.toLowerCase()}-1:1`, body, who)).toBe('invalid-edit');
        expect(partsOf(doc)).toEqual(before);
      }
    }
  );

  test('without a tracked change it lands, for the organiser and for a party', () => {
    for (const [who, edit] of [
      ['Olive', DIRECT],
      ['Ana', EDIT],
    ] as const) {
      const doc = load();
      const body = {
        ...edit,
        kind: 'insert-fragment' as const,
        place: at(doc, 0, 2),
        blocks: [{ e: 'w:p', c: [run('x'), { e: 'w:r', c: [{ e: 'w:rPr', c: [{ e: 'w:b' }] }] }] }],
      };
      expect(apply(doc, `${who.toLowerCase()}-1:1`, body, who).suggestions.length).toBe(1);
      expect(textOf(doc, 0).slice(0, 3)).toBe(`${textOf(load(), 0).slice(0, 2)}x`);
    }
  });

  // Word turns a link into a HYPERLINK field across paragraphs pasted inside it (paste
  // evidence 2, `~$paste-out-e2-link-2p.docx`); not written yet, so refused by name
  test('paragraphs inside a link are refused by name, and one line lands in the link', () => {
    const doc = load();
    const lines = [
      { e: 'w:p', c: [run('x')] },
      { e: 'w:p', c: [run('y')] },
    ];
    const inLink = { ...EDIT, kind: 'insert-fragment' as const, place: at(doc, 1, 7) };
    expect(refusal(doc, 'ana-1:1', { ...inLink, blocks: lines })).toBe('paragraphs-in-link');
    expect(refusal(doc, 'ana-1:1', { ...inLink, blocks: [lines[0]!], whole: true })).toBe(
      'paragraphs-in-link'
    );
    apply(doc, 'ana-1:2', { ...inLink, blocks: [lines[0]!] }, 'Ana');
    expect(textOf(doc, 1)).toBe('See thex terms.');
  });

  // An embedded object's part and an ink part travel as media by hash, never opened; the room
  // gives each only its own content types, whatever a client sends
  test('an embedded object and ink land with their parts; media for another use is refused', () => {
    const media = (sha: string, contentType: string) => ({
      sha256: sha.repeat(64),
      contentType,
      size: 10,
    });
    const OLE = media('a', 'application/vnd.openxmlformats-officedocument.oleObject');
    const DOCX = media(
      'b',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    );
    const INK = media('c', 'application/inkml+xml');
    const PNG = media('d', 'image/png');
    const object = (k: number) => ({
      e: 'w:r',
      c: [{ e: 'w:object', c: [{ e: 'o:OLEObject', a: { Type: 'Embed', 'r:id': `#${k}` } }] }],
    });
    const ink = (k: number) => ({
      e: 'w:r',
      c: [{ e: 'w14:contentPart', a: { 'r:id': `#${k}` } }],
    });
    const picture = (k: number) => ({
      e: 'w:r',
      c: [
        { e: 'w:pict', c: [{ e: 'v:shape', c: [{ e: 'v:imagedata', a: { 'r:id': `#${k}` } }] }] },
      ],
    });
    // A change as a client may send it: the reader checks its shape, the apply its media's use
    const paste = (
      doc: ReturnType<typeof load>,
      edit: typeof DIRECT | typeof EDIT,
      runs: readonly object[],
      list: readonly object[]
    ) =>
      ({
        ...edit,
        kind: 'insert-fragment',
        place: at(doc, 0, 2),
        blocks: [{ e: 'w:p', c: runs }],
        media: list,
      }) as unknown as Parameters<typeof apply>[2];
    for (const [who, edit] of [
      ['Olive', DIRECT],
      ['Ana', EDIT],
    ] as const) {
      const doc = load();
      apply(
        doc,
        `${who.toLowerCase()}-1:1`,
        paste(doc, edit, [object(0), object(1), ink(2)], [OLE, DOCX, INK]),
        who
      );
      const added = [...doc.resources.added.values()];
      expect(
        added.flatMap((r) => (r.kind === 'media' ? [r.part.replace(/ordered_[^.]+/, '*')] : []))
      ).toEqual(['/word/embeddings/*.bin', '/word/embeddings/*.docx', '/word/ink/*.xml']);
      expect(added.flatMap((r) => (r.kind === 'relationship' ? [r.type] : []))).toEqual([
        'oleObject',
        'package',
        'ink',
      ]);
    }
    for (const [runs, list] of [
      [[object(0)], [PNG]],
      [[object(0)], [INK]],
      [[ink(0)], [OLE]],
      [[picture(0)], [INK]],
      [[picture(0)], [DOCX]],
    ] as const) {
      const doc = load();
      const before = partsOf(doc);
      expect(refusal(doc, 'ana-1:1', paste(doc, EDIT, [...runs], [...list]))).toBe('invalid-edit');
      expect(partsOf(doc)).toEqual(before);
    }
  });

  test('holding a tracked change in a pasted style is refused', () => {
    const doc = load();
    const body = {
      ...EDIT,
      kind: 'insert-fragment' as const,
      place: at(doc, 0, 2),
      blocks: [{ e: 'w:p', c: [run('x')] }],
      styles: [
        {
          e: 'w:style',
          a: { 'w:type': 'paragraph', 'w:styleId': 'Clause' },
          c: [{ e: 'w:pPr', c: [{ e: 'w:pPrChange', a: CARA, c: [{ e: 'w:pPr' }] }] }],
        },
      ],
    };
    expect(refusal(doc, 'ana-1:1', body)).toBe('invalid-edit');
  });
});
