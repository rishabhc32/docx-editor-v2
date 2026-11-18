// What a pasted fragment may not hold (group 9.10): the change reader refuses each case by its
// code and path, the shared budget holds across calls, and the client's wire refuses an object
// and a namespace the format does not have before a change is made. A picture, a shape, a VML
// picture and an SVG version travel, their images named by index.

import { describe, expect, test } from 'bun:test';
import {
  readOoxmlPart,
  type OoxmlElement,
  type OoxmlNode,
} from '../../../store/package/ooxml-tree.ts';
import {
  fragmentBudget,
  fragmentElements,
  fragmentParts,
  ORDERED_FRAGMENT_LIMITS as LIMITS,
  type OrderedFragmentBudget,
  type OrderedFragmentRoot,
} from '../change/fragment.ts';
import {
  A16_NAMESPACE_URI,
  DRAWINGML_FRAGMENT_ELEMENTS,
  MATH_FRAGMENT_ELEMENTS,
  MC_FRAGMENT_ELEMENTS,
  OFFICE_FRAGMENT_ELEMENTS,
  PICTURE_FRAGMENT_ELEMENTS,
  SVG_BLIP_FRAGMENT_ELEMENTS,
  VML_FRAGMENT_ELEMENTS,
  W10_FRAGMENT_ELEMENTS,
  W14_FRAGMENT_ELEMENTS,
  W15_FRAGMENT_ELEMENTS,
  WML_FRAGMENT_ELEMENTS,
  WP_FRAGMENT_ELEMENTS,
  WPG_FRAGMENT_ELEMENTS,
  WPS_FRAGMENT_ELEMENTS,
} from '../change/fragment-names.ts';
import { attempt } from '../change/read.ts';
import { fragmentWire } from '../client/fragment-wire.ts';
import { OrderedClientRefusal } from '../client/refusal.ts';
import { p, pictureRun, R, t, V, W, WPS } from './word-fixtures.ts';

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

const REFS = { lists: 1, links: 1, notes: 1, media: 1 };

type Tree = { e: string; a?: Record<string, string>; c?: unknown[] };
const el = (e: string, ...c: unknown[]): Tree => (c.length ? { e, c } : { e });
const run = (...c: unknown[]): Tree => el('w:r', ...c);
const para = (...c: unknown[]): Tree => el('w:p', ...c);
const text = (value: string): Tree => el('w:t', value);

function refusal(
  value: unknown,
  root: OrderedFragmentRoot = 'block',
  budget: OrderedFragmentBudget = fragmentBudget(),
  path = 'blocks'
) {
  const result = attempt(() => fragmentElements(value, path, root, REFS, budget));
  if (result.ok) throw new Error('the reader took it');
  return result.error;
}

const accepts = (value: unknown, root: OrderedFragmentRoot = 'block'): void => {
  expect(attempt(() => fragmentElements(value, 'blocks', root, REFS, fragmentBudget())).ok).toBe(
    true
  );
};

const chartPart = (rels?: unknown[], extra: unknown[] = []) => ({
  type: 'chart',
  root: { e: 'c:chartSpace', c: [el('c:chart'), ...extra] },
  ...(rels ? { rels } : {}),
});
const readParts = (value: unknown, references = { ...REFS, parts: 2 }) =>
  attempt(() => fragmentParts(value, 'parts', references, fragmentBudget()));

describe("the fragment reader checks a chart's and a diagram's parts", () => {
  test('a chart with its colors and workbook, and a diagram drawing named by the data part', () => {
    const colors = { type: 'chartColorStyle', root: el('cs:colorStyle', el('cs:variation')) };
    const external = { e: 'c:externalData', a: { 'r:id': '#1' } };
    expect(readParts([chartPart([{ part: 1 }, { media: 0 }], [external]), colors]).ok).toBe(true);
    const data = {
      type: 'diagramData',
      root: el(
        'dgm:dataModel',
        el('dgm:extLst', el('a:ext', { e: 'dsp:dataModelExt', a: { relId: '#1' } }))
      ),
    };
    const drawing = { type: 'diagramDrawing', root: el('dsp:drawing', el('dsp:spTree')) };
    expect(readParts([data, drawing]).ok).toBe(true);
    // An empty relationship names nothing (a layout's `r:blip=""`)
    const layout = {
      type: 'diagramLayout',
      root: el('dgm:layoutDef', { e: 'dgm:shape', a: { 'r:blip': '' } }),
    };
    expect(readParts([layout]).ok).toBe(true);
  });

  test('a wrong root, a loop, a child of a wrong type, and a reference out of range', () => {
    const fails = (value: unknown) => expect(readParts(value).ok).toBe(false);
    fails([{ type: 'chart', root: el('dgm:dataModel') }]);
    fails([{ type: 'gauge', root: el('c:chartSpace') }]);
    // A part names only a later part, of a type its own may name
    fails([chartPart([{ part: 0 }])]);
    fails([chartPart([{ part: 1 }]), { type: 'diagramColors', root: el('dgm:colorsDef') }]);
    fails([chartPart([{ media: 1 }])]);
    fails([chartPart([], [{ e: 'c:externalData', a: { 'r:id': '#0' } }])]);
    fails([chartPart(undefined, [{ e: 'c:userShapes', a: { 'r:id': 'rId3' } }])]);
    fails([
      {
        type: 'diagramData',
        root: el('dgm:dataModel', { e: 'dsp:dataModelExt', a: { relId: '#4' } }),
      },
    ]);
  });
});

describe('the fragment reader refuses', () => {
  test('an element off the allowlist, and a prefix the format does not have', () => {
    expect(refusal([para(run(el('w:frobnicate')))])).toEqual({
      code: 'invalid-value',
      path: 'blocks[0].c[0].c[0].e',
      detail: 'not an element of w',
    });
    expect(refusal([para(el('x:foo'))])).toMatchObject({
      code: 'invalid-namespace',
      path: 'blocks[0].c[0].e',
    });
    expect(refusal([para(el('w14:9effect'))])).toMatchObject({
      code: 'invalid-value',
      path: 'blocks[0].c[0].e',
    });
    expect(refusal([para(el('p'))])).toMatchObject({
      code: 'invalid-value',
      path: 'blocks[0].c[0].e',
    });
  });

  test('a root of the wrong kind', () => {
    expect(refusal([run()])).toMatchObject({ code: 'invalid-value', path: 'blocks[0].e' });
    expect(refusal([para()], 'abstractNum')).toMatchObject({
      code: 'invalid-value',
      path: 'blocks[0].e',
    });
    expect(refusal([el('w:style')], 'lvlOverride')).toMatchObject({ code: 'invalid-value' });
    accepts([el('w:abstractNum')], 'abstractNum');
    accepts([el('w:lvlOverride')], 'lvlOverride');
    accepts([el('w:style')], 'style');
    accepts([para(), el('w:tbl'), el('w:sdt')]);
  });

  test('a namespace declaration as an attribute', () => {
    expect(refusal([{ e: 'w:p', a: { 'xmlns:w': 'urn:evil' } }])).toEqual({
      code: 'invalid-attribute',
      path: 'blocks[0].a.xmlns:w',
      detail: 'a namespace declaration',
    });
    expect(refusal([{ e: 'w:p', a: { xmlns: 'urn:evil' } }])).toMatchObject({
      code: 'invalid-attribute',
      path: 'blocks[0].a.xmlns',
    });
  });

  test('a prototype key in an element and in its attributes', () => {
    expect(refusal(JSON.parse('[{"e":"w:p","__proto__":{"e":"w:r"}}]'))).toEqual({
      code: 'forbidden-key',
      path: 'blocks[0].__proto__',
    });
    expect(refusal(JSON.parse('[{"e":"w:p","a":{"__proto__":"x"}}]'))).toEqual({
      code: 'forbidden-key',
      path: 'blocks[0].a.__proto__',
    });
    expect(refusal(JSON.parse('[{"e":"w:p","c":[{"e":"w:r","constructor":1}]}]'))).toEqual({
      code: 'forbidden-key',
      path: 'blocks[0].c[0].constructor',
    });
  });

  test('a field the shape does not have, a wrong type, an empty list', () => {
    expect(refusal([{ e: 'w:p', x: 1 }])).toEqual({ code: 'unknown-field', path: 'blocks[0].x' });
    expect(refusal({ e: 'w:p' })).toMatchObject({ code: 'invalid-type', path: 'blocks' });
    expect(refusal([])).toEqual({ code: 'empty', path: 'blocks' });
    expect(refusal([{ e: 'w:p', c: [] }])).toEqual({ code: 'empty', path: 'blocks[0].c' });
    expect(refusal([{ e: 'w:p', a: {} }])).toEqual({ code: 'empty', path: 'blocks[0].a' });
    expect(refusal([{ e: 'w:p', a: { 'w:rsidR': 7 } }])).toMatchObject({
      code: 'invalid-type',
      path: 'blocks[0].a.w:rsidR',
    });
    expect(refusal([{ c: [] }])).toEqual({ code: 'missing-field', path: 'blocks[0].e' });
    const getter = [{}];
    Object.defineProperty(getter, 0, { get: () => para(), enumerable: true });
    expect(refusal(getter)).toMatchObject({ code: 'invalid-type', path: 'blocks[0]' });
  });

  test('a text child of an element that holds no text', () => {
    expect(refusal([{ e: 'w:p', c: ['hello'] }])).toEqual({
      code: 'invalid-type',
      path: 'blocks[0].c[0]',
      detail: 'text in w:p',
    });
    expect(refusal([para(run('loose'))])).toMatchObject({
      code: 'invalid-type',
      path: 'blocks[0].c[0].c[0]',
    });
    accepts([para(run(text('a'), el('w:instrText', ' PAGE '), el('w:delText', 'b')))]);
    accepts([para(el('m:oMath', el('m:r', el('m:t', 'x'))))]);
  });

  test('a control character, a tab, a line feed, two texts in a row and an empty text', () => {
    const at = 'blocks[0].c[0].c[0].c[0]';
    expect(refusal([para(run(text('a\u0007b')))])).toEqual({ code: 'invalid-text', path: at });
    expect(refusal([para(run(text('a\tb')))])).toEqual({ code: 'invalid-text', path: at });
    expect(refusal([para(run(text('a\nb')))])).toEqual({ code: 'invalid-text', path: at });
    expect(refusal([para(run(text('a\ud800b')))])).toEqual({ code: 'invalid-text', path: at });
    expect(refusal([para(run(text('')))])).toEqual({ code: 'empty', path: at });
    expect(refusal([para(run(el('w:t', 'a', 'b')))])).toMatchObject({
      code: 'invalid-value',
      path: 'blocks[0].c[0].c[0].c[1]',
    });
    expect(refusal([{ e: 'w:p', a: { 'w:rsidR': 'a\u0085' } }])).toEqual({
      code: 'invalid-text',
      path: 'blocks[0].a.w:rsidR',
    });
  });

  test('xml attributes other than xml:space, and xml:space other than its two values', () => {
    expect(refusal([para(run({ e: 'w:t', a: { 'xml:lang': 'en' } }))])).toMatchObject({
      code: 'invalid-attribute',
      path: 'blocks[0].c[0].c[0].a.xml:lang',
    });
    expect(refusal([para(run({ e: 'w:t', a: { 'xml:space': 'keep' } }))])).toMatchObject({
      code: 'invalid-value',
      path: 'blocks[0].c[0].c[0].a.xml:space',
    });
    expect(refusal([{ e: 'w:p', a: { 'x:ext': 'edit' } }])).toMatchObject({
      code: 'invalid-namespace',
      path: 'blocks[0].a.x:ext',
    });
    expect(refusal([{ e: 'w:p', a: { 'w:rsid-R': '1' } }])).toMatchObject({
      code: 'invalid-attribute',
      path: 'blocks[0].a.w:rsid-R',
    });
  });

  test('a tree deeper than the cap', () => {
    const chain = (depth: number): Tree => {
      let tree = para();
      for (let k = 1; k < depth; k += 1) tree = para(el('w:customXml', ...(tree.c ?? [])));
      return tree;
    };
    accepts([chain(LIMITS.depth)]);
    expect(refusal([chain(LIMITS.depth + 1)])).toEqual({
      code: 'too-many',
      path: `blocks[0]${'.c[0]'.repeat(LIMITS.depth)}`,
      detail: `deeper than ${LIMITS.depth}`,
    });
  });

  test('more nodes than the cap, with the budget shared across two calls', () => {
    const budget = fragmentBudget();
    const runs = (count: number) => para(...Array.from({ length: count }, () => run()));
    expect(
      attempt(() => fragmentElements([runs(10_000)], 'blocks', 'block', REFS, budget)).ok
    ).toBe(true);
    expect(budget.nodes).toBe(10_001);
    expect(refusal([runs(7000)], 'block', budget, 'notes[0].blocks')).toEqual({
      code: 'too-many',
      path: `notes[0].blocks[0].c[${LIMITS.nodes - 10_002}]`,
      detail: `over ${LIMITS.nodes} fragment nodes`,
    });
    accepts([runs(7000)]);
  });

  test('more text than the cap, with the budget shared across two calls', () => {
    const budget = fragmentBudget();
    const long = (units: number) => [{ e: 'w:p', a: { 'w:rsidR': 'x'.repeat(units) } }];
    for (let k = 0; k < LIMITS.text / LIMITS.attributeValue; k += 1)
      expect(
        attempt(() => fragmentElements(long(LIMITS.attributeValue), 'b', 'block', REFS, budget)).ok
      ).toBe(true);
    expect(refusal(long(1), 'block', budget)).toMatchObject({
      code: 'too-long',
      path: 'blocks[0].a.w:rsidR',
    });
  });

  test('an attribute value over the cap, and more attributes than the cap', () => {
    expect(
      refusal([para({ e: 'w:pPr', c: [{ e: 'w:pStyle', a: { 'w:val': 'x'.repeat(2049) } }] })])
    ).toEqual({
      code: 'too-long',
      path: 'blocks[0].c[0].c[0].a.w:val',
      detail: `over ${LIMITS.attributeValue}`,
    });
    const many = Object.fromEntries(
      Array.from({ length: LIMITS.attributes + 1 }, (_, k) => [`w:a${k}`, '1'])
    );
    expect(refusal([{ e: 'w:p', a: many }])).toMatchObject({
      code: 'too-many',
      path: 'blocks[0].a',
    });
    expect(refusal(Array.from({ length: LIMITS.blocks + 1 }, () => para()))).toMatchObject({
      code: 'too-many',
      path: 'blocks',
    });
  });

  test('a reference out of range or of the wrong form', () => {
    const numId = (value: string) => [
      para({ e: 'w:pPr', c: [{ e: 'w:numPr', c: [{ e: 'w:numId', a: { 'w:val': value } }] }] }),
    ];
    const at = 'blocks[0].c[0].c[0].c[0].a.w:val';
    expect(refusal(numId('#1'))).toEqual({
      code: 'invalid-value',
      path: at,
      detail: 'not a list of the change',
    });
    expect(refusal(numId('#00'))).toMatchObject({ code: 'invalid-value', path: at });
    expect(refusal(numId('1234567890'))).toMatchObject({ code: 'invalid-value', path: at });
    expect(refusal(numId('-1'))).toMatchObject({ code: 'invalid-value', path: at });
    accepts(numId('#0'));
    accepts(numId('0'));
    accepts(numId('123456789'));

    const link = (value: string) => [para({ e: 'w:hyperlink', a: { 'r:id': value } })];
    expect(refusal(link('rId4'))).toMatchObject({
      code: 'invalid-value',
      path: 'blocks[0].c[0].a.r:id',
    });
    expect(refusal(link('#1'))).toMatchObject({
      code: 'invalid-value',
      path: 'blocks[0].c[0].a.r:id',
    });
    accepts(link('#0'));

    const note = (name: string, value: string) => [para(run({ e: name, a: { 'w:id': value } }))];
    expect(refusal(note('w:footnoteReference', '1'))).toMatchObject({
      code: 'invalid-value',
      path: 'blocks[0].c[0].c[0].a.w:id',
    });
    expect(refusal(note('w:endnoteReference', '#1'))).toMatchObject({ code: 'invalid-value' });
    accepts(note('w:footnoteReference', '#0'));
  });

  test('an object: a relationship other than a link or an image, and an object element', () => {
    expect(refusal([para(run({ e: 'w:r', a: { 'r:embed': 'rId5' } }))])).toEqual({
      code: 'invalid-value',
      path: 'blocks[0].c[0].c[0].a.r:embed',
      detail: 'object',
    });
    expect(refusal([para({ e: 'w:altChunk', a: { 'r:id': '#0' } })])).toMatchObject({
      code: 'invalid-value',
      detail: 'object',
    });
    // An ActiveX control is a program: no copy carries it
    expect(refusal([para(run(el('w:object', el('w:control'))))])).toEqual({
      code: 'invalid-value',
      path: 'blocks[0].c[0].c[0].c[0].e',
      detail: 'object',
    });
    // An embedded object's part and an ink part are media, named by index: only in range
    const ole = (id: string) => ({ e: 'o:OLEObject', a: { Type: 'Embed', 'r:id': id } });
    const ink = (e: string, id: string) => ({ e, a: { bwMode: 'auto', 'r:id': id } });
    accepts([para(run(el('w:object', ole('#0'))))]);
    accepts([para(run(ink('w14:contentPart', '#0')))]);
    accepts([para(run({ e: 'w:contentPart', a: { 'r:id': '#0' } }))]);
    for (const [tree, path] of [
      [run(el('w:object', ole('#1'))), 'blocks[0].c[0].c[0].c[0].a.r:id'],
      [run(ink('w14:contentPart', 'rId5')), 'blocks[0].c[0].c[0].a.r:id'],
    ] as const)
      expect(refusal([para(tree)])).toMatchObject({ code: 'invalid-value', path });
    // Ink's DrawingML-typed elements take attributes without a prefix; other w14 elements none
    expect(refusal([para(run({ e: 'w14:ligatures', a: { val: 'all' } }))])).toMatchObject({
      code: 'invalid-attribute',
      detail: 'no prefix',
    });
    // A VML picture by address: Word loads it when it opens the file, so no room writes one,
    // even beside an embedded copy (design D18)
    for (const a of [
      { 'o:href': 'file:///C:/seal.png' },
      { 'r:id': '#0', 'o:href': 'https://images.example.com/seal.png' },
      { 'o:althref': 'seal.png' },
      { src: 'https://images.example.com/seal.png' },
    ])
      for (const e of ['v:imagedata', 'v:fill'])
        expect(refusal([para(run(el('w:pict', el('v:shape', { e, a }))))])).toMatchObject({
          code: 'invalid-value',
          detail: 'a picture by address',
        });
    // A VML picture's image by the Office relationship id (`o:relid`) or `r:pict`: media, by
    // index only; its link to the file (`r:href`) is an address
    const imagedata = (a: Record<string, string>) =>
      para(run(el('w:pict', el('v:shape', { e: 'v:imagedata', a }))));
    for (const key of ['o:relid', 'r:pict']) {
      accepts([imagedata({ [key]: '#0' })]);
      expect(refusal([imagedata({ [key]: 'rId4' })])).toMatchObject({
        code: 'invalid-value',
        detail: 'not a picture of the change',
      });
    }
    expect(refusal([imagedata({ 'r:id': '#0', 'r:href': '#0' })])).toMatchObject({
      code: 'invalid-value',
      detail: 'a picture by address',
    });
  });

  test("a list's picture bullet: `#<j>` into the change's bullets, its id the apply's", () => {
    const level = (value: string) => ({
      e: 'w:abstractNum',
      c: [el('w:lvl', { e: 'w:lvlPicBulletId', a: { 'w:val': value } })],
    });
    const read = (value: unknown, root: OrderedFragmentRoot, bullets: number) =>
      attempt(() =>
        fragmentElements(value, 'blocks', root, { ...REFS, bullets }, fragmentBudget())
      );
    expect(refusal([level('0')], 'abstractNum')).toMatchObject({
      code: 'invalid-value',
      path: 'blocks[0].c[0].c[0].a.w:val',
      detail: 'not a picture bullet of the change',
    });
    expect(read([level('#1')], 'abstractNum', 1).ok).toBe(false);
    expect(read([level('#0')], 'abstractNum', 1).ok).toBe(true);
    const bullet = (a?: Record<string, string>) => ({
      e: 'w:numPicBullet',
      ...(a ? { a } : {}),
      c: [
        el(
          'w:pict',
          el('v:shape', {
            e: 'v:imagedata',
            a: { 'r:id': '#0' },
          })
        ),
      ],
    });
    expect(read([bullet()], 'numPicBullet', 0).ok).toBe(true);
    expect(refusal([bullet({ 'w:numPicBulletId': '3' })], 'numPicBullet')).toMatchObject({
      code: 'invalid-attribute',
      detail: 'the apply gives the id',
    });
    // A bullet is no block, and a block is no bullet
    expect(refusal([bullet()])).toMatchObject({ code: 'invalid-value', path: 'blocks[0].e' });
    expect(refusal([para()], 'numPicBullet')).toMatchObject({ code: 'invalid-value' });
  });

  test("a picture's image out of range, a picture by address, a drawing id that is no decimal", () => {
    const picture = (blip: Record<string, string>, docPr: Record<string, string> = { id: '1' }) => [
      para(
        run(
          el(
            'w:drawing',
            el(
              'wp:inline',
              { e: 'wp:docPr', a: docPr },
              el(
                'a:graphic',
                el('a:graphicData', el('pic:pic', el('pic:blipFill', { e: 'a:blip', a: blip })))
              )
            )
          )
        )
      ),
    ];
    accepts(picture({ 'r:embed': '#0' }));
    expect(refusal(picture({ 'r:embed': '#1' }))).toMatchObject({
      code: 'invalid-value',
      detail: 'not a picture of the change',
    });
    // A picture by address: no room writes one, with or without an embedded copy (design D18)
    const blips: Record<string, string>[] = [
      { 'r:link': '#0' },
      { 'r:embed': '#0', 'r:link': '#0' },
    ];
    for (const blip of blips)
      expect(refusal(picture(blip))).toMatchObject({
        code: 'invalid-value',
        detail: 'a picture by address',
      });
    expect(refusal(picture({ 'r:embed': '#0' }, { id: 'x' }))).toMatchObject({
      code: 'invalid-value',
      detail: 'not a decimal',
    });
  });

  test('an attribute without a prefix outside a drawing, and a namespace declaration in one', () => {
    expect(refusal([para(run({ e: 'w:t', a: { space: 'preserve' }, c: ['a'] }))])).toMatchObject({
      code: 'invalid-attribute',
      detail: 'no prefix',
    });
    expect(
      refusal([para(run(el('w:drawing', { e: 'wp:inline', a: { xmlns: 'urn:x' } })))])
    ).toMatchObject({ code: 'invalid-attribute', detail: 'a namespace declaration' });
    accepts([para(run(el('w:drawing', { e: 'wp:inline', a: { distT: '0' } })))]);
  });

  test("a shape's markup compatibility: a choice's Requires only, naming fragment prefixes", () => {
    const choice = (a: Record<string, string>) => [
      para(
        run(
          el(
            'mc:AlternateContent',
            { e: 'mc:Choice', a, c: [el('w:drawing')] },
            el('mc:Fallback', el('w:pict'))
          )
        )
      ),
    ];
    accepts(choice({ Requires: 'wps' }));
    accepts(choice({ Requires: 'wpg wps' }));
    // A drawing canvas travels, and so do ink (`wpi`) and an Office 2016 chart (`cx1`)
    accepts(choice({ Requires: 'wpc' }));
    accepts(choice({ Requires: 'wpi' }));
    accepts(choice({ Requires: 'cx1' }));
    for (const value of ['cx9', 'wps  wpg', '', 'w14 a b c d e f g h'])
      expect(refusal(choice({ Requires: value }))).toMatchObject({
        code: 'invalid-value',
        path: 'blocks[0].c[0].c[0].c[0].a.Requires',
      });
    expect(refusal(choice({ Requires: 'wps', Other: 'x' }))).toMatchObject({
      code: 'invalid-attribute',
      detail: 'Requires of mc:Choice only',
    });
    expect(
      refusal([para(run({ e: 'mc:AlternateContent', a: { Requires: 'wps' } }))])
    ).toMatchObject({ code: 'invalid-attribute' });
  });

  test("VML: values within the change's text budget, Word's DrawingML copy as base64 lines, a safe click link", () => {
    const shape = (a: Record<string, string>) => [para(run(el('w:pict', { e: 'v:shape', a })))];
    accepts(
      shape({
        path: 'm,l'.repeat(1000),
        'o:gfxdata': 'UEsDBBQ\nABgAIAAAAIQ==\r\n',
        'v:ext': 'edit',
      })
    );
    expect(refusal(shape({ 'o:gfxdata': 'UEs<x>' }))).toMatchObject({
      code: 'invalid-text',
      detail: 'not base64',
    });
    expect(refusal(shape({ path: 'm\tl' }))).toMatchObject({ code: 'invalid-text' });
    accepts(shape({ href: 'https://example.com/terms' }));
    expect(refusal(shape({ href: 'javascript:alert(1)' }))).toEqual({
      code: 'invalid-href',
      path: 'blocks[0].c[0].c[0].c[0].a.href',
    });
  });

  test("a VML picture's and an SVG version's images are indexes into the change's media", () => {
    const vml = (owner: string, value: string) => [
      para(run(el('w:pict', el('v:shape', { e: owner, a: { 'r:id': value } })))),
    ];
    accepts(vml('v:imagedata', '#0'));
    accepts(vml('v:fill', '#0'));
    expect(refusal(vml('v:imagedata', '#1'))).toMatchObject({
      code: 'invalid-value',
      detail: 'not a picture of the change',
    });
    expect(refusal(vml('v:imagedata', 'rId4'))).toMatchObject({ code: 'invalid-value' });
    expect(refusal(vml('v:shape', '#0'))).toMatchObject({ detail: 'object' });
    const svg = (value: string) => [
      para(
        run(
          el(
            'w:drawing',
            el(
              'a:blip',
              el('a:extLst', el('a:ext', { e: 'asvg:svgBlip', a: { 'r:embed': value } }))
            )
          )
        )
      ),
    ];
    accepts(svg('#0'));
    expect(refusal(svg('#1'))).toMatchObject({ detail: 'not a picture of the change' });
  });
});

describe('the fragment name lists', () => {
  test('come from the schemas, with the names Word writes, and do not change', () => {
    expect(WML_FRAGMENT_ELEMENTS.size).toBe(551);
    expect(MATH_FRAGMENT_ELEMENTS.size).toBe(124);
    expect(W14_FRAGMENT_ELEMENTS.size).toBe(17);
    expect(W15_FRAGMENT_ELEMENTS.size).toBe(10);
    expect(WP_FRAGMENT_ELEMENTS.size).toBe(43);
    expect(DRAWINGML_FRAGMENT_ELEMENTS.size).toBe(321);
    expect(PICTURE_FRAGMENT_ELEMENTS.size).toBe(6);
    expect(WPS_FRAGMENT_ELEMENTS.size).toBe(10);
    expect(WPG_FRAGMENT_ELEMENTS.size).toBe(9);
    expect(MC_FRAGMENT_ELEMENTS.size).toBe(3);
    expect(VML_FRAGMENT_ELEMENTS.size).toBe(23);
    expect(OFFICE_FRAGMENT_ELEMENTS.size).toBe(32);
    expect(W10_FRAGMENT_ELEMENTS.size).toBe(6);
    expect(SVG_BLIP_FRAGMENT_ELEMENTS.size).toBe(1);
    for (const name of ['inline', 'anchor', 'extent', 'docPr', 'wrapSquare'])
      expect(WP_FRAGMENT_ELEMENTS.has(name)).toBe(true);
    for (const name of ['graphic', 'graphicData', 'blip', 'hlinkClick', 'srcRect', 'extLst'])
      expect(DRAWINGML_FRAGMENT_ELEMENTS.has(name)).toBe(true);
    for (const name of [
      'p',
      'tbl',
      'sdt',
      'numberingChange',
      'pict',
      'txbxContent',
      'object',
      'softHyphen',
      'noBreakHyphen',
      'lastRenderedPageBreak',
      'proofErr',
      'permStart',
      'permEnd',
      'fldData',
      'legacy',
    ])
      expect(WML_FRAGMENT_ELEMENTS.has(name)).toBe(true);
    expect(MATH_FRAGMENT_ELEMENTS.has('oMathPara')).toBe(true);
    expect(() => (WML_FRAGMENT_ELEMENTS as Set<string>).add('script')).toThrow(TypeError);
    expect(() => (W14_FRAGMENT_ELEMENTS as Set<string>).clear()).toThrow(TypeError);
    expect(WML_FRAGMENT_ELEMENTS.has('script')).toBe(false);
  });
});

/** Body children of literal XML. */
function nodesOf(body: string): OoxmlElement[] {
  const read = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:r="${R}" ` +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
      `xmlns:v="${V}" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w14="${W14}" ` +
      'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
      'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
      'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
      `xmlns:x="urn:example:thing" xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"><w:body>${body}</w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  const holder = read.part.root.children[0];
  if (holder === undefined || holder.kind === 'textValue') throw new Error('no body');
  return (holder.children as readonly OoxmlNode[]).filter(
    (child): child is OoxmlElement => child.kind !== 'textValue'
  );
}

const IDENTITY = {
  numId: (v: string) => v,
  relationship: (v: string) => v,
  media: (v: string) => v,
  note: (_: string, v: string) => v,
};

function wireRefusal(body: string): { code: string; reason: string; detail?: string } {
  try {
    fragmentWire(nodesOf(body), IDENTITY);
  } catch (error) {
    if (error instanceof OrderedClientRefusal)
      return error.detail === undefined
        ? { code: error.code, reason: error.reason }
        : { code: error.code, reason: error.reason, detail: error.detail };
    throw error;
  }
  throw new Error('the wire took it');
}

const vmlShape = (inner: string, a = '') => `<w:pict><v:shape${a}>${inner}</v:shape></w:pict>`;
const pict = (inner: string) => `<w:r>${vmlShape(inner)}</w:r>`;

describe('the fragment wire refuses', () => {
  test('an ActiveX control, an Office 2016 chart, a picture by address alone, a choice the format does not hold', () => {
    const object = { code: 'ordered-unsupported-op', reason: 'insertFragment.object' };
    expect(wireRefusal(p('<w:r><w:object><w:control r:id="rIdX"/></w:object></w:r>'))).toEqual(
      object
    );
    // An embedded object travels only with its part: a wire that carries none refuses it
    expect(wireRefusal(p(pict('<o:OLEObject r:id="rIdOle"/>')))).toEqual(object);
    // A picture that only an address names: Word loads it when it opens the file (design D18)
    const byAddress = {
      code: 'ordered-refused',
      reason: 'external-picture-link',
      detail: 'a picture by address',
    };
    expect(wireRefusal(p(pict('<v:imagedata o:href="file:///C:/seal.png"/>')))).toEqual(byAddress);
    expect(wireRefusal(p(pict('<v:imagedata src="https://images.example.com/s.png"/>')))).toEqual(
      byAddress
    );
    expect(wireRefusal(p(pictureRun(1).replace('r:embed=', 'r:link=')))).toEqual(byAddress);
    expect(wireRefusal(p(pict('<v:imagedata r:href="rIdSeal"/>')))).toEqual(byAddress);
    // An Office 2016 chart (`cx`, a box and whisker, a waterfall) travels only with its parts: a
    // wire that carries none refuses it
    const chartEx = pictureRun(1).replace(
      /<pic:pic>.*<\/pic:pic>/,
      '<cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" r:id="rIdChart"/>'
    );
    expect(wireRefusal(p(chartEx))).toEqual(object);
    // A chart or a diagram travels only with its parts: a wire that carries none refuses it
    const chart = pictureRun(1).replace(/<pic:pic>.*<\/pic:pic>/, '<c:chart r:id="rIdChart"/>');
    expect(wireRefusal(p(chart))).toEqual(object);
    // An element of a namespace the format does not hold, outside an extension and a choice
    expect(
      wireRefusal(
        p(
          pictureRun(1).replace(
            '<pic:cNvPicPr/>',
            '<am3d:model3d xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"/>'
          )
        )
      )
    ).toEqual(object);
  });

  test('what a namespace the format does not hold extends is left out, as Word older than it reads it', () => {
    const json = (body: string) => JSON.stringify(fragmentWire(nodesOf(body), IDENTITY));
    // A DrawingML extension: the picture travels without it
    const sketchy =
      '<a:extLst><a:ext uri="{C807C97D-BFC1-408E-A445-0C87EB9F89A2}"><ask:lineSketchStyleProps ' +
      'xmlns:ask="http://schemas.microsoft.com/office/drawing/2018/sketchyshapes" sd="1219033472"/>' +
      '</a:ext><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId ' +
      `xmlns:a16="${A16_NAMESPACE_URI}" id="{0}"/></a:ext></a:extLst>`;
    const picture = json(
      p(
        pictureRun(1).replace(
          '<a:blip r:embed="rIdImg1"/>',
          `<a:blip r:embed="rIdImg1">${sketchy}</a:blip>`
        )
      )
    );
    expect(picture).toContain('{"e":"a16:creationId"');
    expect(picture).not.toContain('lineSketchStyleProps');
    expect(picture).toContain('{"e":"a:blip","a":{"r:embed":"rIdImg1"}');
    // A choice that requires one: the other choice, or the fallback in the holder's place
    const choice = (requires: string, inner: string) =>
      `<mc:Choice xmlns:x="urn:x" Requires="${requires}">${inner}</mc:Choice>`;
    const holder = (choices: string) =>
      `<w:r><mc:AlternateContent>${choices}<mc:Fallback><w:t>fallback</w:t></mc:Fallback>` +
      '</mc:AlternateContent></w:r>';
    const kept = json(p(holder(choice('x', '<w:t>3D</w:t>') + choice('w14', '<w:t>known</w:t>'))));
    expect(kept).toContain('{"e":"mc:Choice","a":{"Requires":"w14"}');
    expect(kept).not.toContain('3D');
    expect(kept).toContain('fallback');
    const fallback = json(p(holder(choice('x', '<w:t>3D</w:t>'))));
    expect(fallback).not.toContain('mc:');
    expect(fallback).toContain('{"e":"w:r","c":[{"e":"w:t","c":["fallback"]}]}');
  });

  test("an embedded object and ink travel, their parts named through the wire's object reference", () => {
    const named: string[] = [];
    const object =
      '<w:r><w:object w:dxaOrig="1440" w:dyaOrig="720"><v:shape id="_x0000_i1025" o:ole="">' +
      '<v:imagedata r:id="rIdPreview" o:title=""/></v:shape><o:OLEObject Type="Embed" ' +
      'ProgID="Excel.Sheet.12" ShapeID="_x0000_i1025" r:id="rIdSheet"/></w:object></w:r>';
    const ink =
      '<w:r><mc:AlternateContent><mc:Choice Requires="wpi"><w14:contentPart bwMode="auto" ' +
      'r:id="rIdInk"><w14:nvContentPartPr><w14:cNvContentPartPr/></w14:nvContentPartPr>' +
      '<w14:xfrm><a:off x="0" y="0"/><a:ext cx="360" cy="360"/></w14:xfrm></w14:contentPart>' +
      '</mc:Choice><mc:Fallback><w:t>x</w:t></mc:Fallback></mc:AlternateContent></w:r>';
    const json = JSON.stringify(
      fragmentWire(nodesOf(p(object + ink)), {
        ...IDENTITY,
        media: () => '#0',
        object: (id, use) => {
          named.push(`${use} ${id}`);
          return use === 'ink' ? '#2' : '#1';
        },
      })
    );
    expect(named).toEqual(['object rIdSheet', 'ink rIdInk']);
    expect(json).toContain('{"e":"o:OLEObject","a":{"Type":"Embed","ProgID":"Excel.Sheet.12"');
    expect(json).toContain('"r:id":"#1"');
    expect(json).toContain('{"e":"w14:contentPart","a":{"bwMode":"auto","r:id":"#2"}');
    expect(json).toContain('{"e":"mc:Choice","a":{"Requires":"wpi"}');
    const read = attempt(() =>
      fragmentElements(JSON.parse(json), 'blocks', 'block', { ...REFS, media: 3 }, fragmentBudget())
    );
    expect(read.ok ? 'read' : read.error).toBe('read');
  });

  test("a shape travels: its choice's prefixes as the change names them, Word's drawing ids left out", () => {
    const named: string[] = [];
    const box =
      `<w:r><mc:AlternateContent><mc:Choice xmlns:s="${WPS}" Requires="s"><w:drawing><wp:inline>` +
      '<wp:docPr id="8" name="Text Box 8"/><a:graphic><a:graphicData uri="x"><s:wsp>' +
      `<s:txbx><w:txbxContent>${p(t('Box text'))}</w:txbxContent></s:txbx><s:bodyPr/></s:wsp>` +
      '</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback>' +
      vmlShape(
        '<v:imagedata r:id="rIdImg1" o:title=""/>',
        ' w14:anchorId="3CEC9D07" o:spid="_x0000_s1026" style="width:72pt"'
      ) +
      '</mc:Fallback></mc:AlternateContent></w:r>';
    const json = JSON.stringify(
      fragmentWire(nodesOf(p(box)), {
        ...IDENTITY,
        media: (id) => {
          named.push(id);
          return '#0';
        },
      })
    );
    expect(named).toEqual(['rIdImg1']);
    expect(json).toContain('{"e":"mc:Choice","a":{"Requires":"wps"}');
    expect(json).toContain('{"e":"wps:wsp"');
    expect(json).toContain('{"e":"w:txbxContent"');
    expect(json).toContain('{"e":"v:shape","a":{"o:spid":"_x0000_s1026","style":"width:72pt"}');
    expect(json).toContain('{"e":"v:imagedata","a":{"r:id":"#0","o:title":""}}');
    expect(json).not.toContain('anchorId');
  });

  test("a picture's SVG version travels, its image through the media reference", () => {
    const named: string[] = [];
    const svg = pictureRun(1).replace(
      '<a:blip r:embed="rIdImg1"/>',
      '<a:blip r:embed="rIdImg1"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
        '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" ' +
        'r:embed="rIdSvg"/></a:ext></a:extLst></a:blip>'
    );
    const json = JSON.stringify(
      fragmentWire(nodesOf(p(svg)), {
        ...IDENTITY,
        media: (id) => {
          named.push(id);
          return `#${named.length - 1}`;
        },
      })
    );
    expect(named).toEqual(['rIdImg1', 'rIdSvg']);
    expect(json).toContain('{"e":"asvg:svgBlip","a":{"r:embed":"#1"}}');
  });

  test('a picture with an embedded copy travels without its address, never fetched', () => {
    const named: string[] = [];
    const media = (id: string) => {
      named.push(id);
      return `#${named.length - 1}`;
    };
    const linked = pictureRun(1).replace('r:embed="rIdImg1"', 'r:embed="rIdImg1" r:link="rIdSeal"');
    const vml = pict('<v:imagedata r:id="rIdImg2" o:href="file:///C:/seal.png" o:title=""/>');
    const json = JSON.stringify(fragmentWire(nodesOf(p(linked + vml)), { ...IDENTITY, media }));
    expect(named).toEqual(['rIdImg1', 'rIdImg2']);
    expect(json).toContain('{"e":"a:blip","a":{"r:embed":"#0"}}');
    expect(json).toContain('{"e":"v:imagedata","a":{"r:id":"#1","o:title":""}}');
    expect(json).not.toContain('seal');
  });

  test("a picture travels, its image through the media reference, Word's drawing ids left out", () => {
    const named: string[] = [];
    const wire = fragmentWire(nodesOf(p(pictureRun(1))), {
      ...IDENTITY,
      media: (id) => {
        named.push(id);
        return '#0';
      },
    });
    expect(named).toEqual(['rIdImg1']);
    const json = JSON.stringify(wire);
    expect(json).toContain('{"e":"a:blip","a":{"r:embed":"#0"}}');
    expect(json).toContain('{"e":"wp:extent","a":{"cx":"457200","cy":"457200"}}');
  });

  test('an element and an attribute of a namespace the format does not have, by a short name', () => {
    expect(wireRefusal(p(`<w:r><x:thing/>${t('a')}</w:r>`))).toEqual({
      code: 'ordered-refused',
      reason: 'invalid-change',
      detail: 'fragment-element example/thing:thing',
    });
    expect(wireRefusal(p('<w:ins w:id="1" w16du:dateUtc="2026-01-01T00:00:00Z"/>'))).toEqual({
      code: 'ordered-refused',
      reason: 'invalid-change',
      detail: 'fragment-attribute wordml/word16du:dateUtc',
    });
    expect(wireRefusal(p('<w:r><w:frobnicate/></w:r>'))).toMatchObject({
      detail: 'fragment-element w:frobnicate',
    });
  });

  test('text outside a text element, and a tab inside one', () => {
    expect(wireRefusal(p('<w:r>loose</w:r>'))).toMatchObject({
      code: 'ordered-refused',
      detail: 'fragment-text w:r',
    });
    expect(wireRefusal(p('<w:r><w:t>a\tb</w:t></w:r>'))).toMatchObject({
      detail: 'fragment-text w:t',
    });
  });
});
