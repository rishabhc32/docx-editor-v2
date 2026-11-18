// Resources outside the story text (task 9.2, design D14): names made from the change id, the
// undo record, names the file or another resource holds, another person's pending resource,
// and the validator's bounds on every field a change gives.

import { describe, expect, test } from 'bun:test';
import { orderedIdMinter } from '../change/ids.ts';
import { attempt } from '../change/read.ts';
import {
  changeMedia,
  changeSetting,
  changeStyle,
  ORDERED_RESOURCE_LIMITS,
} from '../change/resources.ts';
import { addList, addMedia, addRelationship, addStyle, resourceFor } from '../edits/resources.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import {
  ORDERED_MEDIA_TYPES,
  type OrderedListLevel,
  orderedListNumber,
  orderedMediaPart,
  orderedRelationshipId,
  orderedStoryPart,
} from '../model/resources.ts';
import { takeBack } from '../model/undo-record.ts';
import { applyAs, decide, media, openFixture, pngBytes } from './resources-fixture.ts';

const LEVEL: OrderedListLevel = {
  level: 0,
  start: 1,
  format: 'decimal',
  text: '%1.',
  justification: 'left',
  indent: { left: 720, hanging: 360 },
};

/** The refusal reason of a reader, or `ok`. */
const read = (reader: () => unknown): string => {
  const result = attempt(reader);
  return result.ok ? 'ok' : `${result.error.code} ${result.error.path}`;
};

const refusedWith = (run: () => void): string => {
  try {
    run();
  } catch (error) {
    if (error instanceof OrderedRefusal) return error.reason;
    throw error;
  }
  return 'applied';
};

describe('resources outside the story text (task 9.2)', () => {
  test('keys and Word names come from the change id, the same on every copy', () => {
    const one = orderedIdMinter('ana-1:12');
    const two = orderedIdMinter('ana-1:12');
    const keys = [one.resource(), one.resource()];
    expect(keys).toEqual(['p:ana-1:12.1', 'p:ana-1:12.2']);
    expect([two.resource(), two.resource()]).toEqual(keys);
    expect(orderedRelationshipId('p:ana-1:12.1')).toBe('rOana-1_12_1');
    expect(orderedMediaPart('p:ana-1:12.1', 'image/png')).toBe(
      '/word/media/ordered_ana-1_12_1.png'
    );
    expect(orderedStoryPart('p:a_b:3.4', 'header')).toBe('/word/header_a_b_3_4.xml');
    expect(orderedListNumber('p:ana-1:12.1', 'num')).toBe(orderedListNumber('p:ana-1:12.1', 'num'));
    expect(orderedListNumber('p:ana-1:12.1', 'num')).not.toBe(
      orderedListNumber('p:ana-1:12.1', 'abstract')
    );
    expect(() => orderedRelationshipId('S:ana-1:12.1')).toThrow(RangeError);
  });

  test('a resource is added inside an undo record, and the record takes it back exactly', () => {
    const { doc } = openFixture();
    const record = applyAs(doc, 'ana-1:1', 'Ana', ({ ctx, suggestion }) => {
      const picture = addMedia(ctx, suggestion, media(pngBytes(1)));
      addRelationship(ctx, suggestion, 'image', { resource: picture.key });
      addList(ctx, suggestion, { levels: [LEVEL] });
    });
    expect([...doc.resources.added.keys()]).toEqual(['p:ana-1:1.1', 'p:ana-1:1.2', 'p:ana-1:1.3']);
    takeBack(doc, record);
    expect(doc.resources.added.size).toBe(0);
    // A change refused midway leaves nothing
    expect(
      refusedWith(() =>
        applyAs(doc, 'ana-1:2', 'Ana', ({ ctx, suggestion }) => {
          addMedia(ctx, suggestion, media(pngBytes(2)));
          throw new OrderedRefusal('invalid-edit', 'midway');
        })
      )
    ).toBe('invalid-edit');
    expect(doc.resources.added.size).toBe(0);
  });

  test('a name the file or another resource holds is refused on every copy alike', () => {
    const { doc } = openFixture();
    // The file's own list numbers and relationship ids are known after load
    expect(doc.resources.fileLists.has('num 1')).toBe(true);
    expect(doc.resources.fileRelationships.get('/word/document.xml')?.has('rIdLink')).toBe(true);
    expect(doc.resources.fileParts.has('/word/document.xml')).toBe(true);
    // A list number the file holds already
    doc.resources.fileLists.add(`num ${orderedListNumber('p:ana-1:3.1', 'num')}`);
    expect(
      refusedWith(() =>
        applyAs(doc, 'ana-1:3', 'Ana', ({ ctx, suggestion }) =>
          addList(ctx, suggestion, { levels: [LEVEL] })
        )
      )
    ).toBe('name-taken');
    // A style id two resources share is fine: the first by key is written (compose test)
    applyAs(doc, 'ana-1:4', 'Ana', ({ ctx, suggestion }) => {
      addStyle(ctx, suggestion, { styleId: 'Quote', type: 'paragraph', name: 'Quote' });
      addStyle(ctx, suggestion, { styleId: 'Quote', type: 'paragraph', name: 'Quote' });
    });
    expect(doc.resources.added.size).toBe(2);
  });

  test("another person's pending resource is not theirs to use; an accepted one is", () => {
    const { doc } = openFixture();
    let key = '';
    let anas = '';
    applyAs(doc, 'ana-1:1', 'Ana', ({ ctx, suggestion }) => {
      key = addMedia(ctx, suggestion, media(pngBytes(1))).key;
      anas = suggestion;
    });
    const use = () =>
      applyAs(doc, 'ben-1:1', 'Ben', ({ ctx, suggestion }) => {
        resourceFor(ctx, key, suggestion);
        addRelationship(ctx, suggestion, 'image', { resource: key });
      });
    expect(refusedWith(use)).toBe('foreign-resource');
    decide(doc, anas, 'accepted');
    expect(refusedWith(use)).toBe('applied');
    // A link points outside, to a safe address only; a picture never does
    expect(
      refusedWith(() =>
        applyAs(doc, 'ben-1:2', 'Ben', ({ ctx, suggestion }) =>
          addRelationship(ctx, suggestion, 'hyperlink', { external: 'javascript:alert(1)' })
        )
      )
    ).toBe('invalid-edit');
    expect(
      refusedWith(() =>
        applyAs(doc, 'ben-1:3', 'Ben', ({ ctx, suggestion }) =>
          addRelationship(ctx, suggestion, 'image', { external: 'https://example.com/a.png' })
        )
      )
    ).toBe('invalid-edit');
  });

  test('the validator bounds every field a change gives', () => {
    const hash = 'a'.repeat(64);
    expect(read(() => changeMedia({ sha256: hash, contentType: 'image/png', size: 10 }, 'm'))).toBe(
      'ok'
    );
    expect(
      read(() => changeMedia({ sha256: 'A'.repeat(64), contentType: 'image/png', size: 1 }, 'm'))
    ).toBe('invalid-value m.sha256');
    expect(read(() => changeMedia({ sha256: hash, contentType: 'text/html', size: 1 }, 'm'))).toBe(
      'invalid-value m.contentType'
    );
    expect(
      read(() =>
        changeMedia(
          { sha256: hash, contentType: 'image/png', size: ORDERED_RESOURCE_LIMITS.mediaSize + 1 },
          'm'
        )
      )
    ).toBe('invalid-number m.size');

    // An SVG has its own, smaller cap
    const svgOfSize = (size: number) =>
      read(() =>
        changeMedia({ sha256: hash, contentType: 'image/svg+xml', size }, 'm', ORDERED_MEDIA_TYPES)
      );
    expect(svgOfSize(ORDERED_RESOURCE_LIMITS.svgSize)).toBe('ok');
    expect(svgOfSize(ORDERED_RESOURCE_LIMITS.svgSize + 1)).toBe('invalid-number m.size');
    expect(
      read(() =>
        changeMedia(
          JSON.parse(`{"sha256":"${hash}","contentType":"image/png","size":1,"__proto__":{}}`),
          'm'
        )
      )
    ).toBe('forbidden-key m.__proto__');
    // A style: an id pattern, a name without control characters, properties on the allowlists
    expect(
      read(() =>
        changeStyle(
          {
            styleId: 'Hyperlink',
            type: 'character',
            name: 'Hyperlink',
            run: [{ localName: 'u', attributes: { val: 'single' } }],
          },
          's'
        )
      )
    ).toBe('ok');
    expect(read(() => changeStyle({ styleId: '1a', type: 'paragraph', name: 'X' }, 's'))).toBe(
      'invalid-value s.styleId'
    );
    expect(
      read(() =>
        changeStyle(
          { styleId: 'A', type: 'paragraph', name: 'X', run: [{ localName: 'script' }] },
          's'
        )
      )
    ).toBe('unknown-property s.run[0].localName');
    expect(read(() => changeSetting({ op: 'setDocumentProtection', enforce: true }, 'x'))).toBe(
      'ok'
    );
    expect(read(() => changeSetting({ op: 'setWriteProtection', enforce: true }, 'x'))).toBe(
      'invalid-value x.op'
    );
  });

  test('a style name that XML cannot hold is refused, so later saves still work', () => {
    const style = (name: string) => ({ styleId: 'Link', type: 'character', name });
    expect(read(() => changeStyle(style('Link \uD800 Style'), 's'))).toBe('invalid-text s.name');
    expect(read(() => changeStyle(style('\uDC00'), 's'))).toBe('invalid-text s.name');
    expect(read(() => changeStyle(style('Link 😀'), 's'))).toBe('ok');
    // A name may use the whole style name limit
    const longest = 'n'.repeat(ORDERED_RESOURCE_LIMITS.styleName);
    expect(read(() => changeStyle(style(longest), 's'))).toBe('ok');
    expect(read(() => changeStyle(style(`${longest}n`), 's'))).toBe('too-long s.name');
  });
});
