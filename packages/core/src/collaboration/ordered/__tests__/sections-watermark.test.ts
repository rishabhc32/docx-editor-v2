// A picture watermark in the room (task 9.8, design D14). The room refuses it while its storage
// lacks the bytes (`missing-blob`) and refuses a party's by the host's party rule; applied, each
// header that shows it takes a relationship to one media part, the save writes the bytes, the
// engine reads it back as the document's watermark, and the organiser's revert takes it all out.
// With real editors: `client/editor-watermark.test.ts`.

import { describe, expect, test } from 'bun:test';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { readPackageWatermark } from '../../../store/package/watermark-package.ts';
import { composeOrderedPackage } from '../compose/package.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { OrderedRoom } from '../room/room.ts';
import { suggestionsOnlyPolicy } from '../room/suggestions-only.ts';
import { media, pngBytes } from './resources-fixture.ts';
import { actor, BODY, realApply } from './room-harness.ts';
import { docx, p, packageOf, R, t, WML_CT } from './word-fixtures.ts';

const HEADER = '/word/header1.xml';
const FILE = docx(
  p(t('Body text.')) +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/></w:sectPr>',
  {
    parts: {
      'word/header1.xml':
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>',
    },
    relationships: `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes: `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>`,
  }
);

const OLGA = actor('Olga', 'olga-1', 'organiser');
const ANA = actor('Ana', 'ana-1', 'party');
const BYTES = pngBytes(5);
const PICTURE = {
  kind: 'picture',
  washout: true,
  widthPt: 300,
  heightPt: 200,
  media: media(BYTES),
} as const;
const blobs = (hash: string) => (hash === media(BYTES).sha256 ? BYTES : undefined);

function roomWith(stored: ReadonlySet<string>) {
  let tick = 0;
  const room = new OrderedRoom({
    document: loadOrderedPackage(packageOf(FILE)),
    applyChange: realApply,
    policy: suggestionsOnlyPolicy(),
    now: () => Date.UTC(2026, 9, 29) + 1000 * tick++,
    blobs: { has: (hash) => stored.has(hash) },
  });
  const counters = new Map<string, number>();
  const send = (who: typeof OLGA, change: Record<string, unknown>): string => {
    const k = (counters.get(who.namespace) ?? 0) + 1;
    counters.set(who.namespace, k);
    const result = room.submit(who, [{ v: 1, id: `${who.namespace}:${k}`, ...change }]);
    const refusal = result.refused[0];
    return refusal ? `${refusal.code}${refusal.reason ? `/${refusal.reason}` : ''}` : 'accepted';
  };
  return { room, send };
}

const watermark = (mode: 'edit' | 'suggest') => ({
  kind: 'watermark',
  part: BODY,
  mode,
  watermark: PICTURE,
});

/** The header's image relationship and the media part it names. */
function pictureOf(pkg: OoxmlPackage) {
  const relationship = pkg.relationships.get(HEADER)?.find((r) => r.type === `${R}/image`);
  const part = [...pkg.partBytes.keys()].find((name) => name.startsWith('/word/media/ordered_'));
  return { relationship, part };
}

describe('a picture watermark in the room', () => {
  test("the room refuses it without its bytes, and a party's by the host's party rule", () => {
    expect(roomWith(new Set()).send(OLGA, watermark('edit'))).toBe('missing-blob');
    const { send } = roomWith(new Set([PICTURE.media.sha256]));
    expect(send(ANA, watermark('suggest'))).toBe('policy-furniture/furniture');
  });

  test("the organiser's: a relationship from the header to one media part, read back as the document's; a revert takes them out", () => {
    const { room, send } = roomWith(new Set([PICTURE.media.sha256]));
    expect(send(OLGA, watermark('edit'))).toBe('accepted');
    const base = packageOf(FILE);
    const pkg = composeOrderedPackage(room.document, base, { blobs });
    const { relationship, part } = pictureOf(pkg);
    expect(relationship).toBeDefined();
    expect(part).toBeDefined();
    expect(pkg.partBytes.get(part ?? '')).toEqual(BYTES);
    const xml = serializeOoxmlPart(pkg.parts.get(HEADER)!);
    expect(xml).toContain(`r:id="${relationship?.id}"`);
    expect(xml).toContain('id="WordPictureWatermark1"');
    expect(xml).toContain('Header text');
    expect(readPackageWatermark(pkg)).toMatchObject({ kind: 'picture', washout: true });
    expect(send(OLGA, { kind: 'revert', suggestions: ['S:olga-1:1.1'], status: 'rejected' })).toBe(
      'accepted'
    );
    const reverted = composeOrderedPackage(room.document, base, { blobs });
    expect(pictureOf(reverted)).toEqual({ relationship: undefined, part: undefined });
    expect(readPackageWatermark(reverted)).toBeNull();
  });
});
