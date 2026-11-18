// Body XML normalized for a comparison: two documents compare on what they say, not on
// numbering noise (spike 4's `compare.ts`). Test code: the input is our own fixtures.
//
// Removed: namespace declarations, `w:date`, `w14:paraId` and `w14:textId`. Renumbered:
// revision `w:id` values, in order of first appearance (two wrappers that shared an id still
// share one), and bookmark `w:id` values, which only pair a start with its end (the engine takes
// the lowest free one, a room mints one from the change). Resolved: `r:id` to its target.
// Attributes are sorted per tag.

import { serializeOoxmlPart, type OoxmlPart } from '../../../../store/package/ooxml-tree.ts';

const REVISION_TAGS = new Set([
  'w:ins',
  'w:del',
  'w:rPrChange',
  'w:pPrChange',
  'w:tblPrChange',
  'w:trPrChange',
  'w:tcPrChange',
  'w:tblGridChange',
  'w:moveFrom',
  'w:moveTo',
  'w:moveFromRangeStart',
  'w:moveFromRangeEnd',
  'w:moveToRangeStart',
  'w:moveToRangeEnd',
  'w:cellIns',
  'w:cellDel',
  'w:cellMerge',
]);
const BOOKMARK_TAGS = new Set(['w:bookmarkStart', 'w:bookmarkEnd']);
const DROPPED = new Set(['w:date', 'w14:paraId', 'w14:textId']);

export type RelationshipResolver = (id: string) => string | null;

/** A part's blocks (the body's, or the root's) as normalized XML. */
/** The attributes that name a picture's image: a blip's, and a VML picture's other ids. */
const MEDIA_KEYS: ReadonlySet<string> = new Set(['r:embed', 'r:link', 'r:pict', 'o:relid']);

export function normalizedBody(part: OoxmlPart, resolve: RelationshipResolver): string {
  const xml = serializeOoxmlPart(part);
  const body = /<w:body>([\s\S]*)<\/w:body>/.exec(xml)?.[1] ?? xml;
  const ids = new Map<string, string>();
  const bookmarks = new Map<string, string>();
  // A new object's image and `wp:docPr/@id` take names each side mints its own way: they read
  // by order of appearance. A loaded image reads by its target, so a missed change shows
  const fresh = new Map<string, string>();
  const freshName = (prefix: string, value: string) => {
    const key = `${prefix}${value}`;
    const known =
      fresh.get(key) ?? `${prefix}#${[...fresh.keys()].filter((k) => k.startsWith(prefix)).length}`;
    fresh.set(key, known);
    return known;
  };
  return body.replace(
    /<([\w:]+)((?:\s+[\w:.-]+="[^"]*")*)\s*(\/?)>/g,
    (_match, name: string, attrs: string, close: string) => {
      const list = [...attrs.matchAll(/([\w:.-]+)="([^"]*)"/g)]
        .map((m): [string, string] => [m[1] ?? '', m[2] ?? ''])
        .filter(([key]) => !key.startsWith('xmlns') && !DROPPED.has(key))
        .map(([key, value]): [string, string] => {
          if (key === 'w:id' && REVISION_TAGS.has(name)) {
            const known = ids.get(value) ?? `#${ids.size}`;
            ids.set(value, known);
            return [key, known];
          }
          if (key === 'w:id' && BOOKMARK_TAGS.has(name)) {
            const known = bookmarks.get(value) ?? `b${bookmarks.size}`;
            bookmarks.set(value, known);
            return [key, known];
          }
          if (key === 'r:id') return [key, resolve(value) ?? `?${value}`];
          if (MEDIA_KEYS.has(key)) return [key, resolve(value) ?? freshName('media', value)];
          if (key === 'id' && name === 'wp:docPr') return [key, freshName('docPr', value)];
          return [key, value];
        })
        .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
      return `<${name}${list.map(([k, v]) => ` ${k}="${v}"`).join('')}${close}>`;
    }
  );
}

/**
 * The same, with run boundaries that change nothing removed: runs with one face merge, wrappers
 * of one revision merge, an empty `w:rPr` goes. Equal at this level: the same meaning.
 */
export function semanticBody(normalized: string): string {
  let out = normalized
    .replace(/<w:rPr\/>/g, '')
    .replace(/<w:rPr><\/w:rPr>/g, '')
    .replace(/<w:pPr\/>/g, '')
    .replace(/<w:pPr><\/w:pPr>/g, '')
    .replace(/<w:(trPr|tcPr)\/>/g, '')
    .replace(/<w:(trPr|tcPr)><\/w:\1>/g, '')
    .replace(/ xml:space="preserve"/g, '');
  for (;;) {
    const next = out
      .replace(/<(w:[A-Za-z]+)((?: [^>]*)?)><\/\1>/g, '<$1$2/>')
      .replace(
        /<w:r>((?:<w:rPr>(?:<w:rPr>(?:(?!<\/w:rPr>).)*<\/w:rPr>|(?!<\/w:rPr>|<w:rPr>).)*<\/w:rPr>)?)((?!<w:rPr>)(?:(?!<w:r[ >]|<\/w:r>).)*)<\/w:r><w:r>\1((?!<w:rPr>)(?:(?!<w:r[ >]|<\/w:r>).)*)<\/w:r>/g,
        '<w:r>$1$2$3</w:r>'
      )
      .replace(/<w:(t|delText)>([^<]*)<\/w:\1><w:\1>([^<]*)<\/w:\1>/g, '<w:$1>$2$3</w:$1>')
      .replace(
        /<\/w:(ins|del)><w:\1( [^>]*)>/g,
        (match, name: string, attrs: string, offset: number, all: string) => {
          const open = all.lastIndexOf(`<w:${name} `, offset);
          const openAttrs = /^<w:(?:ins|del)( [^>]*)>/.exec(all.slice(open))?.[1];
          return openAttrs === attrs ? '' : match;
        }
      );
    if (next === out) return out;
    out = next;
  }
}

/** Without revision ids: equal when only the grouping of revisions into ids differs. */
export function withoutRevisionIds(normalized: string): string {
  return semanticBody(normalized.replace(/ w:id="#\d+"/g, ''));
}

/**
 * Word's move markup read as the struck and inserted text it stands for: the engine has no
 * move, so its cut and paste is the reference for a move.
 */
export function asDeleteAndInsert(normalized: string): string {
  const seen = new Map<string, string>();
  return normalized
    .replace(/<w:move(From|To)Range(Start|End)[^>]*\/>/g, '')
    .replace(
      /<w:moveFrom( [^>]*)?>([\s\S]*?)<\/w:moveFrom>/g,
      (_m, attrs: string | undefined, inner: string) =>
        `<w:del${attrs ?? ''}>${inner
          .replace(/<w:t( [^>]*)?>/g, '<w:delText$1>')
          .replace(/<\/w:t>/g, '</w:delText>')}</w:del>`
    )
    .replace(/<w:moveTo( [^>]*)?>/g, '<w:ins$1>')
    .replace(/<\/w:moveTo>/g, '</w:ins>')
    .replace(/ w:id="#\d+"/g, (match) => {
      const known = seen.get(match) ?? ` w:id="#${seen.size}"`;
      seen.set(match, known);
      return known;
    });
}

export type Grade = 'exact' | 'same' | 'ids' | 'WRONG';

function gradeOf(a: string, b: string): Grade {
  if (a === b) return 'exact';
  if (semanticBody(a) === semanticBody(b)) return 'same';
  if (withoutRevisionIds(a) === withoutRevisionIds(b)) return 'ids';
  return 'WRONG';
}

/** How two normalized bodies compare, moves read as deletions and insertions. */
export const grade = (a: string, b: string): Grade =>
  gradeOf(asDeleteAndInsert(a), asDeleteAndInsert(b));

export const agrees = (g: Grade): boolean => g === 'exact' || g === 'same';

/** The first place two strings differ, with context. */
export function firstDifference(left: string, right: string, context = 140): string {
  let at = 0;
  while (at < left.length && left[at] === right[at]) at += 1;
  const from = Math.max(0, at - context);
  return `model:  …${left.slice(from, at + context)}…\nengine: …${right.slice(from, at + context)}…`;
}
