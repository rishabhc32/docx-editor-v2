// A comment's `w:comment` element, built in Word's own shape: each body paragraph carries its
// `w14:paraId`, the `CommentText` style and a run in the `CommentReference` style.

import { COMMENTS_PART } from '../package/comment-lifecycle-scan.ts';
import { WML_NAMESPACE_URI, type OoxmlElement, type OoxmlNode } from '../package/ooxml-tree.ts';
import { W14_NAMESPACE_URI, XML_NAMESPACE_URI } from '../package/ooxml-shared.ts';
import type { AddCommentRequest } from './comment-writes.ts';

function wmlAttr(localName: string, value: string): OoxmlElement['attributes'][number] {
  return {
    kind: 'genericExtension' as const,
    namespaceUri: WML_NAMESPACE_URI,
    localName,
    prefix: 'w',
    value,
  };
}

/**
 * Build an element with an explicit KIND.
 *
 * Typed rather than generic on purpose: everything that reads comments matches on kind, so a
 * body written as generic nodes would serialize correctly and be invisible to the reader in the
 * same session — the comment would appear only after a save and reopen.
 */
export function element(
  id: string,
  kind: OoxmlElement['kind'],
  namespaceUri: string,
  prefix: string,
  localName: string,
  attributes: OoxmlElement['attributes'],
  children: readonly OoxmlNode[],
  bindings: readonly { prefix: string; namespaceUri: string }[] = []
): OoxmlElement {
  return {
    id,
    kind,
    namespaceUri,
    localName,
    prefix,
    namespaceBindings: bindings,
    attributes,
    children,
  } as OoxmlElement;
}

/** One paragraph of a comment's body: its plain text and its minted `w14:paraId`. */
export interface CommentParagraphText {
  readonly paraId: string;
  readonly text: string;
}

/**
 * `<w:comment>` with one paragraph of plain text for each of `paragraphs`, each carrying its
 * minted `w14:paraId`. The first paragraph's nodes take the ids a one-paragraph comment's take.
 */
export function commentElementOf(
  commentId: string,
  request: Pick<AddCommentRequest, 'author' | 'initials' | 'date'>,
  paragraphs: readonly CommentParagraphText[],
  paraIdPrefix: string
): OoxmlElement {
  const base = `${COMMENTS_PART}#comment-${commentId}`;
  return element(
    base,
    'comment',
    WML_NAMESPACE_URI,
    'w',
    'comment',
    [
      wmlAttr('id', commentId),
      wmlAttr('author', request.author),
      ...(request.initials === undefined ? [] : [wmlAttr('initials', request.initials)]),
      ...(request.date === undefined ? [] : [wmlAttr('date', request.date)]),
    ],
    paragraphs.map((paragraph, k) =>
      commentParagraph(k === 0 ? base : `${base}.${k}`, paragraph, paraIdPrefix)
    )
  );
}

/** One `<w:p>` of a comment's body, its node ids under `base`. */
function commentParagraph(
  base: string,
  { paraId, text: value }: CommentParagraphText,
  paraIdPrefix: string
): OoxmlElement {
  // `xml:space="preserve"` when the body has an edge space, which is what a conformant
  // reader is entitled to drop — Word writes the attribute for exactly this case. Omitted
  // otherwise, so an ordinary reply adds no attribute the file did not need.
  const preservesSpace = value !== value.trim();
  const text = element(
    `${base}.t`,
    'text',
    WML_NAMESPACE_URI,
    'w',
    't',
    preservesSpace
      ? [
          {
            kind: 'xmlSpace' as const,
            namespaceUri: XML_NAMESPACE_URI,
            localName: 'space' as const,
            prefix: 'xml' as const,
            value: 'preserve' as const,
          },
        ]
      : [],
    [{ id: `${base}.tv`, kind: 'textValue', value }]
  );
  // Word's own shape: the body carries `CommentText`, and the run carries
  // `CommentReference`. Without them a comment renders in body formatting and the pane loses
  // the styling every other editor gives it.
  const runProperties = element(
    `${base}.rPr`,
    'runProperties',
    WML_NAMESPACE_URI,
    'w',
    'rPr',
    [],
    [
      element(
        `${base}.rStyle`,
        'generic',
        WML_NAMESPACE_URI,
        'w',
        'rStyle',
        [wmlAttr('val', 'CommentReference')],
        []
      ),
    ]
  );
  const run = element(`${base}.r`, 'run', WML_NAMESPACE_URI, 'w', 'r', [], [runProperties, text]);
  const paragraphProperties = element(
    `${base}.pPr`,
    'paragraphProperties',
    WML_NAMESPACE_URI,
    'w',
    'pPr',
    [],
    [
      element(
        `${base}.pStyle`,
        'generic',
        WML_NAMESPACE_URI,
        'w',
        'pStyle',
        [wmlAttr('val', 'CommentText')],
        []
      ),
    ]
  );
  return element(
    `${base}.p`,
    'paragraph',
    WML_NAMESPACE_URI,
    'w',
    'p',
    [
      {
        kind: 'genericExtension' as const,
        namespaceUri: W14_NAMESPACE_URI,
        localName: 'paraId',
        prefix: paraIdPrefix,
        value: paraId,
      },
    ],
    [paragraphProperties, run]
  );
}
