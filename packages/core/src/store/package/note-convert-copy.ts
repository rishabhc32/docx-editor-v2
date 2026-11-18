// The copy a converted note puts in the other kind's notes part (`convertNote`).

import { createNodeIdAllocator } from './ooxml-edit.ts';
import type { OoxmlNode, OoxmlPart } from './ooxml-tree.ts';
import { WML_NAMESPACE_URI } from './ooxml-shared.ts';
import {
  noteIdOf,
  noteKindOf,
  noteRefKindOf,
  noteReferenceKindOf,
  type NoteKind,
} from './note-nodes.ts';
import { usedParaIds, w14PrefixInScopeAt, withFreshParaIds } from './para-id.ts';
import { actorScopedSeed } from './actor-scoped-ids.ts';

const W = WML_NAMESPACE_URI;

/** Built-in Word character styles for note marks, not user theme aliases. */
const BUILTIN_NOTE_REFERENCE_STYLE: Readonly<Record<NoteKind, string>> = {
  footnote: 'FootnoteReference',
  endnote: 'EndnoteReference',
};

function isWmlValAttribute(attr: {
  readonly localName: string;
  readonly namespaceUri: string;
}): boolean {
  return (
    attr.localName === 'val' &&
    (attr.namespaceUri === WML_NAMESPACE_URI || attr.namespaceUri === '')
  );
}

/**
 * Rewrite `w:rStyle/@w:val` on a run that directly holds a note citation or body mark.
 *
 * Only a run that holds a note reference or note mark (typed or generic) and a `w:rPr/w:rStyle`
 * whose `w:val` (or unprefixed `val`) is the source kind's built-in style changes. Any other
 * style id, and a run with no style, stays: Word needs no character style on a note mark, and
 * layout resolves marks through the element type.
 */
export function rewriteRunBuiltinNoteReferenceStyle(
  run: OoxmlNode,
  fromKind: NoteKind,
  toKind: NoteKind
): OoxmlNode {
  if (run.kind !== 'run' || fromKind === toKind) return run;
  const holdsNoteMark = run.children.some(
    (child) => noteReferenceKindOf(child) !== null || noteRefKindOf(child) !== null
  );
  if (!holdsNoteMark) return run;

  const fromStyle = BUILTIN_NOTE_REFERENCE_STYLE[fromKind];
  const toStyle = BUILTIN_NOTE_REFERENCE_STYLE[toKind];

  let changed = false;
  const children = run.children.map((child) => {
    if (child.kind !== 'runProperties') return child;
    let propsChanged = false;
    const props = child.children.map((prop) => {
      if (
        prop.kind !== 'generic' ||
        prop.namespaceUri !== WML_NAMESPACE_URI ||
        prop.localName !== 'rStyle'
      ) {
        return prop;
      }

      const hasFromVal = prop.attributes.some(
        (attr) => isWmlValAttribute(attr) && attr.value === fromStyle
      );
      if (!hasFromVal) return prop;
      propsChanged = true;
      changed = true;
      return {
        ...prop,
        attributes: prop.attributes.map((attr) =>
          isWmlValAttribute(attr) && attr.value === fromStyle ? { ...attr, value: toStyle } : attr
        ),
      };
    });
    return propsChanged ? { ...child, children: props } : child;
  });
  return changed ? ({ ...run, children } as OoxmlNode) : run;
}

function rewriteNoteNode(
  source: OoxmlNode,
  fromKind: NoteKind,
  toKind: NoteKind,
  newId: number,
  nextId: () => string
): OoxmlNode {
  const mapNode = (node: OoxmlNode): OoxmlNode => {
    if (node.kind === 'textValue') return { ...node, id: nextId() };
    let kind = node.kind;
    let name = node.localName;
    let attributes = node.attributes;
    if (node.id === source.id || (node.kind === 'note' && noteIdOf(node) === noteIdOf(source))) {
      kind = 'note';
      name = toKind;
      attributes = node.attributes.map((attr) =>
        attr.localName === 'id' ? { ...attr, value: String(newId) } : attr
      );
      if (!attributes.some((attr) => attr.localName === 'id')) {
        attributes = [
          ...attributes,
          {
            kind: 'genericExtension',
            namespaceUri: W,
            localName: 'id',
            prefix: 'w',
            value: String(newId),
          },
        ];
      }
    } else if (
      node.kind === 'noteRef' ||
      (noteKindOf(node) === null &&
        (node.localName === 'footnoteRef' || node.localName === 'endnoteRef'))
    ) {
      kind = 'noteRef';
      name = `${toKind}Ref`;
    }

    const mappedChildren = node.children.map(mapNode);
    let result = {
      ...node,
      id: nextId(),
      kind,
      localName: name,
      attributes,
      children: mappedChildren,
    } as OoxmlNode;
    if (node.kind === 'run') {
      result = rewriteRunBuiltinNoteReferenceStyle(result, fromKind, toKind);
    }

    return result;
  };
  return mapNode(source);
}

/**
 * The source note as the target part's note `newId`: the other kind's names, new node ids, and a
 * new `w14:paraId` on each paragraph, unique in `targetPart`. A target part that binds no `w14`
 * prefix gets ids only on the paragraphs that already carried one. The source's paragraph ids are
 * unique only in the part it came from, and Word writes a unique id on every paragraph.
 */
export function convertedNoteCopy(
  targetPart: OoxmlPart,
  source: OoxmlNode,
  fromKind: NoteKind,
  toKind: NoteKind,
  newId: number
): OoxmlNode {
  const copy = rewriteNoteNode(source, fromKind, toKind, newId, createNodeIdAllocator(targetPart));
  return withFreshParaIds(
    copy,
    new Set(usedParaIds(targetPart.root)),
    actorScopedSeed(`${targetPart.name}:converted-note:${newId}`),
    { value: 0 },
    w14PrefixInScopeAt(targetPart, targetPart.root)
  );
}
