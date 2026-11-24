// A file's tracked changes outside the stories: the change records of the styles part (a
// style's `w:pPrChange` and `w:rPrChange`, a table style's `w:tblPrChange`, `w:trPrChange` and
// `w:tcPrChange`, also in a conditional format, and the document defaults'). Each review card
// is one pending suggestion with no items, so a decision only turns its status, and compose
// writes the part with the decided records resolved (`compose/part-revisions.ts`).

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { stylesPartOf } from '../../../store/package/ooxml-indexes.ts';
import {
  collectRevisionSites,
  revisionGroupKey,
  type RevisionSite,
} from '../../../store/store/tree-op-revisions.ts';
import { recordedProperties } from '../../../store/store/tree-op-tracked-properties.ts';
import { newSuggestion } from '../model/mutations.ts';
import { reserveFileRevisionId } from '../model/revision-ids.ts';
import type { OrderedDocument, OrderedPartRevision } from '../model/types.ts';
import type { OrderedLoadOptions } from './load.ts';
import { attributeOf, isWml } from '../ooxml-nodes.ts';

/** The style (`w:styleId`) each node of the part sits in, by node id. */
function stylesOfNodes(part: OoxmlPart): Map<string, string> {
  const out = new Map<string, string>();
  for (const style of part.root.children) {
    if (style.kind === 'textValue' || !isWml(style, 'style')) continue;
    const id = attributeOf(style, 'styleId');
    if (id === undefined) continue;
    const stack: OoxmlNode[] = [style];
    for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
      if (node.kind === 'textValue') continue;
      out.set(node.id, id);
      for (const child of node.children) if (child.kind !== 'textValue') stack.push(child);
    }
  }

  return out;
}

/** A change record's `w:id`, author and date; undefined for one with no `w:id` or author. */
export function recordRevision(node: OoxmlElement): OrderedPartRevision['revision'] | undefined {
  const id = attributeOf(node, 'id');
  const author = attributeOf(node, 'author');
  if (id === undefined || author === undefined) {
    return undefined;
  }

  const date = attributeOf(node, 'date');
  return { id, author, ...(date === undefined ? {} : { date }) };
}

/** A site the engine decides both ways: a property record with an author and old properties. */
const isDecidable = (site: RevisionSite): boolean =>
  site.propertyChange && !site.refused && recordedProperties(site.node) !== null;

/** One record group: its first site's identity, and whether every site of it is decidable. */
interface Group {
  readonly revision: OrderedPartRevision;
  decidable: boolean;
}

/**
 * Load the styles part's change records as suggestions, one for each card the editor shows: the
 * records of one element with one `w:id`, author and date (`revisionGroupKey`). Call after the
 * stories: in the part's order, every copy mints the same suggestion ids and the stories' ids
 * stay as they were. A record the engine cannot decide both ways (no author, no old properties)
 * makes none: it stays as the file has it, and the editor's Accept All in a room names it
 * (`not-in-room`). A group a room checkpoint lists (`record`) is that listed suggestion, under
 * its own id.
 */
export function loadPartRevisions(
  doc: OrderedDocument,
  pkg: OoxmlPackage,
  listed: OrderedLoadOptions['suggestions'] = []
): void {
  const part = stylesPartOf(pkg);
  if (!part) return;
  const kept = new Map<string, string>();
  for (const { id, attribution, revisions, record } of listed) {
    const revision = revisions[0];
    if (record !== undefined && revision !== undefined) {
      const { author, date } = attribution;
      kept.set(revisionGroupKey({ id: revision, author, date }, record), id);
    }
  }

  const groups = new Map<string, Group>();
  let styles: Map<string, string> | null = null;
  for (const site of collectRevisionSites(part)) {
    const id = attributeOf(site.node, 'id');
    if (id !== undefined) reserveFileRevisionId(doc, id);
    const revision = recordRevision(site.node);
    if (revision === undefined) {
      continue;
    }

    const key = revisionGroupKey(revision, site.node.localName);
    const known = groups.get(key);
    if (known) {
      known.decidable &&= isDecidable(site);
      continue;
    }

    styles ??= stylesOfNodes(part);
    const styleId = styles.get(site.node.id);
    groups.set(key, {
      revision: {
        part: part.name,
        element: site.node.localName,
        revision,
        ...(styleId === undefined ? {} : { styleId }),
      },
      decidable: isDecidable(site),
    });
  }

  for (const [key, group] of groups) {
    if (!group.decidable) continue;
    const { author, date } = group.revision.revision;
    const id =
      kept.get(key) ??
      newSuggestion(doc, {
        author,
        person: null,
        ...(date === undefined ? {} : { date }),
      }).id;
    doc.partRevisions.set(id, group.revision);
  }
}
