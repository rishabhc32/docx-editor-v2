// The attributes only a checkpoint's parts hold (design D14): a label on each element the
// manifest names, the typer on each typed run, and the room run's number on each run drawn from
// units. The writer puts the labels on after compose; every open takes all of them off before the
// load.

import { glossaryPartOf } from '../../../store/package/building-blocks.ts';
import { withPart, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type {
  OoxmlElement,
  OoxmlNamespaceBinding,
  OoxmlNode,
} from '../../../store/package/ooxml-tree.ts';
import { orderedStoryParts } from '../load/load.ts';
import { withAttributes, withChildren } from '../ooxml-nodes.ts';
import { CHECKPOINT_NAMESPACE, CHECKPOINT_PREFIX, checkpointAttribute } from './checkpoint.ts';

/** The declaration of the namespace a checkpoint writes on each part root that holds it. */
const BINDING: OoxmlNamespaceBinding = {
  prefix: CHECKPOINT_PREFIX,
  namespaceUri: CHECKPOINT_NAMESPACE,
};

/** Whether an attribute or a declaration is of a checkpoint's namespace. */
const isOurs = (item: { readonly namespaceUri: string }): boolean =>
  item.namespaceUri === CHECKPOINT_NAMESPACE;

/** A part root that declares the namespace a checkpoint writes, once. */
export const withCheckpointNamespace = (root: OoxmlElement): OoxmlElement =>
  root.namespaceBindings.some(isOurs)
    ? root
    : { ...root, namespaceBindings: [...root.namespaceBindings, BINDING] };

/** What the label walk gives each element: the labels so far, and whether the part holds ours. */
interface Labelling {
  readonly wanted: ReadonlySet<string>;
  readonly labels: Map<string, number>;
  marked: boolean;
}

/**
 * A composed package as a checkpoint writes it: a label on each element `wanted` names by node
 * id, each a number unique in the checkpoint, and the namespace declared on the root of each
 * part (a story or the glossary) that holds a label or another attribute compose wrote
 * (`marked`: it wrote one, a typer or a room run). Only the paths to a labelled element are made
 * again, and a package with nothing to mark is the same package. A text box's other copy takes the labels of its story when the package is
 * written.
 */
export function withCheckpointLabels(
  pkg: OoxmlPackage,
  wanted: ReadonlySet<string>,
  marked: boolean
): { readonly pkg: OoxmlPackage; readonly labels: ReadonlyMap<string, number> } {
  const labelling: Labelling = { wanted, labels: new Map(), marked: false };
  if (wanted.size === 0 && !marked) {
    return { pkg, labels: labelling.labels };
  }

  const glossary = glossaryPartOf(pkg);
  let out = pkg;
  for (const part of [...orderedStoryParts(pkg).values(), ...(glossary ? [glossary] : [])]) {
    labelling.marked = false;
    const root = labelled(part.root, labelling);
    if (labelling.marked) {
      out = withPart(out, { ...part, root: withCheckpointNamespace(root) });
    }
  }

  return { pkg: out, labels: labelling.labels };
}

/** An element with a label on each element below it, itself too, that the walk wants. */
function labelled(element: OoxmlElement, labelling: Labelling): OoxmlElement {
  let children: OoxmlNode[] | null = null;
  for (let k = 0; k < element.children.length; k += 1) {
    const child = element.children[k];
    if (child === undefined || child.kind === 'textValue') {
      continue;
    }

    const next = labelled(child, labelling);
    if (next !== child) {
      children ??= [...element.children];
      children[k] = next;
    }
  }

  const { wanted, labels } = labelling;
  const out = children === null ? element : withChildren(element, children);
  if (element.attributes.some(isOurs)) {
    labelling.marked = true;
  }

  if (!wanted.has(element.id) || labels.has(element.id)) {
    return out;
  }

  const label = labels.size;
  labels.set(element.id, label);
  labelling.marked = true;
  return withAttributes(out, [...out.attributes, checkpointAttribute('label', String(label))]);
}

/** What the parts of a package held of a checkpoint's attributes. */
export interface CheckpointMarks {
  /** The package with no attribute or namespace declaration of the namespace left. */
  readonly pkg: OoxmlPackage;
  /** The elements that carry each label, by its value: each copy of a text box's story too. */
  readonly labels: ReadonlyMap<string, readonly OoxmlElement[]>;
  /** The typer each element carries, as written, by the element's node id. */
  readonly typers: ReadonlyMap<string, string>;
  /** The room run's number each run and text element carries, by the element's node id. */
  readonly runs: ReadonlyMap<string, string>;
}

/**
 * A package with every attribute and namespace declaration of a checkpoint's namespace taken off,
 * in every part, and what they were. The elements given back are the package's: the ones a load
 * of it reads. Only the paths to an element that held one are made again.
 */
export function withoutCheckpointMarks(pkg: OoxmlPackage): CheckpointMarks {
  const labels = new Map<string, OoxmlElement[]>();
  const typers = new Map<string, string>();
  const runs = new Map<string, string>();
  let out = pkg;
  for (const part of pkg.parts.values()) {
    const root = unmarked(part.root, labels, typers, runs);
    if (root !== part.root) {
      out = withPart(out, { ...part, root });
    }
  }

  return { pkg: out, labels, typers, runs };
}

/** An element without the namespace's attributes and declarations, below it too. */
function unmarked(
  element: OoxmlElement,
  labels: Map<string, OoxmlElement[]>,
  typers: Map<string, string>,
  runs: Map<string, string>
): OoxmlElement {
  let children: OoxmlNode[] | null = null;
  for (let k = 0; k < element.children.length; k += 1) {
    const child = element.children[k];
    if (child === undefined || child.kind === 'textValue') {
      continue;
    }

    const next = unmarked(child, labels, typers, runs);
    if (next !== child) {
      children ??= [...element.children];
      children[k] = next;
    }
  }

  const kept = children === null ? element : withChildren(element, children);
  const marks = element.attributes.filter(isOurs);
  if (marks.length === 0 && !element.namespaceBindings.some(isOurs)) {
    return kept;
  }

  const attributes = element.attributes.filter((attribute) => !isOurs(attribute));
  const out = {
    ...withAttributes(kept, attributes),
    namespaceBindings: element.namespaceBindings.filter((binding) => !isOurs(binding)),
  };
  for (const { localName, value } of marks) {
    if (localName === 'label') {
      labels.set(value, [...(labels.get(value) ?? []), out]);
    } else if (localName === 'typer') {
      typers.set(out.id, value);
    } else if (localName === 'run') {
      runs.set(out.id, value);
    }
  }

  return out;
}
