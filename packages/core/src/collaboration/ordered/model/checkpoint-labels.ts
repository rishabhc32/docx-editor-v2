// The attributes only a checkpoint's parts hold (design D14): a label on each element the
// manifest names, the typer on each typed run, the room run's number on each run drawn from
// units, and a mark on properties drawn only for change records. The version 5 writer put them
// on; every open takes all of them off before the load.

import { withPart, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type {
  OoxmlElement,
  OoxmlNamespaceBinding,
  OoxmlNode,
} from '../../../store/package/ooxml-tree.ts';
import { withAttributes, withChildren } from '../ooxml-nodes.ts';
import { CHECKPOINT_NAMESPACE, CHECKPOINT_PREFIX } from './checkpoint.ts';

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
  /** The node ids of the properties drawn only for change records (`oc:drawn`). */
  readonly drawn: ReadonlySet<string>;
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
  const drawn = new Set<string>();
  let out = pkg;
  for (const part of pkg.parts.values()) {
    const root = unmarked(part.root, labels, typers, runs, drawn);
    if (root !== part.root) {
      out = withPart(out, { ...part, root });
    }
  }

  return { pkg: out, labels, typers, runs, drawn };
}

/** An element without the namespace's attributes and declarations, below it too. */
function unmarked(
  element: OoxmlElement,
  labels: Map<string, OoxmlElement[]>,
  typers: Map<string, string>,
  runs: Map<string, string>,
  drawn: Set<string>
): OoxmlElement {
  let children: OoxmlNode[] | null = null;
  for (let k = 0; k < element.children.length; k += 1) {
    const child = element.children[k];
    if (child === undefined || child.kind === 'textValue') {
      continue;
    }

    const next = unmarked(child, labels, typers, runs, drawn);
    if (next !== child) {
      children ??= [...element.children];
      children[k] = next;
    }
  }

  const kept = children === null ? element : withChildren(element, children);
  if (!element.attributes.some(isOurs) && !element.namespaceBindings.some(isOurs)) {
    return kept;
  }

  const marks = element.attributes.filter(isOurs);
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
    } else if (localName === 'drawn') {
      drawn.add(out.id);
    }
  }

  return out;
}
