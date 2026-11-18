// An equation is one atom of the list.

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { OFFICE_MATH_NAMESPACE_URI } from '../../../store/package/omml-equation.ts';
import { REVISION_ROLE } from '../model/revision-ids.ts';
import { attributeOf, isWml, withChildren } from '../ooxml-nodes.ts';
import type { LoadState } from './state.ts';

/** An equation atom's node and the suggestion that inserted or struck it whole. */
export interface EquationAtom {
  readonly node: OoxmlElement;
  readonly pieces: undefined;
  readonly by: string | null;
  readonly dels: string[];
  readonly fmts: [];
}

const isMath = (node: OoxmlNode, localName: string): node is OoxmlElement =>
  node.kind !== 'textValue' &&
  node.namespaceUri === OFFICE_MATH_NAMESPACE_URI &&
  node.localName === localName;

const isMark = (node: OoxmlNode): node is OoxmlElement => isWml(node, 'ins') || isWml(node, 'del');

/** The mark of each math run, when every run's content is one mark and no other mark is in it. */
function runMarks(equation: OoxmlElement): OoxmlElement[] | null {
  const marks: OoxmlElement[] = [];
  let all = 0;
  const stack: OoxmlNode[] = [equation];
  for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
    if (node.kind === 'textValue') continue;
    if (isMark(node)) all += 1;
    if (isMath(node, 'r')) {
      const content = node.children.filter((child) => !isMath(child, 'rPr'));
      const [mark] = content;
      if (content.length !== 1 || mark === undefined || !isMark(mark)) return null;
      marks.push(mark);
    }

    // In document order: the first run's mark keeps its `w:id`
    for (let k = node.children.length - 1; k >= 0; k -= 1) stack.push(node.children[k]!);
  }

  return marks.length > 0 && marks.length === all ? marks : null;
}

/** The node without the marks `marks`: their content in their place. */
function unmarked(node: OoxmlNode, marks: ReadonlySet<OoxmlNode>): OoxmlNode {
  if (node.kind === 'textValue' || node.children.length === 0) return node;
  return withChildren(
    node,
    node.children.flatMap((child) =>
      marks.has(child) && child.kind !== 'textValue'
        ? child.children.map((inner) => unmarked(inner, marks))
        : [unmarked(child, marks)]
    )
  );
}

/**
 * An equation (`m:oMath` or `m:oMathPara`) read as one atom of a paragraph. Word writes a whole
 * equation deleted with tracking on as one `w:del` inside each math run (`m:r`), the `m:oMath`
 * kept (Word 16 for Mac, 2026-10-29), and an equation inserted the same way with `w:ins`
 * ([MS-OI29500] 22.1.2.87(e)). When every math run of the equation is marked so, in one
 * direction by one author at one date, the atom is struck (or inserted) by one suggestion, and
 * its node is the equation without the marks: compose writes the mark around the equation, the
 * form Word 16 reads as one deletion by its author (task 0.4). Marks on some runs only are file
 * content: the list has no unit inside an equation.
 */
export function equationAtom(
  state: LoadState,
  equation: OoxmlElement,
  markKey: string
): EquationAtom {
  const plain: EquationAtom = { node: equation, pieces: undefined, by: null, dels: [], fmts: [] };
  const marks = runMarks(equation);
  const [first] = marks ?? [];
  if (!marks || !first) return plain;
  const stamp = (mark: OoxmlElement) => {
    const author = attributeOf(mark, 'author') ?? '';
    return `${mark.localName}\u0000${author}\u0000${attributeOf(mark, 'date') ?? ''}`;
  };
  if (!marks.every((mark) => stamp(mark) === stamp(first))) return plain;
  const struck = first.localName === 'del';
  const role = struck ? REVISION_ROLE.deleteIn(markKey) : state.markRoles.run;
  const suggestion = state.revisions.suggestionForAll(marks, role);
  const node = unmarked(equation, new Set(marks)) as OoxmlElement;
  return {
    node,
    pieces: undefined,
    by: struck ? null : suggestion,
    dels: struck ? [suggestion] : [],
    fmts: [],
  };
}
