// Picture and shape ops (group 9.4): insert (the bytes travel once, named by their content
// hash, through the room server: design D14), delete, resize, crop, position, wrap, transform,
// locks, metadata and replace.

import type {
  DrawingTreeDocOp,
  TreeDocOp,
  TreeDocOpKind,
} from '../../../../store/store/tree-op-types.ts';
import type { OrderedObjectChangeOp } from '../../edits/pictures.ts';
import { addressOf, placeAt } from '../paragraphs.ts';
import { newObjectOf, objectOf } from '../pictures.ts';
import { OrderedClientRefusal } from '../refusal.ts';
import {
  translator,
  type OrderedOpOption,
  type OrderedOpTable,
  type OrderedPartyRule,
} from './types.ts';

/**
 * Word has no tracked form for these: a party's, made with tracking off, is not a suggestion.
 * What Word writes decides what a party may do (design D13). A new picture or text box is a
 * tracked insertion of its run, and a deletion a tracked deletion (the editor sends a tracked
 * `deleteText` over the picture's unit; `deleteDrawing` is its untracked form). No revision
 * element of ECMA-376 can hold an earlier DrawingML property or image (`wml.xsd` lines 803 to
 * 955; `CT_Drawing`, `CT_Inline`, `CT_Anchor` and `CT_NonVisualDrawingProps` have no change
 * child). So those are direct edits: the organiser's, taken back with `revert`; the editor
 * refuses them in suggesting mode, and a party's untracked one is refused by the room's rule
 * (`package`).
 */
const untracked = (what: string): OrderedPartyRule => ({
  code: 'policy-direct',
  reason: 'package',
  why: `Word has no tracked form: ${what}`,
});

const PROPERTY = "an object's frame or its properties: one `object-property` change";

/** The property ops, which one `object-property` change carries. */
const PROPERTY_OPS: ReadonlySet<TreeDocOpKind> = new Set([
  'resizeDrawing',
  'cropDrawing',
  'positionDrawing',
  'setDrawingWrap',
  'setDrawingMetadata',
  'setDrawingLocks',
  'transformDrawing',
]);

type PropertyOp = Extract<
  DrawingTreeDocOp,
  {
    readonly op:
      | 'resizeDrawing'
      | 'cropDrawing'
      | 'positionDrawing'
      | 'setDrawingWrap'
      | 'setDrawingMetadata'
      | 'setDrawingLocks'
      | 'transformDrawing';
  }
>;

const isProperty = (op: TreeDocOp | undefined): op is PropertyOp =>
  op !== undefined && PROPERTY_OPS.has(op.op);

/** A whole EMU: the engine writes the value rounded, so the change says the same. */
function emu(value: number, what: string): number {
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded) || rounded <= 0)
    throw new OrderedClientRefusal('ordered-refused', 'invalid-change', what);
  return rounded;
}

/** One property op as the change format says it. */
function changeOpOf(op: PropertyOp): OrderedObjectChangeOp {
  switch (op.op) {
    case 'resizeDrawing':
      return { op: 'resize', cx: emu(op.extentEmu.cx, 'cx'), cy: emu(op.extentEmu.cy, 'cy') };
    case 'cropDrawing':
      // The engine writes each edge rounded (`cropPermille`)
      return {
        op: 'crop',
        crop: {
          left: Math.round(op.crop.left),
          top: Math.round(op.crop.top),
          right: Math.round(op.crop.right),
          bottom: Math.round(op.crop.bottom),
        },
      };
    case 'positionDrawing': {
      const { mode, horizontalEmu, verticalEmu, relativeToH, relativeToV } = op.position;
      return {
        op: 'position',
        position: {
          ...(mode === undefined ? {} : { mode }),
          ...(horizontalEmu === undefined ? {} : { horizontalEmu }),
          ...(verticalEmu === undefined ? {} : { verticalEmu }),
          ...(relativeToH === undefined ? {} : { relativeToH }),
          ...(relativeToV === undefined ? {} : { relativeToV }),
        },
      };
    }
    case 'setDrawingWrap':
      return { op: 'wrap', wrap: op.wrap };
    case 'setDrawingMetadata':
      return {
        op: 'metadata',
        title: op.title,
        description: op.description,
        ...(op.hyperlink === undefined ? {} : { link: op.hyperlink }),
      };
    case 'setDrawingLocks': {
      const { select, move, resize, changeAspect } = op.locks;
      return {
        op: 'locks',
        locks: {
          ...(select === undefined ? {} : { select }),
          ...(move === undefined ? {} : { move }),
          ...(resize === undefined ? {} : { resize }),
          ...(changeAspect === undefined ? {} : { changeAspect }),
        },
      };
    }
    case 'transformDrawing':
      return { op: 'transform', action: op.action };
  }
}

/**
 * A property op's translator. The property ops of a batch on one object are one change: one
 * direct edit and one undo step, as the editor's one command is (a wrap and the position that
 * keeps the object in place; a size and a crop). The first op of a run takes the ops after it on
 * the same object; they make none.
 */
const property = (kind: PropertyOp['op']) =>
  translator(kind, (op, batch) => {
    const ops = batch.ops;
    const at = ops.indexOf(op);
    const before = ops[at - 1];
    if (at > 0 && isProperty(before) && before.drawingNodeId === op.drawingNodeId) return () => [];
    const run: PropertyOp[] = [op];
    for (let k = at + 1; at >= 0 && k < ops.length; k += 1) {
      const next = ops[k];
      if (!isProperty(next) || next.drawingNodeId !== op.drawingNodeId) break;
      run.push(next);
    }

    // Read now: an op the format cannot say refuses the batch before anything applies
    const changeOps = run.map(changeOpOf);
    return () => {
      const object = objectOf(batch, op.drawingNodeId);
      return [
        {
          kind: 'object-property',
          part: object.story.name,
          object: addressOf(object.atom),
          ops: changeOps,
        },
      ];
    };
  });

export const PICTURE_OPS: OrderedOpTable<
  | 'insertDrawing'
  | 'replaceDrawingResource'
  | 'deleteDrawing'
  | 'resizeDrawing'
  | 'cropDrawing'
  | 'positionDrawing'
  | 'setDrawingWrap'
  | 'setDrawingMetadata'
  | 'setDrawingLocks'
  | 'transformDrawing'
> = {
  insertDrawing: {
    status: 'translated',
    group: '9.4',
    changes: ['insert-object'],
    party: null,
    note: 'A picture or a text box at a place, tracked as an insertion of its run; its media a resource named by hash.',
    // Read in the step: an earlier op of the batch may make the paragraph (a paste)
    translate: translator('insertDrawing', (op, batch) => () => {
      const p = batch.paragraph(op.paragraphId);
      const part = p.story.name;
      const object = newObjectOf(batch, op.drawing, part);
      return [{ kind: 'insert-object', part, place: placeAt(p, op.offset), object }];
    }),
  },
  replaceDrawingResource: {
    status: 'translated',
    group: '9.4',
    changes: ['object-resource'],
    party: untracked("a picture's image"),
    note: "A picture's image replaced: new media for the same object, a direct edit.",
    translate: translator('replaceDrawingResource', (op, batch) => () => {
      const object = objectOf(batch, op.drawingNodeId);
      const part = object.story.name;
      const media = batch.media(part, op.relationshipId);
      if (media === null)
        throw new OrderedClientRefusal('ordered-unsupported-op', 'replaceDrawingResource.media');
      return [{ kind: 'object-resource', part, object: addressOf(object.atom), media }];
    }),
  },
  deleteDrawing: {
    status: 'translated',
    group: '9.4',
    changes: ['delete'],
    party: {
      code: 'policy-direct',
      reason: 'text',
      why: "a picture removed with tracking off; suggesting mode sends `deleteText` over the picture's unit",
    },
    note: "A picture or a shape removed: a deletion of the object's unit.",
    // With a revision (the engine refuses it; the editor sends `deleteText`) it is the same
    // deletion as a suggestion
    translate: translator('deleteDrawing', (op, batch) => {
      return () => {
        const { paragraph, offset } = objectOf(batch, op.drawingNodeId);
        return [
          {
            kind: 'delete',
            part: paragraph.story.name,
            from: placeAt(paragraph, offset),
            to: placeAt(paragraph, offset + 1),
          },
        ];
      };
    }),
  },
  resizeDrawing: {
    status: 'translated',
    group: '9.4',
    changes: ['object-property'],
    party: untracked("a picture's size"),
    note: `A new extent. ${PROPERTY}.`,
    translate: property('resizeDrawing'),
  },
  cropDrawing: {
    status: 'translated',
    group: '9.4',
    changes: ['object-property'],
    party: untracked("a picture's crop"),
    note: `A new crop. ${PROPERTY}.`,
    translate: property('cropDrawing'),
  },
  positionDrawing: {
    status: 'translated',
    group: '9.4',
    changes: ['object-property'],
    party: untracked("a floating object's position"),
    note: `A floating object's position. ${PROPERTY}.`,
    translate: property('positionDrawing'),
  },
  setDrawingWrap: {
    status: 'translated',
    group: '9.4',
    changes: ['object-property'],
    party: untracked("an object's wrap"),
    note: `Inline or a wrap around text. ${PROPERTY}.`,
    translate: property('setDrawingWrap'),
  },
  setDrawingMetadata: {
    status: 'translated',
    group: '9.4',
    changes: ['object-property'],
    party: untracked("an object's title, description and link"),
    note: `Title, description and a link (a relationship resource). ${PROPERTY}.`,
    translate: property('setDrawingMetadata'),
  },
  setDrawingLocks: {
    status: 'translated',
    group: '9.4',
    changes: ['object-property'],
    party: untracked("an object's locks"),
    note: `What a reader may change of the object. ${PROPERTY}.`,
    translate: property('setDrawingLocks'),
  },
  transformDrawing: {
    status: 'translated',
    group: '9.4',
    changes: ['object-property'],
    party: untracked("an object's rotation and flips"),
    note: `A quarter turn or a flip. ${PROPERTY}.`,
    translate: property('transformDrawing'),
  },
};

/** Options of the picture ops that change their meaning. */
export const PICTURE_OPTIONS: readonly OrderedOpOption[] = [];
