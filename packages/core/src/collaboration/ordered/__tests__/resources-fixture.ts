// What the resource tests share: a change applied as the room applies one, made of the edits
// a group's change kind will make (a suggestion, its resources, and the content that uses them).

import { buildInlinePictureDrawing } from '../../../store/package/drawing-package-edit.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlDrawingNode } from '../../../store/package/ooxml-tree.ts';
import { orderedIdMinter } from '../change/ids.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import { runOrderedEdit, type OrderedEdit } from '../edits/apply.ts';
import { contextOf, withNewIds, type Ctx, type OrderedEditContext } from '../edits/context.ts';
import { insertText } from '../edits/text.ts';
import { newSuggestion, setStatus } from '../model/mutations.ts';
import type { OrderedDocument } from '../model/types.ts';
import { withUndoRecord, type OrderedUndoRecord } from '../model/undo-record.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { APPLY_FILE, BODY, DATE } from './change-apply-fixture.ts';
import { testAttribution } from './test-persons.ts';

/** A copy of the apply fixture, with the package it was loaded from. */
export function openFixture(): { doc: OrderedDocument; base: OoxmlPackage } {
  const opened = openOrderedWordFile(APPLY_FILE, { file: 1 });
  return { doc: opened.document, base: opened.base };
}

/** What a change's edits get: the working context, its suggestion, and the edit context. */
export interface ChangeRun {
  readonly ctx: Ctx;
  readonly suggestion: string;
  /** The edit context that joins the change's suggestion. */
  readonly edit: OrderedEditContext;
  /** Run one model edit in the change's suggestion. */
  run(edit: OrderedEdit): void;
}

/**
 * Apply one change as the room does: one undo record, one minter for its id, one suggestion.
 * `direct` makes it the organiser's direct edit: its suggestion is accepted at once.
 */
export function applyAs(
  doc: OrderedDocument,
  changeId: string,
  author: string,
  body: (run: ChangeRun) => void,
  direct = false
): OrderedUndoRecord {
  return withUndoRecord(doc, () => {
    const ids = orderedIdMinter(changeId);
    const attribution = testAttribution(author, DATE);
    const base: OrderedEditContext = { doc, story: BODY, attribution, ids };
    const ctx = contextOf(base, []);
    const suggestion = newSuggestion(doc, attribution, ids.suggestion()).id;
    const edit: OrderedEditContext = { ...base, join: suggestion };
    body({ ctx, suggestion, edit, run: (one) => void runOrderedEdit(edit, one) });
    if (direct) setStatus(doc, suggestion, 'accepted');
  }).record;
}

/** A decision on one suggestion, in its own record. */
export function decide(
  doc: OrderedDocument,
  suggestion: string,
  status: 'accepted' | 'rejected'
): void {
  withUndoRecord(doc, () => setStatus(doc, suggestion, status));
}

/** Bytes that read as a PNG: its signature and some data. */
export function pngBytes(seed: number): Uint8Array {
  const bytes = new Uint8Array(64);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  for (let k = 8; k < bytes.length; k += 1) bytes[k] = (seed * 31 + k) & 255;
  return bytes;
}

/** A PNG's change-format media value. */
export const media = (bytes: Uint8Array) =>
  ({ sha256: orderedBlobHash(bytes), contentType: 'image/png', size: bytes.byteLength }) as const;

/** Type a picture whose image is relationship `relationshipId` at a paragraph's start. */
export function typePicture(run: ChangeRun, index: number, relationshipId: string): void {
  const view = orderedViewParagraph(run.ctx.doc, BODY, index);
  if (!view) throw new Error(`no paragraph ${index}`);
  const drawing = buildInlinePictureDrawing({
    docPrId: 77,
    relationshipId,
    extentEmu: { cx: 457_200, cy: 457_200 },
  });
  if (drawing.kind === 'textValue') throw new Error('no drawing');
  const content = withNewIds(run.ctx, drawing as OoxmlDrawingNode);
  insertText({ ...run.ctx, join: run.suggestion }, view, 0, '', {
    pieces: [{ text: '￼', content }],
  });
}
