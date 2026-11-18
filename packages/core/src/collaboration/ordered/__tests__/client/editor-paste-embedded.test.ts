// An embedded object (a workbook in `w:object`) and ink (`w14:contentPart`) pasted (group 9.10).
// Word copies both with their parts. The editor alone does too: the parts travel as bytes that
// nothing opens. So does an ordered room: each part is media by hash, checked by its signature
// and never opened. The organiser's paste is a direct edit that undo and redo turn back and
// again; a party's is a tracked insertion that a reject takes out of the file with its part.
// The saved file holds the part, the room's body equals the editor alone's, and the file reopens
// with the same view.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import type { OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { resolveInternalTarget } from '../../../../store/package/opc-names.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { relationshipsOf } from '../../../../store/package/package-edit.ts';
import { orderedBlobHash } from '../../protocol/blobs.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { checkSaveAndLoad } from '../sessions/session.ts';
import { packageOf } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  FILE,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
} from './editor-harness.ts';
import { copiedEmbedded, EMBEDDED, EMBEDDED_WORKBOOK, INK } from './paste-embedded-harness.ts';
import {
  agreesAlone,
  allChecked,
  changeKinds,
  pastedAlone,
  resolvesAll,
  revisionCards,
  savedFile,
} from './paste-room-checks.ts';
import { reviewTestModule } from './review-module.ts';

const review = reviewTestModule();

afterEach(closeAll);

const BEFORE = ['The fee is 100 USD.', 'Due in 30 days.', 'Signed here.'];

/** The bytes of the part the body's first match of `pattern` names by its relationship. */
function partNamed(pkg: OoxmlPackage, pattern: RegExp): Uint8Array | undefined {
  const xml = serializeOoxmlPart(pkg.parts.get(BODY)!);
  const id = /r:id="([^"]+)"/.exec(pattern.exec(xml)?.[0] ?? '')?.[1];
  const record = relationshipsOf(pkg, BODY).find((r) => r.id === id);
  const target = record ? resolveInternalTarget(BODY, record.rawTarget) : null;
  return target?.ok ? pkg.partBytes.get(target.partName) : undefined;
}

/** The parts kept as bytes of a package: embedded objects' and ink. */
const keptParts = (pkg: OoxmlPackage): string[] =>
  [...pkg.partBytes.keys()]
    .map((name) => (name.startsWith('/') ? name : `/${name}`))
    .filter((name) => /^\/word\/(embeddings|ink)\//.test(name));

/** The saved file resolves every relationship, and its pasted part holds `bytes`. */
function savedHolds(
  room: ReturnType<typeof localRoom>,
  kind: keyof typeof EMBEDDED,
  bytes: Uint8Array
) {
  const pkg = savedFile(room);
  expect(resolvesAll(pkg)).toBe(true);
  expect(partNamed(pkg, EMBEDDED[kind].shows)).toEqual(bytes);
  expect(keptParts(pkg)).toHaveLength(1);
}

const KINDS = [
  { kind: 'object', name: 'an embedded workbook', bytes: EMBEDDED_WORKBOOK },
  { kind: 'ink', name: 'ink', bytes: INK },
  { kind: 'inkWpi', name: "ink in Word 2010's fuller form", bytes: INK },
] as const;

for (const { kind, name, bytes } of KINDS)
  describe(`${name} pasted`, () => {
    test('the editor alone pastes it with its part, as Word does', () => {
      const container = document.createElement('div');
      document.body.append(container);
      const editor = createDocxEditor({ container, document: FILE, author: 'Olive' });
      try {
        const s = editor.surface;
        if (!s) throw new Error('no surface');
        const id = s.session.paragraphIds()[1] ?? '';
        s.setSelection({
          anchor: { paragraphId: id, offset: 0 },
          head: { paragraphId: id, offset: 0 },
        });
        const flavours = copiedEmbedded(kind);
        expect(s.pasteRich(flavours.text, flavours.html)).toBe(true);
        expect(partNamed(s.session.currentPackage(), EMBEDDED[kind].shows)).toEqual(bytes);
      } finally {
        editor.destroy();
        container.remove();
      }
    });

    test("the organiser's: every copy shows it as the editor alone does; undo and redo", async () => {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
      const ana = await join(room, 'Ana', { review });
      const flavours = copiedEmbedded(kind);
      caretAt(olive, 1, 3);
      expect(surface(olive).pasteRich(flavours.text, flavours.html)).toBe(true);
      await settle(16);
      expect(changeKinds(room)).toEqual(['insert-fragment']);
      allChecked(room, [olive, ana]);
      expect(room.blob(orderedBlobHash(bytes))).toEqual(bytes);
      expect(editorXml(ana)).toMatch(EMBEDDED[kind].shows);
      expect(editorXml(ana)).not.toContain('<w:ins ');
      // Every copy holds the part's bytes, read back from the room by hash
      expect(partNamed(surface(ana).session.currentPackage(), EMBEDDED[kind].shows)).toEqual(bytes);
      savedHolds(room, kind, bytes);
      const alone = pastedAlone(flavours, false, 1, 3, review, (pkg) =>
        expect(partNamed(pkg, EMBEDDED[kind].shows)).toEqual(bytes)
      );
      expect(agreesAlone(room, alone, false)).toBe('agrees');
      checkSaveAndLoad(room.room.document, kind, packageOf(FILE), {
        blobs: (hash) => room.blob(hash),
      });

      expect(olive.handle.session.undo()).toBe(true);
      await settle(16);
      allChecked(room, [olive, ana]);
      expect(editorTexts(ana)).toEqual(BEFORE);
      expect(keptParts(savedFile(room))).toEqual([]);
      expect(olive.handle.session.redo()).toBe(true);
      await settle(16);
      allChecked(room, [olive, ana]);
      expect(editorXml(ana)).toMatch(EMBEDDED[kind].shows);
      savedHolds(room, kind, bytes);
    });

    test("a party's: a tracked insertion; a reject takes it and its part out of the file", async () => {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
      const ana = await join(room, 'Ana', { review });
      const flavours = copiedEmbedded(kind);
      caretAt(ana, 0, 4);
      expect(surface(ana).pasteRich(flavours.text, flavours.html)).toBe(true);
      await settle(16);
      expect(changeKinds(room)).toEqual(['insert-fragment']);
      allChecked(room, [olive, ana]);
      expect(editorXml(olive)).toMatch(/<w:ins [^>]*w:author="Ana"/);
      expect(editorXml(olive)).toMatch(EMBEDDED[kind].shows);
      savedHolds(room, kind, bytes);
      expect(agreesAlone(room, pastedAlone(flavours, true, 0, 4, review), true)).toBe('agrees');
      checkSaveAndLoad(room.room.document, kind, packageOf(FILE), {
        blobs: (hash) => room.blob(hash),
      });

      const [card] = revisionCards(olive);
      if (card === undefined) throw new Error('no card');
      expect(olive.editor.rejectReviewItem(card).ok).toBe(true);
      await settle(16);
      allChecked(room, [olive, ana]);
      expect(editorTexts(olive)).toEqual(BEFORE);
      expect(editorXml(olive)).not.toMatch(EMBEDDED[kind].shows);
      expect(keptParts(savedFile(room))).toEqual([]);
    });
  });
