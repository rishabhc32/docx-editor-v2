// Checks the paste tests of group 9.10 share for a paste in an ordered room: the room's saved
// file and its relationships, each copy against the room, the room's body against what the
// editor alone writes for the same paste, and the review cards a copy shows.

import { expect } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import type { EditorModule } from '../../../../contracts/modules.ts';
import { readOoxmlPackage, type OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { resolveInternalTarget } from '../../../../store/package/opc-names.ts';
import type { OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { relationshipsOf } from '../../../../store/package/package-edit.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import { grade, normalizedBody } from '../matrix/compare.ts';
import { packageOf } from '../word-fixtures.ts';
import {
  BODY,
  FILE,
  sameAsRoom,
  screenFollowsList,
  type localRoom,
  type Person,
} from './editor-harness.ts';
import type { Copied } from './paste-harness.ts';

/** A room of the editor harness. */
export type PasteRoom = ReturnType<typeof localRoom>;

/** The kinds of the changes the room ordered, in order. */
export const changeKinds = (room: PasteRoom): string[] =>
  room.room.joinSnapshot(0).map((change) => change.kind);

/** Each person refused nothing, holds the room's list and shows it on screen. */
export function allChecked(room: PasteRoom, people: readonly Person[]): void {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
  }
}

/** The room's saved file, with its media read from the room's storage. */
export function savedFile(room: PasteRoom): OoxmlPackage {
  const bytes = orderedPackageBytes(room.room.document, packageOf(FILE), {
    blobs: (hash) => room.blob(hash),
  });
  const pkg = readOoxmlPackage(bytes);
  if (!pkg.ok) throw new Error('the saved file does not open');
  return pkg.package;
}

/**
 * Whether every relationship the body names resolves to a part the package holds, and each
 * named part's own relationships do too.
 */
export function resolvesAll(pkg: OoxmlPackage): boolean {
  const pending = [BODY];
  const seen = new Set<string>();
  for (let owner = pending.pop(); owner !== undefined; owner = pending.pop()) {
    if (seen.has(owner)) continue;
    seen.add(owner);
    for (const record of relationshipsOf(pkg, owner)) {
      if (record.targetMode === 'External') continue;
      const target = resolveInternalTarget(owner, record.rawTarget);
      if (!target.ok || !(pkg.partBytes.has(target.partName) || pkg.parts.has(target.partName)))
        return false;
      if (pkg.parts.has(target.partName)) pending.push(target.partName);
    }
  }
  return true;
}

/** The revision cards a copy shows, by key. */
export const revisionCards = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .filter((placement) => placement.kind === 'revision')
    .map((placement) => placement.key);

/** A body without its relationship ids, which each copy mints. */
const comparable = (part: OoxmlPart): string =>
  normalizedBody(part, () => null).replace(/ r:(id|dm|lo|qs|cs)="[^"]*"/g, ' r:$1="?"');

/**
 * What the editor alone writes for the same paste at the same place: a party's in suggesting
 * mode, the organiser's in editing mode. `holds` checks the editor's own package.
 */
export function pastedAlone(
  flavours: Copied,
  party: boolean,
  paragraph: number,
  offset: number,
  review: EditorModule,
  holds: (pkg: OoxmlPackage) => void = () => {}
): string {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: FILE,
    author: party ? 'Ana' : 'Olive',
    mode: party ? 'suggesting' : 'edit',
    modules: [review],
  });
  try {
    const s = editor.surface;
    if (!s) throw new Error('no surface');
    const id = s.session.paragraphIds()[paragraph] ?? '';
    s.setSelection({ anchor: { paragraphId: id, offset }, head: { paragraphId: id, offset } });
    expect(s.pasteRich(flavours.text, flavours.html)).toBe(true);
    holds(s.session.currentPackage());
    return comparable(s.session.part());
  } finally {
    editor.destroy();
    container.remove();
  }
}

/** Whether the room's body agrees with the editor alone's (a party's may differ in ids only). */
export function agreesAlone(room: PasteRoom, engine: string, party: boolean): string {
  const g = grade(comparable(composeOrderedPart(room.room.document, BODY).part), engine);
  return g === 'exact' || g === 'same' || (party && g === 'ids') ? 'agrees' : g;
}
