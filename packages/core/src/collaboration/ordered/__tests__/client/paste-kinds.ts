// Paste with formatting at every Word item kind of the matrix (task 9.11 for group 9.10). At the
// caret inside each kind of the body (after it, for a kind with no inside), a party pastes a
// title and two list items, and the organiser pastes a line with bold text and a link. Checks:
//
//  - the room's list composes what the editor alone writes for the same paste (`compare.ts`;
//    a party's paste may number Word's revision ids per role, where the engine gives one);
//  - every copy equals the room and its screen shows its list;
//  - the saved file reopens with the same view and the same decisions (`sessions/session.ts`);
//  - a reject of the party's paste and an undo of the organiser's give the text back, and a
//    redo makes the paste again.
//
// A kind that does not pass every check is on `KNOWN` with its cause and owner, and must give
// exactly the refusals it lists.

import { describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import type { OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { relationshipsOf } from '../../../../store/package/package-edit.ts';
import { paragraphTextOf } from '../../../../store/store/paragraph-model-text.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { AFTER } from '../matrix/actions.ts';
import { bodyOf, extrasOf, KINDS, paragraphOf, type MatrixKind } from '../matrix/catalogue.ts';
import { firstDifference, grade, normalizedBody, semanticBody } from '../matrix/compare.ts';
import { caretOf } from '../matrix/engine.ts';
import { checkSaveAndLoad } from '../sessions/session.ts';
import { docx, packageOf } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { copied, copiedLine, type Copied } from './paste-harness.ts';
import { reviewTestModule } from './review-module.ts';

type Check = 'alone' | 'reopen' | 'rejected-text' | 'decided-screen';

interface KnownPaste {
  readonly cause: string;
  readonly owner: 'both' | 'host' | 'engine' | 'model' | 'harness';
  readonly kinds: readonly string[];
  /** The refusals the party's paste gets, in order; the paste then changes nothing. */
  readonly party?: readonly string[];
  /** The refusals the organiser's paste gets; the paste then changes nothing. */
  readonly organiser?: readonly string[];
  /** The checks the cause makes fail, for the party's paste or the organiser's. */
  readonly skips?: { readonly party?: readonly Check[]; readonly organiser?: readonly Check[] };
}

const KNOWN: readonly KnownPaste[] = [
  {
    cause:
      'A caret between the two halves of a character outside the basic plane: both refuse (`splits-surrogate-pair`); the editor never puts a caret there',
    owner: 'both',
    kinds: ['emoji'],
    party: ['splits-surrogate-pair'],
    organiser: ['splits-surrogate-pair'],
  },
  {
    cause:
      'Paragraphs pasted into a control that holds one line: both refuse (`single-line-control`), as Word refuses Enter there; a line without a paragraph mark goes in',
    owner: 'both',
    kinds: ['plain', 'placeholder', 'dropdown', 'combo', 'date', 'dropprompt', 'dateprompt'],
    party: ['single-line-control'],
  },
  {
    cause:
      'A control bound to custom XML: both refuse an edit of its content (`not-covered`; the engine keeps it from edits)',
    owner: 'both',
    kinds: ['bound'],
    party: ['not-covered'],
    organiser: ['not-covered'],
  },
  {
    cause:
      "A group control's own content: both refuse an edit of it (`locked`), as Word does; the controls nested in a group take edits",
    owner: 'both',
    kinds: ['group'],
    party: ['locked'],
    organiser: ['locked'],
  },
  {
    cause:
      "A party's paste inside another person's pending insertion: the nesting rule (`nested`), reported once for the paste; its later changes name the refused suggestion and go with it",
    owner: 'host',
    kinds: ['ins', 'rowins'],
    party: ['nested'],
    skips: { party: ['alone'] },
  },
  {
    cause:
      "A file with a text box does not reopen with the same paragraph ids, with no edit at all: the text box's two copies (group 9.9)",
    owner: 'harness',
    kinds: ['textbox'],
    skips: { party: ['reopen'] },
  },
];

const known = (kind: MatrixKind) => KNOWN.filter((entry) => entry.kinds.includes(kind.key));
const refusalsOf = (kind: MatrixKind, who: 'party' | 'organiser'): readonly string[] =>
  known(kind).find((entry) => entry[who] !== undefined)?.[who] ?? [];
const skips = (kind: MatrixKind, who: 'party' | 'organiser', check: Check): boolean =>
  known(kind).some((entry) => entry.skips?.[who]?.includes(check) === true);

const review = reviewTestModule();

/** The body kinds: a kind in a header, a footer or a note has no body paste here. */
const BODY_KINDS = KINDS.filter((kind) => kind.story === undefined);

/** List numbers renumbered by first use: the room numbers a new list from the change id. */
function renumbered(xml: string): string {
  const seen = new Map<string, number>();
  return xml.replace(/(<w:numId w:val=")(\d+)/g, (_match, head: string, id: string) => {
    if (!seen.has(id)) seen.set(id, seen.size + 1);
    return `${head}${seen.get(id)}`;
  });
}

/** Relationship targets of the body: the package's, and the room's new ones. */
function resolverOf(pkg: OoxmlPackage, doc?: OrderedDocument) {
  const records = relationshipsOf(pkg, BODY);
  const added = new Map<string, string>();
  for (const resource of doc?.resources.added.values() ?? [])
    if (
      resource.kind === 'relationship' &&
      resource.owner === BODY &&
      'external' in resource.target
    )
      added.set(resource.id, resource.target.external);
  return (id: string) => added.get(id) ?? records.find((r) => r.id === id)?.rawTarget ?? null;
}

/** The caret offset inside the kind's paragraph, in a paragraph's text. */
function caretIn(kind: MatrixKind, text: string): number {
  const selector = kind.inside ?? AFTER;
  return text.includes(selector.find) ? caretOf(text, selector) : text.length;
}

function putCaret(person: Person, kind: MatrixKind): void {
  const s = surface(person);
  const id = s.session.paragraphIds()[paragraphOf(kind)];
  const text = id === undefined ? '' : (paragraphTextOf(s.session.part(), id) ?? '');
  caretAt(person, paragraphOf(kind), caretIn(kind, text));
}

/** What the editor alone writes for the same paste, normalized. */
function alone(file: Uint8Array, kind: MatrixKind, party: boolean, flavours: Copied): string {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: file,
    author: party ? 'Ana' : 'Olive',
    mode: party ? 'suggesting' : 'edit',
    modules: [review],
  });
  try {
    const s = editor.surface;
    if (!s) throw new Error('no surface');
    const id = s.session.paragraphIds()[paragraphOf(kind)] ?? '';
    const offset = caretIn(kind, paragraphTextOf(s.session.part(), id) ?? '');
    s.setSelection({ anchor: { paragraphId: id, offset }, head: { paragraphId: id, offset } });
    s.pasteRich(flavours.text, flavours.html);
    return renumbered(normalizedBody(s.session.part(), resolverOf(s.session.currentPackage())));
  } finally {
    editor.destroy();
    container.remove();
  }
}

/** The room's list, normalized as the editor's is. */
function listed(room: ReturnType<typeof localRoom>, file: Uint8Array): string {
  const doc = room.room.document;
  return renumbered(
    normalizedBody(composeOrderedPart(doc, BODY).part, resolverOf(packageOf(file), doc))
  );
}

function expectAgrees(label: string, list: string, engine: string, ids: boolean): void {
  const g = grade(list, engine);
  const agreed = g === 'exact' || g === 'same' || (ids && g === 'ids');
  expect({
    label,
    grade: agreed ? 'agrees' : g,
    difference: agreed ? '' : firstDifference(semanticBody(list), semanticBody(engine)),
  }).toEqual({
    label,
    grade: 'agrees',
    difference: '',
  });
}

function expectFollows(
  label: string,
  room: ReturnType<typeof localRoom>,
  people: Person[],
  screen = true
): void {
  for (const person of people)
    expect({
      label,
      person: person.name,
      same: sameAsRoom(person, room),
      screen: !screen || screenFollowsList(person),
    }).toEqual({
      label,
      person: person.name,
      same: true,
      screen: true,
    });
}

async function open(file: Uint8Array) {
  const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
  const olive = await join(room, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    review,
    wordFile: file,
  });
  const ana = await join(room, 'Ana', { review, wordFile: file });
  return { room, olive, ana };
}

/** A party pastes a title and two list items; the organiser rejects it. */
async function partyPaste(kind: MatrixKind, file: Uint8Array, flavours: Copied): Promise<void> {
  const label = `${kind.key}, party`;
  const { room, olive, ana } = await open(file);
  const before = editorTexts(olive);
  putCaret(ana, kind);
  surface(ana).pasteRich(flavours.text, flavours.html);
  await settle(12);
  const refused = refusalsOf(kind, 'party');
  expect({ label, refusals: ana.refusals.map((r) => r.reason) }).toEqual({
    label,
    refusals: [...refused],
  });
  expectFollows(label, room, [olive, ana]);
  if (refused.length > 0) {
    expect(editorTexts(olive)).toEqual(before);
    return;
  }
  if (!skips(kind, 'party', 'alone'))
    expectAgrees(label, listed(room, file), alone(file, kind, true, flavours), true);
  if (!skips(kind, 'party', 'reopen')) checkSaveAndLoad(room.room.document, label, packageOf(file));
  for (const item of olive.editor.getReviewItems())
    if (item.kind === 'revision' && item.author === 'Ana') olive.editor.rejectReviewItem(item.key);
  await settle(12);
  expectFollows(`${label}, rejected`, room, [olive, ana], !skips(kind, 'party', 'decided-screen'));
  if (!skips(kind, 'party', 'rejected-text'))
    expect({ label, texts: editorTexts(olive) }).toEqual({ label, texts: before });
}

/** The organiser pastes a line with bold text and a link, undoes it and does it again. */
async function organiserPaste(kind: MatrixKind, file: Uint8Array, flavours: Copied): Promise<void> {
  const label = `${kind.key}, organiser`;
  const { room, olive, ana } = await open(file);
  const before = editorTexts(olive);
  putCaret(olive, kind);
  surface(olive).pasteRich(flavours.text, flavours.html);
  await settle(12);
  const refused = refusalsOf(kind, 'organiser');
  expect({ label, refusals: olive.refusals.map((r) => r.reason) }).toEqual({
    label,
    refusals: [...refused],
  });
  expectFollows(label, room, [olive, ana]);
  if (refused.length > 0 || room.room.order === 0) {
    // Refused by the room, or by the editor before it (a table of contents)
    expect(editorTexts(olive)).toEqual(before);
    return;
  }
  if (!skips(kind, 'organiser', 'alone'))
    expectAgrees(label, listed(room, file), alone(file, kind, false, flavours), false);
  const pasted = editorTexts(ana);
  expect(olive.handle.session.undo()).toBe(true);
  await settle(12);
  expectFollows(
    `${label}, undone`,
    room,
    [olive, ana],
    !skips(kind, 'organiser', 'decided-screen')
  );
  expect({ label, texts: editorTexts(ana) }).toEqual({ label, texts: before });
  expect(olive.handle.session.redo()).toBe(true);
  await settle(12);
  expectFollows(`${label}, done again`, room, [olive, ana]);
  expect({ label, texts: editorTexts(ana) }).toEqual({ label, texts: pasted });
}

/** One shard of the body kinds, as tests. */
export function pasteKindsShard(index: number, count: number): void {
  describe(`paste with formatting at every item kind, shard ${index + 1} of ${count}`, () => {
    let lines: { multi: Copied; line: Copied } | null = null;
    const flavours = () => (lines ??= { multi: copied(3), line: copiedLine() });
    for (const kind of BODY_KINDS.filter((_, k) => k % count === index))
      test(`${kind.key}: ${kind.name}`, async () => {
        try {
          const file = docx(bodyOf(kind), extrasOf(kind));
          await partyPaste(kind, file, flavours().multi);
          closeAll();
          await organiserPaste(kind, file, flavours().line);
        } finally {
          closeAll();
        }
      }, 30_000);
  });
}

/** Every entry of `KNOWN` names kinds of the body. */
export function knownKindsExist(): void {
  test('every known difference names kinds of the body', () => {
    const keys = new Set(BODY_KINDS.map((kind) => kind.key));
    for (const entry of KNOWN)
      for (const key of entry.kinds)
        expect({ key, known: keys.has(key) }).toEqual({ key, known: true });
  });
}
