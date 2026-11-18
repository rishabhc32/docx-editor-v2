// Random sessions with content controls (group 9.6, task 9.11): the organiser and two parties
// type (over prompts too), delete, press Enter, Backspace and Delete, add controls around words
// and at a caret, remove them with or without their words, change their tag, title and lock and
// a list's value, with undo and redo, while the room's frames arrive in random amounts. After
// each round: every copy equals the room, every screen shows its list, and every refusal is one
// of the client's named reasons. At the end each editor's save reopens as the room's document.

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import type { OoxmlNode } from '../../../../store/package/ooxml-tree.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedParagraphCount, orderedViewParagraph } from '../../compose/compose.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { seededRandom } from '../random.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
  join,
  localRoom,
  ORGANISER,
  PARTY,
  sameAsRoom,
  screenFollowsList,
  surface,
  type Person,
} from './editor-harness.ts';

afterEach(closeAll);

const PROMPT_FACE = '<w:rStyle w:val="PlaceholderText"/>';
const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

/** A contract template: inline and block controls, with prompts (one of two paragraphs), a value and a list. */
const FILE = docx(
  p(
    t('Name: ') +
      sdt(
        '<w:alias w:val="Name"/><w:tag w:val="name"/><w:showingPlcHdr/><w:text/>',
        t('Click here.', PROMPT_FACE)
      ) +
      t(' signs.')
  ) +
    p(t('The fee is ') + sdt('<w:tag w:val="fee"/><w:richText/>', t('100 USD')) + t(' a month.')) +
    p(
      t('Plan: ') +
        sdt(
          '<w:tag w:val="plan"/><w:dropDownList>' +
            '<w:listItem w:displayText="Basic" w:value="basic"/>' +
            '<w:listItem w:displayText="Pro" w:value="pro"/></w:dropDownList>',
          t('Basic')
        ) +
        t('.')
    ) +
    sdt(
      '<w:tag w:val="terms"/><w:showingPlcHdr/>',
      p(t('Enter the terms.', PROMPT_FACE)) + p(t('Add more terms.', PROMPT_FACE))
    ) +
    sdt('<w:tag w:val="notes"/>', p(t('Notes apply.'))) +
    p(t('Signed by both parties.'))
);

/** Reasons a refusal may have in a random session: each names a real conflict. */
const NAMED = new Set([
  'pending-refused',
  'no-inverse',
  'redo-refused',
  'hidden-target',
  'empty-range',
  'no-paragraph-before',
  'text-differs',
  'out-of-step',
  'unknown-address',
  'unknown-suggestion',
  'invalid-range',
  'stale-revert',
  'break-in-container',
  'single-line-control',
  // A lock the organiser set on a control another person changed
  'locked',
  // The party rules that the host's earlier check also applies (the room runs the suggestions-only
  // policy)
  'nested',
  'text',
  'package',
  'others_suggestions',
]);

/**
 * Control edits a concurrent edit makes impossible on the copy they reach, by reason and detail:
 * a range that came to cut a field or control, and words a concurrent Enter put in two
 * paragraphs.
 */
const CONFLICTS: readonly (readonly [string, RegExp])[] = [
  ['engine-refused', /^insertContentControl: indivisible-content$/],
  ['invalid-edit', /^a new control around words of one paragraph$/],
];

/** A refusal that names a real conflict. */
const named = (refusal: { readonly reason: string; readonly detail?: string }): boolean =>
  NAMED.has(refusal.reason) ||
  CONFLICTS.some(
    ([reason, detail]) => refusal.reason === reason && detail.test(refusal.detail ?? '')
  );

async function drain(queue: (() => void)[]): Promise<void> {
  for (let guard = 0; guard < 10_000; guard += 1) {
    await Promise.resolve();
    await Promise.resolve();
    const next = queue.shift();
    if (next) {
      next();
      continue;
    }
    await new Promise((resolve) => setTimeout(resolve, 0));
    if (queue.length === 0) return;
  }
  throw new Error('the session does not settle');
}

/** The ids of the content controls an editor shows, in document order. */
function controlIds(person: Person): string[] {
  const ids: string[] = [];
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'contentControl') ids.push(node.id);
    for (const child of node.children) visit(child);
  };
  visit(surface(person).session.part().root);
  return ids;
}

/** An op through the editor's gated write, as the editing API sends it. */
function applyOp(person: Person, op: TreeDocOp): void {
  const gated = surface(person) as unknown as {
    applyGatedTreeOps(ops: readonly TreeDocOp[]): unknown;
  };
  gated.applyGatedTreeOps([op]);
}

function act(person: Person, rand: () => number): void {
  const s = surface(person);
  const texts = editorTexts(person);
  const index = Math.floor(rand() * texts.length);
  const length = texts[index]?.length ?? 0;
  const at = Math.floor(rand() * (length + 1));
  const pick = <T>(values: readonly T[]): T | undefined =>
    values[Math.floor(rand() * values.length)];
  const roll = rand();
  if (roll < 0.22) {
    caretAt(person, index, at);
    s.insertPlainText(pick(['a', 'bc', 'Ana ', 'X']) ?? 'a');
  } else if (roll < 0.3) {
    caretAt(person, index, at);
    s.deleteBackward();
  } else if (roll < 0.35) {
    caretAt(person, index, at);
    s.deleteForward();
  } else if (roll < 0.4) {
    const end = Math.min(length, at + 1 + Math.floor(rand() * 6));
    caretAt(person, index, at, end);
    s.deleteSelection();
  } else if (roll < 0.45) {
    caretAt(person, index, at);
    s.splitParagraph();
  } else if (roll < 0.53) {
    // A new control around words, or at the caret
    const end = rand() < 0.5 ? at : Math.min(length, at + 1 + Math.floor(rand() * 5));
    caretAt(person, index, at, end);
    person.editor.exec({
      type: 'insertContentControl',
      subtype: pick(['richText', 'plainText'] as const) ?? 'richText',
      tag: `t${at}`,
    });
  } else if (roll < 0.6) {
    const control = pick(controlIds(person));
    if (control !== undefined)
      applyOp(person, {
        op: 'removeContentControl',
        controlId: control,
        keepContent: rand() < 0.6,
      });
  } else if (roll < 0.67) {
    const control = pick(controlIds(person));
    if (control !== undefined)
      applyOp(person, {
        op: 'setContentControlProperties',
        controlId: control,
        ...(rand() < 0.5 ? { tag: `tag${at}` } : { alias: `Title ${at}` }),
        ...(rand() < 0.3 ? { lock: pick(['sdtLocked', 'unlocked'] as const) ?? 'unlocked' } : {}),
      });
  } else if (roll < 0.73) {
    // The list's value, from a caret in it
    const plan = texts.findIndex((text) => text.startsWith('Plan: '));
    if (plan < 0) return;
    caretAt(person, plan, 'Plan: '.length + 1);
    person.editor.exec({ type: 'setContentControlValue', value: pick(['basic', 'pro']) ?? 'pro' });
  } else if (roll < 0.87) s.undo();
  else s.redo();
}

function textsOf(doc: OrderedDocument): string[] {
  const out: string[] = [];
  for (let k = 0; k < orderedParagraphCount(doc, BODY); k += 1)
    out.push(
      orderedViewParagraph(doc, BODY, k)
        ?.chars.map((c) => c.ch)
        .join('') ?? ''
    );
  return out;
}

async function session(seed: number, steps: number): Promise<void> {
  const rand = seededRandom(seed);
  const queue: (() => void)[] = [];
  const room = localRoom(FILE, {
    deliver: (run) => queue.push(run),
    policy: suggestionsOnlyPolicy(),
  });
  const people: Person[] = [];
  for (const name of ['Olga', 'Ana', 'Ben']) {
    // Olga is the organiser: she edits directly
    const joining = join(room, name, {
      wordFile: FILE,
      ...(name === 'Olga' ? { mode: 'edit' as const, login: ORGANISER } : { login: PARTY }),
    });
    await drain(queue);
    people.push(await joining);
  }
  for (let round = 0; round < 3; round += 1) {
    for (let step = 0; step < steps; step += 1) {
      const person = people[Math.floor(rand() * people.length)];
      if (person) act(person, rand);
      await Promise.resolve();
      await Promise.resolve();
      // Some frames arrive now, the rest wait: people edit on copies that are behind
      for (let k = Math.floor(rand() * 4); k > 0; k -= 1) queue.shift()?.();
      if (rand() < 0.3) await new Promise((resolve) => setTimeout(resolve, 0));
    }
    await drain(queue);
    await new Promise((resolve) => setTimeout(resolve, 0));
    await drain(queue);
    const label = `seed ${seed} round ${round}`;
    for (const person of people) {
      const unnamed = person.refusals.filter((r) => !named(r));
      expect({
        label,
        who: person.name,
        unnamed,
        status: person.handle.session.status(),
        room: sameAsRoom(person, room),
        screen: screenFollowsList(person),
      }).toEqual({
        label,
        who: person.name,
        unnamed: [],
        status: 'ready',
        room: true,
        screen: true,
      });
    }
  }
  expect(room.room.order).toBeGreaterThan(steps);
  // Each editor's save reopens as the editors show the room's document (an emptied control's
  // prompt too, which the written part shows again)
  for (const person of people) {
    const saved = new Uint8Array(await person.editor.save());
    expect(readOoxmlPackage(saved).ok).toBe(true);
    expect({
      seed,
      who: person.name,
      texts: textsOf(openOrderedWordFile(saved, { file: 2 }).document),
    }).toEqual({ seed, who: person.name, texts: editorTexts(person) });
  }
}

/** More seeds for a local sweep: ORDERED_SWEEP=50 ORDERED_SWEEP_FROM=200. */
const SWEEP = Number(process.env.ORDERED_SWEEP ?? 0);
const FROM = Number(process.env.ORDERED_SWEEP_FROM ?? 100);

describe('random sessions with content controls', () => {
  for (let seed = FROM; seed < FROM + SWEEP; seed += 1)
    test(`sweep, seed ${seed}`, async () => {
      await session(seed, 30);
    });
  for (const seed of [1, 2, 3, 4])
    test(`the organiser and two parties, seed ${seed}`, async () => {
      await session(seed, 30);
    });
});
