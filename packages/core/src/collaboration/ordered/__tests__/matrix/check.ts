// One matrix scenario through the model and the engine, side by side, with every check:
//
//   V  the view equals the engine's (exactly, or with the same meaning);
//   O  every choice of decisions gives one result in every order (up to 6 orders);
//   D  every choice equals the engine deciding the same revisions;
//   W  the saved Word file reopens as the file, and its choices equal the engine on the file;
//   L  every choice made live equals the same choice made after the save and reopen;
//   I  the composed part breaks the tree's rules.
//
// A choice decides groups of suggestions: the suggestions of each action, joined when one
// review card holds suggestions of several actions (a decision applies to a whole card).
// Refusals: S by the model only, E by the engine only, R by both. N: a later step's words are
// gone after the first edit, so it does not apply.

import type { OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { validateOoxmlPart, type OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { textboxFallbackExportPackage } from '../../../../store/package/textbox-fallback-export.ts';
import {
  applyOrderedChange,
  applyOrderedEdit,
  composeOrderedPart,
  loadOrderedPackage,
  orderedCards,
  orderedIdMinter,
  OrderedRefusal,
  setOrderedStatus,
  takeBackOrdered,
  withOrderedUndoRecord,
  type OrderedChange,
  type OrderedDocument,
  type OrderedEdit,
} from '../../index.ts';
import { settleDecidedTables } from '../../change/settle-columns.ts';
import { isTextBoxStoryName } from '../../model/text-boxes.ts';
import { docx } from '../word-fixtures.ts';
import type { Scenario } from './actions.ts';
import {
  agrees,
  asDeleteAndInsert,
  firstDifference,
  grade,
  normalizedBody,
  semanticBody,
  type Grade,
  type RelationshipResolver,
} from './compare.ts';
import {
  DATE,
  engineDecide,
  engineText,
  enterStory,
  mount,
  performEngine,
  storyPart,
  togglesOff,
  triplesOf,
  type Triple,
} from './engine.ts';
import { BODY, editFor, textOf, viewOf, type ModelChange } from './model.ts';
import { reopened, resolverOf, saved } from './word.ts';
import { testAttribution, testPerson } from '../test-persons.ts';

export interface CellResult {
  readonly name: string;
  /** The failed checks and refusals, in `VODWLISERN` order; '' when every check passes. */
  readonly codes: string;
  readonly view?: Grade;
  /** Live choices; with one result in every order; equal to the engine's. */
  readonly choices: number;
  readonly stable: number;
  readonly agree: number;
  /** After a save and reopen: the view's grade; choices; one result; equal to the engine's. */
  readonly word?: {
    readonly view: Grade;
    readonly choices: number;
    readonly stable: number;
    readonly agree: number;
  };
  /** Live choices compared after the reopen, and how many gave the same result. */
  readonly live?: { readonly choices: number; readonly same: number };
  readonly refusal?: string;
  readonly details: readonly string[];
}

/** Decision groups: labels whose suggestions share a review card decide together. */
interface Group {
  readonly labels: readonly string[];
  readonly suggestions: readonly string[];
  readonly triples: readonly Triple[];
  readonly authors: string;
}

type Decision = 'accepted' | 'rejected';
const MAX_ORDERS = 6;

function permutations<T>(items: readonly T[]): T[][] {
  if (items.length <= 1) return [[...items]];
  return items.flatMap((item, i) =>
    permutations([...items.slice(0, i), ...items.slice(i + 1)]).map((rest) => [item, ...rest])
  );
}

/** The story as normalized XML after these decisions, in this order; the document is restored. */
function decided(
  doc: OrderedDocument,
  name: string,
  order: readonly (readonly [readonly string[], Decision])[],
  resolve: RelationshipResolver
): string {
  const { record } = withOrderedUndoRecord(doc, () => {
    const made: string[] = [];
    for (const [suggestions, status] of order)
      for (const id of suggestions)
        if (doc.suggestions.get(id)?.status === 'pending') {
          setOrderedStatus(doc, id, status);
          made.push(id);
        }

    settleDecidedTables(doc, made, orderedIdMinter('check:1'));
  });
  const xml = normalizedBody(composeOrderedPart(doc, name).part, resolve);
  takeBackOrdered(doc, record);
  return xml;
}

function groupsOf(
  doc: OrderedDocument,
  labels: ReadonlyMap<string, ReadonlySet<string>>,
  triples: ReadonlyMap<string, readonly Triple[]>
): Group[] {
  const pending = (id: string) => doc.suggestions.get(id)?.status === 'pending';
  const labelOf = new Map<string, string>();
  for (const [label, ids] of labels) for (const id of ids) if (pending(id)) labelOf.set(id, label);
  const parent = new Map<string, string>();
  const find = (x: string): string => {
    const up = parent.get(x);
    return up === undefined || up === x ? x : find(up);
  };
  for (const label of labels.keys()) parent.set(label, label);
  for (const card of orderedCards(doc)) {
    const touched = [
      ...new Set(
        card.suggestions.map((id) => labelOf.get(id)).filter((l): l is string => l !== undefined)
      ),
    ];
    const head = touched[0];
    if (head !== undefined)
      for (const other of touched.slice(1)) parent.set(find(other), find(head));
  }
  const groups = new Map<string, string[]>();
  for (const label of labels.keys()) {
    const root = find(label);
    groups.set(root, [...(groups.get(root) ?? []), label]);
  }
  const out: Group[] = [];
  for (const members of groups.values()) {
    const suggestions = members.flatMap((label) => [...(labels.get(label) ?? [])].filter(pending));
    if (suggestions.length === 0) continue;
    const authors = [
      ...new Set(suggestions.map((id) => doc.suggestions.get(id)?.attribution.author ?? '')),
    ]
      .sort()
      .join('+');
    out.push({
      labels: members,
      suggestions,
      triples: members.flatMap((label) => triples.get(label) ?? []),
      authors,
    });
  }
  return out;
}

/** Each choice of accept and reject over the groups. */
function choicesOver<T>(groups: readonly T[]): (readonly (readonly [T, Decision])[])[] {
  const out: (readonly (readonly [T, Decision])[])[] = [];
  for (let mask = 0; mask < 1 << groups.length; mask += 1)
    out.push(
      groups.map((g, i): readonly [T, Decision] => [g, (mask >> i) & 1 ? 'accepted' : 'rejected'])
    );
  return out;
}

const action = (status: Decision) => (status === 'accepted' ? 'accept' : 'reject');

/** The triples a reopened suggestion was loaded from. */
function triplesOfSuggestion(doc: OrderedDocument, id: string): Triple[] {
  const attribution = doc.suggestions.get(id)?.attribution;
  if (!attribution) return [];
  const ids = new Set<string>();
  for (const [key, value] of doc.revisionIds.byRole) if (key.startsWith(`${id}/`)) ids.add(value);
  return [...ids].flatMap((revision) => [
    {
      id: revision,
      author: attribution.author,
      ...(attribution.date === undefined ? {} : { date: attribution.date }),
    },
    // A `w:tblGridChange` beside the table's record: its `w:id` alone (`CT_Markup`)
    { id: revision, author: '' },
  ]);
}

function refused(name: string, model: string | null, engine: string | null, details: string[]) {
  const code = model !== null && engine !== null ? 'R' : model !== null ? 'S' : 'E';
  const refusal = [model && `model ${model}`, engine && `engine ${engine}`]
    .filter(Boolean)
    .join('; ');
  return { name, codes: code, choices: 0, stable: 0, agree: 0, refusal, details };
}

/** The name of the document's one text box story. */
function textBoxStoryOf(doc: OrderedDocument): string {
  const found = [...doc.stories.keys()].find(isTextBoxStoryName);
  if (found === undefined) throw new Error('the model has no text box story');
  return found;
}

/** Run one scenario and grade it. */
export function runScenario(scenario: Scenario): CellResult {
  const details: string[] = [];
  const name = scenario.story?.part ?? BODY;
  const first = scenario.actions[0];
  if (!first) throw new Error(`${scenario.name} has no action`);
  const { surface, close } = mount(docx(scenario.body, scenario.extras), first.author);
  try {
    enterStory(surface, scenario.story);
    const pkg: OoxmlPackage = surface.session.currentPackage();
    const doc = loadOrderedPackage(pkg);
    // A text box's edits go to its own story; compose writes it back into the body's part
    const story = scenario.story?.scope === 'textbox' ? textBoxStoryOf(doc) : name;
    const labels = new Map<string, Set<string>>();
    const triples = new Map<string, Triple[]>();
    if (doc.suggestions.size > 0) {
      labels.set('file', new Set(doc.suggestions.keys()));
      triples.set('file', triplesOf(storyPart(surface, scenario.story)));
    }
    for (const [k, step] of scenario.actions.entries()) {
      surface.setAuthor(step.author);
      const seen = textOf(viewOf(doc, story, step.p));
      const wanted = engineText(surface, scenario.story, step.p);
      if (seen !== wanted) {
        details.push(`before ${step.label}: model "${seen}" vs engine "${wanted}"`);
        return { name: scenario.name, codes: 'V', choices: 0, stable: 0, agree: 0, details };
      }
      let edit: OrderedEdit | ModelChange;
      try {
        const off = togglesOff(surface, scenario.story, step);
        edit = editFor(doc, story, step, off);
      } catch (error) {
        details.push(error instanceof Error ? error.message : String(error));
        return { name: scenario.name, codes: 'N', choices: 0, stable: 0, agree: 0, details };
      }
      let model: string | null = null;
      try {
        const id = `m:${k + 1}`;
        const outcome =
          'change' in edit
            ? withOrderedUndoRecord(doc, () =>
                applyOrderedChange(
                  doc,
                  {
                    v: 1,
                    id,
                    mode: step.mode === 'edit' ? 'edit' : 'suggest',
                    ...edit.change,
                  } as unknown as OrderedChange,
                  { person: testPerson(step.author), date: DATE, ids: orderedIdMinter(id) }
                )
              ).value
            : applyOrderedEdit(
                {
                  doc,
                  story,
                  attribution: testAttribution(step.author, DATE),
                  ids: orderedIdMinter(id),
                  ...(step.mode === 'edit' ? { direct: true } : {}),
                },
                edit
              ).value;
        if (step.mode !== 'edit') {
          const set = labels.get(step.label) ?? new Set<string>();
          for (const id of outcome.suggestions) set.add(id);
          labels.set(step.label, set);
        }
      } catch (error) {
        if (!(error instanceof OrderedRefusal)) throw error;
        model = `${step.label}: ${error.message}`;
      }
      surface.setEditingMode(step.mode === 'edit' ? 'edit' : 'suggest');
      const part = () => storyPart(surface, scenario.story);
      const before = new Set(triplesOf(part()).map((t) => `${t.id}|${t.author}`));
      const rejectedBefore = surface.state().lastRejection;
      let engine: string | null = null;
      try {
        performEngine(surface, scenario.story, step);
        const rejection = surface.state().lastRejection;
        if (rejection && rejection !== rejectedBefore) engine = `${step.label}: ${rejection}`;
      } catch (error) {
        engine = `${step.label}: ${error instanceof Error ? error.message : String(error)}`;
      }
      if (model !== null || engine !== null) return refused(scenario.name, model, engine, details);
      const added = triplesOf(part()).filter((t) => !before.has(`${t.id}|${t.author}`));
      triples.set(step.label, [...(triples.get(step.label) ?? []), ...added]);
    }
    const engine = {
      part: storyPart(surface, scenario.story),
      pkg: surface.session.currentPackage(),
    };
    return graded(scenario.name, name, doc, pkg, engine, labels, triples, details);
  } finally {
    close();
  }
}

function graded(
  scenario: string,
  name: string,
  doc: OrderedDocument,
  pkg: OoxmlPackage,
  engine: { readonly part: OoxmlPart; readonly pkg: OoxmlPackage },
  labels: ReadonlyMap<string, ReadonlySet<string>>,
  triples: ReadonlyMap<string, readonly Triple[]>,
  details: string[]
): CellResult {
  let codes = '';
  const resolve = resolverOf(engine.pkg, name, doc);
  const xml = (part: OoxmlPart, r = resolve) => normalizedBody(part, r);
  // The engine edits a text box's live copy and brings its VML fallback in step when it saves;
  // the model's compose draws both in step at once (`withTextboxFallbackInStep`)
  const engineXml = (part: OoxmlPart) =>
    xml(
      textboxFallbackExportPackage({
        ...engine.pkg,
        parts: new Map([[part.name, part]]),
      }).parts.get(part.name) ?? part
    );
  const composed = composeOrderedPart(doc, name).part;
  const valid = validateOoxmlPart(composed);
  if (!valid.ok) details.push(`invalid view: ${JSON.stringify(valid.issues.slice(0, 2))}`);
  const view = grade(xml(composed), engineXml(engine.part));
  if (!agrees(view)) {
    codes += 'V';
    details.push(
      `view (${view})\n${firstDifference(semanticBody(xml(composed)), semanticBody(engineXml(engine.part)))}`
    );
  }
  // Live choices against the engine
  const groups = groupsOf(doc, labels, triples);
  const choices = choicesOver(groups);
  let stable = 0;
  let agree = 0;
  for (const choice of choices) {
    const finals = permutations(choice)
      .slice(0, MAX_ORDERS)
      .map((order) =>
        decided(
          doc,
          name,
          order.map(([g, s]) => [g.suggestions, s] as const),
          resolve
        )
      );
    const final = finals[0] ?? '';
    if (finals.every((f) => f === final)) stable += 1;
    let part = engine.part;
    for (const [g, s] of choice) part = engineDecide(part, g.triples, action(s), engine.pkg);
    const g = grade(final, engineXml(part));
    if (agrees(g)) agree += 1;
    else
      details.push(
        `choice ${choice.map(([x, s]) => `${s === 'accepted' ? '+' : '-'}${x.labels.join('')}`).join(' ')} (${g})\n${firstDifference(semanticBody(final), semanticBody(engineXml(part)))}`
      );
  }
  if (stable < choices.length) codes += 'O';
  if (agree < choices.length) codes += 'D';
  // The saved file, reopened
  const file = saved(pkg, doc);
  const again = reopened(file);
  const filePart = file.parts.get(name);
  if (!filePart) throw new Error(`the saved file has no ${name}`);
  const fileResolve = resolverOf(file, name);
  const reopenedXml = () => xml(composeOrderedPart(again, name).part, fileResolve);
  const wordView = grade(reopenedXml(), xml(filePart, fileResolve));
  const cards = orderedCards(again).slice(0, 4);
  const wordChoices = choicesOver(cards);
  let wordStable = 0;
  let wordAgree = 0;
  for (const choice of wordChoices) {
    const finals = permutations(choice)
      .slice(0, MAX_ORDERS)
      .map((order) =>
        decided(
          again,
          name,
          order.map(([c, s]) => [c.suggestions, s] as const),
          fileResolve
        )
      );
    const final = finals[0] ?? '';
    if (finals.every((f) => f === final)) wordStable += 1;
    let part = filePart;
    for (const [c, s] of choice)
      part = engineDecide(
        part,
        c.suggestions.flatMap((id) => triplesOfSuggestion(again, id)),
        action(s),
        file
      );
    const g = grade(final, xml(part, fileResolve));
    if (agrees(g)) wordAgree += 1;
    else
      details.push(
        `reopened choice (${g})\n${firstDifference(semanticBody(final), semanticBody(xml(part, fileResolve)))}`
      );
  }
  if (!agrees(wordView) || wordStable < wordChoices.length || wordAgree < wordChoices.length) {
    codes += 'W';
    if (!agrees(wordView))
      details.push(
        `reopened view (${wordView})\n${firstDifference(semanticBody(reopenedXml()), semanticBody(xml(filePart, fileResolve)))}`
      );
  }
  // Each live choice, made again after the reopen: cards matched by their authors
  const allCards = orderedCards(again);
  const byAuthors = new Map<string, string[]>();
  for (const card of allCards) {
    const authors = [
      ...new Set(card.suggestions.map((id) => again.suggestions.get(id)?.attribution.author ?? '')),
    ]
      .sort()
      .join('+');
    byAuthors.set(authors, [...(byAuthors.get(authors) ?? []), ...card.suggestions]);
  }
  let live: CellResult['live'];
  const mapped =
    groups.length === new Set(groups.map((g) => g.authors)).size &&
    allCards.length > 0 &&
    [...byAuthors.keys()].every((authors) => groups.some((g) => g.authors === authors));
  if (mapped) {
    let same = 0;
    for (const choice of choices) {
      const now = decided(
        doc,
        name,
        choice.map(([g, s]) => [g.suggestions, s] as const),
        resolve
      );
      const later = decided(
        again,
        name,
        choice.map(([g, s]) => [byAuthors.get(g.authors) ?? [], s] as const),
        fileResolve
      );
      const a = semanticBody(asDeleteAndInsert(now));
      const b = semanticBody(asDeleteAndInsert(later));
      if (a === b) same += 1;
      else details.push(`live vs reopened\n${firstDifference(a, b)}`);
    }
    live = { choices: choices.length, same };
    if (same < choices.length) codes += 'L';
  }
  if (!valid.ok) codes += 'I';
  return {
    name: scenario,
    codes,
    view,
    choices: choices.length,
    stable,
    agree,
    word: { view: wordView, choices: wordChoices.length, stable: wordStable, agree: wordAgree },
    ...(live ? { live } : {}),
    details,
  };
}
