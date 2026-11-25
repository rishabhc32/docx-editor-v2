// A link on a prompt's words, once a party's typing fills the control, with two real editors and
// one ordered room. The glossary keeps the prompt with its link (`content-control-prompt-
// glossary.ts`), and relationship ids are per part (OPC): the glossary's link names a
// relationship of the glossary's own, so every saved file resolves each id where it is named, as
// the editor alone saves it. A reject of the typing gives the prompt back with its link.

import { afterEach, describe, expect, test } from 'bun:test';
import {
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { unresolvedRelationships } from '../oracle.ts';
import { docx, p, R, t, W, WML_CT, type PackageExtras } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { alone, converged, roomFile, select } from './links-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const GLOSSARY = '/word/glossary/document.xml';
const FACE = '<w:rStyle w:val="PlaceholderText"/>';
const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

/** "Name: Click here to enter. end." with an inline control showing its prompt. */
const prompt = (here: string, extras?: PackageExtras) =>
  docx(
    p(
      t('Name: ') +
        sdt('<w:showingPlcHdr/><w:text/>', t('Click ', FACE) + here + t(' to enter.', FACE)) +
        t(' end.')
    ) + p(t('Signed.')),
    extras
  );

/** The file's own link on "here", through the main part's `rIdLink`. */
const LINKED = prompt(`<w:hyperlink r:id="rIdLink">${t('here', FACE)}</w:hyperlink>`);
const PLAIN = prompt(t('here', FACE));
const FILLED = ['Name: Q end.', 'Signed.'];
const PROMPT = ['Name: Click here to enter. end.', 'Signed.'];

/** A saved file, read again: a refused read fails the test. */
function reread(bytes: Uint8Array): OoxmlPackage {
  const read = readOoxmlPackage(bytes, {});
  if (!read.ok) {
    throw new Error(`the saved file does not open: ${read.reason}`);
  }

  return read.package;
}

/** Each saved copy names only relationships its parts hold, and keeps the link in the glossary. */
async function savesResolve(
  room: ReturnType<typeof localRoom>,
  file: Uint8Array,
  people: Person[]
) {
  const saves = [roomFile(room.room.document, file)];
  for (const person of people) {
    saves.push(reread(new Uint8Array(await person.editor.save())));
  }

  for (const pkg of saves) {
    expect(unresolvedRelationships(pkg)).toEqual([]);
    expect(
      pkg.externalTargets.map((target) => [target.ownerPart, target.rawTarget])
    ).toContainEqual([GLOSSARY, expect.stringContaining('example.com')]);
  }
}

function rejectAll(person: Person): void {
  for (const card of person.editor.getReviewItems()) {
    expect(person.editor.rejectReviewItem(card.key).ok).toBe(true);
  }
}

describe("a link on a prompt's words, once a party's typing fills the control", () => {
  test("the editor alone keeps the file's link in the glossary with a relationship of its own", async () => {
    const pkg = await alone(LINKED, 'Cy', 'suggesting', (editor) => {
      select(editor, 0, 8);
      editor.surface?.insertPlainText('Q');
    });
    const saved = reread(writeOoxmlPackage(pkg));
    expect(unresolvedRelationships(saved)).toEqual([]);
    expect(saved.externalTargets.filter((target) => target.ownerPart === GLOSSARY)).toHaveLength(1);

    // The glossary's root binds the link's prefix, as the room's editors check it
    expect(serializeOoxmlPart(saved.parts.get(GLOSSARY)!)).toContain(`xmlns:r="${R}"`);
  });

  for (const [label, file] of [
    ["the file's link", LINKED],
    ["the organiser's link", PLAIN],
  ] as const) {
    test(`${label}: every save resolves its ids, and a reject gives the link back`, async () => {
      const review = reviewTestModule();
      const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
      const olga = await join(room, 'Olga', {
        mode: 'edit',
        login: ORGANISER,
        wordFile: file,
        review,
      });
      const cy = await join(room, 'Cy', { wordFile: file, review });
      if (file === PLAIN) {
        select(olga.editor, 0, 12, 16);
        expect(
          olga.editor.exec({ type: 'insertHyperlink', href: 'https://example.com/fees' })
        ).toMatchObject({ ok: true });
        await settle();
      }

      expect(editorXml(cy)).toContain('<w:hyperlink');
      caretAt(cy, 0, 8);
      surface(cy).insertPlainText('Q');
      await settle();
      converged(room, [olga, cy]);
      expect(editorTexts(cy)).toEqual(FILLED);
      await savesResolve(room, file, [olga, cy]);

      rejectAll(olga);
      await settle();
      converged(room, [olga, cy]);
      expect(editorTexts(cy)).toEqual(PROMPT);
      expect(editorXml(cy)).toContain('<w:hyperlink');
      for (const pkg of [
        roomFile(room.room.document, file),
        reread(new Uint8Array(await cy.editor.save())),
      ]) {
        expect(unresolvedRelationships(pkg)).toEqual([]);
      }
    });
  }
});

/**
 * A glossary with relationships of its own, as Word writes one: `rIdLink` names `held` there, the
 * id the main part gives the prompt's link.
 */
const glossaryHolding = (held: string): PackageExtras => ({
  parts: {
    'word/glossary/document.xml': `<w:glossaryDocument xmlns:w="${W}"><w:docParts/></w:glossaryDocument>`,
    'word/glossary/styles.xml': `<w:styles xmlns:w="${W}"/>`,
    'word/glossary/_rels/document.xml.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdLink" ${held}/></Relationships>`,
  },
  contentTypes:
    `<Override PartName="/word/glossary/document.xml" ContentType="${WML_CT}.document.glossary+xml"/>` +
    `<Override PartName="/word/glossary/styles.xml" ContentType="${WML_CT}.styles+xml"/>`,
  relationships: `<Relationship Id="rIdGlossary" Type="${R}/glossaryDocument" Target="glossary/document.xml"/>`,
});
const OTHER = glossaryHolding(`Type="${R}/styles" Target="styles.xml"`);
const SAME = glossaryHolding(
  `Type="${R}/hyperlink" Target="https://example.com/old" TargetMode="External"`
);
const linkedWith = (extras: PackageExtras) =>
  prompt(`<w:hyperlink r:id="rIdLink">${t('here', FACE)}</w:hyperlink>`, extras);

/**
 * The glossary keeps the prompt with its link only where `rIdLink` names the same link there: a
 * restore puts the block back where its ids resolve. Where the glossary's `rIdLink` names its
 * styles, the block is not kept, and its `rIdLink` still names them.
 */
function keepsLinkOnlyWhereSame(pkg: OoxmlPackage, same: boolean): void {
  expect(unresolvedRelationships(pkg)).toEqual([]);
  const held = (pkg.relationships.get(GLOSSARY) ?? []).filter((r) => r.id === 'rIdLink');
  expect(held.map((r) => r.type.split('/').at(-1))).toEqual([same ? 'hyperlink' : 'styles']);
  expect(serializeOoxmlPart(pkg.parts.get(GLOSSARY)!).includes('r:id="rIdLink"')).toBe(same);
}

describe('a link whose id the glossary already holds', () => {
  for (const [label, extras, same] of [
    ['for its styles', OTHER, false],
    ['for the same link', SAME, true],
  ] as const) {
    test(`${label}: the editor alone and the room keep the prompt only with the same link`, async () => {
      const file = linkedWith(extras);
      const pkg = await alone(file, 'Cy', 'suggesting', (editor) => {
        select(editor, 0, 8);
        editor.surface?.insertPlainText('Q');
      });
      keepsLinkOnlyWhereSame(reread(writeOoxmlPackage(pkg)), same);

      const review = reviewTestModule();
      const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
      const olga = await join(room, 'Olga', {
        mode: 'edit',
        login: ORGANISER,
        wordFile: file,
        review,
      });
      const cy = await join(room, 'Cy', { wordFile: file, review });
      caretAt(cy, 0, 8);
      surface(cy).insertPlainText('Q');
      await settle();
      converged(room, [olga, cy]);
      expect(editorTexts(cy)).toEqual(FILLED);
      keepsLinkOnlyWhereSame(roomFile(room.room.document, file), same);
      for (const person of [olga, cy]) {
        keepsLinkOnlyWhereSame(reread(new Uint8Array(await person.editor.save())), same);
      }
    });
  }
});
