// Real editors without a browser window (happy-dom), each with an ordered client, around one
// in-process room. Frames go through the protocol codec, delivered in microtasks.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { stubReviewModule } from '../../../../editor/__tests__/review-test-module.ts';
import { paragraphTextOf } from '../../../../store/store/paragraph-model-text.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import {
  composeOrderedPart,
  orderedParagraphCount,
  orderedViewParagraph,
} from '../../compose/compose.ts';
import {
  createOrderedCollaboration,
  type OrderedCollaborationHandle,
  type OrderedCollaborationOptions,
} from '../../client/create.ts';
import { OrderedLocalRoom, type OrderedLocalLogin } from '../../client/local-room.ts';
import type { OrderedClientRefusalReport } from '../../client/refusal.ts';
import type { OrderedSchedule } from '../../client/link.ts';
import type { OrderedMediaUploader, OrderedMediaUploadResult } from '../../client/media-uploads.ts';
import type { OrderedMediaReader } from '../../client/media-reads.ts';
import type { OrderedTransport } from '../../client/transport.ts';
import type { OrderedPerson } from '../../model/person.ts';
import type { OrderedDocument } from '../../model/types.ts';
import type { OrderedRoomPolicy } from '../../room/policy.ts';
import type { EditorModule } from '../../../../contracts/modules.ts';
import type { ImageDecodePort } from '../../../../store/package/image-resources.ts';
import type { ContentLockPolicy } from '../../../../store/package/ooxml-edit.ts';
import { docx, p, t } from '../word-fixtures.ts';
import { testPerson } from '../test-persons.ts';

export const BODY = '/word/document.xml';
export const CODE = 'test-code';
export const ROOM = 'room-1';

/** Three plain paragraphs. */
export const FILE = docx(
  p(t('The fee is 100 USD.')) + p(t('Due in 30 days.')) + p(t('Signed here.'))
);

export type Editor = ReturnType<typeof createDocxEditor>;

export interface Person {
  readonly name: string;
  readonly handle: OrderedCollaborationHandle;
  readonly editor: Editor;
  readonly refusals: OrderedClientRefusalReport[];
  readonly container: HTMLElement;
}

const open: { close(): void }[] = [];

export function closeAll(): void {
  for (const item of open.splice(0).reverse()) item.close();
}

export function localRoom(
  wordFile = FILE,
  options: {
    autoSave?: boolean;
    deliver?: (run: () => void) => void;
    /** The room's clock. Default: one second later on every batch. */
    now?: () => number;
    /** The room's policy. Default: every writer may write. */
    policy?: OrderedRoomPolicy;
    /** The room's relaxation of content-control locks. Default: none, Word's lock. */
    contentLockPolicy?: ContentLockPolicy;
    /** The media the room's storage holds from the start. Default: none. */
    blobs?: ReadonlyMap<string, Uint8Array>;
    /** The room's Word file number. Default: 1. */
    file?: number;
  } = {}
): OrderedLocalRoom {
  let tick = 0;
  return new OrderedLocalRoom({
    room: ROOM,
    wordFile,
    file: options.file ?? 1,
    code: CODE,
    now: options.now ?? (() => Date.UTC(2026, 9, 28, 9, 0, 0) + 1000 * tick++),
    ...(options.autoSave === undefined ? {} : { autoSave: options.autoSave }),
    ...(options.deliver === undefined ? {} : { deliver: options.deliver }),
    ...(options.policy === undefined ? {} : { policy: options.policy }),
    ...(options.contentLockPolicy === undefined
      ? {}
      : { contentLockPolicy: options.contentLockPolicy }),
    ...(options.blobs === undefined ? {} : { blobs: options.blobs }),
  });
}

export const ORGANISER: Omit<OrderedLocalLogin, 'person'> = {
  suggests: false,
  decides: true,
  readOnly: false,
  comments: true,
};
export const PARTY: Omit<OrderedLocalLogin, 'person'> = {
  suggests: true,
  decides: false,
  readOnly: false,
  comments: true,
};

/** A person joins the room with an editor in suggesting mode (or editing). */
export async function join(
  room: OrderedLocalRoom,
  name: string,
  options: {
    readonly mode?: 'suggesting' | 'edit';
    readonly login?: Omit<OrderedLocalLogin, 'person'>;
    /** The login's person, which the session stamps too. Default: `testPerson(name)`. */
    readonly person?: OrderedPerson;
    readonly wordFile?: Uint8Array;
    readonly transport?: OrderedTransport;
    /** How the host stores new pictures. Default: the room's storage, as this login. */
    readonly upload?: OrderedMediaUploader;
    /** How the host reads pictures. Default: the room's storage. */
    readonly media?: OrderedMediaReader;
    /** The person's clock. Default: a fixed time an hour before the room's. */
    readonly now?: () => number;
    /** The editor's review module. Default: a stub that derives no cards. */
    readonly review?: EditorModule;
    /** How the editor decodes pictures, for picture commands. Default: none. */
    readonly imageDecodePort?: ImageDecodePort;
    /** The client's `suggests` option: refuse a party's retyping break before sending. */
    readonly suggests?: boolean;
    /** The client's `offlineEditing` option. Default: on. */
    readonly offlineEditing?: boolean;
    /** The client's timers, for reconnects. Default: none, so the client never connects again. */
    readonly schedule?: OrderedSchedule | null;
    /** The client's `onApplyError` option. */
    readonly onApplyError?: OrderedCollaborationOptions['onApplyError'];
    /** The editor's `contentLockPolicy` option. Default: none, Word's lock. */
    readonly contentLockPolicy?: ContentLockPolicy;
    /** The session's namespace, which names its change ids. Default: a new one. */
    readonly namespace?: string;
  } = {}
): Promise<Person> {
  const refusals: OrderedClientRefusalReport[] = [];
  const person = options.person ?? testPerson(name);
  const login = { person, ...(options.login ?? PARTY) };
  const handle = await createOrderedCollaboration({
    transport: options.transport ?? room.transport(login),
    room: ROOM,
    code: CODE,
    identity: { actorId: person.id, name },
    person,
    wordFile: async () => options.wordFile ?? FILE,
    onRefused: (report) => refusals.push(report),
    // Pictures read back from the room's storage, by hash (design D14)
    media:
      options.media ??
      (async (_location, sha256) => {
        const bytes = room.blob(sha256);
        if (!bytes) throw new Error(`the room holds no blob ${sha256}`);
        return bytes;
      }),
    upload: options.upload ?? room.uploader(login),
    schedule: options.schedule ?? null,
    now: options.now ?? (() => Date.UTC(2026, 9, 28, 8, 0, 0)),
    ...(options.suggests === undefined ? {} : { suggests: options.suggests }),
    ...(options.offlineEditing === undefined ? {} : { offlineEditing: options.offlineEditing }),
    ...(options.onApplyError === undefined ? {} : { onApplyError: options.onApplyError }),
    ...(options.namespace === undefined ? {} : { namespace: options.namespace }),
  });
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: handle.document,
    author: name,
    mode: options.mode ?? 'suggesting',
    ...(options.imageDecodePort === undefined ? {} : { imageDecodePort: options.imageDecodePort }),
    ...(options.contentLockPolicy === undefined
      ? {}
      : { contentLockPolicy: options.contentLockPolicy }),
    modules: [
      options.review ?? stubReviewModule(),
      { id: 'collaboration', collaboration: { session: handle.session } },
    ],
  });
  open.push({
    close() {
      editor.destroy();
      container.remove();
      handle.destroy();
    },
  });
  await settle();
  return { name, handle, editor, refusals, container };
}

/** Let every queued frame and publish run. */
export async function settle(rounds = 8): Promise<void> {
  for (let k = 0; k < rounds; k += 1) await new Promise((resolve) => setTimeout(resolve, 0));
}

/**
 * Type `text` as a person types it: one key per task, through the editor's input, so the keys
 * typed at one caret without a pause are one typing run, one undo step.
 */
export async function typeKeys(person: Person, text: string): Promise<void> {
  for (const data of text) {
    const pages = person.container.querySelector('.docx-pages');
    const init = { bubbles: true, cancelable: true, inputType: 'insertText', data };
    pages?.dispatchEvent(new InputEvent('beforeinput', init));
    await settle();
  }
}

export function surface(person: Person) {
  const found = person.editor.surface;
  if (!found) throw new Error('no surface');
  return found;
}

/** The editor's paragraph texts. */
export function editorTexts(person: Person): string[] {
  const session = surface(person).session;
  const part = session.part();
  return session.paragraphIds().map((id) => paragraphTextOf(part, id) ?? '<none>');
}

/** The list's paragraph texts. */
export function listTexts(person: Person, story = BODY): string[] {
  const doc = person.handle.session.orderedDocument();
  const out: string[] = [];
  for (let k = 0; k < orderedParagraphCount(doc, story); k += 1)
    out.push(
      orderedViewParagraph(doc, story, k)
        ?.chars.map((c) => c.ch)
        .join('') ?? '<none>'
    );
  return out;
}

/** The list's story as XML. */
export function listXml(person: Person, story = BODY): string {
  return serializeOoxmlPart(
    composeOrderedPart(person.handle.session.orderedDocument(), story).part
  );
}

/** Every story of a copy as XML: the list's own compose, with no node ids. */
export function storiesOf(doc: OrderedDocument): string[] {
  return [...doc.stories.keys()].map((name) =>
    serializeOoxmlPart(composeOrderedPart(doc, name).part)
  );
}

/** The person's copy is the room's document: nothing pending, every story the same. */
export function sameAsRoom(person: Person, room: OrderedLocalRoom): boolean {
  return (
    person.handle.session.unsentChanges().length === 0 &&
    person.handle.session.order() === room.room.order &&
    JSON.stringify(storiesOf(person.handle.session.orderedDocument())) ===
      JSON.stringify(storiesOf(room.room.document))
  );
}

/** The editor shows the list exactly. */
export function screenFollowsList(person: Person): boolean {
  return editorXml(person) === listXml(person);
}

/** The editor's main part as XML. */
export function editorXml(person: Person): string {
  return serializeOoxmlPart(surface(person).session.part());
}

/** Put the caret in the index-th paragraph. */
export function caretAt(
  person: Person,
  index: number,
  offset: number,
  end = offset,
  endIndex = index
): void {
  const ids = surface(person).session.paragraphIds();
  const paragraphId = ids[index];
  const headId = ids[endIndex];
  if (paragraphId === undefined || headId === undefined) throw new Error(`no paragraph ${index}`);
  surface(person).setSelection({
    anchor: { paragraphId, offset },
    head: { paragraphId: headId, offset: end },
  });
}

/** The editor's caret: the paragraph's index and the offset. */
export function caretOf(person: Person): { index: number; offset: number; text: string } {
  const s = surface(person);
  const head = s.state().selection.head;
  const ids = s.session.paragraphIds();
  const index = ids.indexOf(head.paragraphId);
  return {
    index,
    offset: head.offset,
    text: paragraphTextOf(s.session.part(), head.paragraphId) ?? '',
  };
}

/**
 * A transport whose frames from the room wait while `held`, until `release`: one person's
 * copy then hears of a change only when the test says so.
 */
export class HeldTransport implements OrderedTransport {
  held = false;
  readonly #inner: OrderedTransport;
  readonly #queue: (() => void)[] = [];

  constructor(inner: OrderedTransport) {
    this.#inner = inner;
  }

  connect(
    events: Parameters<OrderedTransport['connect']>[0]
  ): ReturnType<OrderedTransport['connect']> {
    return this.#inner.connect({
      open: () => events.open(),
      close: (reason) => events.close(reason),
      message: (frame) => {
        if (this.held) this.#queue.push(() => events.message(frame));
        else events.message(frame);
      },
    });
  }

  /** Deliver the frames that waited, in order. */
  release(): void {
    for (const deliver of this.#queue.splice(0)) deliver();
  }
}

/** A host's upload that waits until the test lets it go to the storage, answers it, or fails it. */
export class HeldUpload {
  readonly waiting: {
    readonly sha256: string;
    go(): void;
    answer(result: OrderedMediaUploadResult): void;
    fail(): void;
  }[] = [];

  readonly #inner: OrderedMediaUploader;

  constructor(inner: OrderedMediaUploader) {
    this.#inner = inner;
  }

  readonly upload: OrderedMediaUploader = (request) =>
    new Promise((resolve, reject) => {
      this.waiting.push({
        sha256: request.sha256,
        go: () => {
          this.#inner(request).then(resolve, reject);
        },
        answer: resolve,
        fail: () => reject(new Error('the host is offline')),
      });
    });

  /** Let every waiting upload go to the storage. */
  goAll(): void {
    for (const upload of this.waiting.splice(0)) {
      upload.go();
    }
  }
}

/** Run every queued delivery of a room that holds its frames, and what they queue, until none is left. */
export async function drain(queue: (() => void)[]): Promise<void> {
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
