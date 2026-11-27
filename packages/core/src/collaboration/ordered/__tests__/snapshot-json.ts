// The body of a checkpoint of version 6 by field name, for tests that change one value in it. A
// record is `[mask, ...values]`, its values in its field table's order (`snapshot-codec.ts`).

import { DOCUMENT, ITEM, STORY } from '../model/snapshot-fields.ts';

/** A record's values by field name. */
export interface RecordView {
  get(key: string): unknown;
  /** Gives the field `value`, written as it is. */
  set(key: string, value: unknown): void;
  delete(key: string): void;
}

/** The record `json` of the field table `fields`. */
export function recordView(json: unknown[], fields: object): RecordView {
  const keys = Object.entries(fields).flatMap(([key, field]) => ('stored' in field ? [key] : []));
  const bit = (key: string): number => {
    const k = keys.indexOf(key);
    if (k === -1) {
      throw new Error(`No stored field ${key}.`);
    }

    return k;
  };
  const held = (k: number) => Math.floor((json[0] as number) / 2 ** k) % 2 === 1;
  const place = (k: number) => 1 + keys.slice(0, k).filter((_, j) => held(j)).length;
  return {
    get: (key) => (held(bit(key)) ? json[place(bit(key))] : undefined),
    set: (key, value) => {
      const k = bit(key);
      if (held(k)) {
        json[place(k)] = value;
      } else {
        json.splice(place(k), 0, value);
        json[0] = (json[0] as number) + 2 ** k;
      }
    },
    delete: (key) => {
      const k = bit(key);
      if (held(k)) {
        json.splice(place(k), 1);
        json[0] = (json[0] as number) - 2 ** k;
      }
    },
  };
}

type Body = Record<string, unknown>;

/** The document of a checkpoint's body. */
export const documentView = (body: Body): RecordView =>
  recordView(body.document as unknown[], DOCUMENT);

/** Every item of a checkpoint's body as written, story by story. */
export function itemJson(body: Body): unknown[][] {
  const stories = documentView(body).get('stories') as [unknown, unknown[]][];
  return stories.flatMap(([, story]) => recordView(story, STORY).get('tree') as unknown[][]);
}

/** Every item of a checkpoint's body, story by story. */
export const itemViews = (body: Body): RecordView[] =>
  itemJson(body).map((item) => recordView(item, ITEM));
