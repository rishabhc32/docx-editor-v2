// Writes a room's saved item list (checkpoint version 6, design D2): `ordered/checkpoint.json`,
// every stored field of the document through the field table, with the Word XML its items hold
// as references into `ordered/base.docx` or once by value.

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { XML_HARD_MAX_BYTES } from '../../../store/package/xml-reader.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { OrderedRefusal } from './refusal.ts';
import type { OrderedFileMedia } from './resources.ts';
import { list, record, type SnapshotWriteContext } from './snapshot-codec.ts';
import { DOCUMENT, FILE_MEDIA } from './snapshot-fields.ts';
import { SnapshotNodeWriter } from './snapshot-nodes.ts';
import type { OrderedDocument, OrderedItem } from './types.ts';

/** The entries of a room checkpoint of version 6, in its zip: the item list and the base. */
export const SNAPSHOT_ENTRY = '/ordered/checkpoint.json';
export const SNAPSHOT_BASE_ENTRY = '/ordered/base.docx';

/** The format a checkpoint names: the same for every version. */
export const SNAPSHOT_FORMAT = 'docx-editor.dev/ordered-checkpoint';

/** The version this engine writes. */
export const SNAPSHOT_VERSION = 6;

/** The lowest reader version that reads what this engine writes. */
export const SNAPSHOT_READS = 6;

export const DOCUMENT_CODEC = record(DOCUMENT);

/** The file media `ordered/base.docx` names by reference. */
export const MEDIA_CODEC = list(record(FILE_MEDIA));

/** The fields before the body, in this order; the checksum covers the body's text. */
export interface SnapshotHeader {
  readonly format: string;
  readonly version: number;
  readonly reads: number;
  /** The Word file number the saved loaded spaces hold. */
  readonly file: number;
  /** SHA-256 of the body's JSON text, lower-case hex. */
  readonly sha256: string;
}

/**
 * The text of `checkpoint.json` before its body: the header's fields in a fixed order, so a
 * reader finds the body's exact text after it.
 */
export function headerText(header: SnapshotHeader): string {
  const { format, version, reads, file, sha256 } = header;
  return (
    `{"format":${JSON.stringify(format)},"version":${version},"reads":${reads},` +
    `"file":${file},"sha256":${JSON.stringify(sha256)},"body":`
  );
}

/**
 * `checkpoint.json` of `doc`: its stored fields, `media` (the file media of `ordered/base.docx`),
 * and the nodes and namespaces they name. `base` is the package `ordered/base.docx` holds.
 *
 * @throws OrderedRefusal `too-many-items` when it is past the engine's limit of one part (64 MiB).
 */
export function snapshotJsonBytes(
  doc: OrderedDocument,
  base: OoxmlPackage,
  media: readonly OrderedFileMedia[]
): Uint8Array {
  const places = new Map<OrderedItem, readonly [number, number]>();
  [...doc.stories.values()].forEach((story, s) => {
    story.tree.toArray().forEach((item, k) => places.set(item, [s, k]));
  });
  const nodes = new SnapshotNodeWriter(base);
  const w: SnapshotWriteContext = {
    nodes,
    item: (item) => {
      const place = places.get(item);
      if (place === undefined) {
        throw new RangeError('A resource names an item no story holds.');
      }

      return place;
    },
  };
  const document = DOCUMENT_CODEC.write(doc, w);
  const body = JSON.stringify({
    media: MEDIA_CODEC.write(media, w),
    namespaces: nodes.namespaces,
    nodes: nodes.list,
    document,
  });
  const encoder = new TextEncoder();
  const bodyBytes = encoder.encode(body);
  const header = encoder.encode(
    headerText({
      format: SNAPSHOT_FORMAT,
      version: SNAPSHOT_VERSION,
      reads: SNAPSHOT_READS,
      file: doc.file,
      sha256: orderedBlobHash(bodyBytes),
    })
  );
  const bytes = new Uint8Array(header.byteLength + bodyBytes.byteLength + 1);
  bytes.set(header);
  bytes.set(bodyBytes, header.byteLength);
  bytes[bytes.byteLength - 1] = 0x7d; // }
  if (bytes.byteLength > XML_HARD_MAX_BYTES) {
    throw new OrderedRefusal('too-many-items', `checkpoint over ${XML_HARD_MAX_BYTES} bytes`);
  }

  return bytes;
}
