import type { CollaborationPackageWrite } from '../../../index.ts';
import { translateBookmarkRemoval } from './fields.ts';
import { mainStoryOf } from './paste-blocks.ts';
import type { OrderedTranslateBatch, OrderedTranslateStep } from './types.ts';

/**
 * One package edit a write describes (`CollaborationPackageWrite`: a script's document properties,
 * a bookmark deleted from a story the script does not write in) as a step of the batch. The gate
 * refuses by name a package edit that does not say what it does.
 */
export function packageWriteStep(
  write: CollaborationPackageWrite,
  batch: OrderedTranslateBatch
): OrderedTranslateStep {
  switch (write.kind) {
    case 'document-properties':
      // Word records no change of a document property: a setting, direct only (D14)
      return () => [
        {
          kind: 'setting',
          part: mainStoryOf(batch.doc).name,
          setting: { op: 'setDocumentProperties', values: write.values },
        },
      ];
    case 'bookmark-removal':
      return () => translateBookmarkRemoval(batch, write.partName, write.name);
  }
}
