// What a change adds outside the story text, across a fold (design D14): a list level, a link's
// address, an embedded object, a chart's part and a list's picture bullet. A fold keeps each one
// with its suggestion: a decision after a fold writes it or leaves it out as the room does live.

import { describe, test } from 'bun:test';
import { openOrderedWordFile } from '../client/open.ts';
import { RESOURCE_ROOMS, resourceBlobs } from './checkpoint-resources-fixture.ts';
import { checkFold } from './fold-check.ts';

describe('a fold keeps what a change added outside the story text with its suggestion', () => {
  for (const room of RESOURCE_ROOMS) {
    test(room.name, async () => {
      const { document: doc, base } = openOrderedWordFile(room.source, { file: 1 });
      room.edit(doc);
      await checkFold(doc, base, resourceBlobs(base), room.name);
    });
  }
});
