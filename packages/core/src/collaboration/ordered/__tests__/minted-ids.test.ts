// The change a minted id was made from: the inverse of the id minter's formats, and nothing for
// a file's id or an id of another prefix.

import { describe, expect, test } from 'bun:test';
import { mintedChange, orderedIdMinter } from '../change/ids.ts';

describe('minted ids', () => {
  test('a minted id gives back its change; any other id gives none', () => {
    const ids = orderedIdMinter('ana-1:7');
    for (const [id, prefix] of [
      [ids.units(), 'i'],
      [ids.node(), 'n'],
      [ids.suggestion(), 'S'],
    ] as const) {
      expect(mintedChange(id, prefix)).toBe('ana-1:7');
    }

    expect(mintedChange('/word/document.xml#0.0.3', 'n')).toBeNull();
    expect(mintedChange(ids.units(), 'n')).toBeNull();
    expect(mintedChange('n:ana-1:7', 'n')).toBeNull();
  });
});
