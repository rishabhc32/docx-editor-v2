// What the section kinds share (task 9.8): a section's holder as a target.

import type { OrderedChangeAddress } from '../../types.ts';
import type { OrderedTargetReader } from '../spec.ts';

/**
 * Name a section's holder as a target: the body's own section item (`z:`) as a block item, else
 * a paragraph mark (`m:`, `k:`).
 */
export function touchSection(reader: OrderedTargetReader, holder: OrderedChangeAddress): void {
  reader.item(holder, holder.space.startsWith('z:') ? 'zero' : 'mark');
}
