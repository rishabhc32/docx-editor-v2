// Development warning for a `headerFooterEditing` prop that changes after the editor exists.

import { isDevelopment } from '../lib/is-development';

let printed = false;

/** `undefined` means the default, `true`, so the two compare equal. */
function normalize(value: unknown): unknown {
  return value === undefined ? true : value;
}

/**
 * Warn once when `headerFooterEditing` differs from the value the editor was created with. The
 * editor reads it once, when it is created, so a changed prop is ignored. Never warns in a
 * production build.
 */
export function warnHeaderFooterEditingChanged(created: unknown, current: unknown): void {
  if (normalize(created) === normalize(current) || printed || !isDevelopment()) return;
  printed = true;
  console.warn(
    '[docx-editor] headerFooterEditing changed after the editor was created, so the ' +
      'change is ignored. The editor reads headerFooterEditing once; create a new editor ' +
      'to change it.'
  );
}
