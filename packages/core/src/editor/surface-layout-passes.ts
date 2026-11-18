// Passes of a surface's live layout: a commit's, the scheduler's, an opening's. A pass in a
// fresh layout session (the footnote reserves found in the background, a save's page numbers)
// is not counted: it does not replace what the user sees.

let livePasses = 0;

/** A pass of the live layout starts. */
export function noteLiveLayoutPass(): void {
  livePasses += 1;
}

/**
 * @internal Deterministic work recorder: live layout passes started, counted across every
 * surface since the last reset.
 */
export function liveLayoutPassTestRecorder(): { readonly passes: number; reset(): void } {
  return {
    get passes() {
      return livePasses;
    },
    reset() {
      livePasses = 0;
    },
  };
}
