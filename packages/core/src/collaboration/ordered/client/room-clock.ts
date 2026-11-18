// The room's clock as this copy reads it from the room's stamps (task 6.8).

/** How many of the latest windows the gap reads. */
const WINDOWS = 8;

export class OrderedRoomClock {
  readonly #now: () => number;
  readonly #windows: { readonly low: number; readonly high: number }[] = [];
  #gap = 0;

  /** `now`: the local clock, in milliseconds since 1970. */
  constructor(now: () => number) {
    this.#now = now;
  }

  /** The local time now, in milliseconds: when a change is made. */
  local(): number {
    return this.#now();
  }

  /**
   * When the room will most likely stamp a change made at local time `local`. A change's date is
   * the room's clock when the room orders it (`at`, written as `w:date`). The copy applies its own
   * change at once with this date; when it is not the room's, the paragraph is drawn again when
   * the room's change comes back, a second layout for one key press. With the room's clock the
   * two agree, except when the room's second turns between the key press and the stamp.
   */
  stampFor(local: number): number {
    return local + this.#gap;
  }

  /**
   * The room stamped `at` a change of this person made at local time `local`. Each such change
   * gives a window: the room stamped it between `at` and `at` + 1 s (stamps are whole seconds),
   * which is the local time it was made plus the gap between the two clocks plus the trip to the
   * room. The gap is the middle of the windows the latest changes share. Windows that share
   * nothing mean a clock moved: the latest ones that do share count, the older ones go.
   */
  observe(local: number, at: string): void {
    const stamped = Date.parse(at);
    if (!Number.isFinite(stamped) || !Number.isFinite(local)) return;
    this.#windows.push({ low: stamped - local, high: stamped - local + 1000 });
    if (this.#windows.length > WINDOWS) this.#windows.shift();
    let low = -Infinity;
    let high = Infinity;
    for (let k = this.#windows.length - 1; k >= 0; k -= 1) {
      const window = this.#windows[k];
      if (!window) continue;
      const nextLow = Math.max(low, window.low);
      const nextHigh = Math.min(high, window.high);
      if (nextLow >= nextHigh) break;
      low = nextLow;
      high = nextHigh;
    }

    this.#gap = (low + high) / 2;
  }
}
