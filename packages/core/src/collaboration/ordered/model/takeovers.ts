// The suggestions a record was taken from (design D5, policy 2). A takeover of a takeover keeps
// each of them, so a take-back gives the record to the latest one that is still pending.

/** A record's links to the suggestions it was taken from. */
interface TakenLinks {
  readonly replaced?: string;
  readonly replacedEarlier?: readonly string[];
}

/** The suggestions a record was taken from, oldest first. */
export function takenFrom(record: TakenLinks): string[] {
  return record.replaced === undefined ? [] : [...(record.replacedEarlier ?? []), record.replaced];
}

/** Whether two records were taken from the same suggestions, in the same order. */
export function sameTakenFrom(a: TakenLinks, b: TakenLinks): boolean {
  const chain = takenFrom(a);
  const other = takenFrom(b);
  return chain.length === other.length && chain.every((id, k) => id === other[k]);
}

/** The links of a record taken from the suggestions `chain` names, oldest first. */
export function takenLinks(chain: readonly string[]): TakenLinks {
  const replaced = chain.at(-1);
  if (replaced === undefined) {
    return {};
  }

  return chain.length === 1 ? { replaced } : { replaced, replacedEarlier: chain.slice(0, -1) };
}

/** `record` held by `suggestion`, taken from the suggestions `chain` names, oldest first. */
export function withTakenFrom<T extends TakenLinks & { readonly suggestion: string }>(
  record: T,
  suggestion: string,
  chain: readonly string[]
): T {
  const { replaced: _replaced, replacedEarlier: _earlier, ...rest } = record;

  // Still a T: both links are optional on every record type
  return { ...rest, suggestion, ...takenLinks(chain) } as T;
}
