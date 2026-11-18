import type { OrderedPartyRule } from './types.ts';

/** The engine writes these with tracking off in suggesting mode: a party's is not a suggestion. */
export function untrackedRule(reason: 'text' | 'package', what: string): OrderedPartyRule {
  return { code: 'policy-direct', reason, why: `the engine has no tracked form: ${what}` };
}
