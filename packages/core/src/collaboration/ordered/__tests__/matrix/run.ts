// One shard of a matrix set as tests: each scenario must pass every check, or give exactly the
// codes its known difference lists (`known.ts`). With `ORDERED_MATRIX_REPORT` set to a file,
// each result is also appended to it as one JSON line, for `report.ts`.

import { appendFileSync } from 'node:fs';
import { describe, expect, test } from 'bun:test';
import { shardOf, type SetName } from './scenarios.ts';
import { runScenario } from './check.ts';
import { knownFor } from './known.ts';

export function matrixShard(set: SetName, index: number, count: number): void {
  const report = process.env.ORDERED_MATRIX_REPORT;
  describe(`ordered matrix: ${set}, shard ${index + 1} of ${count}`, () => {
    for (const scenario of shardOf(set, index, count)) {
      test(scenario.name, () => {
        const result = runScenario(scenario);
        const known = knownFor(set, scenario.name);
        const allowed = known.length > 0 ? known.map((entry) => entry.codes) : [''];
        if (report !== undefined)
          appendFileSync(
            report,
            `${JSON.stringify({
              set,
              name: scenario.name,
              kinds: scenario.kinds,
              codes: result.codes,
              known: known.map((entry) => entry.cause),
              choices: result.choices,
              stable: result.stable,
              agree: result.agree,
              word: result.word ?? null,
              live: result.live ?? null,
            })}\n`
          );
        if (allowed.includes(result.codes)) return;
        expect({
          name: scenario.name,
          codes: result.codes,
          refusal: result.refusal,
          details: result.details.slice(0, 3),
        }).toEqual({
          name: scenario.name,
          codes: allowed[0] ?? '',
          refusal: undefined,
          details: [],
        });
      });
    }
  });
}
