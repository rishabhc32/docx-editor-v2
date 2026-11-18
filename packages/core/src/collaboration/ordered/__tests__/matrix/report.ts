// The full matrix report: runs every shard file, one process at a time (each stays under 1 GB),
// then prints the results by set, the same counts on spike 5's own kinds next to the spike's,
// and the coverage of the editor's feature list.
//
//   bun packages/core/src/collaboration/ordered/__tests__/matrix/report.ts
//   bun packages/core/src/collaboration/ordered/__tests__/matrix/report.ts --from results.jsonl

import { spawnSync } from 'node:child_process';
import { mkdtempSync, readdirSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { SPIKE_KINDS } from './catalogue.ts';
import { COVERAGE } from './coverage.ts';
import { MIX_KEYS } from './scenarios.ts';

interface Row {
  readonly set: string;
  readonly name: string;
  readonly kinds: readonly string[];
  readonly codes: string;
  readonly known: readonly string[];
  readonly choices: number;
  readonly stable: number;
  readonly agree: number;
  readonly word: { view: string; choices: number; stable: number; agree: number } | null;
  readonly live: { choices: number; same: number } | null;
}

const here = fileURLToPath(new URL('.', import.meta.url));
const fromAt = process.argv.indexOf('--from');
let file = fromAt >= 0 ? (process.argv[fromAt + 1] ?? '') : '';
if (!file) {
  file = join(mkdtempSync(join(tmpdir(), 'ordered-matrix-')), 'results.jsonl');
  for (const shard of readdirSync(here)
    .filter((f) => /^matrix-.*\.test\.ts$/.test(f))
    .sort()) {
    const run = spawnSync('bun', ['test', join(here, shard)], {
      env: { ...process.env, ORDERED_MATRIX_REPORT: file },
      encoding: 'utf8',
    });
    process.stderr.write(`${shard}: ${run.status === 0 ? 'pass' : 'FAIL'}\n`);
  }
}
const rows: Row[] = readFileSync(file, 'utf8')
  .split('\n')
  .filter(Boolean)
  .map((line) => JSON.parse(line) as Row);

const refused = (r: Row) => /^[SER]$/.test(r.codes);
const sum = (list: readonly Row[], f: (r: Row) => number) => list.reduce((n, r) => n + f(r), 0);

function line(label: string, list: readonly Row[]): string {
  const run = list.filter((r) => !refused(r) && r.codes !== 'N');
  const view = run.filter((r) => !r.codes.includes('V')).length;
  const choices = sum(run, (r) => r.choices);
  const word = run.filter((r) => r.word && !r.codes.includes('W')).length;
  const live = run.filter((r) => r.live);
  return [
    label,
    `${list.length}`,
    `${list.filter(refused).length}`,
    `${view} of ${run.length}`,
    `${sum(run, (r) => r.stable)} of ${choices}`,
    `${sum(run, (r) => r.agree)} of ${choices}`,
    `${word} of ${run.length}`,
    `${sum(live, (r) => r.live?.same ?? 0)} of ${sum(live, (r) => r.live?.choices ?? 0)}`,
    `${list.filter((r) => r.codes === '').length}`,
    `${list.filter((r) => r.codes !== '').length}`,
  ].join(' | ');
}

const head =
  '| Set | Scenarios | Refused (S, E, R) | View = engine | Choices: one result in every order | Choices = engine | Word file reopens as the file | Live = reopened | Every check passes | On the known list |\n|---|---|---|---|---|---|---|---|---|---|';
const sets = [...new Set(rows.map((r) => r.set))];
console.log('## Every kind\n');
console.log(head);
for (const set of sets)
  console.log(
    `| ${line(
      set,
      rows.filter((r) => r.set === set)
    )} |`
  );
console.log(`| ${line('all', rows)} |`);

// The spike's own scenarios: its 65 kinds, its 22 mix kinds, its 7 containers
const spike = new Set(SPIKE_KINDS.map((k) => k.key));
const spikeMix = new Set(MIX_KEYS.slice(0, 22));
const spikeContainers = new Set(['link', 'bookmark', 'comment', 'smart', 'xml', 'plain', 'rich']);
const spikeRows = rows.filter((r) => {
  if (r.set === 'single' || r.set === 'pairs') return spike.has(r.kinds[0] ?? '');
  if (r.set === 'mix') return r.kinds.every((k) => spikeMix.has(k));
  if (r.set === 'nested')
    return spikeMix.has(r.kinds[0] ?? '') && spikeContainers.has(r.kinds[1] ?? '');
  return false;
});
console.log("\n## Spike 5's own scenarios\n");
console.log(head);
for (const set of ['single', 'pairs', 'mix', 'nested'])
  console.log(
    `| ${line(
      set,
      spikeRows.filter((r) => r.set === set)
    )} |`
  );
console.log(`| ${line('all', spikeRows)} |`);

console.log('\n## Known differences by cause\n\n| Cause | Cells |\n|---|---|');
const causes = new Map<string, number>();
for (const r of rows)
  if (r.codes !== '') for (const c of r.known) causes.set(c, (causes.get(c) ?? 0) + 1);
for (const [cause, n] of [...causes].sort((a, b) => b[1] - a[1]))
  console.log(`| ${cause} | ${n} |`);

console.log(
  '\n## The feature list\n\n| Entry | Kinds | Pass | Known difference | Refused by name |\n|---|---|---|---|---|'
);
for (const [id, entry] of Object.entries(COVERAGE)) {
  if (entry.outside) {
    console.log(`| ${id} | - | - | - | - (${entry.outside}) |`);
    continue;
  }
  const kinds = new Set(entry.kinds ?? []);
  const mine = rows.filter((r) => r.kinds.some((k) => kinds.has(k)));
  const pass = mine.filter((r) => r.codes === '').length;
  const byModel = (r: Row) => r.codes === 'S' || r.codes === 'R';
  const known = mine.filter((r) => r.codes !== '' && !byModel(r)).length;
  const byName = mine.filter(byModel).length + (entry.refusals?.length ?? 0);
  console.log(
    `| ${id} | ${[...kinds].join(', ')} | ${pass} | ${known} | ${byName}${entry.refusals ? ` (${entry.refusals.join(', ')})` : ''} |`
  );
}
