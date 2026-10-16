#!/usr/bin/env node
// Stamp a fork release version, such as `2.31.0-sp.1`, on the packages that the fork
// release ships, and pin the links between them to that version. Run it on a clean
// release checkout before `bun run build:packages`, because `react` embeds its version
// at build time. FORK-UPGRADE.md describes the full release.
import { readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';

const version = process.argv[2];
if (!/^\d+\.\d+\.\d+-sp\.\d+$/.test(version ?? '')) {
  console.error('Usage: node scripts/fork/stamp-release.mjs <release>-sp.<n>');
  process.exit(1);
}

const root = path.resolve(import.meta.dirname, '../..');
const shipped = ['core', 'react', 'pro', 'editor-api', 'i18n'];
const names = new Set(shipped.map((directory) => `@docx-editor.dev/${directory}`));

for (const directory of shipped) {
  const file = path.join(root, 'packages', directory, 'package.json');
  const manifest = JSON.parse(readFileSync(file, 'utf8'));
  manifest.version = version;
  // Other @docx-editor.dev peers, such as the optional `vue`, keep their upstream range.
  for (const field of ['dependencies', 'peerDependencies', 'optionalDependencies']) {
    for (const name of Object.keys(manifest[field] ?? {})) {
      if (names.has(name)) manifest[field][name] = version;
    }
  }
  writeFileSync(file, `${JSON.stringify(manifest, null, 2)}\n`);
  const links = ['dependencies', 'peerDependencies'].flatMap((field) =>
    Object.entries(manifest[field] ?? {})
      .filter(([name]) => name.startsWith('@docx-editor.dev/'))
      .map(([name, range]) => `${name}@${range}`)
  );
  console.log(`${manifest.name}@${manifest.version}: ${links.join(', ') || 'no internal links'}`);
}
