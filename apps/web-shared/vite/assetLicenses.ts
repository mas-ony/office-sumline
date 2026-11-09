import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import type { Plugin } from 'vite';

/** The package a file belongs to: the nearest package.json above it that has a name. */
function packageOf(file: string): { dir: string; name: string; version: string; license?: string } | null {
  for (let dir = dirname(file); dir !== dirname(dir); dir = dirname(dir)) {
    const manifest = join(dir, 'package.json');
    if (!existsSync(manifest)) continue;
    const { name, version = '0.0.0', license } = JSON.parse(readFileSync(manifest, 'utf8')) as { name?: string; version?: string; license?: string };
    if (name) return { dir, name, version, license };
  }
  return null;
}

/**
 * Vite's licence file (`build.license`) lists the packages of the bundled scripts only. The fonts
 * come in through CSS `@import`, so this adds the packages of the other files the build took from
 * node_modules, in the same format.
 */
export function assetLicenses(fileName = 'THIRD-PARTY-LICENSES.md'): Plugin {
  let root = '';
  return {
    name: 'asset-licenses',
    apply: 'build',
    configResolved(config) {
      root = config.root;
    },
    // After Vite's own licence plugin has written the list.
    generateBundle: {
      order: 'post',
      handler(_, bundle) {
        const list = bundle[fileName];
        if (list?.type !== 'asset') return;
        const text = String(list.source);
        const added = new Map<string, string>();
        for (const file of Object.values(bundle)) {
          if (file.type !== 'asset') continue;
          for (const original of file.originalFileNames) {
            const path = resolve(root, original);
            if (!path.includes('/node_modules/')) continue;
            const pkg = packageOf(path);
            const key = pkg && `${pkg.name}@${pkg.version}`;
            if (!pkg || !key || added.has(key) || text.includes(`\n## ${pkg.name} - ${pkg.version}`)) continue;
            const licence = readdirSync(pkg.dir).find((f) => /^(licen[cs]e|copying)/i.test(f));
            const body = licence ? `\n${readFileSync(join(pkg.dir, licence), 'utf8').trim()}\n` : '';
            added.set(key, `\n## ${pkg.name} - ${pkg.version}${pkg.license ? ` (${pkg.license})` : ''}\n${body}`);
          }
        }
        const sorted = [...added].sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
        list.source = text + sorted.map(([, entry]) => entry).join('');
      },
    },
  };
}
