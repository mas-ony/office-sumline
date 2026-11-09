import { existsSync, readFileSync } from 'node:fs';
import type { Plugin } from 'vite';

const dir = new URL('../../../tests/fixtures/', import.meta.url);
const TYPES: Record<string, string> = {
  csv: 'text/csv',
  xls: 'application/vnd.ms-excel',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

/**
 * Serves the repository's test workbooks at /tests/fixtures/ in development, so the end-to-end
 * tests can open them with `?src=/tests/fixtures/…`.
 */
export function testFixtures(): Plugin {
  return {
    name: 'sumline-test-fixtures',
    apply: 'serve',
    configureServer(server) {
      server.middlewares.use('/tests/fixtures/', (req, res, next) => {
        const name = decodeURIComponent((req.url ?? '').split('?')[0].replace(/^\//, ''));
        const file = new URL(name, dir);
        if (!/^[\w.-]+$/.test(name) || !existsSync(file)) return next();
        res.setHeader('Content-Type', TYPES[name.split('.').pop()!.toLowerCase()] ?? 'application/octet-stream');
        res.end(readFileSync(file));
      });
    },
  };
}
