import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';
import { assetLicenses } from '../web-shared/vite/assetLicenses';
import { testFixtures } from '../web-shared/vite/testFixtures';

export default defineConfig({
  plugins: [react(), testFixtures(), assetLicenses()],
  base: './',
  // The icons, manifest and service worker, shared with the Vue app.
  publicDir: '../web-shared/public',
  build: {
    target: 'es2020',
    sourcemap: true,
    chunkSizeWarningLimit: 2000,
    // The licences of the bundled packages, icons and fonts, shipped with the built app.
    license: { fileName: 'THIRD-PARTY-LICENSES.md' },
  },
  // Listen on every interface so phones on the LAN can open the dev server.
  server: { host: true, port: 5175 },
  preview: { host: true, port: 4175 },
});
