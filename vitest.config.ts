import { defineConfig } from 'vitest/config';

// Each package has its own test setup.
export default defineConfig({
  test: {
    projects: ['packages/*'],
  },
});
