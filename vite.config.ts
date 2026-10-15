import { defineConfig } from 'vite';

// Relative base so the build works under https://<user>.github.io/<repo>/
export default defineConfig({
  base: './',
  build: { outDir: 'dist', target: 'es2020' },
  test: { environment: 'node', setupFiles: ['tests/setup.ts'] },
} as never);
