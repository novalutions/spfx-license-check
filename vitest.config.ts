import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vitest/config';

const local = (path: string): string => fileURLToPath(new URL(path, import.meta.url));

// SPFx provides @microsoft/sp-http at runtime; tests only need AadHttpClient.configurations.v1
const spHttp = { find: '@microsoft/sp-http', replacement: local('./test/fakes/sp-http.ts') };

export default defineConfig({
  test: {
    projects: [
      {
        resolve: { alias: [spHttp] },
        test: { name: 'source', environment: 'node', include: ['test/**/*.test.ts'] }
      },
      {
        // the same tests against the ES5 output (npm run build:es5-test), so ES5-only behaviour is covered too
        resolve: { alias: [spHttp, { find: /^\.\.\/src$/, replacement: local('./temp/es5/index.js') }] },
        test: { name: 'es5', environment: 'node', include: ['test/**/*.test.ts'] }
      }
    ]
  }
});
