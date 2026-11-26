import { defineConfig, devices } from '@playwright/test';

// The tests drive the production build in real browsers: editing relies on contenteditable, which
// differs between engines, so a DOM emulation would not tell us whether a saved file is faithful.
// Run one browser with `npm test -- --project=chromium` (or firefox, or webkit).
// The specs Safari's engine runs (see the webkit project below).
const WEBKIT_SPECS = ['roundtrip', 'fixtures', 'open', 'editing', 'lists', 'tables', 'track', 'paste', 'starters', 'start-page-size']
  .map((name) => `**/${name}.spec.js`);

export default defineConfig({
  testDir: 'tests',
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  reporter: process.env.CI ? [['list'], ['github']] : 'list',
  use: {
    baseURL: 'http://localhost:4173/',
    acceptDownloads: true,
    trace: 'off',
    // The tests click the toolbar's tools wherever they are: all of them in one toolbar, as "All tools"
    // shows them. The ribbon (the tabs, the default) is tested on its own, in tests/ribbon.spec.js.
    storageState: { cookies: [], origins: [{ origin: 'http://localhost:4173', localStorage: [{ name: 'intact.ribbon', value: '0' }] }] },
  },
  projects: [
    { name: 'chromium', use: { ...devices['Desktop Chrome'] } },
    { name: 'firefox', use: { ...devices['Desktop Firefox'] } },
    // Safari's engine (WebKit), for the core of the promises: documents come back as they went in
    // (round-trip and the real-world fixtures), open, take typing, lists, tables, tracked changes and
    // pasting, and starters are written as Word writes them. The rest of the suite follows spec by
    // spec, as each passes here: add it to this list, never skip a case to get it in.
    {
      name: 'webkit',
      use: { ...devices['Desktop Safari'] },
      testMatch: WEBKIT_SPECS,
    },
  ],
  webServer: {
    command: 'npm run build && npx vite preview --port 4173 --strictPort',
    url: 'http://localhost:4173/',
    reuseExistingServer: !process.env.CI,
    timeout: 120_000,
  },
});
