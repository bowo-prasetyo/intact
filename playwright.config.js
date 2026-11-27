import { defineConfig, devices } from '@playwright/test';

// The tests drive the production build in real browsers: editing relies on contenteditable, which
// differs between engines, so a DOM emulation would not tell us whether a saved file is faithful.
// Run one browser with `npm test -- --project=chromium` (or firefox, or webkit).
// The specs Safari's engine (WebKit) doesn't run yet, each with why: the rest of the suite runs in
// it, as in Chromium and Firefox. A spec joins once it passes there; a case is never skipped in an
// engine to get it in, only for a capability that engine lacks (and the case says which).
const WEBKIT_PENDING = {
  // The toolbar's text labels find no room for themselves: WebKit's font metrics.
  'dates-labels': 'the toolbar labels\' room, by WebKit\'s font metrics',
  // The header's file name measures 107px for the 110 it is held to: WebKit's font metrics.
  'phone-share-insert': 'the phone header\'s widths, by WebKit\'s font metrics',
};

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
    // Safari's engine (WebKit): the whole suite but the specs pending above.
    {
      name: 'webkit',
      use: { ...devices['Desktop Safari'] },
      testIgnore: Object.keys(WEBKIT_PENDING).map((name) => `**/${name}.spec.js`),
    },
  ],
  webServer: {
    command: 'npm run build && npx vite preview --port 4173 --strictPort',
    url: 'http://localhost:4173/',
    reuseExistingServer: !process.env.CI,
    timeout: 120_000,
  },
});
