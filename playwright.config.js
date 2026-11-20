import { defineConfig, devices } from '@playwright/test';

// The tests drive the production build in real browsers: editing relies on contenteditable, which
// differs between engines, so a DOM emulation would not tell us whether a saved file is faithful.
// Run one browser with `npm test -- --project=chromium` (or firefox).
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
  ],
  webServer: {
    command: 'npm run build && npx vite preview --port 4173 --strictPort',
    url: 'http://localhost:4173/',
    reuseExistingServer: !process.env.CI,
    timeout: 120_000,
  },
});
