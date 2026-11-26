// "Instant": the start page shows at once. What it loads before anything else (the page's own
// script and stylesheet, and any module they preload) stays within a budget, and leaves out the
// document's engine (DocxDocument, JSZip and the rest), which is loaded the first time a document is
// opened or started, or once the start page is idle (see src/store.js). A change that pulls the
// engine back into the start page, or grows it past the budget, fails here.
import { test, expect } from '@playwright/test';

// Uncompressed bytes of the start page's script and styles. It was 609 kB with the engine in it,
// about 220 kB without; the rest is headroom for the start page's own growth, not for the engine.
const BUDGET = 280 * 1024;

test('the start page loads within its budget, without the document engine', async ({ request }) => {
  const html = await (await request.get('./')).text();
  const urls = [
    ...[...html.matchAll(/<script[^>]+type="module"[^>]+src="([^"]+)"/g)].map((m) => m[1]),
    ...[...html.matchAll(/<link[^>]+rel="(?:stylesheet|modulepreload)"[^>]+href="([^"]+)"/g)].map((m) => m[1]),
  ];
  expect(urls.length).toBeGreaterThan(0);
  let bytes = 0;
  for (const url of urls) {
    const body = await (await request.get(url)).body();
    bytes += body.length;
    // The engine's chunk, and JSZip in it, aren't among them.
    expect(url).not.toMatch(/DocxDocument|EditorView/);
    expect(body.toString('utf8')).not.toContain('JSZip');
  }
  expect(bytes).toBeLessThanOrEqual(BUDGET);
});

test('a document opens from the start page with the engine loaded on demand', async ({ page }) => {
  const engine = [];
  page.on('request', (r) => { if (/DocxDocument/.test(r.url())) engine.push(r.url()); });
  await page.goto('./');
  await page.getByRole('button', { name: /^Letter:/ }).click();
  await expect(page.locator('.docx-page')).toBeVisible();
  expect(engine.length).toBeGreaterThan(0);
});
