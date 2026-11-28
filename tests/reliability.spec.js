// P0 reliability regressions: large-document preservation and the local-only privacy contract.
// These run against the production build in Chromium, Firefox and WebKit.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf } from './helpers.js';

test('a large document opens and saves with only the edited paragraph changed', async ({ page }) => {
  test.setTimeout(120_000);
  const count = 400;
  const body = Array.from({ length: count }, (_, i) =>
    `<w:p><w:r><w:t>Section ${i + 1}: representative text that must survive a round trip.</w:t></w:r></w:p>`,
  ).join('');
  const source = await buildDocx(body);
  await openInEditor(page, source, 'large-reliability-fixture.docx');

  const paragraphs = page.locator('.docx-page > p');
  await expect(paragraphs).toHaveCount(count);
  const editedIndex = 199;
  await paragraphs.nth(editedIndex).click();
  await page.keyboard.press('End');
  await page.keyboard.type(' [verified edit]');

  const { xml } = await download(page);
  const before = await canonicalize(page, await documentXmlOf(source));
  const after = await canonicalize(page, xml);

  expect(after, 'no paragraph is lost or added').toHaveLength(before.length);
  for (let i = 0; i < before.length; i++) {
    if (i === editedIndex) {
      expect(JSON.stringify(after[i])).toContain('[verified edit]');
    } else {
      expect(after[i], `paragraph ${i + 1} is preserved`).toEqual(before[i]);
    }
  }
});

test('opening and saving a local document makes no cross-origin requests', async ({ page }) => {
  const crossOriginRequests = [];
  const appOrigin = new URL('http://localhost:4173').origin;
  page.on('request', (request) => {
    const requested = new URL(request.url());
    if (requested.protocol === 'http:' || requested.protocol === 'https:') {
      if (requested.origin !== appOrigin) crossOriginRequests.push(request.url());
    }
  });

  const source = await buildDocx('<w:p><w:r><w:t>Private local document</w:t></w:r></w:p>');
  await openInEditor(page, source, 'private-local.docx');

  const csp = await page.locator('meta[http-equiv="Content-Security-Policy"]').getAttribute('content');
  expect(csp).toContain("default-src 'none'");
  expect(csp).toContain("connect-src 'none'");
  expect(csp).toContain("object-src 'none'");
  expect(csp).toContain("base-uri 'none'");

  await download(page);
  expect(crossOriginRequests, 'the app must not contact another origin while handling a local file').toEqual([]);
});
