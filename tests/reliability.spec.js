// Large documents: a document of 400 paragraphs opens, takes an edit and saves with that paragraph
// exactly as edited and every other one unchanged. How long opening and saving took is recorded as
// test annotations (in the report) and in the log, so baselines collect on CI's runners before any budget is set.
// (Privacy, the other half of reliability's browser checks, is in tests/deployment.spec.js.)
import { test, expect } from '@playwright/test';
import fs from 'node:fs';
import JSZip from 'jszip';
import { buildDocx, openInEditor, canonicalize, documentXmlOf } from './helpers.js';

test('a large document opens and saves with only the edited paragraph changed, exactly', async ({ page }, testInfo) => {
  test.setTimeout(120_000);
  const count = 400;
  const text = (i) => `Section ${i + 1}: representative text that must survive a round trip.`;
  const source = await buildDocx(Array.from({ length: count }, (_, i) => `<w:p><w:r><w:t>${text(i)}</w:t></w:r></w:p>`).join(''));

  let started = Date.now();
  await openInEditor(page, source, 'large-reliability-fixture.docx');
  const paragraphs = page.locator('.docx-page > p');
  await expect(paragraphs).toHaveCount(count);
  testInfo.annotations.push({ type: 'open (ms)', description: String(Date.now() - started) });

  const edited = 199;
  await paragraphs.nth(edited).click();
  await page.keyboard.press('End');
  await page.keyboard.type(' [verified edit]');

  // Saving, as the person waits for it: from Download (in the review) until the status bar says it
  // is saved. (Writing the downloaded file to disk is the browser's, and slower in WebKit's test
  // build than anywhere a person downloads: it isn't timed.)
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review).toBeVisible();
  const downloaded = page.waitForEvent('download');
  started = Date.now();
  await review.getByRole('button', { name: 'Download' }).click();
  await expect(page.locator('.statusbar [role=status]')).toContainText('Saved');
  testInfo.annotations.push({ type: 'save (ms)', description: String(Date.now() - started) });
  const xml = await (await JSZip.loadAsync(fs.readFileSync(await (await downloaded).path()))).file('word/document.xml').async('string');
  // In CI's log too (its list reporter doesn't show annotations).
  console.log(`Large document, ${testInfo.project.name}: ${testInfo.annotations.map((a) => `${a.type} ${a.description}`).join(', ')}`);

  const before = await canonicalize(page, await documentXmlOf(source));
  const after = await canonicalize(page, xml);
  expect(after, 'no paragraph is lost or added').toHaveLength(before.length);
  expect(after[edited].p.inlines).toEqual([{ rPr: '', text: `${text(edited)} [verified edit]` }]);
  for (let i = 0; i < before.length; i++) if (i !== edited) expect(after[i], `paragraph ${i + 1} is unchanged`).toEqual(before[i]);
});
