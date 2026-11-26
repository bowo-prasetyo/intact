// Feedback: a plain link on the About page to the project's public form (.github/ISSUE_TEMPLATE/
// feedback.yml), opened in a new tab without a referrer, so it doesn't say where Intact is hosted.
// Intact sends nothing itself: no request leaves the page for the form's host, and there is no link
// in the editor, where a document is open.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor } from './helpers.js';

test('the About page links to the feedback form, and sends nothing itself', async ({ page }) => {
  const outside = [];
  page.on('request', (r) => { if (!r.url().startsWith('http://localhost:4173/') && !r.url().startsWith('data:') && !r.url().startsWith('blob:')) outside.push(r.url()); });
  await page.goto('./#/about');
  const link = page.getByRole('link', { name: 'a short public form' });
  await expect(link).toHaveAttribute('href', 'https://github.com/bowo-prasetyo/intact/issues/new?template=feedback.yml');
  await expect(link).toHaveAttribute('target', '_blank');
  await expect(link).toHaveAttribute('rel', 'noopener noreferrer');
  await expect(page.getByRole('heading', { name: 'Feedback' })).toBeVisible();
  expect(outside).toEqual([]);
});

test('the editor has no feedback link, while a document is open', async ({ page }) => {
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>Text.</w:t></w:r></w:p>'));
  await expect(page.locator('a[href*="issues/new"]')).toHaveCount(0);
});

test('the form asks for no document, and says so', () => {
  const form = fs.readFileSync(new URL('../.github/ISSUE_TEMPLATE/feedback.yml', import.meta.url), 'utf8');
  expect(form).toContain("Don't paste or attach a document");
  expect(form).not.toMatch(/type: (upload|file)/);
});
