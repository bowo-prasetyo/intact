import { test, expect } from '@playwright/test';
import { editorialChecks, repurpose, seoChecks, wordsIn } from '../src/editor/creator.js';
import { buildDocx, openInEditor } from './helpers.js';

test('editorial checks identify repeated words, long sentences and paragraphs', () => {
  const longSentence = `This sentence is deliberately long and contains enough individual words to cross the thirty word threshold so that the editor can suggest splitting it into smaller and clearer sentences for readers.`;
  const result = editorialChecks(`A useful useful introduction. ${longSentence}\n\n${'paragraph '.repeat(155)}`);
  expect(result.issues.map((issue) => issue.type)).toContain('Repeated word');
  expect(result.issues.map((issue) => issue.type)).toContain('Long sentence');
  expect(result.issues.map((issue) => issue.type)).toContain('Long paragraph');
  expect(wordsIn('one two three')).toBe(3);
});

test('SEO checks report title, description, keyword, structure and link signals', () => {
  const result = seoChecks({
    title: 'A practical guide to sustainable urban gardening',
    description: 'Learn how to plan a small urban garden, choose suitable plants, prepare healthy soil and maintain a productive growing space throughout the year.',
    keyword: 'urban gardening',
    text: '# Urban gardening\n\nUrban gardening can make a small space productive.\n\n## Getting started\n\nRead https://example.com for more.',
  });
  expect(result.checks.find((x) => x.label === 'Keyword in title').ok).toBe(true);
  expect(result.checks.find((x) => x.label === 'Keyword near the beginning').ok).toBe(true);
  expect(result.checks.find((x) => x.label === 'Links included').ok).toBe(true);
  expect(result.occurrences).toBe(2);
});

test('repurposing produces bounded, editable platform drafts', () => {
  const drafts = repurpose('A practical story begins here. It explains the central idea. More detail follows in the article.', 'A practical story');
  expect(drafts.linkedin).toContain('What would you add');
  expect(drafts.x.length).toBeLessThanOrEqual(280);
  expect(drafts.instagram).toContain('#Apracticalstory');
  expect(drafts.newsletter).toContain('In this edition');
});

test('Content Creator Studio opens from the editor and persists the daily writing log locally', async ({ page }) => {
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>A useful article for readers.</w:t></w:r></w:p>'));
  await page.getByRole('button', { name: 'Creator tools' }).click();
  const dialog = page.getByRole('dialog', { name: 'Content Creator Studio' });
  await expect(dialog).toBeVisible();
  await expect(dialog).toContainText('Current document');
  await dialog.getByLabel('Words written today').fill('420');
  await dialog.getByRole('button', { name: 'Save goals and today\'s log' }).click();
  await expect(dialog.getByRole('status')).toContainText('Saved on this device.');
  expect(await page.evaluate(() => JSON.parse(localStorage.getItem('intact.creator.dailyLog')))).toBeTruthy();
  await dialog.getByRole('button', { name: 'Quality check' }).click();
  await expect(dialog).toContainText('Editorial quality check');
  await dialog.getByRole('button', { name: 'SEO' }).click();
  await expect(dialog).toContainText('SEO preparation checklist');
  await dialog.getByRole('button', { name: 'Repurpose' }).click();
  await expect(dialog.getByLabel('LinkedIn post draft')).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(dialog).toHaveCount(0);
});
