// What an editor looks at again before filing a piece, in the document's own tools rather than a
// workspace beside it: Word count lists long paragraphs and a word typed twice in a row ("the the")
// with the long sentences, each to show one by one, reading the text as it will read (deleted and
// hidden text left out); and Copy as Markdown puts the document's own title, description (Comments)
// and keywords, from its properties, at the top as front matter for a site's page. The content
// creator workspace that used to sit in the header is gone, and what it kept in the browser with it.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor } from './helpers.js';
import { repeatedWords } from '../src/editor/readability.js';
import { frontMatter, toMarkdown } from '../src/editor/webCopy.js';

const para = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const status = (page) => page.locator('.statusbar [role=status]');
const count = (page) => page.getByRole('button', { name: 'Word count, length and readability' });
const LONG_PARAGRAPH = Array.from({ length: 16 }, (_, i) => `Sentence ${i + 1} of the long paragraph adds a few more words.`).join(' ');

test.beforeEach(async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  await page.addInitScript(() => {
    const set = DataTransfer.prototype.setData;
    DataTransfer.prototype.setData = function (type, data) {
      (window.copied ||= {})[type] = data;
      return set.call(this, type, data);
    };
  });
});

test('a word typed twice in a row is found; one doubled on purpose, a number or across punctuation is not', () => {
  const text = 'We went to the the shop. It had had its day. Well, well. Pages 10 10 and so so. In in the end.';
  const found = repeatedWords(text).map(({ start, end }) => text.slice(start, end));
  expect(found).toEqual(['the the', 'In in']);
});

test('front matter holds the title, description and keywords that are set, quoted for YAML', () => {
  expect(frontMatter({ title: 'Flood levy: what it costs', description: 'The council’s "temporary" levy, explained.', keywords: 'flood, levy; council' }))
    .toBe('---\ntitle: "Flood levy: what it costs"\ndescription: "The council’s \\"temporary\\" levy, explained."\nkeywords: ["flood", "levy", "council"]\n---\n\n');
  expect(frontMatter({ title: '', description: null, keywords: undefined })).toBe('');
  expect(toMarkdown([{ type: 'paragraph', inlines: [{ text: 'Hi.' }] }], { title: 'T' })).toBe('---\ntitle: "T"\n---\n\nHi.\n');
});

test('Word count lists long paragraphs and words typed twice, and Show selects each in turn', async ({ page }) => {
  const deleted = '<w:p><w:r><w:t xml:space="preserve">Kept text, </w:t></w:r><w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>gone gone gone</w:delText></w:r></w:del></w:p>';
  await openInEditor(page, await buildDocx(para('We went to the the shop.') + para(LONG_PARAGRAPH) + deleted + para('In in the end.')));
  await count(page).click();
  const dialog = page.getByRole('dialog', { name: 'Word count' });
  await expect(dialog).toContainText('1 long paragraph (more than 150 words)');
  // The tracked deletion ("gone gone gone") is text the author already took out: not listed.
  await expect(dialog).toContainText('2 words typed twice in a row');
  await expect(dialog).toContainText('Suggestions only: nothing is changed.');

  await dialog.getByRole('button', { name: 'Show: 2 words typed twice in a row' }).click();
  await expect(dialog).toBeHidden();
  expect(await page.evaluate(() => getSelection().toString())).toBe('the the');
  await expect(status(page)).toHaveText('Typed twice: “the the” (1 of 2) – Word count shows the next');
  await count(page).click();
  await dialog.getByRole('button', { name: 'Show: 2 words typed twice in a row' }).click();
  expect(await page.evaluate(() => getSelection().toString())).toBe('In in');

  await count(page).click();
  await dialog.getByRole('button', { name: 'Show: 1 long paragraph (more than 150 words)' }).click();
  expect(await page.evaluate(() => getSelection().toString())).toBe(LONG_PARAGRAPH);
  await expect(status(page)).toHaveText(/^Long paragraph: 1\d\d words \(1 of 1\)/);
});

test('a clean text says there is nothing to look at again', async ({ page }) => {
  await openInEditor(page, await buildDocx(para('A short and tidy paragraph.')));
  await count(page).click();
  const dialog = page.getByRole('dialog', { name: 'Word count' });
  await expect(dialog).toContainText('No paragraphs of more than 150 words');
  await expect(dialog).toContainText('No words typed twice in a row');
  await expect(dialog.getByRole('button', { name: /^Show/ })).toHaveCount(0);
});

test('Copy as Markdown starts with the title, description and keywords set in Document properties', async ({ page }) => {
  await openInEditor(page, await buildDocx(para('The council voted on Thursday.')));
  await page.getByRole('button', { name: /Insert a check box or date/ }).click();
  await page.getByRole('menuitem', { name: /Document property/ }).click();
  const props = page.getByRole('dialog', { name: 'Document properties' });
  await props.getByLabel('Title').fill('Council approves flood levy');
  await props.getByLabel('Keywords').fill('flood levy, council');
  await props.getByLabel('Comments').fill('What the levy costs riverside owners, and when it starts.');
  await expect(props).toContainText('Title 27 characters · Comments 57');
  await props.getByRole('button', { name: 'OK' }).click();
  await expect(props).toBeHidden();

  await page.evaluate(() => { window.copied = {}; });
  await page.getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menu', { name: 'Save' }).getByRole('menuitem', { name: /^Copy as Markdown/ }).click();
  expect(await page.evaluate(() => window.copied['text/plain'])).toBe('---\ntitle: "Council approves flood levy"\n'
    + 'description: "What the levy costs riverside owners, and when it starts."\nkeywords: ["flood levy", "council"]\n---\n\n'
    + 'The council voted on Thursday.\n');
  await expect(status(page)).toContainText('after front matter from the document’s properties');
});

test('the creator workspace is gone from the header, and what it kept in the browser is cleared', async ({ page }) => {
  await page.addInitScript(() => {
    if (!sessionStorage.getItem('seeded')) {
      localStorage.setItem('intact.creator.plan', '{"title":"A private brief"}');
      localStorage.setItem('intact.creator.dailyLog', '{"2026-11-26":420}');
      sessionStorage.setItem('seeded', '1');
    }
  });
  await openInEditor(page, await buildDocx(para('Text.')));
  await expect(page.getByRole('button', { name: 'Creator tools' })).toHaveCount(0);
  expect(await page.evaluate(() => Object.keys(localStorage).filter((k) => k.startsWith('intact.creator.')))).toEqual([]);
});
