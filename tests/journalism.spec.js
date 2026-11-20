// What a reporter needs from a story file before filing it: Word count (a click on the status bar's
// count, as in Word) with a length target in words and a headline's limit in characters, shown in
// the status bar as they write, and Word's Readability Statistics with the long sentences to show;
// Redact, which blacks out a source's name for good (never tracked, a link's address gone with it)
// and "Before you release" saying where a redacted name is still in the file; and the accessibility
// check's plain-language tip.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const status = (page) => page.locator('.statusbar [role=status]');
const count = (page) => page.getByRole('button', { name: 'Word count, length and readability' });
const para = (text, style = '') => `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const HEADLINE = 'Council approves flood levy after heated debate'; // 47 characters
const LONG = 'The city council voted on Thursday night to impose a temporary levy on riverside properties, ending months of argument over who should pay for the new flood walls that engineers say the town will need within ten years.';
const STORY = para(HEADLINE, 'Heading1')
  + para(`${LONG} The vote was close.`)
  + para('“We cannot keep asking the rest of the city to pay,” said Councillor Ana Ruiz. Ana Ruiz chairs the finance committee.');

/** Selects `text` (its first occurrence) in the page. */
async function select(page, text) {
  await page.evaluate((w) => {
    const walker = document.createTreeWalker(document.querySelector('.docx-page'), NodeFilter.SHOW_TEXT);
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      const at = t.data.indexOf(w);
      if (at < 0) continue;
      document.querySelector('.docx-page').focus();
      getSelection().setBaseAndExtent(t, at, t, at + w.length);
      return;
    }
    throw new Error(`Not found: ${w}`);
  }, text);
}

async function openReview(page) {
  await page.locator('.review-toggle').click();
  return page.getByRole('dialog', { name: 'Review your changes' });
}

test.beforeEach(async ({ page }) => page.setViewportSize({ width: 1280, height: 800 }));

test('Word count shows the counts and Word’s readability statistics, from the status bar', async ({ page }) => {
  await openInEditor(page, await buildDocx(STORY));
  await count(page).click();
  const dialog = page.getByRole('dialog', { name: 'Word count' });
  const value = (label) => dialog.locator('dt', { hasText: new RegExp(`^${label}$`) }).locator('xpath=following-sibling::dd[1]');
  await expect(value('Words')).toHaveText('70');
  await expect(value('Paragraphs')).toHaveText('3');
  await expect(value('Sentences')).toHaveText('5');
  await expect(value('Words per sentence')).toHaveText('14');
  await expect(value('Flesch Reading Ease')).toHaveText(/^\d+(\.\d)? \(/);
  await expect(value('Flesch-Kincaid Grade Level')).toHaveText(/^\d+(\.\d)?$/);
  await expect(value('Passive sentences')).toHaveText('0%');
  await expect(dialog).toContainText('Headline: 47 characters, with spaces – “Council approves flood levy after heated debate”');
  await expect(dialog).toContainText('1 long sentence (more than 30 words)');

  // Show selects the long sentence, to cut or split it.
  await dialog.getByRole('button', { name: 'Show' }).click();
  await expect(dialog).toBeHidden();
  expect(await page.evaluate(() => getSelection().toString())).toBe(LONG);
  await expect(status(page)).toHaveText('Long sentence: 38 words (1 of 1) – Word count shows the next');
});

test('a length target and a headline limit show in the status bar, over them in amber', async ({ page }) => {
  await openInEditor(page, await buildDocx(STORY));
  await count(page).click();
  const dialog = page.getByRole('dialog', { name: 'Word count' });
  await dialog.getByLabel('Target, in words').fill('100');
  await expect(dialog).toContainText('30 words to go');
  await dialog.getByLabel('Headline limit, in characters').fill('40');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(count(page)).toHaveText('70 / 100 words · 402 characters');
  await expect(count(page)).not.toHaveClass(/over/);
  const headline = page.locator('.headline-count');
  await expect(headline).toHaveText('Headline 47 / 40');
  await expect(headline).toHaveClass(/over/);
  await expect(headline).toHaveAttribute('title', /7 over the limit$/);

  // Cutting the headline brings it under the limit, as you type.
  await select(page, ' after heated debate');
  await page.keyboard.press('Delete');
  await expect(headline).toHaveText('Headline 27 / 40');
  await expect(headline).not.toHaveClass(/over/);

  // A lower target: over it, in amber; a selection's count is shown as before.
  await count(page).click();
  await dialog.getByLabel('Target, in words').fill('50');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(count(page)).toHaveText(/^67 \/ 50 words/);
  await expect(count(page)).toHaveClass(/over/);
  await select(page, 'Ana Ruiz');
  await expect(count(page)).toHaveText(/^2 of 67 words/);

  // Cleared, the status bar is as it was, and the target isn't written into the file.
  await page.evaluate(() => getSelection().collapseToEnd());
  await count(page).click();
  await dialog.getByLabel('Target, in words').fill('');
  await dialog.getByLabel('Headline limit, in characters').fill('');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(count(page)).toHaveText(/^67 words/);
  await expect(headline).toBeHidden();
  const { zip } = await download(page);
  expect(Object.keys(zip.files).some((p) => /custom\.xml$/.test(p))).toBe(false);
});

test('the scores and the passive voice are for English; another language gets the counts', async ({ page }) => {
  const styles = `${XML}<w:styles ${W_NS}><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="id-ID"/></w:rPr></w:rPrDefault></w:docDefaults>`
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>';
  await openInEditor(page, await buildDocx(para('Rapat dimulai pukul 09.00. Semua peserta hadir.'), { parts: { 'word/styles.xml': styles } }));
  await count(page).click();
  const dialog = page.getByRole('dialog', { name: 'Word count' });
  await expect(dialog.locator('dt', { hasText: 'Sentences per paragraph' })).toBeVisible();
  await expect(dialog.locator('dt', { hasText: 'Flesch Reading Ease' })).toHaveCount(0);
  await expect(dialog).toContainText('The passive voice and the Flesch scores are worked out for English text, as in Word.');
});

test('Redact blacks out the selection for good, also while tracking changes', async ({ page }) => {
  await openInEditor(page, await buildDocx(STORY));
  // Track changes on: the redaction still isn't tracked, so the name isn't kept as a deletion.
  await page.getByRole('button', { name: 'Track changes' }).click();
  const who = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await who.getByRole('textbox').fill('Desk');
  await who.getByRole('button', { name: 'Start tracking' }).click();
  await select(page, 'Councillor Ana Ruiz');
  await page.getByRole('button', { name: 'Redact', exact: true }).click();
  await expect(page.locator('.docx-page')).toContainText('said ███████████████████. Ana Ruiz chairs');
  await expect(status(page)).toContainText('Redacted (not tracked: a tracked deletion would keep the text in the file)');
  const { xml } = await download(page);
  expect(xml).toContain('███████████████████');
  expect(xml).not.toContain('Councillor');
  expect(xml).not.toContain('<w:del ');
});

test('Redact refuses a selection with a tracked change in it, whose earlier wording would stay', async ({ page }) => {
  const body = '<w:p><w:r><w:t xml:space="preserve">Said by </w:t></w:r><w:del w:id="1" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Ana Ruiz</w:delText></w:r></w:del><w:r><w:t xml:space="preserve"> to us.</w:t></w:r></w:p>';
  await openInEditor(page, await buildDocx(body));
  await select(page, 'Said by ');
  await page.evaluate(() => {
    const ps = document.querySelector('.docx-page p');
    const last = ps.lastChild.nodeType === 3 ? ps.lastChild : ps.lastChild.lastChild;
    const sel = getSelection();
    sel.setBaseAndExtent(sel.anchorNode, 0, last, last.length);
  });
  await page.getByRole('button', { name: 'Redact', exact: true }).click();
  await expect(status(page)).toHaveText('The selection holds tracked changes: accept or reject them first (Review › Accept), so no earlier wording stays in the file');
  await expect(page.locator('.docx-page')).not.toContainText('█');
});

test('a redacted link goes, address and all; Undo brings it back with its address', async ({ page }) => {
  // rId3 (in the helper's relationships) is the link's address: https://example.com/.
  const body = `<w:p><w:r><w:t xml:space="preserve">Her profile: </w:t></w:r><w:hyperlink r:id="rId3" xmlns:r="${REL}"><w:r><w:t>Ana Ruiz online</w:t></w:r></w:hyperlink><w:r><w:t>.</w:t></w:r></w:p>`;
  await openInEditor(page, await buildDocx(body));
  await select(page, 'Ana Ruiz online');
  await page.getByRole('button', { name: 'Redact', exact: true }).click();
  await expect(page.locator('.docx-page a[href]')).toHaveCount(0);
  let { zip, xml } = await download(page);
  expect(xml).not.toContain('w:hyperlink');
  expect(xml).toContain('Her profile: ███████████████.');
  expect(await zip.file('word/_rels/document.xml.rels').async('string')).not.toContain('example.com');

  await page.locator('.docx-page').focus();
  await page.keyboard.press('Control+z');
  await expect(page.locator('.docx-page a[href]')).toHaveText('Ana Ruiz online');
  ({ zip, xml } = await download(page));
  expect(xml).toContain('w:hyperlink r:id="rId3"');
  expect(await zip.file('word/_rels/document.xml.rels').async('string')).toContain('Target="https://example.com/"');
});

test('Before you release says where a redacted name still is, and redacts the rest of the text', async ({ page }) => {
  const comments = `${XML}<w:comments ${W_NS}><w:comment w:id="0" w:author="Desk" w:date="2026-01-05T10:00:00Z" w:initials="D"><w:p><w:r><w:t>Is Ana Ruiz on the record?</w:t></w:r></w:p></w:comment></w:comments>`;
  const body = '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t xml:space="preserve">A source, Ana Ruiz, spoke to us.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>'
    + para('Ana Ruiz asked not to be named.');
  await openInEditor(page, await buildDocx(body, { parts: { 'word/comments.xml': comments }, rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>` }));
  await select(page, 'Ana Ruiz');
  await page.getByRole('button', { name: 'Redact', exact: true }).click();
  await expect(status(page)).toContainText('“Ana Ruiz” is in the document 1 more times: Before you release (Review) redacts them all');

  let review = await openReview(page);
  await expect(review).toContainText('1 redacted passage (█). Redacting here isn’t tracked, so the text under it isn’t in the file.');
  await expect(review).toContainText('What you redacted is still in the file: “Ana Ruiz” once in the text and in 1 comment (edit or delete it).');
  await review.getByRole('button', { name: 'Redact them all' }).click();
  review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review).toContainText('2 redacted passages (█).');
  await expect(review).toContainText('What you redacted is still in the file: “Ana Ruiz” in 1 comment (edit or delete it).');
  await expect(review.getByRole('button', { name: 'Redact them all' })).toHaveCount(0);
  await review.getByRole('button', { name: 'Keep editing' }).click();
  await expect(page.locator('.docx-page')).not.toContainText('Ana Ruiz');
});

test('the accessibility check suggests splitting long sentences, and shows them', async ({ page }) => {
  await openInEditor(page, await buildDocx(STORY));
  const review = await openReview(page);
  const tip = review.locator('li', { hasText: 'Hard to read' });
  await expect(tip).toContainText('1 sentence of more than 30 words.');
  await tip.getByRole('button', { name: 'Show' }).click();
  await expect(status(page)).toHaveText('Hard to read');
});

test('a length target stays with the document’s unsaved copy, through a reload', async ({ page }) => {
  await openInEditor(page, await buildDocx(STORY), 'levy.docx');
  await count(page).click();
  const dialog = page.getByRole('dialog', { name: 'Word count' });
  await dialog.getByLabel('Target, in words').fill('600');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await page.locator('.docx-page p').last().click();
  await page.keyboard.press('End');
  await page.keyboard.type(' She declined to comment.');
  await expect(page.locator('.statusbar .draft-status')).toHaveText('Changes kept on this device');
  await page.reload();
  await page.getByRole('region', { name: 'Unsaved changes on this device' }).getByRole('button', { name: 'Restore levy.docx' }).click();
  await expect(count(page)).toHaveText(/^74 \/ 600 words/);
});
