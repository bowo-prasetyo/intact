// The document's language is the page's (and a run's own language its own), so browsers that pick a
// spelling dictionary by language (Firefox) check the text in it. The status bar counts the words of
// the selection, as Word's does: "3 of 9 words".
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx, openInEditor } from './helpers.js';

const BODY = '<w:p><w:r><w:t xml:space="preserve">The colour of the sea. </w:t></w:r><w:r><w:rPr><w:lang w:val="fr-FR"/></w:rPr><w:t>La mer est bleue.</w:t></w:r></w:p>';

async function withLanguage(lang) {
  const zip = await JSZip.loadAsync(await buildDocx(BODY));
  const styles = (await zip.file('word/styles.xml').async('string'))
    .replace('<w:rPrDefault><w:rPr>', `<w:rPrDefault><w:rPr><w:lang w:val="${lang}" w:eastAsia="en-US" w:bidi="ar-SA"/>`);
  zip.file('word/styles.xml', styles);
  return zip.generateAsync({ type: 'nodebuffer' });
}

test('the page is in the document’s language, and a run in another language in its own', async ({ page }) => {
  await openInEditor(page, await withLanguage('en-GB'));
  await expect(page.locator('.docx-sheet')).toHaveAttribute('lang', 'en-GB');
  await expect(page.locator('.docx-page [lang="fr-FR"]')).toHaveText('La mer est bleue.');
});

test('the status bar counts the words selected', async ({ page }) => {
  await openInEditor(page, await withLanguage('en-GB'));
  const words = page.locator('.statusbar > span').first();
  await expect(words).toHaveText('9 words');
  await page.evaluate(() => {
    const t = document.querySelector('.docx-page p').firstChild;
    const r = document.createRange();
    r.setStart(t, 4);
    r.setEnd(t, 21); // "colour of the sea"
    getSelection().removeAllRanges();
    getSelection().addRange(r);
  });
  await expect(words).toHaveText('4 of 9 words');
  await page.evaluate(() => getSelection().collapseToEnd());
  await expect(words).toHaveText('9 words');
});
