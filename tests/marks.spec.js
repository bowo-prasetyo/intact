// "Show formatting marks" (¶): · for spaces, ° for non-breaking spaces, → for tabs, ↵ for line
// breaks and ¶ for paragraph ends, drawn over the page. Showing them never changes the document.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf } from './helpers.js';

const DOC = '<w:p><w:r><w:t xml:space="preserve">One two three</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t xml:space="preserve">Tab</w:t><w:tab/><w:t xml:space="preserve">here and§ 5</w:t><w:br/><w:t>next line</w:t></w:r></w:p>'
  + '<w:p/>'
  + '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  + '<w:p><w:r><w:t>End.</w:t></w:r></w:p>';

const marks = (page, kind) => page.locator(`.marks-layer .fmark.${kind}`);
const toggle = (page) => page.getByTitle('Show formatting marks (Ctrl+Shift+8)');

test('the ¶ button shows marks for spaces, tabs, line breaks and paragraph ends', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await expect(page.locator('.marks-layer .fmark')).toHaveCount(0);
  await toggle(page).click();
  await expect(toggle(page)).toHaveAttribute('aria-pressed', 'true');
  await expect(marks(page, 'para')).toHaveCount(5); // four paragraphs and the table cell's
  await expect(marks(page, 'space')).toHaveText(['·', '·', '·', '°', '·']); // "§\u00a05" comes before "next line"
  await expect(marks(page, 'tab')).toHaveText(['→']);
  await expect(marks(page, 'break')).toHaveText(['↵']);

  // Each ¶ sits just after its paragraph's text, on its line.
  const end = await page.locator('.docx-page p').first().evaluate((p) => {
    const range = document.createRange();
    range.selectNodeContents(p);
    const rects = range.getClientRects();
    return rects[rects.length - 1].right;
  });
  const first = await marks(page, 'para').first().boundingBox();
  expect(Math.abs(first.x - end)).toBeLessThan(4);

  await toggle(page).click();
  await expect(page.locator('.marks-layer .fmark')).toHaveCount(0);
});

test('showing marks changes neither the page nor the saved document', async ({ page }) => {
  const docx = await buildDocx(DOC);
  await openInEditor(page, docx);
  const before = await page.locator('.docx-page').innerHTML();
  await page.keyboard.press('ControlOrMeta+Shift+Digit8');
  await expect(marks(page, 'para')).toHaveCount(5);
  expect(await page.locator('.docx-page').innerHTML()).toBe(before);
  const { xml } = await download(page);
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
});

test('marks follow typing, and stay on for the next document', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await toggle(page).click();
  await expect(marks(page, 'space')).toHaveCount(5);
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await page.keyboard.type(' four');
  await expect(marks(page, 'space')).toHaveCount(6);
  await page.keyboard.press('Enter');
  await expect(marks(page, 'para')).toHaveCount(6);

  // Clicking through the layer still places the caret in the text.
  const mark = await marks(page, 'space').first().boundingBox();
  await page.mouse.click(mark.x + mark.width / 2, mark.y + mark.height / 2);
  await expect.poll(() => page.evaluate(() => document.activeElement?.classList.contains('docx-page'))).toBe(true);

  // Another tab of the same browser: the choice is remembered.
  const other = await page.context().newPage();
  await openInEditor(other, await buildDocx(DOC));
  await expect(toggle(other)).toHaveAttribute('aria-pressed', 'true');
  await expect(marks(other, 'para')).toHaveCount(5);
});
