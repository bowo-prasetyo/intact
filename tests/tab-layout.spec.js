// Tabs where Word puts them: to a right stop (a table of contents' page numbers, with a dot leader),
// a decimal stop (prices lined up at their points), and Insert > Answer line: an underlined tab to a
// right tab stop at the margin, as teachers make "Name: ______" in Word, drawn and printed alike.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const tabs = (list) => `<w:tabs>${list.map(([val, pos, leader]) => `<w:tab w:val="${val}"${leader ? ` w:leader="${leader}"` : ''} w:pos="${pos}"/>`).join('')}</w:tabs>`;
const run = (t, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${t === '\t' ? '<w:tab/>' : `<w:t xml:space="preserve">${t}</w:t>`}</w:r>`;
const para = (pPr, ...runs) => `<w:p><w:pPr>${pPr}</w:pPr>${runs.join('')}</w:p>`;

// A table of contents' line as Word writes it: the tab and page number in runs of their own.
const TOC = para(tabs([['right', 9000, 'dot']]), run('Introduction'), run('\t', '<w:noProof/><w:webHidden/>'), run('5', '<w:noProof/><w:webHidden/>'));
const TOTAL = para(tabs([['right', 6000]]), run('Total'), run('\t'), run('$100.00'));
const PRICE = para(tabs([['decimal', 4500]]), run('Price'), run('\t'), run('12.50'));
const NAME = '<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r></w:p>';

/** Where text `text` (in the paragraph holding it) ends and starts, in CSS px from where its paragraph's text starts. */
function edges(page, paraText, text) {
  return page.evaluate(([pt, t]) => {
    const p = [...document.querySelectorAll('.docx-page p')].find((x) => x.textContent.includes(pt));
    const cs = getComputedStyle(p);
    const x0 = p.getBoundingClientRect().left + parseFloat(cs.paddingLeft) + parseFloat(cs.borderLeftWidth);
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const at = n.data.indexOf(t);
      if (at < 0) continue;
      const r = document.createRange();
      r.setStart(n, at);
      r.setEnd(n, at + t.length);
      const q = r.getBoundingClientRect();
      return { left: q.left - x0, right: q.right - x0 };
    }
    return null;
  }, [paraText, text]);
}

test('a right tab with a dot leader: the page number ends at the stop, the dots drawn across the tab', async ({ page }) => {
  await openInEditor(page, await buildDocx(TOC + TOTAL));
  await expect.poll(async () => (await edges(page, 'Introduction', '5')).right).toBeCloseTo(600, 0); // 9000 twips
  const tab = page.locator('.docx-page p', { hasText: 'Introduction' }).locator('[data-tab-laid]');
  await expect(tab).toHaveCount(1);
  await expect(tab).toHaveCSS('background-image', /radial-gradient/);
  // A right tab alone in its paragraph: the amount ends at the stop too.
  await expect.poll(async () => (await edges(page, 'Total', '$100.00')).right).toBeCloseTo(400, 0);
  // Typing before it: the amount stays where it ends.
  await page.locator('.docx-page p', { hasText: 'Total' }).click();
  await page.keyboard.press('Home');
  await page.keyboard.type('Grand ');
  await expect.poll(async () => (await edges(page, 'Total', '$100.00')).right).toBeCloseTo(400, 0);
  // Nothing of it is saved: the file has the tab as it was.
  const { xml } = await download(page);
  expect(xml).toContain('<w:tab w:val="right" w:leader="dot" w:pos="9000"/>');
  expect(xml).not.toContain('tab-laid');
});

test('a decimal tab: the price lined up at its point', async ({ page }) => {
  await openInEditor(page, await buildDocx(PRICE));
  // 4500 twips (300px), to within a pixel: WebKit places the text at a whole pixel (299).
  await expect.poll(async () => Math.abs((await edges(page, 'Price', '.')).left - 300)).toBeLessThanOrEqual(1);
});

test('Insert > Answer line: an underlined tab to a right stop at the margin, drawn and printed', async ({ page }) => {
  await openInEditor(page, await buildDocx(NAME));
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await page.getByRole('button', { name: /Insert a check box or date/ }).click();
  await page.getByRole('menuitem', { name: /Answer line/ }).click();
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Answer line added: an underlined tab to a right tab stop at the margin, as Word makes one');
  const line = page.locator('.docx-page p [data-tab-laid]');
  await expect(line).toHaveCount(1);
  await expect(line).toHaveCSS('background-image', /linear-gradient/);
  // From after "Name:" to the paragraph's right edge.
  const reach = await page.evaluate(() => {
    const p = document.querySelector('.docx-page p');
    const el = p.querySelector('[data-tab-laid]');
    return { right: el.getBoundingClientRect().right, edge: p.getBoundingClientRect().right - parseFloat(getComputedStyle(p).paddingRight) };
  });
  expect(Math.abs(reach.right - reach.edge)).toBeLessThan(3);
  // It prints: the line is a background printed as it is.
  await page.emulateMedia({ media: 'print' });
  await expect(line).toHaveCSS('print-color-adjust', 'exact');
  await page.emulateMedia({ media: 'screen' });
  // Typing goes on after it.
  await page.keyboard.type('X');
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:tabs><w:tab w:val="right" w:pos="\d+"\/><\/w:tabs>/);
  expect(xml).toMatch(/<w:r><w:rPr><w:u w:val="single"\/><\/w:rPr><w:tab\/><\/w:r>/);
  expect(xml).toContain('Name: </w:t>');
});
