// Templates (.dotx) and files with macros (.docm, .dotm) open and are saved as what they are; macros
// are never run, and are kept. Line numbering (w:lnNumType) is shown in the margin, as Word prints it.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx, openInEditor } from './helpers.js';

const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const TYPES = {
  dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
  docm: 'application/vnd.ms-word.document.macroEnabled.main+xml',
  dotm: 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml',
};

/** A document saved as `ext`: its main part's content type, and for macros a VBA project. */
async function asFormat(ext, body = '<w:p><w:r><w:t>Body text.</w:t></w:r></w:p>') {
  const zip = await JSZip.loadAsync(await buildDocx(body));
  const ct = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', ct.replace(MAIN, TYPES[ext]).replace('</Types>', ext.endsWith('m') ? '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>' : '</Types>'));
  if (ext.endsWith('m')) zip.file('word/vbaProject.bin', Buffer.from('VBA project bytes'));
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function downloaded(page) {
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  const [dl] = await Promise.all([page.waitForEvent('download'), review.getByRole('button', { name: 'Download', exact: true }).click()]);
  return { name: dl.suggestedFilename(), zip: await JSZip.loadAsync(fs.readFileSync(await dl.path())) };
}

for (const ext of ['dotx', 'docm', 'dotm']) {
  test(`a .${ext} opens and is saved as a .${ext}`, async ({ page }) => {
    await openInEditor(page, await asFormat(ext), `Letter.${ext}`);
    await expect(page.locator('.app-header .ext')).toHaveText(`.${ext}`);
    await page.locator('.docx-page p').first().click();
    await page.keyboard.press('End');
    await page.keyboard.type(' Edited');
    const { name, zip } = await downloaded(page);
    expect(name).toBe(`Letter.${ext}`);
    expect(await zip.file('word/document.xml').async('string')).toContain('Edited');
    expect(await zip.file('[Content_Types].xml').async('string')).toContain(TYPES[ext]);
    if (ext.endsWith('m')) expect(Buffer.from(await zip.file('word/vbaProject.bin').async('uint8array')).toString()).toBe('VBA project bytes');
  });
}

test('a file with macros says that they are kept and never run', async ({ page }) => {
  await openInEditor(page, await asFormat('docm'), 'Form.docm');
  const bar = page.getByRole('region', { name: 'Restrictions on this document' });
  await expect(bar).toContainText('Contains macros. Intact never runs them: they are kept in the file as they are.');
  await expect(bar.getByRole('button', { name: 'Edit anyway' })).toHaveCount(0); // nothing is locked
  await page.locator('.docx-page p').first().click();
  await page.keyboard.type('x');
  await expect(page.locator('.docx-page p').first()).toContainText('x');
});

// ------------------------------------------------------------------ line numbers

const numbers = (page) => page.locator('.line-numbers-layer .line-number').allTextContents();
const para = (text, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;

test('line numbers are shown in the margin, every line, as Word numbers them', async ({ page }) => {
  const sectPr = '<w:lnNumType w:countBy="1" w:restart="continuous"/>';
  const body = [para('One'), para('Two'), para('Not counted', '<w:suppressLineNumbers/>'), para('Three')].join('')
    + '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>In a table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    + para('Four');
  await openInEditor(page, await buildDocx(body, { sectPr }));
  await expect.poll(() => numbers(page)).toEqual(['1', '2', '3', '4']);
  // Each number is in the margin, level with its line.
  const [n1, line1] = await Promise.all([page.locator('.line-number').first().boundingBox(), page.locator('.docx-page p').first().boundingBox()]);
  expect(n1.x + n1.width).toBeLessThan(line1.x + 1);
  expect(Math.abs(n1.y - line1.y)).toBeLessThan(6);
  // They follow the text as it changes.
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.type('New line');
  await expect.poll(() => numbers(page)).toEqual(['1', '2', '3', '4', '5']);
});

test('line numbers count by, start at, and restart where Word began a page', async ({ page }) => {
  const sectPr = '<w:lnNumType w:countBy="2" w:start="10"/>';
  const body = [para('a'), para('b'), para('c'), '<w:p><w:r><w:lastRenderedPageBreak/><w:t>d (page 2)</w:t></w:r></w:p>', para('e'), para('f')].join('');
  await openInEditor(page, await buildDocx(body, { sectPr }));
  // 11 12 13 | 11 12 13: every second number shown.
  await expect.poll(() => numbers(page)).toEqual(['12', '12']);
});

test('without line numbering, none are shown', async ({ page }) => {
  await openInEditor(page, await buildDocx(para('Plain')));
  await expect(page.locator('.line-numbers-layer')).toHaveCount(0);
});
