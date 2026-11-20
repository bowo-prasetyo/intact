// The rest of the writing track: Page setup for one section ("Apply to": this section, this point
// forward, the selected text — a landscape page for a wide table), a different first page and
// different odd and even pages for headers and footers, and Modify Style. Each written as Word
// writes it, and undone with Ctrl+Z.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx, openInEditor, download, downloadFile, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const part = (root, inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${root} ${W_NS}>${inner}</w:${root}>`;
const para = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const BODY = para('Intro.') + para('Wide table.') + para('More wide table.') + para('Closing.');
const dialog = (page) => page.getByRole('dialog', { name: 'Page setup' });
const status = (page) => page.locator('.statusbar [role=status]');
const sections = (xml) => xml.match(/<w:sectPr\b.*?<\/w:sectPr>/gs);
const p = (page, text) => page.locator('.docx-page > p', { hasText: text });

async function selectParagraphs(page, from, to) {
  await page.evaluate(([a, b]) => {
    const ps = [...document.querySelectorAll('.docx-page > p')];
    const start = ps.find((x) => x.textContent === a).firstChild;
    const end = ps.find((x) => x.textContent === b).firstChild;
    document.querySelector('.docx-page').focus();
    getSelection().setBaseAndExtent(start, 0, end, end.length);
  }, [from, to]);
}

async function pageSetup(page) {
  await page.getByRole('button', { name: 'Page setup' }).click();
  await expect(dialog(page)).toBeVisible();
  return dialog(page);
}

// ------------------------------------------------------------------ one section's page setup

test('the selected text on a landscape page: section breaks before and after it, the rest as it was', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await selectParagraphs(page, 'Wide table.', 'More wide table.');
  const d = await pageSetup(page);
  const scope = d.getByRole('combobox', { name: 'Apply paper and margins to' });
  await expect(scope).toHaveValue('whole');
  await scope.selectOption('selected');
  await expect(d.getByText('Section breaks go before and after the selected paragraphs')).toBeVisible();
  await d.getByLabel('Orientation').selectOption('landscape');
  await d.getByRole('button', { name: 'OK' }).click();
  await expect(status(page)).toHaveText('Page setup changed for the selected text (2 section breaks added)');
  // The breaks are named on the page with their section's paper.
  await expect(p(page, 'Intro.')).toHaveAttribute('data-section-note', 'Section 1: A4 portrait · section break (next page)');
  await expect(p(page, 'More wide table.')).toHaveAttribute('data-section-note', 'Section 2: A4 landscape · section break (next page)');

  const { xml } = await download(page);
  const [first, middle, last] = sections(xml);
  expect(first).toContain('<w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838"/>');
  expect(middle).toContain('<w:type w:val="nextPage"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>');
  expect(last).toContain('<w:pgSz w:w="11906" w:h="16838"/>');
  expect(last).not.toContain('landscape');
  // The breaks are in the paragraphs before and at the end of the selection.
  expect(xml).toMatch(/<w:p><w:pPr><w:sectPr>.*?<\/w:sectPr><\/w:pPr><w:r><w:t[^>]*>Intro\.<\/w:t>/s);
  expect(xml).toMatch(/<w:p><w:pPr><w:sectPr>.*?landscape.*?<\/w:sectPr><\/w:pPr><w:r><w:t[^>]*>More wide table\.<\/w:t>/s);

  // Ctrl+Z takes it all back.
  await page.locator('.docx-page').focus();
  await page.keyboard.press('Control+z');
  await expect(page.locator('.docx-page [data-section-end]')).toHaveCount(0);
});

test('this point forward: one break before the caret’s paragraph; this section changes only its own', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await p(page, 'Closing.').click();
  let d = await pageSetup(page);
  await d.getByRole('combobox', { name: 'Apply paper and margins to' }).selectOption('forward');
  await d.getByLabel('Left (cm)').fill('4');
  await d.getByRole('button', { name: 'OK' }).click();
  await expect(status(page)).toHaveText('Page setup changed from this point forward (1 section break added)');

  // Now two sections: in the first, "This section" is the default and changes it alone.
  await p(page, 'Intro.').click();
  d = await pageSetup(page);
  await expect(d.getByRole('combobox', { name: 'Apply paper and margins to' })).toHaveValue('section');
  await expect(d.getByRole('combobox', { name: 'Apply paper and margins to' }).locator('option:checked')).toHaveText('This section (1 of 2)');
  await expect(d.getByLabel('Left (cm)')).toHaveValue('2.54');
  await d.getByLabel('Top (cm)').fill('3');
  await d.getByRole('button', { name: 'OK' }).click();

  const { xml } = await download(page);
  const [first, last] = sections(xml);
  expect(first).toContain('w:top="1701"');
  expect(first).toContain('w:left="1440"');
  expect(last).toContain('w:top="1440"');
  expect(last).toContain('w:left="2268"');
});

// ------------------------------------------------------------------ headers and footers

const HF = {
  parts: {
    'word/header1.xml': part('hdr', para('Acme Ltd')),
    'word/footer1.xml': part('ftr', para('Confidential')),
  },
  rels: `<Relationship Id="rIdH1" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdF1" Type="${REL}/footer" Target="footer1.xml"/>`,
  sectPr: '<w:headerReference w:type="default" r:id="rIdH1"/><w:footerReference w:type="default" r:id="rIdF1"/>',
};

test('a different first page: a first-page header and footer to type in, written as Word writes them', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, HF));
  await expect(page.locator('.hf-tabs')).toHaveCount(0);
  const d = await pageSetup(page);
  await d.getByLabel('Different first page').check();
  await d.getByRole('button', { name: 'OK' }).click();
  const tabs = page.getByRole('tablist', { name: 'Headers' });
  await expect(tabs.getByRole('tab')).toHaveText(['Other pages', 'First page']);
  await tabs.getByRole('tab', { name: 'First page' }).click();
  await page.keyboard.type('Cover');
  await expect(page.locator('.docx-hf-band.header .docx-hf-part:visible')).toHaveText('Cover');

  const zip = await JSZip.loadAsync(fs.readFileSync(await downloadFile(page)));
  const doc = await zip.file('word/document.xml').async('string');
  const sect = sections(doc)[0];
  expect(sect).toMatch(/<w:headerReference w:type="first" r:id="rIdHeader1"\/>/);
  expect(sect).toMatch(/<w:footerReference w:type="first" r:id="rIdFooter1"\/>/);
  expect(sect).toContain('<w:titlePg/>');
  expect(await zip.file('word/header2.xml').async('string')).toContain('>Cover<');
  expect(await zip.file('word/header1.xml').async('string')).toContain('Acme Ltd');
});

test('different odd and even pages: Word’s setting, and even-page ones; switched off again, nothing is left', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, HF));
  let d = await pageSetup(page);
  await d.getByLabel('Different odd and even pages').check();
  await d.getByRole('button', { name: 'OK' }).click();
  await expect(page.getByRole('tablist', { name: 'Footers' }).getByRole('tab')).toHaveText(['Odd pages', 'Even pages']);
  let zip = await JSZip.loadAsync(fs.readFileSync(await downloadFile(page)));
  expect(await zip.file('word/settings.xml').async('string')).toContain('<w:evenAndOddHeaders/>');
  expect(sections(await zip.file('word/document.xml').async('string'))[0]).toContain('<w:headerReference w:type="even" r:id="rIdHeader1"/>');

  d = await pageSetup(page);
  await expect(d.getByLabel('Different odd and even pages')).toBeChecked();
  await d.getByLabel('Different odd and even pages').uncheck();
  await d.getByRole('button', { name: 'OK' }).click();
  await expect(page.locator('.hf-tabs')).toHaveCount(0);
  zip = await JSZip.loadAsync(fs.readFileSync(await downloadFile(page)));
  expect(await zip.file('word/settings.xml')?.async('string') ?? '').not.toContain('evenAndOddHeaders');
  expect(await zip.file('word/document.xml').async('string')).not.toContain('w:type="even"');
});

test('a first-page header the file has but doesn’t use shows when a different first page is switched on', async ({ page }) => {
  const opts = {
    parts: { ...HF.parts, 'word/header2.xml': part('hdr', para('Cover page header')) },
    rels: `${HF.rels}<Relationship Id="rIdH2" Type="${REL}/header" Target="header2.xml"/>`,
    sectPr: `${HF.sectPr}<w:headerReference w:type="first" r:id="rIdH2"/>`,
  };
  await openInEditor(page, await buildDocx(BODY, opts));
  // Unused (no w:titlePg): not in the page, so not found by Find and the like.
  await expect(page.locator('.docx-page')).not.toContainText('Cover page header');
  const d = await pageSetup(page);
  await d.getByLabel('Different first page').check();
  await d.getByRole('button', { name: 'OK' }).click();
  await page.getByRole('tablist', { name: 'Headers' }).getByRole('tab', { name: 'First page' }).click();
  await expect(page.locator('.docx-hf-band.header .docx-hf-part:visible')).toHaveText('Cover page header');
  const zip = await JSZip.loadAsync(fs.readFileSync(await downloadFile(page)));
  expect(await zip.file('word/document.xml').async('string')).toContain('<w:titlePg/>');
  // Untouched, its part goes back as it was.
  expect(await zip.file('word/header2.xml').async('string')).toBe(opts.parts['word/header2.xml']);
});

// ------------------------------------------------------------------ Modify Style

const STYLED = '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Results</w:t></w:r></w:p>' + para('Plain text.');

test('Modify Style changes the style itself: every paragraph in it follows, and Ctrl+Z puts it back', async ({ page }) => {
  await openInEditor(page, await buildDocx(STYLED));
  await page.locator('.docx-page h1, .docx-page [data-style="Heading1"]').first().click();
  await page.getByRole('combobox', { name: 'Paragraph style' }).selectOption('__modify');
  const d = page.getByRole('dialog', { name: /Modify style/ });
  await expect(d.getByRole('heading')).toHaveText(/Modify style: (Heading 1|heading 1)/);
  await expect(d.getByLabel('Style font size (pt)')).toHaveValue('16');
  await expect(d.getByLabel('Bold')).toBeChecked();
  await d.getByLabel('Style font size (pt)').fill('20');
  await d.getByLabel('Bold').uncheck();
  await d.getByLabel('Automatic').uncheck();
  await d.getByLabel('Style colour').fill('#0e6467');
  await d.getByLabel('Style alignment').selectOption('center');
  await d.getByRole('button', { name: 'OK' }).click();
  const heading = page.locator('.docx-page > :first-child');
  await expect(heading).toHaveCSS('font-size', '26.6667px'); // 20 pt
  await expect(heading).toHaveCSS('text-align', 'center');
  await expect(status(page)).toContainText('changed: every paragraph in it follows');

  const { zip } = await download(page);
  const styles = await zip.file('word/styles.xml').async('string');
  const h1 = styles.match(/<w:style [^>]*w:styleId="Heading1".*?<\/w:style>/s)[0];
  expect(h1).toContain('<w:jc w:val="center"/>');
  expect(h1).toContain('<w:b w:val="0"/>');
  expect(h1).toContain('<w:color w:val="0E6467"/>');
  expect(h1).toContain('<w:sz w:val="40"/><w:szCs w:val="40"/>');
  // The paragraph itself isn't changed: it is the style.
  expect(await zip.file('word/document.xml').async('string')).toContain('<w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Results</w:t>');

  await page.locator('.docx-page').focus();
  await page.keyboard.press('Control+z');
  await expect(heading).toHaveCSS('font-size', '21.3333px'); // 16 pt again
});

test('Modify style… from the right-click menu changes Normal, the body text’s style', async ({ page }) => {
  await openInEditor(page, await buildDocx(STYLED));
  await p(page, 'Plain text.').click({ button: 'right' });
  await page.getByRole('menu', { name: 'Context menu' }).getByRole('menuitem', { name: 'Modify style…' }).click();
  const d = page.getByRole('dialog', { name: /Modify style/ });
  await d.getByLabel('Style font', { exact: true }).selectOption('Times New Roman');
  await d.getByLabel('Style space after (pt)').fill('6');
  await d.getByRole('button', { name: 'OK' }).click();
  await expect(p(page, 'Plain text.')).toHaveCSS('font-family', /Times New Roman/);
  const { zip } = await download(page);
  const normal = (await zip.file('word/styles.xml').async('string')).match(/<w:style [^>]*w:styleId="Normal".*?<\/w:style>/s)[0];
  expect(normal).toContain('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>');
  expect(normal).toContain('<w:spacing w:after="120"/>');
});
