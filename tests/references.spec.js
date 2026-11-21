// References, as on Word's References tab, from the Insert menu: a table of contents, captions,
// cross-references, equations (typed as LaTeX, written as Office Math), citations in APA or MLA
// style with the document's sources (Word's b:Sources part in customXml), and a bibliography;
// Update (F9) brings citations and the bibliography up to date. Also the Navigation pane's
// headings, and the review's notes left out where they don't apply (a document started here, a
// document that isn't a CV).
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const H1 = (t) => `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
const P = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const ESSAY = P('My Essay') + H1('Introduction') + P('Memory matters.') + H1('Method') + P('We measured it.');
const status = (page) => page.locator('.statusbar [role=status]');

async function caretAt(page, text, where = 'End') {
  await page.locator('.docx-page > :is(p, h1, h2, h3)', { hasText: text }).first().click();
  await page.keyboard.press(where);
}
async function insertMenu(page, item) {
  await page.getByRole('button', { name: /Insert a check box or date/ }).click();
  await page.getByRole('menuitem', { name: item }).click();
}

test('Table of contents: "Contents" and Word\'s TOC field, built from the headings and linked to them', async ({ page }) => {
  await openInEditor(page, await buildDocx(ESSAY));
  await caretAt(page, 'My Essay');
  await page.keyboard.press('Enter');
  await insertMenu(page, /Table of contents/);
  await expect(status(page)).toContainText('Table of contents inserted: 2 headings');
  const blocks = page.locator('.docx-page > :is(p, h1, h2)');
  await expect(blocks.nth(1)).toHaveText('Contents');
  await expect(blocks.nth(2)).toContainText('Introduction');
  await expect(blocks.nth(3)).toContainText('Method');
  const { zip, xml } = await download(page);
  expect(xml).toMatch(/<w:pStyle w:val="TOCHeading"\/><\/w:pPr><w:r><w:t xml:space="preserve">Contents<\/w:t><\/w:r>/);
  expect(xml).toContain(' TOC \\o "1-3" \\h \\z \\u ');
  expect(xml).toMatch(/<w:hyperlink w:anchor="_Toc\d+" w:history="1">.*?Introduction/);
  expect(xml).toMatch(/<w:bookmarkStart w:id="\d+" w:name="_Toc\d+"\/><w:r><w:t xml:space="preserve">Introduction<\/w:t>/);
  const styles = await zip.file('word/styles.xml').async('string');
  expect(styles).toContain('<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/>');
  expect(await zip.file('word/settings.xml').async('string')).toContain('<w:updateFields w:val="true"/>');
});

test('a table of contents goes in the text itself, and not while tracking', async ({ page }) => {
  await openInEditor(page, await buildDocx(ESSAY));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill('Rita');
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await caretAt(page, 'My Essay');
  await insertMenu(page, /Table of contents/);
  await expect(status(page)).toHaveText('A table of contents can’t be tracked here: turn off Track changes to insert it');
  await expect(page.locator('.docx-page > p', { hasText: 'Contents' })).toHaveCount(0);
});

test('Caption: "Figure 1: text" in Word\'s Caption style, numbered with a SEQ field; a cross-reference points to it', async ({ page }) => {
  await openInEditor(page, await buildDocx(ESSAY));
  await caretAt(page, 'Memory matters.');
  await insertMenu(page, /Caption/);
  const cap = page.getByRole('dialog', { name: 'Caption' });
  await expect(cap.getByLabel('Label')).toHaveValue('Figure');
  await cap.getByLabel('Caption text').fill('Results');
  await expect(cap.locator('.preview')).toHaveText('Figure 1: Results');
  await cap.getByRole('button', { name: 'Insert' }).click();
  await expect(page.locator('.docx-page > p', { hasText: 'Results' })).toHaveText('Figure 1: Results');
  // A second one above the first: Update numbers them in order.
  await caretAt(page, 'Introduction');
  await insertMenu(page, /Caption/);
  await cap.getByLabel('Caption text').fill('Setup');
  await cap.getByRole('button', { name: 'Insert' }).click();
  await expect(page.locator('.docx-page > p', { hasText: 'Setup' })).toHaveText('Figure 1: Setup');
  await expect(page.locator('.docx-page > p', { hasText: 'Results' })).toHaveText('Figure 2: Results');

  // A cross-reference to "Figure 2", as its label and number.
  await caretAt(page, 'We measured it.');
  await page.keyboard.type(' See ');
  await insertMenu(page, /Cross-reference/);
  const xref = page.getByRole('dialog', { name: 'Cross-reference' });
  await expect(xref.getByLabel('Reference type')).toHaveValue('Figure');
  await xref.getByRole('option', { name: 'Figure 2: Results' }).click();
  await xref.getByRole('button', { name: 'Insert' }).click();
  await expect(page.locator('.docx-page > p', { hasText: 'We measured' })).toHaveText('We measured it. See Figure 2');
  const { zip, xml } = await download(page);
  expect(xml).toMatch(/<w:pStyle w:val="Caption"\/><\/w:pPr><w:r><w:t xml:space="preserve">Figure <\/w:t><\/w:r><w:r><w:fldChar w:fldCharType="begin"\/><\/w:r><w:r><w:instrText xml:space="preserve"> SEQ Figure \\\* ARABIC <\/w:instrText>/);
  const name = /w:name="(_Ref\d+)"\/><w:r><w:t xml:space="preserve">Figure <\/w:t><\/w:r>.*?SEQ Figure.*?<w:t xml:space="preserve">2<\/w:t><\/w:r><w:r><w:fldChar w:fldCharType="end"\/><\/w:r><w:bookmarkEnd/.exec(xml)?.[1];
  expect(name).toBeTruthy();
  expect(xml).toContain(`<w:instrText xml:space="preserve"> REF ${name} \\h </w:instrText>`);
  expect(await zip.file('word/styles.xml').async('string')).toContain('w:styleId="Caption"');
});

test('a cross-reference to a heading shows its text, and Update (F9) follows the heading when it changes', async ({ page }) => {
  await openInEditor(page, await buildDocx(ESSAY));
  await caretAt(page, 'We measured it.');
  await page.keyboard.type(' As in ');
  await insertMenu(page, /Cross-reference/);
  const xref = page.getByRole('dialog', { name: 'Cross-reference' });
  await xref.getByLabel('Reference type').selectOption('Heading');
  await xref.getByRole('option', { name: 'Introduction' }).click();
  await xref.getByRole('button', { name: 'Insert' }).click();
  const line = page.locator('.docx-page > p', { hasText: 'We measured' });
  await expect(line).toHaveText('We measured it. As in Introduction');
  // Changed within its text (at its very end, a browser may put the caret after the bookmark, as
  // Word does after a bookmark's end, and the words typed there aren't the bookmark's).
  await page.evaluate(() => {
    const h = document.querySelector('.docx-page > h1');
    const text = [...h.childNodes].map((n) => (n.nodeType === 3 ? n : n.firstChild)).find((n) => n?.nodeType === 3 && n.data.includes('Introduction'));
    document.querySelector('.docx-page').focus();
    getSelection().collapse(text, text.data.indexOf('Introduction') + 5);
  });
  await page.keyboard.type('-');
  await expect(page.locator('.docx-page > h1').first()).toHaveText('Intro-duction');
  await page.keyboard.press('F9');
  await expect(line).toHaveText('We measured it. As in Intro-duction');
});

test('Equation: typed as LaTeX, previewed, and written as Word\'s Office Math', async ({ page }) => {
  await openInEditor(page, await buildDocx(ESSAY));
  await caretAt(page, 'We measured it.');
  await page.keyboard.type(' ');
  await insertMenu(page, /Equation/);
  const eq = page.getByRole('dialog', { name: 'Equation' });
  await eq.getByLabel('Equation, as in LaTeX').fill('\\frac{1}{');
  await expect(eq.locator('.error')).toHaveText('A { has no } to close it');
  await expect(eq.getByRole('button', { name: 'Insert' })).toBeDisabled();
  await eq.getByLabel('Equation, as in LaTeX').fill('E = mc^2 + \\frac{a}{b}');
  await expect(eq.locator('.maths math')).toHaveCount(1);
  await eq.getByRole('button', { name: 'Insert' }).click();
  await expect(page.locator('.docx-page > p', { hasText: 'We measured' }).locator('.equation math msup')).toHaveCount(1);
  const { xml } = await download(page);
  expect(xml).toMatch(/<m:oMath( xmlns:m="[^"]+")?><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"\/><\/w:rPr><m:t xml:space="preserve">E<\/m:t><\/m:r>/);
  expect(xml).toContain('<m:sSup><m:e>');
  expect(xml).toMatch(/<m:f><m:num>.*?<m:t xml:space="preserve">a<\/m:t>.*?<\/m:num><m:den>.*?b.*?<\/m:den><\/m:f>/);
});

test('an equation on its own line in an empty paragraph: an equation paragraph (m:oMathPara)', async ({ page }) => {
  await openInEditor(page, await buildDocx(ESSAY + P('')));
  await page.locator('.docx-page > p').last().click();
  await insertMenu(page, /Equation/);
  const eq = page.getByRole('dialog', { name: 'Equation' });
  await eq.getByLabel('Equation, as in LaTeX').fill('\\sum_{i=1}^{n} x_i');
  await eq.getByLabel(/On its own line/).check();
  await eq.getByRole('button', { name: 'Insert' }).click();
  await expect(page.locator('.docx-page .equation.display')).toHaveCount(1);
  const { xml } = await download(page);
  expect(xml).toMatch(/<m:oMathPara( xmlns:m="[^"]+")?><m:oMath><m:nary><m:naryPr><m:chr m:val="∑"\/>/);
});

test('Citation and bibliography: a new source kept as Word keeps sources, cited in APA, listed under References', async ({ page }) => {
  await openInEditor(page, await buildDocx(ESSAY + P('')));
  await caretAt(page, 'Memory matters.');
  await page.keyboard.press('ArrowLeft');
  await page.keyboard.type(' ');
  await insertMenu(page, /Citation/);
  const cite = page.getByRole('dialog', { name: 'Citation' });
  await cite.getByLabel('Authors').fill('Smith, Jane; Lee, Kim');
  await cite.getByLabel('Title', { exact: true }).fill('Remembering');
  await cite.getByLabel('Year').fill('2020');
  await cite.getByLabel('Publisher').fill('North Press');
  await cite.getByLabel('Page cited').fill('4');
  await expect(cite.locator('.preview')).toHaveText('(Smith & Lee, 2020, p. 4)');
  await cite.getByRole('button', { name: 'Insert' }).click();
  await expect(page.locator('.docx-page > p', { hasText: 'Memory' })).toHaveText('Memory matters (Smith & Lee, 2020, p. 4).');

  await page.locator('.docx-page > p').last().click();
  await insertMenu(page, /Bibliography/);
  await expect(page.locator('.docx-page > h1', { hasText: 'References' })).toHaveCount(1);
  const entry = page.locator('.docx-page > p', { hasText: 'Remembering' });
  await expect(entry).toHaveText('Smith, J., & Lee, K. (2020). Remembering. North Press.');
  await expect(entry.locator('[data-rstyle], i, em, span').filter({ hasText: /^Remembering\.$/ }).first()).toBeVisible();

  const { zip, xml } = await download(page);
  expect(xml).toContain('<w:instrText xml:space="preserve"> CITATION Smi20 \\p 4 \\l 1033 </w:instrText>');
  expect(xml).toContain('<w:instrText xml:space="preserve"> BIBLIOGRAPHY </w:instrText>');
  expect(xml).toMatch(/<w:pStyle w:val="Bibliography"\/><w:ind w:left="720" w:hanging="720"\/>/);
  const sources = await zip.file('customXml/item1.xml').async('string');
  expect(sources).toContain('StyleName="APA"');
  expect(sources).toContain('<b:Tag>Smi20</b:Tag><b:SourceType>Book</b:SourceType>');
  expect(sources).toContain('<b:NameList><b:Person><b:Last>Smith</b:Last><b:First>Jane</b:First></b:Person><b:Person><b:Last>Lee</b:Last><b:First>Kim</b:First></b:Person></b:NameList>');
  expect(sources).toContain('<b:Title>Remembering</b:Title><b:Year>2020</b:Year><b:Publisher>North Press</b:Publisher>');
  expect(await zip.file('customXml/itemProps1.xml').async('string')).toContain('<ds:schemaRef ds:uri="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"/>');
  expect(await zip.file('customXml/_rels/item1.xml.rels').async('string')).toContain('Target="itemProps1.xml"');
  expect(await zip.file('word/_rels/document.xml.rels').async('string')).toMatch(/Type="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships\/customXml" Target="..\/customXml\/item1.xml"/);
  expect(await zip.file('[Content_Types].xml').async('string')).toContain('<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>');
});

test('MLA: citing again in MLA restyles the citations and the bibliography, which lists a source added later on F9', async ({ page }) => {
  await openInEditor(page, await buildDocx(ESSAY + P('')));
  const cite = page.getByRole('dialog', { name: 'Citation' });
  await caretAt(page, 'Memory matters.');
  await insertMenu(page, /Citation/);
  await cite.getByLabel('Authors').fill('Smith, Jane');
  await cite.getByLabel('Title', { exact: true }).fill('Remembering');
  await cite.getByLabel('Year').fill('2020');
  await cite.getByLabel('Publisher').fill('North Press');
  await cite.getByRole('button', { name: 'Insert' }).click();
  await page.locator('.docx-page > p').last().click();
  await insertMenu(page, /Bibliography/);
  // A website, cited in MLA: the first citation and the list follow.
  await caretAt(page, 'We measured it.');
  await insertMenu(page, /Citation/);
  await cite.getByLabel('Citation style').selectOption('MLA');
  await cite.getByLabel(/Smith \(2020\)/).uncheck(); // suggested: the last source cited
  await cite.getByLabel('A new source…').check();
  await cite.getByLabel('Type of source').selectOption('InternetSite');
  await cite.getByLabel('Authors').fill('Park, Sun');
  await cite.getByLabel('Title', { exact: true }).fill('How We Recall');
  await cite.getByLabel('Year').fill('2023');
  await cite.getByLabel('Website name').fill('Mind Notes');
  await cite.getByLabel('Web address').fill('https://example.org/recall');
  await cite.getByRole('button', { name: 'Insert' }).click();
  await expect(page.locator('.docx-page > p', { hasText: 'Memory' })).toHaveText('Memory matters.(Smith)');
  await expect(page.locator('.docx-page > p', { hasText: 'We measured' })).toHaveText('We measured it.(Park)');
  await expect(page.locator('.docx-page > h1', { hasText: 'Works Cited' })).toHaveCount(0); // the heading stays as inserted
  await expect(page.locator('.docx-page > p', { hasText: 'Remembering' })).toHaveText('Smith, Jane. Remembering. North Press, 2020.');
  await expect(page.locator('.docx-page > p', { hasText: 'Recall' })).toHaveText('Park, Sun. “How We Recall.” Mind Notes, 2023, https://example.org/recall.');
  const { zip } = await download(page);
  expect(await zip.file('customXml/item1.xml').async('string')).toContain('StyleName="MLA"');
});

test('a document\'s own sources (Word\'s) are offered, and its part is kept byte for byte until one is added', async ({ page }) => {
  const SOURCES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<b:Sources SelectedStyle="\\APASixthEditionOfficeOnline.xsl" StyleName="APA" Version="6" xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">'
    + '<b:Source><b:Tag>Doe19</b:Tag><b:SourceType>JournalArticle</b:SourceType><b:Author><b:Author><b:NameList><b:Person><b:Last>Doe</b:Last><b:First>Ann</b:First></b:Person></b:NameList></b:Author></b:Author>'
    + '<b:Title>On Memory</b:Title><b:JournalName>Mind</b:JournalName><b:Year>2019</b:Year><b:Volume>12</b:Volume><b:Issue>3</b:Issue><b:Pages>45-67</b:Pages></b:Source></b:Sources>';
  await openInEditor(page, await buildDocx(ESSAY, { parts: { 'customXml/item1.xml': SOURCES } }));
  // Saved without citing: the sources part is as it was.
  await caretAt(page, 'We measured it.');
  await page.keyboard.type(' Again.');
  let { zip } = await download(page);
  expect(await zip.file('customXml/item1.xml').async('string')).toBe(SOURCES);
  await insertMenu(page, /Citation/);
  const cite = page.getByRole('dialog', { name: 'Citation' });
  await expect(cite.getByLabel(/Doe \(2019\) – On Memory/)).toBeChecked();
  await cite.getByRole('button', { name: 'Insert' }).click();
  await expect(page.locator('.docx-page > p', { hasText: 'We measured' })).toHaveText('We measured it. Again.(Doe, 2019)');
  await page.locator('.docx-page > p').last().click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await insertMenu(page, /Bibliography/);
  await expect(page.locator('.docx-page > p', { hasText: 'On Memory' })).toHaveText('Doe, A. (2019). On Memory. Mind, 12(3), 45–67.');
  ({ zip } = await download(page));
  expect(await zip.file('customXml/item1.xml').async('string')).toBe(SOURCES); // citing an existing source changes nothing in it
});

test('the Navigation pane lists the headings by level and goes to one; it stays open in this browser', async ({ page }) => {
  await openInEditor(page, await buildDocx(ESSAY));
  const pane = page.getByRole('navigation', { name: 'Headings' });
  await expect(pane).toHaveCount(0);
  await page.getByRole('button', { name: 'Headings', exact: true }).click();
  await expect(pane.getByRole('listitem')).toHaveText(['Introduction', 'Method']);
  await pane.getByRole('button', { name: 'Method' }).click();
  await page.keyboard.type('New ');
  await expect(page.locator('.docx-page > h1', { hasText: 'Method' })).toHaveText('New Method');
  await expect(pane.getByRole('button', { name: 'New Method' })).toHaveAttribute('aria-current', 'location');
  // A new heading appears as it is made.
  await caretAt(page, 'We measured it.');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ControlOrMeta+Alt+2');
  await page.keyboard.type('Results');
  await expect(pane.getByRole('listitem')).toHaveText(['Introduction', 'New Method', 'Results']);
  await expect(pane.locator('li.level2')).toHaveText('Results');
  // Remembered: the next document opens with it.
  await page.goto('./');
  await page.setInputFiles('input[type=file]', { name: 'again.docx', mimeType: 'application/octet-stream', buffer: await buildDocx(ESSAY) });
  await expect(page.getByRole('button', { name: 'Headings', exact: true })).toHaveAttribute('aria-pressed', 'true');
  await expect(pane.getByRole('listitem')).toHaveText(['Introduction', 'Method']);
});

test('a document started here doesn\'t ask for its edits to be tracked', async ({ page }) => {
  await page.goto('./');
  await page.getByRole('button', { name: /Blank document/ }).click();
  await page.locator('.docx-page').click();
  await page.keyboard.type('A new essay.');
  await page.keyboard.press('ControlOrMeta+s');
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review).toBeVisible();
  await expect(review).not.toContainText('aren’t tracked');
  await expect(review.locator('section.screening')).toHaveCount(0);
});
