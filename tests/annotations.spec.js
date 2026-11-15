// Comments, headers, footers and notes are shown read-only, in plain language, and never change.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const part = (root, inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${root} ${W_NS} xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">${inner}</w:${root}>`;

const HEADERS = {
  parts: {
    'word/header1.xml': part('hdr', '<w:p><w:r><w:t>Default header</w:t></w:r></w:p>'),
    'word/header2.xml': part('hdr', '<w:p><w:r><w:t xml:space="preserve">First page header </w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="95250" cy="95250"/><wp:docPr id="9" name="Logo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdLogo"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'),
    'word/_rels/header2.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLogo" Type="${REL}/image" Target="media/image1.png"/></Relationships>`,
    'word/footer1.xml': part('ftr', '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>'),
  },
  rels: `<Relationship Id="rIdH1" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdH2" Type="${REL}/header" Target="header2.xml"/><Relationship Id="rIdF1" Type="${REL}/footer" Target="footer1.xml"/>`,
  sectPr: '<w:headerReference w:type="default" r:id="rIdH1"/><w:headerReference w:type="first" r:id="rIdH2"/><w:footerReference w:type="default" r:id="rIdF1"/>',
};

const COMMENTS = {
  parts: {
    'word/comments.xml': part('comments', `
      <w:comment w:id="0" w:author="Ann Lee" w:initials="AL" w:date="2026-01-05T10:00:00Z"><w:p w14:paraId="00000001"><w:r><w:t>Check the amount.</w:t></w:r></w:p></w:comment>
      <w:comment w:id="1" w:author="Bob" w:date="2026-01-06T10:00:00Z"><w:p w14:paraId="00000002"><w:r><w:t>Confirmed with finance.</w:t></w:r></w:p></w:comment>
      <w:comment w:id="2" w:author="Ann Lee" w:date="2026-01-07T10:00:00Z"><w:p w14:paraId="00000003"><w:r><w:t>Typo here?</w:t></w:r></w:p></w:comment>`),
    'word/commentsExtended.xml': part('commentsEx', '').replace('<w:commentsEx', '<w15:commentsEx').replace('</w:commentsEx>', '</w15:commentsEx>')
      .replace('></w15:commentsEx>', '><w15:commentEx w15:paraId="00000001" w15:done="0"/><w15:commentEx w15:paraId="00000002" w15:paraIdParent="00000001" w15:done="0"/><w15:commentEx w15:paraId="00000003" w15:done="1"/></w15:commentsEx>'),
    'word/footnotes.xml': part('footnotes', '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:t>Source: annual report.</w:t></w:r></w:p></w:footnote>'),
  },
  rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rIdCx" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/><Relationship Id="rIdFn" Type="${REL}/footnotes" Target="footnotes.xml"/>`,
};
const COMMENTED = `<w:p><w:r><w:t xml:space="preserve">The fee is </w:t></w:r><w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/><w:r><w:t>EUR 1,200</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> per month</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:commentRangeStart w:id="2"/><w:r><w:t>Recieved</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r><w:r><w:t xml:space="preserve"> on time. Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>4</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r></w:p>`;

const highlighted = (page, name) => page.evaluate((n) => [...(CSS.highlights.get(n) || [])].map((r) => r.toString()), name);

test('the first page\'s header and footer are shown, read-only and outside the editable text', async ({ page }) => {
  const docx = await buildDocx('<w:p><w:r><w:t>Body text</w:t></w:r></w:p>', { ...HEADERS, sectPr: `${HEADERS.sectPr}<w:titlePg/>` });
  await openInEditor(page, docx);
  const header = page.locator('.docx-hf.header .hf-variant:not([hidden])');
  await expect(page.getByRole('tablist', { name: 'Headers' }).getByRole('tab')).toHaveText(['First page', 'Other pages']);
  await expect(header).toHaveText('First page header ');
  await expect(page.locator('.docx-hf.header')).toHaveAttribute('title', 'First page header – shown for reference; edit it in Word');
  await expect(header.locator('img')).toHaveCount(1); // the header's own relationships resolve its image
  await expect(page.locator('.docx-hf.footer')).toContainText('Page 1');
  await expect(page.locator('.docx-page')).toHaveText('Body text');
  await expect(page.locator('.docx-hf [contenteditable]')).toHaveCount(0);

  const { zip, xml } = await download(page);
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
  expect(await zip.file('word/header2.xml').async('string')).toBe(HEADERS.parts['word/header2.xml']);
});

test('without a different first page, the default header is shown', async ({ page }) => {
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>Body</w:t></w:r></w:p>', HEADERS));
  await expect(page.locator('.docx-hf.header .hf-variant')).toHaveText('Default header');
  await expect(page.getByRole('tablist', { name: 'Headers' })).toHaveCount(0);
  await expect(page.locator('.docx-hf.header')).toHaveAttribute('title', /^Header – /);
});

test('comments are listed with replies and resolved state, and point at their text', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  const docx = await buildDocx(COMMENTED, COMMENTS);
  await openInEditor(page, docx);
  const panel = page.getByRole('complementary', { name: 'Comments' });
  await expect(panel).toBeVisible(); // open by default on a wide screen
  const threads = panel.locator('.thread');
  await expect(threads).toHaveCount(2);
  await expect(threads.nth(0)).toContainText('Ann Lee');
  await expect(threads.nth(0)).toContainText('Check the amount.');
  await expect(threads.nth(0).locator('.reply')).toContainText('Confirmed with finance.');
  await expect(threads.nth(1)).toContainText('Resolved');

  // Only top-level comments get a marker in the text, numbered in order.
  await expect(page.locator('.docx-page .comment-ref')).toHaveCount(2);
  expect(await page.locator('.docx-page .comment-ref').evaluateAll((els) => els.map((e) => e.dataset.n))).toEqual(['1', '2']);
  expect(await highlighted(page, 'comment')).toEqual(['EUR 1,200', 'Recieved']);

  await threads.nth(1).click();
  expect(await highlighted(page, 'comment-active')).toEqual(['Recieved']);
  await page.locator('.docx-page .comment-ref').first().click();
  await expect(threads.nth(0)).toHaveClass(/active/);
  expect(await highlighted(page, 'comment-active')).toEqual(['EUR 1,200']);

  await page.getByTitle('Comments (2)').click();
  await expect(panel).toHaveCount(0);

  const { xml } = await download(page);
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
});

test('editing commented text keeps its comment anchored', async ({ page }) => {
  await openInEditor(page, await buildDocx(COMMENTED, COMMENTS));
  await page.locator('.docx-page p').nth(1).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.type('ei');
  expect(await highlighted(page, 'comment')).toEqual(['EUR 1,200', 'Received']);
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:commentRangeStart w:id="2"\/><w:r><w:t xml:space="preserve">Received<\/w:t><\/w:r><w:commentRangeEnd w:id="2"\/>/);
});

test('preserved content is described in plain language', async ({ page }) => {
  await openInEditor(page, await buildDocx(COMMENTED, COMMENTS));
  const note = page.locator('.docx-page .note-ref');
  await expect(note).toHaveAttribute('data-n', '1');
  await expect(note).toHaveAttribute('title', 'Footnote 1: Source: annual report. – edit it in Word');
  const field = page.locator('.docx-page .field');
  await expect(field).toHaveText('4');
  await expect(field).toHaveAttribute('title', 'Page number – Word updates it');
  // Field codes and markers are not shown as clutter.
  await expect(page.locator('.docx-page p').nth(1)).toHaveText('Recieved on time. Page 4ü');
  await expect(page.locator('.docx-page .symbol')).toHaveAttribute('title', 'Symbol (Wingdings)');
  await expect(page.locator('.docx-page [title*="Preserved content"]')).toHaveCount(0);
});

test('documents with several headers offer them as tabs: sections, first page, even pages', async ({ page }) => {
  const h = (n) => part('hdr', `<w:p><w:r><w:t>Header ${n}</w:t></w:r></w:p>`);
  const docx = await buildDocx(
    `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdA"/><w:headerReference w:type="even" r:id="rIdB"/></w:sectPr></w:pPr><w:r><w:t>Section one</w:t></w:r></w:p><w:p><w:r><w:t>Section two</w:t></w:r></w:p>`,
    {
      parts: {
        'word/header1.xml': h('A'), 'word/header2.xml': h('B'), 'word/header3.xml': h('C'), 'word/header4.xml': h('D'),
        'word/settings.xml': part('settings', '<w:evenAndOddHeaders/>'),
      },
      rels: `<Relationship Id="rIdA" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdB" Type="${REL}/header" Target="header2.xml"/><Relationship Id="rIdC" Type="${REL}/header" Target="header3.xml"/><Relationship Id="rIdD" Type="${REL}/header" Target="header4.xml"/><Relationship Id="rIdS" Type="${REL}/settings" Target="settings.xml"/>`,
      // Section 2 has its own first-page and default headers, and inherits section 1's even one.
      sectPr: '<w:headerReference w:type="default" r:id="rIdC"/><w:headerReference w:type="first" r:id="rIdD"/><w:titlePg/>',
    },
  );
  await openInEditor(page, docx);
  const tabs = page.getByRole('tablist', { name: 'Headers' }).getByRole('tab');
  await expect(tabs).toHaveText(['Section 1: odd pages', 'Section 1: even pages', 'Section 2: first page', 'Section 2: odd pages']);
  const shown = page.locator('.docx-hf.header .hf-variant:not([hidden])');
  await expect(shown).toHaveText('Header A');
  await tabs.nth(2).click();
  await expect(shown).toHaveText('Header D');
  await expect(page.locator('.docx-hf.header')).toHaveAttribute('title', /^Section 2: first page header/);
});

test('Find also searches headers, footers and comments, read-only', async ({ page }) => {
  await page.setViewportSize({ width: 900, height: 700 }); // comments start closed
  const docx = await buildDocx(COMMENTED.replace('The fee is ', 'The fee per Page is '), {
    parts: { ...HEADERS.parts, ...COMMENTS.parts },
    rels: HEADERS.rels + COMMENTS.rels,
    sectPr: `${HEADERS.sectPr}<w:titlePg/>`,
  });
  await openInEditor(page, docx);
  await expect(page.getByRole('complementary', { name: 'Comments' })).toHaveCount(0);
  await page.locator('.docx-page p').first().click();
  await page.keyboard.press('ControlOrMeta+f');
  await page.keyboard.type('page');
  // Header (first page, "First page header"), the body twice, the footer ("Page 1"), no comment.
  const status = page.locator('.findbar .status');
  await expect(status).toHaveText(/ of 4 · 2 read-only$/);

  await page.getByRole('textbox', { name: 'Find' }).fill('Default');
  await expect(status).toHaveText('1 of 1 · 1 read-only');
  // The match is in the "Other pages" header, so that tab is shown.
  await expect(page.locator('.docx-hf.header .hf-variant:not([hidden])')).toHaveText('Default header');

  await page.getByRole('textbox', { name: 'Find' }).fill('finance');
  await expect(status).toHaveText('1 of 1 · 1 read-only');
  await expect(page.getByRole('complementary', { name: 'Comments' })).toBeVisible(); // opened to show it
});
