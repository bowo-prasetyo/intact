// Moving a heading with its text from the Headings pane, as Word's Navigation pane: a chapter (a
// Heading 1 and all under it, its scenes included) goes up or down past another chapter, or where it
// is dragged; the paragraphs are moved as they are, so the file keeps them exactly, and Undo puts
// them back. Refused, with the reason, while tracking or where a range would be cut at its edge.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const status = (page) => page.locator('.statusbar [role=status]');
const pane = (page) => page.getByRole('navigation', { name: 'Headings' });
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W_NS}>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>
</w:styles>`;
const h = (level, text) => `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const para = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const NOVEL = para('Title page.')
  + h(1, 'Chapter One') + para('One begins.') + h(2, 'Scene A') + para('Scene A text.') + h(2, 'Scene B') + para('Scene B text.')
  + h(1, 'Chapter Two') + para('Two begins.')
  + h(1, 'Chapter Three') + para('Three begins.');

async function open(page, body = NOVEL, opts = {}) {
  await openInEditor(page, await buildDocx(body, { ...opts, parts: { 'word/styles.xml': STYLES, ...(opts.parts || {}) } }));
  await page.getByRole('button', { name: /^Headings/ }).click();
  await expect(pane(page).locator('li')).toHaveCount(5);
}

const order = (page) => page.locator('.docx-page > :is(p, h1, h2, h3)').allInnerTexts();
const item = (page, name) => pane(page).locator('li', { has: page.getByRole('button', { name, exact: true }) });

test('▼ moves a chapter below the next one, with its scenes; Ctrl+Z puts it back; the file has the paragraphs as they were', async ({ page }) => {
  await open(page);
  await item(page, 'Chapter One').hover();
  await item(page, 'Chapter One').getByRole('button', { name: 'Move down' }).click();
  await expect(status(page)).toContainText('Moved “Chapter One” below “Chapter Two”, with its text – Undo (Ctrl+Z) puts it back');
  expect(await order(page)).toEqual(['Title page.', 'Chapter Two', 'Two begins.', 'Chapter One', 'One begins.', 'Scene A', 'Scene A text.', 'Scene B', 'Scene B text.', 'Chapter Three', 'Three begins.']);
  await expect(pane(page).locator('li .go')).toHaveText(['Chapter Two', 'Chapter One', 'Scene A', 'Scene B', 'Chapter Three']);

  const { xml } = await download(page);
  const at = (t) => xml.indexOf(`>${t}</w:t>`);
  expect(at('Chapter Two')).toBeLessThan(at('Chapter One'));
  expect(at('Scene B text.')).toBeLessThan(at('Chapter Three'));
  expect(xml).toContain('<w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">Chapter One</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">One begins.</w:t>'); // moved as it was

  await page.locator('.docx-page').focus();
  await page.keyboard.press('Control+z');
  await expect.poll(() => order(page)).toEqual(['Title page.', 'Chapter One', 'One begins.', 'Scene A', 'Scene A text.', 'Scene B', 'Scene B text.', 'Chapter Two', 'Two begins.', 'Chapter Three', 'Three begins.']);
});

test('a chapter dragged onto another goes before it', async ({ page }) => {
  await open(page);
  await pane(page).getByRole('button', { name: 'Chapter Three', exact: true }).dragTo(pane(page).getByRole('button', { name: 'Chapter One', exact: true }));
  await expect(status(page)).toContainText('Moved “Chapter Three” above “Chapter One”');
  await expect(pane(page).locator('li .go')).toHaveText(['Chapter Three', 'Chapter One', 'Scene A', 'Scene B', 'Chapter Two']);
});

test('Alt+Shift+↑ in the pane moves the heading up, and the focus stays on it to move it again', async ({ page }) => {
  await open(page);
  await pane(page).getByRole('button', { name: 'Chapter Three', exact: true }).focus();
  await page.keyboard.press('Alt+Shift+ArrowUp');
  await expect(pane(page).locator('li .go')).toHaveText(['Chapter One', 'Scene A', 'Scene B', 'Chapter Three', 'Chapter Two']);
  await expect(pane(page).getByRole('button', { name: 'Chapter Three', exact: true })).toBeFocused();
  await page.keyboard.press('Alt+Shift+ArrowUp');
  await expect(pane(page).locator('li .go')).toHaveText(['Chapter Three', 'Chapter One', 'Scene A', 'Scene B', 'Chapter Two']);
  await page.keyboard.press('Alt+Shift+ArrowUp');
  await expect(status(page)).toContainText('“Chapter Three” is already the first');
});

test('a scene moves among its chapter’s scenes, not out of its chapter', async ({ page }) => {
  await open(page);
  await item(page, 'Scene B').hover();
  await item(page, 'Scene B').getByRole('button', { name: 'Move up' }).click();
  expect(await order(page)).toEqual(['Title page.', 'Chapter One', 'One begins.', 'Scene B', 'Scene B text.', 'Scene A', 'Scene A text.', 'Chapter Two', 'Two begins.', 'Chapter Three', 'Three begins.']);
  await item(page, 'Scene A').hover();
  await item(page, 'Scene A').getByRole('button', { name: 'Move down' }).click();
  await expect(status(page)).toContainText('“Scene A” is already the last under its heading');
});

test('while tracking, a chapter isn’t moved, and it says why', async ({ page }) => {
  await open(page);
  await page.getByRole('button', { name: 'Track changes' }).click();
  const named = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await named.getByRole('textbox').fill('Rita');
  await named.getByRole('button', { name: 'Start tracking' }).click();
  await item(page, 'Chapter One').hover();
  await item(page, 'Chapter One').getByRole('button', { name: 'Move down' }).click();
  await expect(status(page)).toContainText('“Chapter One” wasn’t moved: moving text isn’t tracked here – turn off Track Changes to move it, or cut and paste it');
  await expect(pane(page).locator('li .go').first()).toHaveText('Chapter One');
});

test('a comment from one chapter into the next keeps them where they are', async ({ page }) => {
  const comments = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W_NS}><w:comment w:id="7" w:author="Ed" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Tighten.</w:t></w:r></w:p></w:comment></w:comments>`;
  const body = para('Title page.')
    + h(1, 'Chapter One') + '<w:p><w:commentRangeStart w:id="7"/><w:r><w:t>One begins.</w:t></w:r></w:p>'
    + h(1, 'Chapter Two') + '<w:p><w:r><w:t>Two begins.</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>'
    + h(1, 'Chapter Three') + para('Three begins.');
  await openInEditor(page, await buildDocx(body, { parts: { 'word/styles.xml': STYLES, 'word/comments.xml': comments }, rels: `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>` }));
  await page.getByRole('button', { name: /^Headings/ }).click();
  await item(page, 'Chapter One').hover();
  await item(page, 'Chapter One').getByRole('button', { name: 'Move down' }).click();
  await expect(status(page)).toContainText('“Chapter One” can’t be moved here: a comment starts on one side of its edge and ends on the other');
  await expect(pane(page).locator('li .go')).toHaveText(['Chapter One', 'Chapter Two', 'Chapter Three']);
  // Chapter Three holds no part of it: it moves (into the commented stretch, which still starts before it and ends after).
  await item(page, 'Chapter Three').hover();
  await item(page, 'Chapter Three').getByRole('button', { name: 'Move up' }).click();
  await expect(status(page)).toContainText('Moved “Chapter Three” above “Chapter Two”');
  const { xml } = await download(page);
  expect(xml.indexOf('<w:commentRangeStart w:id="7"/>')).toBeLessThan(xml.indexOf('>Chapter Three</w:t>'));
  expect(xml.indexOf('>Chapter Three</w:t>')).toBeLessThan(xml.indexOf('<w:commentRangeEnd w:id="7"/>'));
});

test('a chapter holding a section break (another page setup) stays where it is', async ({ page }) => {
  const body = h(1, 'Chapter One') + '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:pPr><w:r><w:t>Landscape next.</w:t></w:r></w:p>'
    + h(1, 'Chapter Two') + para('Two begins.');
  await openInEditor(page, await buildDocx(body, { parts: { 'word/styles.xml': STYLES } }));
  await page.getByRole('button', { name: /^Headings/ }).click();
  await item(page, 'Chapter One').hover();
  await item(page, 'Chapter One').getByRole('button', { name: 'Move down' }).click();
  await expect(status(page)).toContainText('“Chapter One” can’t be moved here: it holds a section break (another page setup): move the text in Word');
});
