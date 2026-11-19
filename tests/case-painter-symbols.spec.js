// Three of Word's everyday tools: Change Case (the Aa menu and Shift+F3), the Format Painter (and
// Ctrl+Shift+C / Ctrl+Shift+V) and Insert Symbol. Each changes only the text it is used on, keeps
// the runs' other formatting, is one undoable step and is tracked while tracking.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const BODY = '<w:p><w:r><w:t xml:space="preserve">the annual </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>general</w:t></w:r><w:r><w:t xml:space="preserve"> meeting. it starts at nine.</w:t></w:r></w:p>'
  + '<w:p><w:r><w:rPr><w:i/><w:color w:val="C00000"/><w:sz w:val="28"/></w:rPr><w:t>Brand red</w:t></w:r><w:r><w:t xml:space="preserve"> and plain words here</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t>Total</w:t></w:r></w:p>';

const status = (page) => page.locator('.statusbar [role=status]');
const para = (page, n) => page.locator('.docx-page > p').nth(n);
const paragraphXml = (xml, n) => [...xml.matchAll(/<w:p>(.*?)<\/w:p>/g)][n][1];

/** Selects from `from` to the end of `to` in paragraph `n` (`to` omitted: just `from`). */
async function select(page, n, from, to = from) {
  await page.evaluate(([i, a, b]) => {
    const p = document.querySelectorAll('.docx-page > p')[i];
    const points = [];
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let t = walker.nextNode(); t; t = walker.nextNode()) points.push(t);
    const find = (w, end) => {
      for (const t of points) {
        const at = t.data.indexOf(w);
        if (at >= 0) return [t, end ? at + w.length : at];
      }
      return null;
    };
    document.querySelector('.docx-page').focus();
    getSelection().setBaseAndExtent(...find(a, false), ...find(b, true));
  }, [n, from, to]);
}

/** Puts the caret `offset` characters into `word` in paragraph `n`. */
async function caretIn(page, n, word, offset = 1) {
  await page.evaluate(([i, w, o]) => {
    const p = document.querySelectorAll('.docx-page > p')[i];
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      const at = t.data.indexOf(w);
      if (at < 0) continue;
      document.querySelector('.docx-page').focus();
      getSelection().collapse(t, at + o);
      return;
    }
  }, [n, word, offset]);
}

async function startTracking(page, name = 'Rita') {
  await page.getByRole('button', { name: 'Track changes' }).click();
  const dialog = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await dialog.getByRole('textbox').fill(name);
  await dialog.getByRole('button', { name: 'Start tracking' }).click();
  await expect(page.getByRole('button', { name: 'Track changes' })).toHaveAttribute('aria-pressed', 'true');
}

// ------------------------------------------------------------------ change case

test('Change case from the Aa menu keeps each run’s formatting', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await select(page, 0, 'the', 'meeting');
  await page.getByRole('button', { name: 'Change case' }).click();
  const menu = page.getByRole('menu', { name: 'Change case' });
  await expect(menu.getByRole('menuitem')).toHaveText(['Sentence case.', 'lowercase', 'UPPERCASE', 'Capitalize Each Word', 'tOGGLE cASE']);
  await menu.getByRole('menuitem', { name: 'UPPERCASE' }).click();
  await expect(para(page, 0)).toHaveText('THE ANNUAL GENERAL MEETING. it starts at nine.');
  await expect(para(page, 0).locator('strong')).toHaveText('GENERAL');
  // Still selected: Capitalize Each Word next.
  await page.getByRole('button', { name: 'Change case' }).click();
  await page.getByRole('menuitem', { name: 'Capitalize Each Word' }).click();
  await expect(para(page, 0)).toHaveText('The Annual General Meeting. it starts at nine.');

  const { xml } = await download(page);
  expect(paragraphXml(xml, 0)).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">General</w:t></w:r>');
  expect(paragraphXml(xml, 0)).toContain('The Annual ');
});

test('Sentence case capitalises each sentence; Shift+F3 cycles UPPERCASE, lowercase, Capitalize Each Word', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await select(page, 0, 'the', 'nine.');
  await page.getByRole('button', { name: 'Change case' }).click();
  await page.getByRole('menuitem', { name: 'Sentence case.' }).click();
  await expect(para(page, 0)).toHaveText('The annual general meeting. It starts at nine.');

  await select(page, 2, 'Total');
  await page.keyboard.press('Shift+F3');
  await expect(para(page, 2)).toHaveText('TOTAL');
  await page.keyboard.press('Shift+F3');
  await expect(para(page, 2)).toHaveText('total');
  await page.keyboard.press('Shift+F3');
  await expect(para(page, 2)).toHaveText('Total');
  // Each is one step for Undo.
  await page.keyboard.press('Control+z');
  await expect(para(page, 2)).toHaveText('total');
});

test('Shift+F3 with the caret in a word changes that word: lowercase, then Capitalize Each Word, then UPPERCASE', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await caretIn(page, 1, 'plain', 2);
  await page.keyboard.press('Shift+F3');
  await expect(para(page, 1)).toHaveText('Brand red and Plain words here');
  await page.keyboard.press('Shift+F3');
  await expect(para(page, 1)).toHaveText('Brand red and PLAIN words here');
});

test('Change case works in a table cell', async ({ page }) => {
  const table = '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>unit price</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
  await openInEditor(page, await buildDocx(table));
  await page.evaluate(() => {
    const t = document.querySelector('.docx-page td p').firstChild;
    document.querySelector('.docx-page').focus();
    getSelection().setBaseAndExtent(t, 0, t, t.length);
  });
  await page.keyboard.press('Shift+F3');
  await expect(page.locator('.docx-page td')).toHaveText('Unit Price');
});

test('while tracking, a change of case is a tracked deletion and insertion', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await startTracking(page);
  await select(page, 2, 'Total');
  await page.keyboard.press('Shift+F3');
  await expect(para(page, 2).locator('[data-tc="del"]')).toHaveText('Total');
  await expect(para(page, 2).locator('[data-tc="ins"]')).toHaveText('TOTAL');
  await expect(status(page)).toHaveText('UPPERCASE (tracked)');
  const { xml } = await download(page);
  expect(paragraphXml(xml, 2)).toMatch(/<w:del [^>]*w:author="Rita"[^>]*>.*<w:delText[^>]*>Total<\/w:delText>/);
  expect(paragraphXml(xml, 2)).toMatch(/<w:ins [^>]*w:author="Rita"[^>]*>.*<w:t[^>]*>TOTAL<\/w:t>/);
});

// ------------------------------------------------------------------ format painter

test('the Format Painter copies the caret’s formatting to the text selected next, then turns off', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await caretIn(page, 1, 'Brand', 2);
  const brush = page.getByRole('button', { name: 'Format Painter' });
  await brush.click();
  await expect(brush).toHaveAttribute('aria-pressed', 'true');
  await expect(page.locator('.docx-page')).toHaveClass(/painting/);
  // Select "plain words" with the mouse.
  const words = await page.evaluate(() => {
    const p = document.querySelectorAll('.docx-page > p')[1];
    const t = [...p.childNodes].map((n) => (n.nodeType === 3 ? n : n.firstChild)).find((n) => n?.data?.includes('plain'));
    const r = document.createRange();
    r.setStart(t, t.data.indexOf('plain'));
    r.setEnd(t, t.data.indexOf('plain') + 'plain words'.length);
    const rects = r.getClientRects();
    const first = rects[0];
    const last = rects[rects.length - 1];
    return { x0: first.left + 1, y: first.top + first.height / 2, x1: last.right - 1 };
  });
  await page.mouse.move(words.x0, words.y);
  await page.mouse.down();
  await page.mouse.move(words.x1, words.y, { steps: 5 });
  await page.mouse.up();
  await expect(brush).toHaveAttribute('aria-pressed', 'false');
  await expect(status(page)).toHaveText('Formatting painted');

  const { xml } = await download(page);
  const p1 = paragraphXml(xml, 1);
  expect(p1).toMatch(/<w:r><w:rPr><w:i\/><w:color w:val="C00000"\/><w:sz w:val="28"\/><\/w:rPr><w:t[^>]*>plain words<\/w:t><\/w:r>/);
  expect(p1).toContain('<w:t xml:space="preserve"> and </w:t>');
  expect(p1).toContain('<w:t xml:space="preserve"> here</w:t>');
});

test('Ctrl+Shift+C and Ctrl+Shift+V copy and paint formatting; painting plain text takes formatting off', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await caretIn(page, 1, 'plain', 1);
  await page.keyboard.press('Control+Shift+C');
  await expect(status(page)).toContainText('Formatting copied');
  await select(page, 1, 'Brand');
  await page.keyboard.press('Control+Shift+V');
  await expect(status(page)).toHaveText('Formatting painted');
  const { xml } = await download(page);
  expect(paragraphXml(xml, 1)).toMatch(/<w:r><w:t xml:space="preserve">Brand<\/w:t><\/w:r>/);
  expect(paragraphXml(xml, 1)).toContain('<w:t xml:space="preserve"> red</w:t>');
});

test('a double-click keeps the Format Painter on for several places, until Esc', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await caretIn(page, 0, 'general', 2);
  const brush = page.getByRole('button', { name: 'Format Painter' });
  await brush.dblclick();
  await expect(brush).toHaveAttribute('aria-pressed', 'true');
  // A click in a word paints the word.
  for (const word of ['annual', 'starts']) {
    const at = await page.evaluate((w) => {
      const p = document.querySelectorAll('.docx-page > p')[0];
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      for (let t = walker.nextNode(); t; t = walker.nextNode()) {
        const i = t.data.indexOf(w);
        if (i < 0) continue;
        const r = document.createRange();
        r.setStart(t, i + 2);
        r.setEnd(t, i + 3);
        const b = r.getBoundingClientRect();
        return { x: b.left + b.width / 2, y: b.top + b.height / 2 };
      }
      return null;
    }, word);
    await page.mouse.click(at.x, at.y);
    await expect(para(page, 0).locator('strong', { hasText: word })).toHaveCount(1);
    await expect(brush).toHaveAttribute('aria-pressed', 'true');
  }
  await page.keyboard.press('Escape');
  await expect(brush).toHaveAttribute('aria-pressed', 'false');
});

// ------------------------------------------------------------------ insert symbol

test('Insert symbol puts the symbol at the caret, by group, by name or by code, and remembers it', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await caretIn(page, 2, 'Total', 5);
  await page.getByRole('button', { name: 'Insert symbol' }).click();
  const dialog = page.getByRole('dialog', { name: 'Symbol' });
  await dialog.getByRole('tab', { name: 'Currency' }).click();
  await dialog.getByRole('button', { name: 'Euro (U+20AC)' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(para(page, 2)).toHaveText('Total€');
  await expect(status(page)).toHaveText('Symbol inserted: €');

  // By name, then Enter inserts the first found; by code.
  await page.getByRole('button', { name: 'Insert symbol' }).click();
  await expect(dialog.getByRole('group', { name: 'Recently used symbols' }).getByRole('button')).toHaveText(['€']);
  await dialog.getByRole('searchbox').fill('check mark');
  await dialog.getByRole('searchbox').press('Enter');
  await expect(para(page, 2)).toHaveText('Total€✓');
  await page.getByRole('button', { name: 'Insert symbol' }).click();
  await dialog.getByRole('searchbox').fill('U+00BD');
  await dialog.getByRole('button', { name: 'One half (U+00BD)' }).click();
  await expect(para(page, 2)).toHaveText('Total€✓½');

  const { xml } = await download(page);
  expect(paragraphXml(xml, 2)).toContain('Total€✓½');
});

test('a symbol inserted while tracking is a tracked insertion; Esc closes the dialog', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await startTracking(page);
  await caretIn(page, 2, 'Total', 5);
  await page.getByRole('button', { name: 'Insert symbol' }).click();
  const dialog = page.getByRole('dialog', { name: 'Symbol' });
  await dialog.getByRole('button', { name: 'Copyright (U+00A9)' }).click();
  await expect(para(page, 2).locator('[data-tc="ins"]')).toHaveText('©');
  await page.getByRole('button', { name: 'Insert symbol' }).click();
  await page.keyboard.press('Escape');
  await expect(dialog).toHaveCount(0);
});
