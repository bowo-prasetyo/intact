// Inserting tables and pictures, as Word does. A table goes at the caret (the paragraph split
// there), across the page, in the Table Grid style (added to the file if need be). A picture goes in
// line with the text at its own size, or narrowed to fit the page; its file, relationship and
// content type are written only if it is still there when saved. Both are tracked while tracking.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, downloadFile } from './helpers.js';

const DOC = '<w:p><w:r><w:t xml:space="preserve">Before the table. After it.</w:t></w:r></w:p><w:p><w:r><w:t>Last.</w:t></w:r></w:p>';
const status = (page) => page.locator('.statusbar [role=status]');
const picker = (page) => page.getByRole('dialog', { name: 'Insert table' });

async function caretAfter(page, text) {
  await page.evaluate((t) => {
    const root = document.querySelector('.docx-page');
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      const i = n.data.indexOf(t);
      if (i < 0) continue;
      root.focus();
      const range = document.createRange();
      range.setStart(n, i + t.length);
      getSelection().removeAllRanges();
      getSelection().addRange(range);
      return;
    }
    throw new Error(`no text "${t}"`);
  }, text);
}

async function startTracking(page) {
  await page.getByRole('button', { name: 'Track changes' }).click();
  await page.getByRole('dialog', { name: 'Your name on tracked changes' }).getByRole('textbox').fill('Rita');
  await page.getByRole('button', { name: 'Start tracking' }).click();
}

/** A PNG of `w` × `h` pixels, drawn in the page. */
async function png(page, w, h) {
  const base64 = await page.evaluate(([cw, ch]) => {
    const c = document.createElement('canvas');
    c.width = cw;
    c.height = ch;
    const g = c.getContext('2d');
    g.fillStyle = '#c00';
    g.fillRect(0, 0, cw, ch);
    return c.toDataURL('image/png').split(',')[1];
  }, [w, h]);
  return Buffer.from(base64, 'base64');
}

async function insertPicture(page, files) {
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.getByTitle('Insert picture').click()]);
  await chooser.setFiles(files);
}

test('a table from the grid goes where the caret is, splitting the paragraph; the caret goes into it', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretAfter(page, 'Before the table.');
  await page.getByTitle('Insert table').click();
  await picker(page).getByRole('button', { name: '3 by 2 table' }).hover();
  await expect(picker(page).locator('.size')).toHaveText('3 × 2 table');
  await picker(page).getByRole('button', { name: '3 by 2 table' }).click();
  await expect(status(page)).toHaveText('Table inserted: 3 × 2');
  await expect(page.locator('table tr')).toHaveCount(2);
  await page.keyboard.type('A1');
  await page.keyboard.press('Tab');
  await page.keyboard.type('B1');
  const blocks = await page.locator('.docx-page > *').evaluateAll((els) => els.map((e) => `${e.tagName}:${e.textContent}`));
  expect(blocks).toEqual(['P:Before the table.', 'TABLE:A1B1', 'P: After it.', 'P:Last.']);

  const { zip, xml } = await download(page);
  // Across the page (9026 twips between the margins), in Table Grid, cells as Word makes them.
  expect(xml).toContain('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="3008"/><w:gridCol w:w="3008"/><w:gridCol w:w="3008"/></w:tblGrid>'
    + '<w:tr><w:tc><w:tcPr><w:tcW w:w="3008" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">A1</w:t></w:r></w:p></w:tc>');
  const styles = await zip.file('word/styles.xml').async('string');
  expect(styles).toContain('<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>');
  expect(styles).toContain('<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>');
});

test('in an empty paragraph, or at the start of one, the table goes before it; any size can be typed', async ({ page }) => {
  await openInEditor(page, await buildDocx(`<w:p/>${DOC}`));
  await page.locator('.docx-page p').first().click();
  await page.getByTitle('Insert table').click();
  await picker(page).getByRole('spinbutton', { name: 'Columns' }).fill('12');
  await picker(page).getByRole('spinbutton', { name: 'Rows' }).fill('9');
  await picker(page).getByRole('button', { name: 'Insert', exact: true }).click();
  await expect(page.locator('table tr')).toHaveCount(9);
  await expect(page.locator('table tr').first().locator('td')).toHaveCount(12);
  const blocks = await page.locator('.docx-page > *').evaluateAll((els) => els.map((e) => e.tagName));
  expect(blocks).toEqual(['TABLE', 'P', 'P', 'P']);
});

test('while tracking, the table\'s rows are tracked insertions, and Reject all takes it out', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await startTracking(page);
  await caretAfter(page, 'Last.');
  await page.getByTitle('Insert table').click();
  await picker(page).getByRole('button', { name: '2 by 2 table' }).press('Enter');
  await expect(status(page)).toHaveText('Table inserted: 2 × 2 (tracked)');
  await expect(page.locator('tr[data-rev-row="ins"]')).toHaveCount(2);
  const { xml } = await download(page);
  expect([...xml.matchAll(/<w:tr><w:trPr><w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"\/><\/w:trPr>/g)]).toHaveLength(2);
  await page.getByTitle('More ways to reject').click();
  await page.getByRole('menuitem', { name: 'Reject all changes' }).click();
  await expect(page.locator('table')).toHaveCount(0);
});

test('the grid picker works from the keyboard: arrows size the table, Escape closes it', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretAfter(page, 'Last.');
  await page.getByTitle('Insert table').click();
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ArrowDown');
  await expect(picker(page).locator('.size')).toHaveText('3 × 2 table');
  await page.keyboard.press('Escape');
  await expect(picker(page)).toHaveCount(0);
  await page.getByTitle('Insert table').click();
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('Enter');
  await expect(page.locator('table tr')).toHaveCount(2);
});

test('a picture goes in line at the caret, at its size; its file, relationship and type are added', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretAfter(page, 'Before');
  await insertPicture(page, { name: 'logo.png', mimeType: 'image/png', buffer: await png(page, 40, 20) });
  await expect(status(page)).toHaveText('Picture inserted');
  const img = page.locator('.docx-page p').first().locator('img');
  await expect(img).toHaveCount(1);
  expect(await img.evaluate((i) => [i.style.width, i.style.height])).toEqual(['40px', '20px']);
  await page.keyboard.type('!'); // the caret is after the picture

  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review.locator('.notes')).toContainText('1 picture added.');
  await review.getByRole('button', { name: 'Keep editing' }).click();
  const { zip, xml } = await download(page);
  // The fixture already has word/media/image1.png (and a picture with drawing id 1).
  expect(xml).toMatch(/<w:t xml:space="preserve">Before<\/w:t><\/w:r><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="381000" cy="190500"\/>/);
  expect(xml).toContain('<wp:docPr id="1" name="Picture 1"/>'); // no other drawing in this file
  expect(xml).toContain('<pic:cNvPr id="1" name="logo.png"/>');
  expect(xml).toContain('<a:blip r:embed="rIdImage1"/>');
  expect(xml).toMatch(/<\/w:drawing><\/w:r><w:r><w:t xml:space="preserve">! the table/);
  expect(zip.file('word/media/image2.png')).toBeTruthy();
  expect(await zip.file('word/_rels/document.xml.rels').async('string')).toContain('<Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>');
});

test('a large picture is narrowed to fit the page; a JPEG gets its content type', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretAfter(page, 'Last.');
  const jpeg = Buffer.from((await page.evaluate(() => {
    const c = document.createElement('canvas');
    c.width = 1203;
    c.height = 400;
    return c.toDataURL('image/jpeg');
  })).split(',')[1], 'base64');
  await insertPicture(page, { name: 'photo.jpg', mimeType: 'image/jpeg', buffer: jpeg });
  // 9026 twips between the margins: 601.7 px wide, so 200 px high.
  const { zip, xml } = await download(page);
  expect(xml).toMatch(/<wp:extent cx="573151\d" cy="190500\d"\/>/);
  expect(zip.file('word/media/image1.jpeg')).toBeTruthy();
  expect(await zip.file('[Content_Types].xml').async('string')).toContain('<Default Extension="jpeg" ContentType="image/jpeg"/>');
});

test('a picture inserted and then undone leaves nothing in the file', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretAfter(page, 'Last.');
  await insertPicture(page, { name: 'dot.png', mimeType: 'image/png', buffer: await png(page, 4, 4) });
  await expect(page.locator('.docx-page img')).toHaveCount(1);
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.docx-page img')).toHaveCount(0);
  const { zip } = await download(page);
  expect(zip.file('word/media/image2.png')).toBeNull();
  expect(await zip.file('word/_rels/document.xml.rels').async('string')).not.toContain('rIdImage');
});

test('while tracking, a picture is a tracked insertion; other files are refused with a reason', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await startTracking(page);
  await caretAfter(page, 'Last.');
  await insertPicture(page, { name: 'notes.webp', mimeType: 'image/webp', buffer: Buffer.from('RIFF') });
  await expect(status(page)).toHaveText('“notes.webp” isn’t a picture Word shows everywhere: insert a PNG, JPEG, GIF or BMP picture');
  await insertPicture(page, { name: 'dot.png', mimeType: 'image/png', buffer: await png(page, 4, 4) });
  await expect(status(page)).toHaveText('Picture inserted (tracked)');
  await expect(page.locator('.docx-page .tc-ins img')).toHaveCount(1);
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"><w:r><w:drawing>/);
});

test('a picture pasted on its own (a screenshot) is inserted', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretAfter(page, 'Last.');
  const base64 = (await png(page, 10, 10)).toString('base64');
  await page.evaluate((b64) => {
    const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
    const data = new DataTransfer();
    data.items.add(new File([bytes], 'image.png', { type: 'image/png' }));
    document.querySelector('.docx-page').dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
  }, base64);
  await expect(page.locator('.docx-page img')).toHaveCount(1);
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Picture inserted');
});

test('the file still opens after inserting both: every part is in the package', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await caretAfter(page, 'Last.');
  await insertPicture(page, { name: 'dot.png', mimeType: 'image/png', buffer: await png(page, 4, 4) });
  await page.getByTitle('Insert table').click();
  await picker(page).getByRole('button', { name: '1 by 1 table' }).click();
  const path = await downloadFile(page);
  const fs = await import('node:fs');
  await openInEditor(page, fs.readFileSync(path));
  await expect(page.locator('.docx-page img')).toHaveCount(1);
  await expect(page.locator('.docx-page table')).toHaveCount(1);
});

test('a picture file dropped on the page goes where it is dropped', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  const base64 = (await png(page, 10, 10)).toString('base64');
  const box = await page.getByText('Last.').boundingBox();
  await page.evaluate(([b64, x, y]) => {
    const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
    const data = new DataTransfer();
    data.items.add(new File([bytes], 'drop.png', { type: 'image/png' }));
    document.elementFromPoint(x, y).dispatchEvent(new DragEvent('drop', { dataTransfer: data, clientX: x, clientY: y, bubbles: true, cancelable: true }));
  }, [base64, box.x + box.width - 1, box.y + box.height / 2]);
  await expect(page.locator('.docx-page p').nth(1).locator('img')).toHaveCount(1);
});
