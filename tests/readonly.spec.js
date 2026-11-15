// Read-only content explains itself. Clicking it says what it is and where to edit it, and puts the
// caret beside it (browsers leave the caret where it was, so typing landed out of sight); keys that
// can't apply say why instead of doing nothing.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, canonicalize, documentXmlOf, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const status = (page) => page.locator('.statusbar [role=status]');
const PICTURE = '<w:r><w:drawing><wp:inline><wp:extent cx="190500" cy="190500"/><wp:docPr id="1" name="Logo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId4"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

const BODY = `<w:p><w:r><w:t xml:space="preserve">The client </w:t></w:r><w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>shall promptly and fully</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> pay.</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Locked: </w:t></w:r><w:sdt><w:sdtPr><w:lock w:val="contentLocked"/><w:alias w:val="Company"/><w:id w:val="9"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>ACME Corporation Limited</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Logo: </w:t></w:r>${PICTURE}<w:r><w:t xml:space="preserve"> and </w:t></w:r><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>E=mc²</m:t></m:r></m:oMath></w:p>
<w:altChunk r:id="rIdChunk"/>
<w:p><w:r><w:t>Plain last paragraph.</w:t></w:r></w:p>`;
const OPTS = {
  parts: {
    'word/header1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W_NS}><w:p><w:r><w:t>Confidential header</w:t></w:r></w:p></w:hdr>`,
    'word/comments.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W_NS}><w:comment w:id="0" w:author="Bob"><w:p><w:r><w:t>Check the fee.</w:t></w:r></w:p></w:comment></w:comments>`,
    'word/chunk.html': '<p>Imported</p>',
  },
  rels: `<Relationship Id="rIdH1" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rIdChunk" Type="${REL}/aFChunk" Target="chunk.html"/>`,
  sectPr: '<w:headerReference w:type="default" r:id="rIdH1"/>',
};
const COMMENTED = '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Fee</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>';

/** Clicks `locator` at `fraction` of its width, halfway down. */
async function clickAt(page, locator, fraction) {
  const box = await locator.boundingBox();
  await page.mouse.click(box.x + box.width * fraction, box.y + box.height / 2);
}

async function expectUnchanged(page, docx) {
  const { xml } = await download(page);
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await documentXmlOf(docx)));
}

test('clicking a tracked change says what it is and puts the caret beside it', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  const inserted = page.locator('.docx-page .inserted');
  await clickAt(page, inserted, 0.9);
  await expect(status(page)).toHaveText('Read-only: Inserted by Alice on 2026-01-01 – tracked change, accept or reject it in Word');
  await page.keyboard.type('X');
  await clickAt(page, inserted, 0.1);
  await page.keyboard.type('Y');
  const { xml } = await download(page);
  expect(xml).toContain('<w:t xml:space="preserve">The client Y</w:t></w:r><w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>shall promptly and fully</w:t></w:r></w:ins><w:r><w:t xml:space="preserve">X pay.</w:t>');
});

test('typing after clicking read-only content never lands where the caret was before', async ({ page }) => {
  const docx = await buildDocx(BODY, OPTS);
  await openInEditor(page, docx);
  await page.locator('.docx-page p', { hasText: 'Plain last paragraph.' }).click();
  await clickAt(page, page.locator('.docx-page .raw-block'), 0.5); // a block: no caret beside it
  await expect(status(page)).toHaveText('Read-only: Embedded content – kept as it is; edit it in Word');
  await page.keyboard.type('X');
  await page.keyboard.press('Enter');
  await page.keyboard.press('Backspace');
  await expect(status(page)).toHaveText('Read-only: Embedded content – kept as it is; edit it in Word');
  await expectUnchanged(page, docx);

  // Clicking back into the text types there again.
  await page.locator('.docx-page p', { hasText: 'Plain last paragraph.' }).click();
  await page.keyboard.press('End');
  await page.keyboard.type('!');
  const { xml } = await download(page);
  expect(xml).toContain('Plain last paragraph.!');
});

test('locked fields, pictures and equations say what they are when clicked', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY, OPTS));
  await clickAt(page, page.locator('.docx-page .sdt-locked'), 0.9);
  await expect(status(page)).toHaveText('Read-only: Content control “Company” – locked in Word');
  await page.keyboard.type('!');
  await clickAt(page, page.locator('.docx-page .raw-inline:has(img)'), 0.5);
  await expect(status(page)).toHaveText('Read-only: Picture – kept as it is; edit it in Word');
  await clickAt(page, page.locator('.docx-page [title^="Equation"]'), 0.5);
  await expect(status(page)).toHaveText('Read-only: Equation – kept as it is; edit it in Word');
  const { xml } = await download(page);
  expect(xml).toContain('</w:sdt><w:r><w:t xml:space="preserve">! end.</w:t>'); // typed beside the field
});

test('typing in a header, footer or comment says it is read-only', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  const docx = await buildDocx(COMMENTED, OPTS);
  await openInEditor(page, docx);
  await page.locator('.docx-hf.header').getByText('Confidential header').click();
  await expect(status(page)).toHaveText('Read-only: Header – shown for reference; edit it in Word');
  await page.locator('.docx-page p').first().click();
  await page.keyboard.type('!'); // an edit clears the message
  await expect(status(page)).toHaveText('');
  await page.locator('.docx-hf.header').getByText('Confidential header').click();
  await page.keyboard.type('X');
  await expect(status(page)).toHaveText('Read-only: Header – shown for reference; edit it in Word');

  await page.getByRole('complementary', { name: 'Comments' }).getByText('Check the fee.').dblclick();
  await page.keyboard.type('X');
  await expect(status(page)).toHaveText('Read-only: comments are shown for reference – reply to or resolve them in Word');
  await expect(page.locator('.docx-hf.header')).toHaveText('Confidential header');
  const { zip } = await download(page);
  expect(await zip.file('word/header1.xml').async('string')).toBe(OPTS.parts['word/header1.xml']);
  expect(await zip.file('word/comments.xml').async('string')).toBe(OPTS.parts['word/comments.xml']);
});

test('a locked table cell says why when clicked, and typing there changes nothing', async ({ page }) => {
  const cell = (text) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const docx = await buildDocx('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>'
    + `<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/><w:alias w:val="Total"/><w:id w:val="7"/><w:text/></w:sdtPr><w:sdtContent>${cell('EUR 1,200')}</w:sdtContent></w:sdt>${cell('Notes')}`
    + '</w:tr></w:tbl><w:p><w:r><w:t>After the table.</w:t></w:r></w:p>');
  await openInEditor(page, docx);
  await page.locator('.docx-page p', { hasText: 'After the table.' }).click();
  await page.locator('.docx-page td.sdt-locked').click();
  await expect(status(page)).toHaveText('Read-only: Content control “Total” – locked in Word');
  await page.keyboard.type('X');
  await page.keyboard.press('Backspace');
  await expect(status(page)).toHaveText('Read-only: Content control “Total” – locked in Word');
  await expectUnchanged(page, docx);
  await page.locator('.docx-page td', { hasText: 'Notes' }).click();
  await page.keyboard.press('End');
  await page.keyboard.type('!');
  await expect(page.locator('.docx-page td', { hasText: 'Notes!' })).toHaveCount(1);
});
