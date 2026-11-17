// Restrictions Word puts on a document are respected: Restrict Editing (w:documentProtection),
// "read-only recommended" (w:writeProtection) and digital signatures. Encrypted, rights-managed and
// legacy files say what they are instead of failing with a ZIP error.
import fs from 'node:fs';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { buildDocx, openInEditor, download, downloadFile, documentXmlOf, FIELDS, W_NS } from './helpers.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const status = (page) => page.locator('.statusbar [role=status]');
const bar = (page) => page.getByRole('region', { name: 'Restrictions on this document' });

const TEXT = '<w:p><w:r><w:t>The agreed text.</w:t></w:r></w:p>';

/** A document whose settings.xml holds `settings`. */
function withSettings(body, settings) {
  return buildDocx(body, {
    parts: { 'word/settings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W_NS}>${settings}</w:settings>` },
    rels: `<Relationship Id="rIdS" Type="${REL}/settings" Target="settings.xml"/>`,
  });
}

const protection = (edit, extra = '') => `<w:documentProtection w:edit="${edit}" w:enforcement="1"${extra}/>`;
const PASSWORD = ' w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="14" w:cryptSpinCount="100000" w:hash="aGFzaA==" w:salt="c2FsdA=="';

/** Clicks into `text` and types at the end of it. */
async function typeIn(page, locator, text) {
  await locator.click();
  await page.keyboard.press('End');
  await page.keyboard.type(text);
}

// ------------------------------------------------------------------ Restrict Editing

test('read-only protection opens the document read-only and says why', async ({ page }) => {
  await openInEditor(page, await withSettings(TEXT, protection('readOnly', PASSWORD)));
  await expect(bar(page)).toContainText('Restrict Editing is on. Word allows no changes (read only), and asks for a password to stop the protection.');
  await expect(bar(page)).toContainText('Opened read-only.');
  await typeIn(page, page.locator('.docx-page p').first(), ' Changed');
  await expect(page.locator('.docx-page p').first()).toHaveText('The agreed text.');
  await expect(status(page)).toContainText('Read-only: Restrict Editing in Word allows no changes – choose Edit anyway above to change it');
  await expect(page.getByTitle('Bold (Ctrl+B)')).toBeDisabled();
  await expect(page.getByRole('button', { name: 'Track changes' })).toBeDisabled();
  await expect(page.getByTitle(/^New comment/)).toBeDisabled();

  // Downloaded as it is: the text and the protection.
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review).toContainText('No changes');
  await expect(review).toContainText('Restrict Editing (no changes (read only)) stays in the downloaded file');
});

test('Edit anyway lifts a restriction after confirming; the protection stays in the file', async ({ page }) => {
  const docx = await withSettings(TEXT, protection('readOnly', PASSWORD));
  await openInEditor(page, docx);
  const asked = [];
  page.removeAllListeners('dialog');
  page.once('dialog', (d) => { asked.push(d.message()); d.dismiss(); });
  await bar(page).getByRole('button', { name: 'Edit anyway' }).click();
  expect(asked[0]).toContain('Editing it here goes around that restriction');
  await expect(bar(page).getByRole('button', { name: 'Edit anyway' })).toBeVisible(); // cancelled: still locked

  page.once('dialog', (d) => d.accept());
  await bar(page).getByRole('button', { name: 'Edit anyway' }).click();
  await expect(bar(page)).toContainText('Editing anyway. The restriction stays in the file you download, so Word still applies it.');
  await expect(bar(page).getByRole('button', { name: 'Edit anyway' })).toHaveCount(0);
  await typeIn(page, page.locator('.docx-page p').first(), ' Changed');
  await expect(page.locator('.docx-page p').first()).toHaveText('The agreed text. Changed');

  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review).toContainText('Restrict Editing in Word allows no changes (read only), and you chose to edit anyway. The restriction stays in the downloaded file, with its password');
  const [dl] = await Promise.all([page.waitForEvent('download'), review.getByRole('button', { name: 'Download' }).click()]);
  const zip = await JSZip.loadAsync(fs.readFileSync(await dl.path()));
  expect(await zip.file('word/document.xml').async('string')).toContain('Changed');
  const settings = await zip.file('word/settings.xml').async('string');
  expect(settings).toContain('w:edit="readOnly"');
  expect(settings).toContain('w:hash="aGFzaA=="');
});

test('protection that isn’t enforced restricts nothing, as in Word', async ({ page }) => {
  await openInEditor(page, await withSettings(TEXT, '<w:documentProtection w:edit="readOnly" w:enforcement="0"/>'));
  await expect(bar(page)).toHaveCount(0);
  await typeIn(page, page.locator('.docx-page p').first(), ' Changed');
  await expect(page.locator('.docx-page p').first()).toHaveText('The agreed text. Changed');
});

test('comments only: comments can be added, the text stays read-only', async ({ page }) => {
  await openInEditor(page, await withSettings(TEXT, protection('comments')));
  await expect(bar(page)).toContainText('You can add and answer comments; the text is read-only.');
  await typeIn(page, page.locator('.docx-page p').first(), ' Changed');
  await expect(status(page)).toContainText('Only comments can be added');
  await expect(page.locator('.docx-page p').first()).toHaveText('The agreed text.');

  await page.evaluate(() => { // "agreed", to comment on
    const t = document.querySelector('.docx-page p').firstChild.firstChild || document.querySelector('.docx-page p').firstChild;
    const node = t.nodeType === 3 ? t : document.createTreeWalker(t, NodeFilter.SHOW_TEXT).nextNode();
    const range = document.createRange();
    range.setStart(node, 4);
    range.setEnd(node, 10);
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  });
  await page.getByTitle('New comment (Ctrl+Alt+M)').click();
  const composer = page.getByRole('form', { name: 'New comment' });
  await composer.getByRole('textbox', { name: 'Your name' }).fill('Rita');
  await composer.getByRole('textbox', { name: 'New comment' }).fill('Please confirm.');
  await composer.getByRole('button', { name: 'Comment' }).click();
  const { zip, xml } = await download(page);
  expect(xml).toMatch(/<w:t xml:space="preserve">The <\/w:t><\/w:r><w:commentRangeStart w:id="\d+"\/><w:r><w:t[^>]*>agreed<\/w:t><\/w:r><w:commentRangeEnd/);
  expect(await zip.file('word/comments.xml').async('string')).toContain('Please confirm.');
});

test('filling in forms only: the fields can be filled in, the rest is read-only', async ({ page }) => {
  const body = `<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r>${FIELDS}</w:p>`
    + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t xml:space="preserve">In a cell: </w:t></w:r>`
    + `<w:sdt><w:sdtPr><w:alias w:val="Unit"/><w:id w:val="201"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Finance</w:t></w:r></w:sdtContent></w:sdt></w:p></w:tc></w:tr></w:tbl>${TEXT}`;
  await openInEditor(page, await withSettings(body, protection('forms')));
  await expect(bar(page)).toContainText('You can fill in its fields; the rest is read-only.');

  await typeIn(page, page.locator('.docx-page > p').last(), ' Changed');
  await expect(status(page)).toContainText('Only the fields can be filled in');
  await page.locator('.sdt[title*="Client name"] .sdt-content').click();
  await page.keyboard.type('Ada Lovelace');
  await page.locator('.sdt[title*="Unit"] .sdt-content').click();
  await page.keyboard.press('End');
  await page.keyboard.type(' and Audit');
  await page.locator('.sdt-checkbox').click();
  const { xml } = await download(page);
  expect(xml).toContain('Ada Lovelace');
  expect(xml).toContain('Finance and Audit');
  expect(xml).toContain('<w14:checked w14:val="1"/>');
  expect(xml).not.toContain('Changed');
});

test('tracked changes only: every edit is tracked, and changes can’t be accepted or rejected', async ({ page }) => {
  const body = `${TEXT}<w:p><w:ins w:id="1" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Ann’s words.</w:t></w:r></w:ins></w:p>`;
  await openInEditor(page, await withSettings(body, protection('trackedChanges')));
  await expect(bar(page)).toContainText('Your edits are tracked, and tracked changes can’t be accepted or rejected.');
  const track = page.getByRole('button', { name: 'Track changes' });
  await expect(track).toHaveAttribute('aria-pressed', 'true');
  await expect(track).toBeDisabled();
  await expect(page.getByTitle('Accept the change (or the changes selected) and go to the next')).toBeDisabled();
  await page.keyboard.press('Control+Shift+E');
  await expect(track).toHaveAttribute('aria-pressed', 'true');

  await typeIn(page, page.locator('.docx-page p').first(), ' Added');
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:ins [^>]*><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:t xml:space="preserve"> Added<\/w:t><\/w:r><\/w:ins>/);
});

test('read-only recommended opens read-only, and Edit anyway needs no warning', async ({ page }) => {
  await openInEditor(page, await withSettings(TEXT, '<w:writeProtection w:recommended="1"/>'));
  await expect(bar(page)).toContainText('Read-only recommended. The author recommends opening it read-only unless you need to change it.');
  let asked = false;
  page.removeAllListeners('dialog');
  page.on('dialog', (d) => { asked = true; d.accept(); });
  await bar(page).getByRole('button', { name: 'Edit anyway' }).click();
  expect(asked).toBe(false);
  await typeIn(page, page.locator('.docx-page p').first(), ' Changed');
  await expect(page.locator('.docx-page p').first()).toHaveText('The agreed text. Changed');
});

// ------------------------------------------------------------------ digital signatures

/** Adds a (dummy) digital signature to a package, as Word stores one. */
async function sign(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  let ct = await zip.file('[Content_Types].xml').async('string');
  ct = ct.replace('</Types>', '<Default Extension="sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/>'
    + '<Override PartName="/_xmlsignatures/sig1.xml" ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/></Types>');
  zip.file('[Content_Types].xml', ct);
  let rels = await zip.file('_rels/.rels').async('string');
  rels = rels.replace('</Relationships>', `<Relationship Id="rIdSig" Type="${PKG}/digital-signature/origin" Target="_xmlsignatures/origin.sigs"/></Relationships>`);
  zip.file('_rels/.rels', rels);
  zip.file('_xmlsignatures/origin.sigs', '');
  zip.file('_xmlsignatures/_rels/origin.sigs.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${PKG}/digital-signature/signature" Target="sig1.xml"/></Relationships>`);
  zip.file('_xmlsignatures/sig1.xml', '<?xml version="1.0" encoding="UTF-8"?><Signature xmlns="http://www.w3.org/2000/09/xmldsig#" Id="idPackageSignature"/>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

test('a signed document opens read-only and downloads unchanged with its signature', async ({ page }) => {
  const docx = await sign(await buildDocx(TEXT));
  await openInEditor(page, docx);
  await expect(bar(page)).toContainText('Digitally signed. The document has 1 digital signature; changing it removes it, as in Word.');
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review).toContainText('Nothing has changed, so the file downloads exactly as it was opened and its digital signature stays valid.');
  const [dl] = await Promise.all([page.waitForEvent('download'), review.getByRole('button', { name: 'Download' }).click()]);
  expect(Buffer.compare(fs.readFileSync(await dl.path()), docx)).toBe(0); // byte for byte
  await expect(status(page)).toContainText('unchanged, so its digital signature is still valid');
});

test('a real Word document, signed and left untouched, comes back byte for byte', async ({ page }) => {
  const docx = await sign(fs.readFileSync('tests/fixtures/word-letter-id.docx'));
  await openInEditor(page, docx);
  const path = await downloadFile(page);
  expect(Buffer.compare(fs.readFileSync(path), docx)).toBe(0);
});

test('editing a signed document removes its signature, as Word does', async ({ page }) => {
  await openInEditor(page, await sign(await buildDocx(TEXT)));
  await bar(page).getByRole('button', { name: 'Edit anyway' }).click(); // confirmed by openInEditor's dialog handler
  await expect(bar(page)).toContainText('If you change anything, the downloaded file has no signature.');
  await typeIn(page, page.locator('.docx-page p').first(), ' Changed');
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  await expect(review).toContainText('The digital signature is removed from the downloaded file');
  const [dl] = await Promise.all([page.waitForEvent('download'), review.getByRole('button', { name: 'Download' }).click()]);
  const zip = await JSZip.loadAsync(fs.readFileSync(await dl.path()));
  expect(Object.keys(zip.files).filter((p) => p.startsWith('_xmlsignatures'))).toEqual([]);
  expect(await zip.file('_rels/.rels').async('string')).not.toContain('digital-signature');
  const ct = await zip.file('[Content_Types].xml').async('string');
  expect(ct).not.toContain('sigs');
  expect(ct).not.toContain('sig1.xml');
  expect(await documentXmlOf(await zip.generateAsync({ type: 'nodebuffer' }))).toContain('Changed');
  await expect(status(page)).toContainText('digital signature removed');
});

// ------------------------------------------------------------------ files that aren't .docx packages

/** A minimal OLE compound file holding the given stream names (as its directory stores them). */
function compoundFile(...names) {
  const header = Buffer.alloc(512);
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(header);
  return Buffer.concat([header, ...names.map((n) => Buffer.from(n, 'utf16le')), Buffer.alloc(512)]);
}

for (const [what, file, says] of [
  ['password-encrypted', compoundFile('EncryptionInfo', 'EncryptedPackage'), 'This document is encrypted with a password, so Intact can’t open it'],
  ['rights-managed', compoundFile('\u0006DataSpaces', 'DRMEncryptedDataSpace', 'EncryptedPackage'), 'This document is protected by rights management: a sensitivity label that encrypts it, or Information Rights Management.'],
  ['a Word 97–2003', compoundFile('WordDocument', '1Table'), 'This is a Word 97–2003 document (.doc) with a .docx name.'],
  ['not a', Buffer.from('Just some text, not a document.'), 'This file isn’t a .docx document. Every .docx is a ZIP package'],
]) {
  test(`${what} file says what it is`, async ({ page }) => {
    await page.goto('./');
    await page.setInputFiles('input[type=file]', { name: 'report.docx', mimeType: 'application/octet-stream', buffer: file });
    const alert = page.getByRole('alert');
    await expect(alert).toContainText(says);
    await expect(alert).not.toContainText('zip file');
  });
}
