// Writing tools from Word's Review tab and AutoCorrect options, for a reporter's copy: your own
// AutoCorrect entries (and common English typos) replaced as you type, in the case typed, with Ctrl+Z
// straight after giving back what was typed; the Thesaurus (Shift+F7), on this device, in English and
// Indonesian; and Copy for the web, which puts the text on the clipboard as a website's editor (CMS)
// takes it (clean HTML and plain text) or as Markdown.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';
import { inflect, formsOf, participleAfter } from '../src/editor/inflect.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const status = (page) => page.locator('.statusbar [role=status]');
const para = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const lastPara = (page) => page.locator('.docx-page > p').last();

/** A document in `lang` (w:lang in the default run properties). */
function inLanguage(lang, body = para('')) {
  const styles = `${XML}<w:styles ${W_NS}><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="${lang}"/></w:rPr></w:rPrDefault></w:docDefaults>`
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    + '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style></w:styles>';
  return buildDocx(body, { parts: { 'word/styles.xml': styles } });
}

async function typeAtEnd(page, text) {
  await lastPara(page).click();
  await page.keyboard.press('End');
  await page.keyboard.type(text);
}

async function addEntry(page, from, to) {
  await page.getByRole('button', { name: 'AutoCorrect options' }).click();
  const dialog = page.getByRole('dialog', { name: 'AutoCorrect' });
  await dialog.getByLabel('Replace', { exact: true }).fill(from);
  await dialog.getByLabel('With').fill(to);
  await dialog.getByRole('button', { name: 'Add' }).click();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toBeHidden();
}

/** Puts the caret in `word` (its first occurrence) in the page. */
async function caretIn(page, word) {
  await page.evaluate((w) => {
    const walker = document.createTreeWalker(document.querySelector('.docx-page'), NodeFilter.SHOW_TEXT);
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      const at = t.data.indexOf(w);
      if (at < 0) continue;
      document.querySelector('.docx-page').focus();
      getSelection().collapse(t, at + 1);
      return;
    }
    throw new Error(`Not found: ${w}`);
  }, word);
}

test.beforeEach(async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 800 });
  // What the page puts on the clipboard (reading the clipboard itself needs a permission Firefox lacks).
  await page.addInitScript(() => {
    const set = DataTransfer.prototype.setData;
    DataTransfer.prototype.setData = function (type, data) {
      (window.copied ||= {})[type] = data;
      return set.call(this, type, data);
    };
  });
});

test('your AutoCorrect entry is replaced as you type, in the case you typed it; Ctrl+Z gives it back', async ({ page }) => {
  await openInEditor(page, await inLanguage('en-GB'));
  await addEntry(page, 'cllr', 'Councillor');
  await typeAtEnd(page, 'cllr Ruiz and CLLR Lee. Cllr');
  await expect(lastPara(page)).toHaveText('Councillor Ruiz and COUNCILLOR Lee. Cllr');
  await page.keyboard.press(',');
  await expect(lastPara(page)).toHaveText('Councillor Ruiz and COUNCILLOR Lee. Councillor,');
  await page.keyboard.press('Control+z');
  await expect(lastPara(page)).toHaveText('Councillor Ruiz and COUNCILLOR Lee. Cllr,');
  await page.keyboard.type(' cllr');
  await page.keyboard.press('Enter');
  await expect(page.locator('.docx-page > p').nth(-2)).toHaveText('Councillor Ruiz and COUNCILLOR Lee. Cllr, Councillor');

  // Kept in this browser: a reload still has it.
  await page.reload();
  await page.getByText('Blank document').click();
  await page.getByRole('button', { name: 'AutoCorrect options' }).click();
  await expect(page.getByRole('dialog', { name: 'AutoCorrect' }).locator('.entries')).toContainText('cllr→Councillor');
});

test('AutoCorrect while tracking changes: the word just typed is your own insertion, so it becomes the entry', async ({ page }) => {
  await openInEditor(page, await inLanguage('en-GB'));
  await addEntry(page, 'mfa', 'Ministry of Foreign Affairs');
  await page.getByRole('button', { name: 'Track changes' }).click();
  const who = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await who.getByRole('textbox').fill('Desk');
  await who.getByRole('button', { name: 'Start tracking' }).click();
  await typeAtEnd(page, 'The mfa said');
  // As in Word, changing your own insertion changes it: no deletion of what was never there.
  await expect(lastPara(page)).toHaveText('The Ministry of Foreign Affairs said');
  await expect(page.locator('.docx-page [data-tc="del"]')).toHaveCount(0);
  const { xml } = await download(page);
  expect(xml).toMatch(/<w:ins [^>]*w:author="Desk"[^>]*>.*Ministry of Foreign Affairs/);
  expect(xml).not.toContain('mfa');
});

test('common English typos are corrected in English text, as typed or with a capital, not in capitals', async ({ page }) => {
  await openInEditor(page, await inLanguage('en-US'));
  await typeAtEnd(page, 'teh vote adn Teh count. TEH ');
  await expect(lastPara(page)).toHaveText('the vote and The count. TEH ');
});

test('“i” on its own becomes “I”, but not a list’s “(i)” or “i.e.”', async ({ page }) => {
  await openInEditor(page, await inLanguage('en-GB'));
  await typeAtEnd(page, "i think i'm sure, i.e. (i) yes ");
  await expect(lastPara(page)).toHaveText('I think I’m sure, i.e. (i) yes ');
});

test('common English typos aren’t corrected in Indonesian text (where “teh” is tea)', async ({ page }) => {
  await openInEditor(page, await inLanguage('id-ID'));
  await typeAtEnd(page, 'minum teh manis ');
  await expect(lastPara(page)).toHaveText('minum teh manis ');
});

test('AutoCorrect switched off types the text as it is', async ({ page }) => {
  await openInEditor(page, await inLanguage('en-GB'));
  await page.getByRole('button', { name: 'AutoCorrect options' }).click();
  const dialog = page.getByRole('dialog', { name: 'AutoCorrect' });
  await dialog.getByLabel('Replace text as you type').uncheck();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(status(page)).toHaveText('AutoCorrect is off: text is typed as it is');
  await typeAtEnd(page, 'teh end ');
  await expect(lastPara(page)).toHaveText('teh end ');
});

test('an entry needs one word to replace, and opened with text selected offers it to replace with', async ({ page }) => {
  await openInEditor(page, await inLanguage('en-GB', para('Royal Institute of Chartered Surveyors')));
  await page.evaluate(() => {
    const t = document.querySelector('.docx-page p').firstChild;
    document.querySelector('.docx-page').focus();
    getSelection().setBaseAndExtent(t, 0, t, t.length);
  });
  await page.getByRole('button', { name: 'AutoCorrect options' }).click();
  const dialog = page.getByRole('dialog', { name: 'AutoCorrect' });
  await expect(dialog.getByLabel('With')).toHaveValue('Royal Institute of Chartered Surveyors');
  await dialog.getByLabel('Replace', { exact: true }).fill('ri cs');
  await dialog.getByRole('button', { name: 'Add' }).click();
  await expect(dialog.getByRole('alert')).toHaveText('The text to replace is one word, without spaces');
});

test('the thesaurus (Shift+F7) offers other words; a click puts one in place, in the word’s case', async ({ page }) => {
  await openInEditor(page, await inLanguage('en-GB', para('Big changes approved.')));
  await caretIn(page, 'Big');
  await page.keyboard.press('Shift+F7');
  const dialog = page.getByRole('dialog', { name: 'Thesaurus' });
  await expect(dialog.getByRole('searchbox', { name: 'Look up' })).toHaveValue('Big');
  const first = dialog.locator('.meaning').first();
  await expect(first).toContainText('large (adjective)');
  await expect(first).toContainText('Opposite');
  await first.getByRole('button', { name: 'large', exact: true }).click();
  await expect(dialog).toBeHidden();
  await expect(page.locator('.docx-page p')).toHaveText('Large changes approved.');
  await expect(status(page)).toHaveText('“Big” replaced with “Large”');

  // Another form of a word: "approved" is looked up as "approve".
  await caretIn(page, 'approved');
  await page.keyboard.press('Shift+F7');
  await expect(dialog.locator('.meaning').first()).toContainText('sanction');
  await dialog.getByRole('searchbox', { name: 'Look up' }).fill('xyzzyq');
  await dialog.getByRole('button', { name: 'Look up' }).click();
  await expect(dialog).toContainText('No meanings found for “xyzzyq”.');
});

test('in Indonesian text the thesaurus is Indonesian', async ({ page }) => {
  await openInEditor(page, await inLanguage('id-ID', para('Rumah itu besar.')));
  await caretIn(page, 'besar');
  await page.getByRole('button', { name: 'Thesaurus' }).click();
  const dialog = page.getByRole('dialog', { name: 'Thesaurus' });
  await expect(dialog.getByRole('combobox', { name: 'Language' })).toHaveValue('id');
  await dialog.getByRole('button', { name: 'gede', exact: true }).first().click();
  await expect(page.locator('.docx-page p')).toHaveText('Rumah itu gede.');
});

const STORY = '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Levy approved</w:t></w:r></w:p>'
  + '<w:p><w:r><w:t xml:space="preserve">The </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>council</w:t></w:r><w:r><w:t xml:space="preserve"> voted </w:t></w:r>'
  + '<w:r><w:rPr><w:i/></w:rPr><w:t>7-2</w:t></w:r><w:r><w:t xml:space="preserve">, said </w:t></w:r>'
  + `<w:hyperlink r:id="rId3" xmlns:r="${REL}"><w:r><w:t>the minutes</w:t></w:r></w:hyperlink>`
  + '<w:del w:id="1" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:delText> (draft)</w:delText></w:r></w:del>'
  + '<w:r><w:rPr><w:vanish/></w:rPr><w:t> HIDDEN</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>'
  + '<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>Riverside &lt;homes&gt;</w:t></w:r></w:p>'
  + '<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>Shops</w:t></w:r></w:p>'
  + '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
  + '<w:tr><w:tc><w:p><w:r><w:t>For</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>7</w:t></w:r></w:p></w:tc></w:tr>'
  + '<w:tr><w:tc><w:p><w:r><w:t>Against</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';

async function copyFromMenu(page, item) {
  await page.evaluate(() => { window.copied = {}; });
  await page.getByRole('button', { name: 'More ways to save' }).click();
  await page.getByRole('menu', { name: 'Save' }).getByRole('menuitem', { name: new RegExp(`^${item}`) }).click();
  return page.evaluate(() => window.copied);
}

test('Copy for the web puts clean HTML and plain text on the clipboard: the text as it will read', async ({ page }) => {
  await openInEditor(page, await buildDocx(STORY));
  const copied = await copyFromMenu(page, 'Copy for the web');
  expect(copied['text/html']).toBe('<h1>Levy approved</h1>'
    + '<p>The <strong>council</strong> voted <em>7-2</em>, said <a href="https://example.com/">the minutes</a>.</p>'
    + '<ul><li>Riverside &lt;homes&gt;</li><li>Shops</li></ul>'
    + '<table><tr><td>For</td><td>7</td></tr><tr><td>Against</td><td>2</td></tr></table>');
  expect(copied['text/plain']).toBe('Levy approved\n\nThe council voted 7-2, said the minutes.\n\n• Riverside <homes>\n• Shops\n\nFor\t7\nAgainst\t2\n');
  await expect(status(page)).toHaveText(/^Copied for the web: 5 paragraphs/);
});

test('Copy as Markdown, and Copy for the web of the selection from the right-click menu', async ({ page }) => {
  await openInEditor(page, await buildDocx(STORY));
  const copied = await copyFromMenu(page, 'Copy as Markdown');
  expect(copied['text/plain']).toBe('# Levy approved\n\nThe **council** voted *7-2*, said [the minutes](https://example.com/).\n\n'
    + '- Riverside \\<homes\\>\n- Shops\n\n| For | 7 |\n| --- | --- |\n| Against | 2 |\n');
  expect(copied['text/html']).toBeUndefined();

  // The selection, from the right-click menu: what is selected, as the web takes it.
  await page.evaluate(() => { window.copied = {}; });
  const box = await page.evaluate(() => {
    const p = [...document.querySelectorAll('.docx-page p')][0];
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    let t = w.nextNode();
    while (!t.data.includes('council')) t = w.nextNode();
    const r = document.createRange();
    r.selectNodeContents(t);
    document.querySelector('.docx-page').focus();
    getSelection().removeAllRanges();
    getSelection().addRange(r);
    const b = r.getBoundingClientRect();
    return { x: b.left + b.width / 2, y: b.top + b.height / 2 };
  });
  await page.mouse.click(box.x, box.y, { button: 'right' });
  await page.getByRole('menu', { name: 'Context menu' }).getByRole('menuitem', { name: 'Copy for the web' }).click();
  expect(await page.evaluate(() => window.copied)).toEqual({ 'text/html': '<p><strong>council</strong></p>', 'text/plain': 'council\n' });
});

test('the thesaurus offers words in the form looked up: approves → sanctions, took → chose, bigger → larger', async ({ page }) => {
  await openInEditor(page, await inLanguage('en-GB', para('The board approves it. She took the job. A bigger levy.')));
  const dialog = page.getByRole('dialog', { name: 'Thesaurus' });
  await caretIn(page, 'approves');
  await page.keyboard.press('Shift+F7');
  await expect(dialog).toContainText('Words for “approve” are shown as “approves” is: okays.');
  await expect(dialog.locator('.meaning').first()).toContainText('(verb · approve)');
  await dialog.getByRole('button', { name: 'sanctions', exact: true }).click();
  await expect(page.locator('.docx-page p')).toHaveText('The board sanctions it. She took the job. A bigger levy.');

  await caretIn(page, 'took');
  await page.keyboard.press('Shift+F7');
  await dialog.getByRole('button', { name: 'chose', exact: true }).click();
  await caretIn(page, 'bigger');
  await page.keyboard.press('Shift+F7');
  await expect(dialog.getByRole('button', { name: 'more prominent' })).toBeVisible();
  await dialog.getByRole('button', { name: 'larger', exact: true }).first().click();
  await expect(page.locator('.docx-page p')).toHaveText('The board sanctions it. She chose the job. A larger levy.');
});

test('after “had” or “was” an -ed form is a past participle: returned → gave, but had returned → given', async ({ page }) => {
  await openInEditor(page, await inLanguage('en-GB', para('She returned the money. He had returned the keys.')));
  const dialog = page.getByRole('dialog', { name: 'Thesaurus' });
  await caretIn(page, 'returned the money');
  await page.keyboard.press('Shift+F7');
  await expect(dialog.getByRole('button', { name: 'gave', exact: true }).first()).toBeVisible();
  await expect(dialog.getByRole('button', { name: 'given', exact: true })).toHaveCount(0);
  await page.keyboard.press('Escape');
  await caretIn(page, 'returned the keys');
  await page.keyboard.press('Shift+F7');
  await expect(dialog.getByRole('button', { name: 'gave', exact: true })).toHaveCount(0);
  await dialog.getByRole('button', { name: 'given', exact: true }).first().click();
  await expect(page.locator('.docx-page p')).toHaveText('She returned the money. He had given the keys.');
});

test('AutoCorrect entries go to a CSV file and come back from one, merged, in another browser', async ({ page }) => {
  await openInEditor(page, await inLanguage('en-GB'));
  await addEntry(page, 'cllr', 'Councillor');
  await addEntry(page, 'rics', 'Royal Institution of Chartered Surveyors, London');
  await page.getByRole('button', { name: 'AutoCorrect options' }).click();
  const dialog = page.getByRole('dialog', { name: 'AutoCorrect' });
  const [file] = await Promise.all([page.waitForEvent('download'), dialog.getByRole('button', { name: 'Export…' }).click()]);
  expect(file.suggestedFilename()).toBe('autocorrect.csv');
  const csv = (await import('node:fs')).readFileSync(await file.path(), 'utf8');
  expect(csv).toBe('﻿Replace,With\r\ncllr,Councillor\r\nrics,"Royal Institution of Chartered Surveyors, London"\r\n');
  await expect(dialog.getByRole('status')).toHaveText('Exported 2 entries to autocorrect.csv, in your downloads');
  await dialog.getByRole('button', { name: 'Cancel' }).click();

  // Another browser: nothing kept, then the list imported (with a list of its own, made in a spreadsheet).
  await page.evaluate(() => localStorage.removeItem('intact.autocorrect'));
  await page.reload();
  await page.getByText('Blank document').click();
  await page.getByRole('button', { name: 'AutoCorrect options' }).click();
  await dialog.locator('input[type=file]').setInputFiles({ name: 'autocorrect.csv', mimeType: 'text/csv', buffer: Buffer.from(csv) });
  await expect(dialog.getByRole('status')).toHaveText('From autocorrect.csv: 2 entries added. Choose OK to keep them.');
  await dialog.locator('input[type=file]').setInputFiles({ name: 'mine.txt', mimeType: 'text/plain', buffer: Buffer.from('cllr\tCllr\ndpr\tDewan Perwakilan Rakyat\nno good\tx\n') });
  await expect(dialog.getByRole('status')).toHaveText('From mine.txt: 1 entry added, 1 changed, 1 row skipped (not an entry: a word to replace with spaces, or nothing to replace it with). Choose OK to keep them.');
  await expect(dialog.locator('.entries li')).toHaveText(['cllr→Cllr×', 'dpr→Dewan Perwakilan Rakyat×', 'rics→Royal Institution of Chartered Surveyors, London×']);
  await dialog.getByRole('button', { name: 'OK' }).click();
  await page.locator('.docx-page').click();
  await page.keyboard.type('The dpr met. ');
  await expect(page.locator('.docx-page p').last()).toHaveText('The Dewan Perwakilan Rakyat met. ');
});

test('English word forms: the base and form of a word, and a word put in that form', () => {
  const forms = (w) => formsOf(w).filter((f) => f.kind !== 'base').map((f) => `${f.base}/${f.pos}/${f.kind}`);
  expect(forms('took')).toEqual(['take/v/simplePast']);
  expect(forms('taken')).toEqual(['take/v/participle']);
  expect(forms('stopped')).toContain('stop/v/past');
  expect(forms('children')).toEqual(['child/n/plural']);
  const cases = [
    ['sanction', 'third', 'v', 'sanctions'], ['carry', 'third', 'v', 'carries'], ['push', 'third', 'v', 'pushes'], ['go', 'third', 'v', 'goes'],
    ['stop', 'past', 'v', 'stopped'], ['carry', 'past', 'v', 'carried'], ['approve', 'past', 'v', 'approved'], ['visit', 'past', 'v', 'visited'],
    ['permit', 'past', 'v', 'permitted'], ['offer', 'past', 'v', 'offered'], ['tell', 'past', 'v', 'told'], ['give', 'simplePast', 'v', 'gave'],
    ['give', 'participle', 'v', 'given'], ['turn over', 'third', 'v', 'turns over'], ['die', 'ing', 'v', 'dying'], ['make', 'ing', 'v', 'making'],
    ['run', 'ing', 'v', 'running'], ['see', 'ing', 'v', 'seeing'], ['levy', 'plural', 'n', 'levies'], ['tax', 'plural', 'n', 'taxes'],
    ['man', 'plural', 'n', 'men'], ['flood wall', 'plural', 'n', 'flood walls'], ['right of way', 'plural', 'n', 'rights of way'],
    ['large', 'er', 'a', 'larger'], ['big', 'est', 'a', 'biggest'], ['happy', 'er', 'a', 'happier'], ['prominent', 'er', 'a', 'more prominent'],
    ['good', 'est', 'a', 'best'], ['quickly', 'base', 'r', 'quickly'],
  ];
  for (const [word, kind, pos, want] of cases) expect(`${word} → ${inflect(word, kind, pos)}`).toBe(`${word} → ${want}`);
  expect(inflect('give', 'past', 'v', true)).toBe('given');
  expect(inflect('O.K.', 'third', 'v')).toBeNull();
  expect(participleAfter('It has ')).toBe(true);
  expect(participleAfter('It was quickly ')).toBe(true);
  expect(participleAfter('They ')).toBe(false);
});
