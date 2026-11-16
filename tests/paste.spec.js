// Pasting keeps bold, italic, underline, strikethrough, paragraphs and lists ("merge formatting"),
// from web pages, Google Docs and Word, and drops everything else. Also: list spacing as in Word.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const DOC = '<w:p><w:r><w:t xml:space="preserve">Start  end</w:t></w:r></w:p><w:p><w:r><w:t>Other</w:t></w:r></w:p>';

/**
 * Defines window.pasteEvent(dataTransfer) in the page: a paste event carrying that data. The data is
 * attached to the event itself, because Firefox ignores `clipboardData` passed to the constructor of
 * a scripted ClipboardEvent (Chromium honours it).
 */
async function definePasteEvent(page) {
  await page.evaluate(() => {
    window.pasteEvent = (dt) => {
      const event = new ClipboardEvent('paste', { bubbles: true, cancelable: true });
      Object.defineProperty(event, 'clipboardData', { value: dt });
      return event;
    };
  });
}

/** Puts the caret after "Start " and pastes `html` (and `text`) as the browser would. */
async function pasteAt(page, { html = '', text = '' }, where = 'Start ') {
  await definePasteEvent(page);
  await page.evaluate(({ html, text, where }) => {
    document.querySelector('.docx-page').focus(); // a real paste happens in the focused editor
    const p = [...document.querySelectorAll('.docx-page p')].find((e) => e.textContent.includes(where.trim()));
    const node = [...p.childNodes].find((n) => n.nodeType === 3 && n.data.includes(where)) || p.firstChild;
    const range = document.createRange();
    range.setStart(node, node.data.indexOf(where) + where.length);
    getSelection().removeAllRanges();
    getSelection().addRange(range);
    const dt = new DataTransfer();
    if (html) dt.setData('text/html', html);
    dt.setData('text/plain', text);
    p.dispatchEvent(pasteEvent(dt));
  }, { html, text, where });
}

const savedParagraph = (xml, text) => xml.match(/<w:p[ >].*?<\/w:p>/gs).find((p) => p.includes(text));
const markers = (page) => page.locator('.docx-page .list').evaluateAll((els) => els.map((e) => `${e.dataset.marker} ${e.textContent}`));

test('pasting from a web page keeps bold, italic, underline and strikethrough, and nothing else', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  // A link to anything but a web or e-mail address is dropped (its text is kept); see links.spec.js.
  await pasteAt(page, {
    html: '<meta charset="utf-8"><span style="font-family: Comic Sans MS; color: red; font-size: 30px">big <b>bold</b> <i>it</i> <u>under</u> <s>gone</s> <a href="javascript:alert(1)" data-rpr="0">link</a></span>',
    text: 'big bold it under gone link',
  });
  const p = page.locator('.docx-page p').first();
  await expect(p).toHaveText('Start big bold it under gone link end');
  await expect(p.locator('[style], [data-rpr], a, font')).toHaveCount(0);

  const { xml } = await download(page);
  const saved = savedParagraph(xml, 'Start');
  expect(saved).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">bold</w:t></w:r>');
  expect(saved).toContain('<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">it</w:t></w:r>');
  expect(saved).toContain('<w:u w:val="single"/>');
  expect(saved).toContain('<w:strike/>');
  expect(saved).not.toMatch(/rFonts|w:color|w:sz|hyperlink/);
});

test('Google Docs: its <b style="font-weight:normal"> wrapper is not bold, its spans are', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await pasteAt(page, {
    html: '<meta charset="utf-8"><b style="font-weight:normal;" id="docs-internal-guid-1"><span style="font-weight:700;">Bold</span><span style="font-weight:400;font-style:italic;"> italic</span></b>',
  });
  const p = page.locator('.docx-page p').first();
  await expect(p.locator('strong')).toHaveText('Bold');
  await expect(p.locator('em')).toHaveText(' italic');
});

test('pasted paragraphs and lists become paragraphs and real list items', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await pasteAt(page, {
    html: '<p>Intro</p><ul><li>Apples</li><li>Pears<ul><li>Conference</li></ul></li></ul><ol><li>First</li><li>Second</li></ol><p>Outro</p>',
  });
  await expect(page.locator('.docx-page > p')).toHaveText(['Start Intro', 'Apples', 'Pears', 'Conference', 'First', 'Second', 'Outro end', 'Other']);
  await expect.poll(() => markers(page)).toEqual(['• Apples', '• Pears', '◦ Conference', '1. First', '2. Second']);

  const { xml } = await download(page);
  expect(savedParagraph(xml, 'Conference')).toContain('<w:ilvl w:val="1"/>');
  expect(savedParagraph(xml, 'Outro')).not.toContain('numPr');
  expect(savedParagraph(xml, 'Start Intro')).not.toContain('numPr');

  await page.getByTitle('Undo (Ctrl+Z)').click(); // one paste, one undo step
  await expect(page.locator('.docx-page > p')).toHaveText(['Start  end', 'Other']);
});

test('Word\'s clipboard lists (mso-list paragraphs) become list items without their typed markers', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  const item = (level, marker, text) => `<p class=MsoListParagraphCxSpMiddle style='text-indent:-.25in;mso-list:l0 level${level} lfo1'><![if !supportLists]><span style='font-family:Symbol;mso-fareast-font-family:Symbol'><span style='mso-list:Ignore'>${marker}<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp; </span></span></span><![endif]><b>${text}</b><o:p></o:p></p>`;
  await pasteAt(page, {
    html: `<html xmlns:o="urn:schemas-microsoft-com:office:office"><body><!--StartFragment--><p class=MsoNormal>Plan:<o:p></o:p></p>${item(1, '·', 'One')}${item(2, 'o', 'Two')}<!--EndFragment--></body></html>`,
  });
  await expect.poll(() => markers(page)).toEqual(['• One', '◦ Two end']);
  await expect(page.locator('.docx-page .list strong')).toHaveText(['One', 'Two']);
});

test('pictures are not pasted, and the status bar says so; plain text still pastes as paragraphs', async ({ page }) => {
  await openInEditor(page, await buildDocx(DOC));
  await pasteAt(page, { html: '<p>Logo: <img src="https://example.com/x.png" onerror="window.pwned=1"></p>' });
  await expect(page.locator('.statusbar')).toContainText('1 picture was not pasted');
  await expect(page.locator('.docx-page img')).toHaveCount(0);
  expect(await page.evaluate(() => window.pwned)).toBeUndefined();

  await pasteAt(page, { text: 'line one\nline two' }, 'Other');
  await expect(page.locator('.docx-page > p').nth(1)).toHaveText('Otherline one');
  await expect(page.locator('.docx-page > p').nth(2)).toHaveText('line two');
});

test('list items copied from Intact paste as list items', async ({ page }) => {
  // Into an empty paragraph, so the first item becomes a list item too.
  await openInEditor(page, await buildDocx('<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p/><w:p><w:r><w:t>After</w:t></w:r></w:p>'));
  await definePasteEvent(page);
  await page.evaluate(() => {
    document.querySelector('.docx-page').focus();
    const empty = document.querySelectorAll('.docx-page p')[1];
    const range = document.createRange();
    range.setStart(empty, 0);
    getSelection().removeAllRanges();
    getSelection().addRange(range);
    const dt = new DataTransfer();
    dt.setData('text/html', '<p class="list" data-marker="•" data-style="ListBullet" data-ppr="5">Copied one</p><p class="list" data-marker="a." data-num-id="7" data-ilvl="1">Copied two</p>');
    empty.dispatchEvent(pasteEvent(dt));
  });
  await expect.poll(() => markers(page)).toEqual(['• Copied one', 'a. Copied two']);
  await expect(page.locator('.docx-page [data-ppr="5"], .docx-page [data-style="ListBullet"]')).toHaveCount(0);
});

test('pasting a list into a list of the same kind continues it', async ({ page }) => {
  const NUM = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';
  await openInEditor(page, await buildDocx(`<w:p>${NUM}<w:r><w:t>Alpha</w:t></w:r></w:p><w:p>${NUM}<w:r><w:t xml:space="preserve">Beta </w:t></w:r></w:p>`));
  await pasteAt(page, { html: '<ol><li>Gamma</li><li>Delta</li></ol>' }, 'Beta ');
  await expect.poll(() => markers(page)).toEqual(['1. Alpha', '2. Beta Gamma', '3. Delta']);
});

test('list items with "no space between paragraphs of the same style" sit tight, as in Word', async ({ page }) => {
  const styles = '<w:pPr><w:pStyle w:val="ListBullet"/><w:contextualSpacing/><w:spacing w:after="400"/></w:pPr>';
  await openInEditor(page, await buildDocx(`<w:p>${styles}<w:r><w:t>One</w:t></w:r></w:p><w:p>${styles}<w:r><w:t>Two</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:after="400"/></w:pPr><w:r><w:t>After</w:t></w:r></w:p>`));
  const margin = (text) => page.locator('.docx-page p', { hasText: text }).evaluate((e) => getComputedStyle(e).marginBottom);
  expect(await margin('One')).toBe('0px'); // next paragraph has the same style
  expect(await margin('Two')).toBe('26.6667px'); // 20pt: the next paragraph has another style
});
