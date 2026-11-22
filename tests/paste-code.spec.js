// Pasting code: text in <code>, <pre> and the like, or in a monospaced font (as VS Code, a terminal or
// a documentation site copies it), keeps its lines and indentation and goes in the document's own
// code styles ("Code", "HTML Code") where it has them, else in Consolas, Word's monospaced font.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, W_NS } from './helpers.js';

const P = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const savedParagraphs = (xml) => xml.match(/<w:p[ >].*?<\/w:p>/gs);
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W_NS}>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="HTMLCode"><w:name w:val="HTML Code"/><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/></w:rPr></w:style>
</w:styles>`;

async function paste(page, html, text, where) {
  await page.evaluate(({ html, text, where }) => {
    document.querySelector('.docx-page').focus();
    const p = [...document.querySelectorAll('.docx-page p')].find((e) => e.textContent.includes(where));
    const node = [...p.childNodes].find((n) => n.nodeType === 3 && n.data.includes(where));
    const range = document.createRange();
    range.setStart(node, node.data.indexOf(where) + where.length);
    getSelection().removeAllRanges();
    getSelection().addRange(range);
    const dt = new DataTransfer();
    dt.setData('text/html', html);
    dt.setData('text/plain', text);
    const event = new ClipboardEvent('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: dt });
    p.dispatchEvent(event);
  }, { html, text, where });
}

// As VS Code copies two lines: each a <div> with white-space: pre, in the editor's font.
const VSCODE = '<meta charset="utf-8"><div style="color: #000; background-color: #fff; font-family: Consolas, \'Courier New\', monospace; font-size: 14px; white-space: pre;">'
  + '<div><span style="color: #0000ff;">if</span><span> (ready) {</span></div><div><span>    start();</span></div><div><span>}</span></div></div>';

test('a code block from a code editor goes in the document\'s Code style, its indentation kept', async ({ page }) => {
  await openInEditor(page, await buildDocx(P('Run this:') + P(''), { parts: { 'word/styles.xml': STYLES } }));
  await paste(page, VSCODE, 'if (ready) {\n    start();\n}', 'Run this:');
  await expect(page.getByRole('status')).toContainText('Code pasted in the document’s Code and HTML Code styles');
  const ps = savedParagraphs((await download(page)).xml);
  expect(ps).toContain('<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr><w:r><w:t xml:space="preserve">    start();</w:t></w:r></w:p>');
  expect(ps).toContain('<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr><w:r><w:t xml:space="preserve">}</w:t></w:r></w:p>');
});

test('inline code in a sentence takes the HTML Code character style, and a <pre> block its lines', async ({ page }) => {
  await openInEditor(page, await buildDocx(P('Intro: '), { parts: { 'word/styles.xml': STYLES } }));
  await paste(page, '<p>Call <code>start()</code> first.</p>', 'Call start() first.', 'Intro: ');
  const { xml } = await download(page);
  expect(xml).toContain('<w:r><w:rPr><w:rStyle w:val="HTMLCode"/></w:rPr><w:t xml:space="preserve">start()</w:t></w:r>');
});

test('without code styles, code is pasted in Consolas, and <pre> keeps each line', async ({ page }) => {
  await openInEditor(page, await buildDocx(P('Example:') + P('')));
  await paste(page, '<p>See:</p><pre>make\n  install\n</pre>', 'See:\nmake\n  install', 'Example:');
  await expect(page.getByRole('status')).toContainText('Consolas');
  const ps = savedParagraphs((await download(page)).xml);
  const mono = '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/></w:rPr>';
  expect(ps.some((p) => p.includes(`${mono}<w:t xml:space="preserve">make</w:t>`))).toBe(true);
  expect(ps.some((p) => p.includes(`${mono}<w:t xml:space="preserve">  install</w:t>`))).toBe(true);
});
