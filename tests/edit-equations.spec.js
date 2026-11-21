// Editing an equation already in a document: double-click it (or the right-click menu's Edit
// equation…) to open it in the Equation dialog as LaTeX; Update writes it again as Word's Office
// Math, in its place (tracked while tracking). Unchanged, or with what LaTeX here can't write (Word's
// letter styles), it stays exactly as Word wrote it.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, documentXmlOf, canonicalize } from './helpers.js';

const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const r = (t, rPr = '') => `<m:r>${rPr ? `<m:rPr>${rPr}</m:rPr>` : ''}<m:t xml:space="preserve">${t}</m:t></m:r>`;
const display = (inner) => `<w:p><m:oMathPara ${M}><m:oMath>${inner}</m:oMath></m:oMathPara></w:p>`;
const CTRL = '<m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr></m:ctrlPr>';
const QUADRATIC = display(`${r('x=')}<m:f><m:fPr>${CTRL}</m:fPr><m:num>${r('-b±')}<m:rad><m:radPr><m:degHide m:val="1"/>${CTRL}</m:radPr><m:deg/><m:e><m:sSup><m:e>${r('b')}</m:e><m:sup>${r('2')}</m:sup></m:sSup>${r('-4ac')}</m:e></m:rad></m:num><m:den>${r('2a')}</m:den></m:f>`);
const SUM = display(`<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/></m:naryPr><m:sub>${r('i=1')}</m:sub><m:sup>${r('n')}</m:sup><m:e><m:sSub><m:e>${r('x')}</m:e><m:sub>${r('i')}</m:sub></m:sSub></m:e></m:nary>${r('=')}<m:nary><m:naryPr><m:limLoc m:val="subSup"/></m:naryPr><m:sub>${r('0')}</m:sub><m:sup>${r('∞')}</m:sup><m:e>${r('f')}<m:d><m:e>${r('x')}</m:e></m:d>${r('dx')}</m:e></m:nary>`);
const MATRIX = display(`<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e><m:m><m:mr><m:e>${r('1')}</m:e><m:e>${r('0')}</m:e></m:mr><m:mr><m:e>${r('0')}</m:e><m:e>${r('1')}</m:e></m:mr></m:m></m:e></m:d>`
  + `<m:func><m:fName>${r('sin', '<m:sty m:val="p"/>')}</m:fName><m:e>${r('θ')}</m:e></m:func><m:acc><m:accPr><m:chr m:val="̂"/></m:accPr><m:e>${r('y')}</m:e></m:acc>`);
const SCRIPT = display(`${r('L', '<m:scr m:val="script"/>')}${r('=1')}`);
const INLINE = `<w:p><w:r><w:t xml:space="preserve">The mean </w:t></w:r><m:oMath ${M}><m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>${r('x')}</m:e></m:bar>${r('=')}<m:f><m:num>${r('1')}</m:num><m:den>${r('n')}</m:den></m:f></m:oMath><w:r><w:t xml:space="preserve"> is in the line.</w:t></w:r></w:p>`;
const BODY = QUADRATIC + SUM + MATRIX + SCRIPT + INLINE;

const equation = (page, i) => page.locator('.docx-page .equation').nth(i);
const status = (page) => page.locator('.statusbar [role=status]');
const dialog = (page) => page.getByRole('dialog', { name: 'Edit equation' });
const latex = (page) => dialog(page).getByLabel('Equation, as in LaTeX');

test('Word\'s equations open in the Equation dialog as LaTeX', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  const expected = [
    'x=\\frac{-b\\pm\\sqrt{b^{2}-4ac}}{2a}',
    '\\sum_{i=1}^{n}{x_{i}}=\\int_{0}^{\\infty}{f\\left( x \\right)dx}',
    '\\begin{bmatrix} 1 & 0 \\\\ 0 & 1 \\end{bmatrix}\\sin{\\theta}\\hat{y}',
  ];
  for (const [i, tex] of expected.entries()) {
    await equation(page, i).dblclick();
    await expect(latex(page)).toHaveValue(tex);
    await expect(dialog(page).getByLabel('Preview').locator('math')).toHaveCount(1);
    await expect(dialog(page).getByRole('checkbox', { name: /On its own line/ })).toHaveCount(0);
    await dialog(page).getByRole('button', { name: 'Close' }).click();
  }
  await equation(page, 4).dblclick();
  await expect(latex(page)).toHaveValue('\\overline{x}=\\frac{1}{n}');
  await dialog(page).getByRole('button', { name: 'Close' }).click();
  // Word's script letters: not something LaTeX here writes, so it stays as Word wrote it.
  await expect(equation(page, 3)).not.toHaveAttribute('data-editable');
  await equation(page, 3).dblclick();
  await expect(dialog(page)).toHaveCount(0);
});

test('an edited equation is written again as Word\'s Office Math, on its own line as it was; the others are kept', async ({ page }) => {
  const buffer = await buildDocx(BODY);
  await openInEditor(page, buffer);
  await equation(page, 0).dblclick();
  await latex(page).fill('x=\\frac{-b\\pm\\sqrt{b^{2}-4ac}}{2c}');
  await dialog(page).getByRole('button', { name: 'Update' }).click();
  await expect(status(page)).toHaveText('Equation updated – written as Word’s own equation');
  await expect(equation(page, 0)).toHaveClass(/display/);
  await expect(equation(page, 0).locator('mfrac > mrow').nth(1)).toHaveText('2c');
  const { xml } = await download(page);
  expect(xml).toMatch(/^.*?<w:body><w:p><m:oMathPara[^>]*><m:oMath>(?:(?!<\/m:oMath>).)*<m:den><m:r>(?:(?!<\/m:r>).)*<m:t xml:space="preserve">2c<\/m:t><\/m:r><\/m:den><\/m:f><\/m:oMath><\/m:oMathPara><\/w:p>/s);
  expect(xml).not.toContain('>2a<');
  // The other equations are as they were.
  const original = await documentXmlOf(buffer);
  for (const kept of ['<m:chr m:val="∑"/>', '<m:scr m:val="script"/>', '<m:pos m:val="top"/>']) expect(original).toContain(kept);
  expect(await canonicalize(page, xml.replace(/<w:p><m:oMathPara[^>]*>.*?<\/m:oMathPara><\/w:p>/, '')))
    .toEqual(await canonicalize(page, original.replace(/<w:p><m:oMathPara [^>]*>.*?<\/m:oMathPara><\/w:p>/, '')));

  // Undo puts the equation back as it was.
  await page.keyboard.press('ControlOrMeta+Z');
  await expect(equation(page, 0).locator('mfrac > mrow').nth(1)).toHaveText('2a');
});

test('unchanged, an equation is kept exactly as Word wrote it', async ({ page }) => {
  const buffer = await buildDocx(BODY);
  await openInEditor(page, buffer);
  await equation(page, 4).dblclick();
  await dialog(page).getByRole('button', { name: 'Update' }).click();
  await expect(status(page)).toHaveText('Equation unchanged – kept as it was');
  await page.locator('.docx-page p').last().click();
  await page.keyboard.press('End');
  await page.keyboard.type(' Edited.');
  const { xml } = await download(page);
  const original = (await documentXmlOf(buffer)).replace('is in the line.', 'is in the line. Edited.');
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, original));
});

test('the right-click menu edits the equation clicked; while tracking, the change is tracked', async ({ page }) => {
  await openInEditor(page, await buildDocx(INLINE));
  await page.getByRole('button', { name: 'Track changes' }).click();
  const name = page.getByRole('dialog', { name: 'Your name on tracked changes' });
  await name.getByRole('textbox').fill('Rita');
  await name.getByRole('button', { name: 'Start tracking' }).click();
  await page.locator('.docx-page p').click(); // a caret in the text
  await equation(page, 0).click({ button: 'right' });
  await page.getByRole('menuitem', { name: 'Edit equation…' }).click();
  await latex(page).fill('\\overline{x}=\\frac{1}{N}');
  await dialog(page).getByRole('button', { name: 'Update' }).click();
  await expect(status(page)).toHaveText('Equation updated – written as Word’s own equation');
  const { xml } = await download(page);
  // The old equation deleted and the new one inserted, each marked with Rita's name.
  expect(xml).toMatch(/<w:del w:id="\d+" w:author="Rita" w:date="[^"]+"><m:oMath[^>]*>.*?<m:t xml:space="preserve">n<\/m:t>.*?<\/m:oMath><\/w:del>/);
  expect(xml).toMatch(/<w:ins w:id="\d+" w:author="Rita" w:date="[^"]+"><m:oMath[^>]*>.*?<m:t xml:space="preserve">N<\/m:t>.*?<\/m:oMath><\/w:ins>/);
});
