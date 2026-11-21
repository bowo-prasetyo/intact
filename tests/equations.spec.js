// Equations (Office Math) are shown as MathML, which the browser draws, instead of a placeholder:
// fractions, roots, scripts, sums and integrals, delimiters, matrices, functions, accents and Word's
// letter styles. They are written back exactly as they were unless edited (see edit-equations.spec.js).
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download, documentXmlOf, canonicalize } from './helpers.js';

const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const r = (t, rPr = '') => `<m:r>${rPr ? `<m:rPr>${rPr}</m:rPr>` : ''}<m:t xml:space="preserve">${t}</m:t></m:r>`;
const display = (inner) => `<w:p><m:oMathPara ${M}><m:oMath>${inner}</m:oMath></m:oMathPara></w:p>`;
const QUADRATIC = display(`${r('x=')}<m:f><m:num>${r('-b±')}<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:sSup><m:e>${r('b')}</m:e><m:sup>${r('2')}</m:sup></m:sSup>${r('-4ac')}</m:e></m:rad></m:num><m:den>${r('2a')}</m:den></m:f>`);
const SUM = display(`<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/></m:naryPr><m:sub>${r('i=1')}</m:sub><m:sup>${r('n')}</m:sup><m:e><m:sSub><m:e>${r('x')}</m:e><m:sub>${r('i')}</m:sub></m:sSub></m:e></m:nary>${r('=')}<m:nary><m:naryPr><m:limLoc m:val="subSup"/></m:naryPr><m:sub>${r('0')}</m:sub><m:sup>${r('∞')}</m:sup><m:e>${r('f')}<m:d><m:e>${r('x')}</m:e></m:d>${r('dx')}</m:e></m:nary>`);
const MIXED = display(`<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e><m:m><m:mr><m:e>${r('1')}</m:e><m:e>${r('0')}</m:e></m:mr><m:mr><m:e>${r('0')}</m:e><m:e>${r('1')}</m:e></m:mr></m:m></m:e></m:d>`
  + `<m:func><m:fName>${r('sin', '<m:sty m:val="p"/>')}</m:fName><m:e>${r('θ')}</m:e></m:func><m:acc><m:accPr><m:chr m:val="̂"/></m:accPr><m:e>${r('y')}</m:e></m:acc>`
  + `${r('L', '<m:scr m:val="script"/>')}${r('R', '<m:scr m:val="double-struck"/><m:sty m:val="p"/>')}${r('v', '<m:sty m:val="b"/>')}<m:rad><m:deg>${r('3')}</m:deg><m:e>${r('x')}</m:e></m:rad>`);
const INLINE = `<w:p><w:r><w:t xml:space="preserve">The mean </w:t></w:r><m:oMath ${M}><m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>${r('x')}</m:e></m:bar>${r('=')}<m:f><m:num>${r('1')}</m:num><m:den>${r('n')}</m:den></m:f></m:oMath><w:r><w:t xml:space="preserve"> is in the line.</w:t></w:r></w:p>`;
const BODY = QUADRATIC + SUM + MIXED + INLINE;

const equation = (page, i) => page.locator('.docx-page .equation').nth(i);

test('equations are drawn as MathML: on their own line, or in the line of text', async ({ page }) => {
  await openInEditor(page, await buildDocx(BODY));
  await expect(page.locator('.docx-page .equation')).toHaveCount(4);
  await expect(page.locator('.docx-page .equation.display math[display="block"]')).toHaveCount(3);
  // The quadratic formula: a fraction over a square root, with a minus sign for Word's hyphen.
  const quadratic = equation(page, 0);
  await expect(quadratic.locator('mfrac msqrt msup')).toHaveCount(1);
  await expect(quadratic.locator('mfrac > mrow').first()).toHaveText('−b±b2−4ac');
  // A sum with its limits under and over it; an integral with them beside it; brackets round x.
  await expect(equation(page, 1).locator('munderover > mo').first()).toHaveText('∑');
  await expect(equation(page, 1).locator('msubsup > mo').first()).toHaveText('∫');
  await expect(equation(page, 1).locator('mrow > mo[fence="true"]')).toHaveText(['(', ')']);
  // A matrix in brackets, a function name upright, an accent, and Word's letter styles.
  const mixed = equation(page, 2);
  await expect(mixed.locator('mtable mtr')).toHaveCount(2);
  await expect(mixed.locator('mi', { hasText: 'sin' })).toHaveCount(1);
  await expect(mixed.locator('mover[accent="true"] > mo')).toHaveText('ˆ');
  await expect(mixed.locator('mi[mathvariant="normal"]')).toHaveText(['ℒ', 'ℝ', '𝐯']);
  await expect(mixed.locator('mroot')).toHaveCount(1);
  // In line: not a block of its own.
  const inline = equation(page, 3);
  await expect(inline).not.toHaveClass(/display/);
  await expect(inline.locator('math')).not.toHaveAttribute('display', 'block');
  await expect(inline.locator('mover > mo')).toHaveText('‾');
  await expect(page.locator('.docx-page p').last()).toContainText('The mean');
});

test('equations are written back exactly as they were; one LaTeX here can\'t write stays read-only', async ({ page }) => {
  const buffer = await buildDocx(BODY);
  await openInEditor(page, buffer);
  await equation(page, 3).click();
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Equation – double-click to edit it');
  await equation(page, 2).click(); // Word's script and double-struck letters
  await expect(page.locator('.statusbar [role=status]')).toHaveText('Read-only: Equation – kept as it is; edit it in Word');
  await page.locator('.docx-page p').last().click();
  await page.keyboard.press('End');
  await page.keyboard.type(' Edited.');
  const { xml } = await download(page);
  const original = (await documentXmlOf(buffer)).replace('is in the line.', 'is in the line. Edited.');
  expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, original));
});
