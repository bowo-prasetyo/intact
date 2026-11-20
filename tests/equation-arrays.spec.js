// Equations a maths teacher writes: matrices in their brackets, cases, working lined up at its =
// signs, and accents, typed as in LaTeX and written as Word's own (m:m in m:d, m:eqArr with its
// alignment points the &s in the text, m:acc, m:bar). An equation array is shown lined up at those
// points, whether it was typed here or Word wrote it.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor, download } from './helpers.js';

const P = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const r = (t) => `<m:r><m:t xml:space="preserve">${t}</m:t></m:r>`;

async function typeEquation(page, latex) {
  await page.locator('.docx-page > p').last().click();
  await page.getByRole('button', { name: /Insert a check box or date/ }).click();
  await page.getByRole('menuitem', { name: /Equation/ }).click();
  const eq = page.getByRole('dialog', { name: 'Equation' });
  await eq.getByLabel('Equation, as in LaTeX').fill(latex);
  return eq;
}

test('a matrix in brackets: written as Word’s matrix in a delimiter, shown as a table', async ({ page }) => {
  await openInEditor(page, await buildDocx(P('Matrices.') + P('')));
  const eq = await typeEquation(page, '\\begin{bmatrix} 1 & 2 \\\\ 3 & 4 \\end{bmatrix}');
  await expect(eq.locator('.maths mtable mtr')).toHaveCount(2);
  await eq.getByRole('button', { name: 'Insert' }).click();
  await expect(page.locator('.docx-page .equation mtable mtd')).toHaveCount(4);
  const { xml } = await download(page);
  expect(xml).toMatch(/<m:d><m:dPr><m:begChr m:val="\["\/><m:endChr m:val="\]"\/><\/m:dPr><m:e><m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="2"\/><m:mcJc m:val="center"\/><\/m:mcPr><\/m:mc><\/m:mcs><\/m:mPr><m:mr>/);
  expect(xml.match(/<m:mr>/g)).toHaveLength(2);
});

test('working lined up at its = signs: an equation array, its alignment points the &s', async ({ page }) => {
  await openInEditor(page, await buildDocx(P('Solve it.') + P('')));
  const eq = await typeEquation(page, '\\begin{aligned} 2x + 3 &= 11 \\\\ 2x &= 8 \\\\ x &= 4 \\end{aligned}');
  await eq.getByRole('button', { name: 'Insert' }).click();
  const rows = page.locator('.docx-page .equation mtable mtr');
  await expect(rows).toHaveCount(3);
  // Two columns: before the = (to the right), from it on (to the left); the = signs line up.
  await expect(rows.first().locator('mtd')).toHaveCount(2);
  await expect(rows.nth(1).locator('mtd').nth(1)).toHaveText('=8');
  const xs = await rows.evaluateAll((trs) => trs.map((tr) => tr.children[1].getBoundingClientRect().left));
  expect(Math.max(...xs) - Math.min(...xs)).toBeLessThan(1);
  const { xml } = await download(page);
  expect(xml.match(/<m:eqArr>/g)).toHaveLength(1);
  expect(xml.match(/<m:t>&amp;<\/m:t>/g)).toHaveLength(3);
});

test('cases and accents', async ({ page }) => {
  await openInEditor(page, await buildDocx(P('Cases.') + P('')));
  const eq = await typeEquation(page, 'f(x) = \\begin{cases} 1 & x > 0 \\\\ 0 & x \\le 0 \\end{cases} + \\vec{v} + \\overline{AB}');
  await eq.getByRole('button', { name: 'Insert' }).click();
  const { xml } = await download(page);
  expect(xml).toContain('<m:d><m:dPr><m:begChr m:val="{"/><m:endChr m:val=""/></m:dPr><m:e><m:eqArr>');
  expect(xml).toContain('<m:acc><m:accPr><m:chr m:val="⃗"/></m:accPr>');
  expect(xml).toContain('<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>');
});

test('what is wrong is said: an environment not closed, or not known', async ({ page }) => {
  await openInEditor(page, await buildDocx(P('Oops.') + P('')));
  const eq = await typeEquation(page, '\\begin{pmatrix} a & b');
  await expect(eq.locator('.error')).toHaveText('A \\begin{pmatrix} has no \\end{pmatrix} to close it');
  await eq.getByLabel('Equation, as in LaTeX').fill('\\begin{tabular} a \\end{tabular}');
  await expect(eq.locator('.error')).toHaveText('{tabular} isn’t one this editor knows: try pmatrix, bmatrix, cases or aligned');
  await expect(eq.getByRole('button', { name: 'Insert' })).toBeDisabled();
});

test('an equation array Word wrote is shown lined up at its alignment points', async ({ page }) => {
  const arr = `<w:p><m:oMathPara ${M}><m:oMath><m:eqArr><m:e>${r('y')}${r('&amp;=(x+1)^2')}</m:e><m:e>${r('&amp;=x^2+2x+1')}</m:e></m:eqArr></m:oMath></m:oMathPara></w:p>`;
  await openInEditor(page, await buildDocx(arr));
  const rows = page.locator('.docx-page .equation mtable mtr');
  await expect(rows).toHaveCount(2);
  await expect(rows.first().locator('mtd')).toHaveCount(2);
  const xs = await rows.evaluateAll((trs) => trs.map((tr) => tr.children[1].getBoundingClientRect().left));
  expect(Math.abs(xs[0] - xs[1])).toBeLessThan(1);
  // No & is shown.
  await expect(page.locator('.docx-page .equation')).not.toContainText('&');
});
