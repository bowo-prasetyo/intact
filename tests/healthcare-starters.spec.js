import { test, expect } from '@playwright/test';

const TEMPLATES = [
  ['Healthcare policy', '1. Purpose'],
  ['Standard operating procedure', '3. Procedure'],
  ['Patient administration form', 'Registration details'],
  ['Discharge administration checklist', 'Completion checklist'],
  ['Incident report', 'Factual account'],
  ['Patient complaint record', 'Complaint summary'],
  ['Healthcare quality audit', 'Findings'],
  ['Corrective action record (CAPA)', 'Cause analysis'],
];

test('the start page groups healthcare administration templates separately', async ({ page }) => {
  await page.goto('./');
  const group = page.getByRole('region', { name: 'Healthcare administration' });
  await expect(group).toBeVisible();
  await expect(group.getByRole('button')).toHaveCount(TEMPLATES.length);
  await expect(group).toContainText('Administrative templates only');
  await expect(group).toContainText('privacy rules');
});

for (const [title, expectedSection] of TEMPLATES) {
  test(`${title} opens as an editable Word document with administrative safeguards`, async ({ page }) => {
    await page.goto('./');
    const group = page.getByRole('region', { name: 'Healthcare administration' });
    await group.getByRole('button', { name: new RegExp(`^${title.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}:`) }).click();
    const document = page.locator('.docx-page');
    await expect(document.first()).toBeVisible();
    await expect(document).toContainText(title);
    await expect(document).toContainText(expectedSection);
    await expect(document).toContainText('Before approval or sharing');
    await expect(document).toContainText('Inspect document');
    await expect(document).toContainText('not clinical guidance or a compliance certification');
  });
}
