// Real documents (tests/fixtures/*.docx, see the README there) go through the core promises:
// opening and downloading without edits changes nothing, saving again changes nothing, and editing
// one paragraph changes only that paragraph.
import { test, expect } from '@playwright/test';
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { openInEditor, download, canonicalize } from './helpers.js';

const DIR = path.join(import.meta.dirname, 'fixtures');
const FIXTURES = fs.readdirSync(DIR).filter((f) => f.endsWith('.docx')).sort();

for (const name of FIXTURES) {
  test.describe(name, () => {
    const buffer = fs.readFileSync(path.join(DIR, name));
    const xmlOf = async (zip) => zip.file(/document\.xml$/).find((f) => !f.name.includes('/_rels/')).async('string');

    test('downloading without edits changes nothing', async ({ page }) => {
      const original = await JSZip.loadAsync(buffer);
      await openInEditor(page, buffer, name);
      const { zip, xml } = await download(page);
      expect(await canonicalize(page, xml)).toEqual(await canonicalize(page, await xmlOf(original)));
      // Every other part comes back byte for byte.
      for (const [part, file] of Object.entries(original.files)) {
        if (file.dir || part.endsWith('word/document.xml')) continue;
        const before = Buffer.from(await file.async('uint8array'));
        const after = Buffer.from(await zip.file(part).async('uint8array'));
        expect(before.equals(after), `${part} is unchanged`).toBe(true);
      }
    });

    test('a second save of a saved file changes nothing', async ({ page }) => {
      await openInEditor(page, buffer, name);
      const first = await download(page);
      await page.getByRole('button', { name: 'Close' }).click();
      await openInEditor(page, await first.zip.generateAsync({ type: 'nodebuffer' }), name);
      const second = await download(page);
      expect(await canonicalize(page, second.xml)).toEqual(await canonicalize(page, first.xml));
    });

    test('editing one paragraph changes only that paragraph', async ({ page }) => {
      await openInEditor(page, buffer, name);
      // The last top-level paragraph with text.
      const target = page.locator('.docx-page > :is(p, h1, h2, h3, h4, h5, h6)').filter({ hasText: /\S/ }).last();
      const text = (await target.textContent()).trim();
      await target.click();
      await page.keyboard.press('End');
      await page.keyboard.type(' [edited]');
      const { xml } = await download(page);
      const before = await canonicalize(page, await xmlOf(await JSZip.loadAsync(buffer)));
      const after = await canonicalize(page, xml);
      expect(after.length).toBe(before.length);
      const changed = after.map((b, i) => (JSON.stringify(b) === JSON.stringify(before[i]) ? null : i)).filter((i) => i !== null);
      expect(changed).toHaveLength(1);
      expect(JSON.stringify(after[changed[0]])).toContain('[edited]');
      expect(JSON.stringify(before[changed[0]])).toContain(JSON.stringify(text.slice(-12)).slice(1, -1));
    });
  });
}
