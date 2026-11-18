// Real documents (tests/fixtures/*.docx, see the README there), written by Word, LibreOffice, pandoc
// and python-docx, go through the core promises:
//  - everything is shown: every paragraph's text (in the body, headers, footers and notes) and every
//    picture is on the page;
//  - opening and downloading without edits changes nothing, and saving again changes nothing;
//  - typing in one place (a paragraph, a heading, a list item, a table cell, a header or footer, a
//    text box, a note) changes only that place: one block of its part, with the typed text in the formatting
//    around it, and every other part byte for byte (bar the "modified" date in docProps/core.xml);
//  - the saved package is sound: every XML part well formed, every part has a content type, every
//    relationship points at a part that is there;
//  - nothing goes wrong in the page while doing it (no errors in the console).
import { test, expect } from '@playwright/test';
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { openInEditor, download, canonicalize } from './helpers.js';

const DIR = path.join(import.meta.dirname, 'fixtures');
const FIXTURES = fs.readdirSync(DIR).filter((f) => f.endsWith('.docx')).sort();
const EDIT = ' [edited]';

/** The main document part's name (word/document.xml, or as the package's root relationship says). */
async function mainPart(zip) {
  const rels = await zip.file('_rels/.rels').async('string');
  const target = rels.match(/Type="[^"]*\/officeDocument"[^>]*Target="\/?([^"]+)"|Target="\/?([^"]+)"[^>]*Type="[^"]*\/officeDocument"/);
  return target[1] || target[2];
}

/** The text of each paragraph in an XML part, as Word shows it (deleted text and field codes aside). */
function paragraphTexts(page, xml) {
  return page.evaluate((text) => {
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    const out = [];
    for (const p of doc.getElementsByTagNameNS(W, 'p')) {
      if (p.getElementsByTagNameNS(W, 'p').length) continue; // a paragraph holding a text box's: count the inner ones
      let s = '';
      for (const t of p.querySelectorAll('*')) {
        if (t.namespaceURI !== W || !['t', 'tab'].includes(t.localName)) continue;
        let inner = false;
        for (let a = t.parentNode; a && a !== p; a = a.parentNode) {
          if (a.namespaceURI === W && a.localName === 'p') inner = true; // a text box's own paragraph
        }
        if (!inner) s += t.localName === 'tab' ? ' ' : t.textContent;
      }
      s = s.replace(/\s+/g, ' ').trim();
      if (s) out.push(s);
    }
    return out;
  }, xml);
}

/** Top-level blocks of a part that differ before and after (by meaning, see canonicalize). */
async function changedBlocks(page, beforeXml, afterXml) {
  const before = await canonicalize(page, beforeXml);
  const after = await canonicalize(page, afterXml);
  expect(after.length, 'no block is added or removed').toBe(before.length);
  return after
    .map((b, i) => ({ i, before: JSON.stringify(before[i]), after: JSON.stringify(b) }))
    .filter((c) => c.before !== c.after);
}

/** Where two strings differ, with some context: for failure messages. */
function difference(a, b) {
  let i = 0;
  while (i < a.length && a[i] === b[i]) i++;
  let j = 0;
  while (j < a.length - i && j < b.length - i && a[a.length - 1 - j] === b[b.length - 1 - j]) j++;
  return `before: …${a.slice(Math.max(0, i - 120), a.length - j + 40)}…\nafter:  …${b.slice(Math.max(0, i - 120), b.length - j + 40)}…`;
}

/** Every part but `except` comes back byte for byte. */
async function samePartsBut(original, saved, except) {
  for (const [name, file] of Object.entries(original.files)) {
    if (file.dir || except.includes(name)) continue;
    const before = Buffer.from(await file.async('uint8array'));
    const after = Buffer.from(await saved.file(name).async('uint8array'));
    expect(before.equals(after), `${name} is unchanged`).toBe(true);
  }
}

/** The saved package is sound: well-formed XML, content types, relationship targets. */
async function expectSoundPackage(page, zip) {
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const types = await zip.file('[Content_Types].xml').async('string');
  const defaults = [...types.matchAll(/<Default [^>]*Extension="([^"]+)"/g)].map((m) => m[1].toLowerCase());
  const overrides = [...types.matchAll(/<Override [^>]*PartName="([^"]+)"/g)].map((m) => decodeURI(m[1]).slice(1));
  for (const name of names) {
    if (name === '[Content_Types].xml') continue;
    const ext = name.split('.').pop().toLowerCase();
    expect(overrides.includes(name) || defaults.includes(ext), `${name} has a content type`).toBe(true);
    if (!/\.(xml|rels)$/i.test(name)) continue;
    const text = await zip.file(name).async('string');
    const wellFormed = await page.evaluate((t) => !new DOMParser().parseFromString(t, 'application/xml').getElementsByTagName('parsererror').length, text);
    expect(wellFormed, `${name} is well-formed XML`).toBe(true);
    if (!name.endsWith('.rels')) continue;
    const base = name.replace(/_rels\/[^/]*\.rels$/, '');
    for (const m of text.matchAll(/<Relationship [^>]*>/g)) {
      if (/TargetMode="External"/.test(m[0])) continue;
      const target = decodeURI(m[0].match(/Target="([^"]+)"/)[1]);
      const resolved = path.posix.normalize(target.startsWith('/') ? target.slice(1) : base + target);
      expect(names.includes(resolved), `${name}: ${target} is in the package`).toBe(true);
    }
  }
}

/** Collects errors raised in the page; checked at the end of each test. */
function watchErrors(page) {
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
  return errors;
}

/**
 * Puts the caret at the end of the paragraph `target` or, with `word`, selects the first `word` in
 * it. With a range (End would go to the end of a line, not of the paragraph).
 */
async function select(page, target, word = null) {
  await target.scrollIntoViewIfNeeded();
  await target.click();
  await target.evaluate((el, w) => {
    // Text the user can edit: its nearest contenteditable isn't "false" (a field, a picture…).
    const editable = (n) => n.parentElement.closest('[contenteditable]')?.getAttribute('contenteditable') !== 'false';
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (editable(n) ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT) });
    const range = document.createRange();
    let last = null;
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      if (w && n.data.includes(w)) {
        range.setStart(n, n.data.indexOf(w));
        range.setEnd(n, n.data.indexOf(w) + w.length);
        getSelection().removeAllRanges();
        getSelection().addRange(range);
        return;
      }
      if (n.data.trim()) last = n;
    }
    if (!last) throw new Error('no editable text in the paragraph');
    range.setStart(last, last.data.replace(/\s+$/, '').length);
    range.collapse(true);
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  }, word);
}

/** Types EDIT at the end of the paragraph `target`. */
async function typeAtEnd(page, target) {
  await select(page, target);
  await page.keyboard.type(EDIT);
}

/**
 * Exactly one block of `part` changed, and it is the old one with EDIT typed at the end of one of
 * its paragraphs in the formatting there (so removing EDIT gives the block back exactly).
 */
/** Whether the document has Track Changes on (w:trackRevisions in its settings). */
async function tracking(zip) {
  return /<w:trackRevisions(?![^>]*w:val="(?:0|false|off)")[^>]*\/>/.test(await zip.file('word/settings.xml')?.async('string') ?? '');
}

/**
 * Saved XML with the insertions made in the test unwrapped (and inserted paragraph marks dropped):
 * with Track Changes on in the document, typing is tracked, as in Word, and is then held to the same
 * rules as untracked typing. New insertions are those dated after the document's own. Returns the
 * XML and how many there were.
 */
function untracked(page, originalXml, savedXml) {
  return page.evaluate(([before, text]) => {
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const dates = new Set([...before.matchAll(/w:date="([^"]+)"/g)].map((m) => m[1]));
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    const ins = [...doc.getElementsByTagNameNS(W, 'ins')].filter((i) => !dates.has(i.getAttributeNS(W, 'date')));
    for (const i of ins) {
      if (i.parentNode.localName === 'rPr') i.remove(); // a paragraph mark inserted
      else i.replaceWith(...i.childNodes);
    }
    for (const rPr of doc.getElementsByTagNameNS(W, 'rPr')) if (!rPr.childNodes.length && rPr.parentNode.localName === 'pPr' && !before.includes('<w:rPr/>') && !before.includes('<w:rPr></w:rPr>')) rPr.remove();
    return { xml: new XMLSerializer().serializeToString(doc), count: ins.length };
  }, [originalXml, savedXml]);
}

async function expectOnlyEdit(page, original, saved, part, text) {
  let afterXml = await saved.file(part).async('string');
  if (await tracking(original)) {
    const { xml, count } = await untracked(page, await original.file(part).async('string'), afterXml);
    expect(count, 'with Track Changes on, the typing is a tracked insertion').toBeGreaterThan(0);
    afterXml = xml;
  }
  const changed = await changedBlocks(page, await original.file(part).async('string'), afterXml);
  expect(changed, `one block of ${part} changed`).toHaveLength(1);
  // (All of it: a text box is written twice, for Word and, as VML, for older readers; both change.)
  const after = changed[0].after.split(JSON.stringify(EDIT).slice(1, -1)).join('');
  expect(after === changed[0].before, `the typed text is all that changed, in the formatting around it:\n${difference(changed[0].before, after)}`).toBe(true);
  expect(changed[0].before, 'the block typed in').toContain(JSON.stringify(text.split(/\s+/).pop().slice(-10)).slice(1, -1));
  await samePartsBut(original, saved, [part, 'docProps/core.xml']);
  await expectSoundPackage(page, saved);
}

const page$ = (page) => page.locator('.docx-page');
/** Paragraphs in the text with words in them: top level, list items, headings, cells. */
const KINDS = {
  'a paragraph': (page) => page$(page).locator(':scope > p:not(.list)'),
  'a heading': (page) => page$(page).locator(':scope > :is(h1, h2, h3, h4, h5, h6)'),
  'a list item': (page) => page$(page).locator(':scope > .list'),
  'a table cell': (page) => page$(page).locator('td > p, th > p'),
};

for (const name of FIXTURES) {
  test.describe(name, () => {
    const buffer = fs.readFileSync(path.join(DIR, name));
    let errors;
    test.beforeEach(({ page }) => { errors = watchErrors(page); });
    test.afterEach(() => { expect(errors, 'no errors in the page').toEqual([]); });

    test('every paragraph and picture is shown', async ({ page }) => {
      const zip = await JSZip.loadAsync(buffer);
      await openInEditor(page, buffer, name);
      const main = await mainPart(zip);
      // What the page shows, tracked deletions aside (shown struck through, they aren't the text).
      const shown = (await page.locator('.docx-page').evaluate((el) => {
        const copy = el.cloneNode(true);
        for (const d of copy.querySelectorAll('.deleted')) d.remove();
        return copy.textContent;
      })).replace(/\s+/g, ' ');
      const parts = Object.keys(zip.files).filter((n) => n === main || /^word\/(header|footer)\d*\.xml$|^word\/(footnotes|endnotes)\.xml$/.test(n));
      for (const part of parts) {
        for (const text of await paragraphTexts(page, await zip.file(part).async('string'))) {
          expect(shown, `${part}: “${text}” is shown`).toContain(text);
        }
      }
      const pictures = ((await zip.file(main).async('string')).match(/<a:blip [^>]*r:embed=|<v:imagedata [^>]*r:id=/g) || []).length;
      expect(await page$(page).locator('img').filter({ hasNot: page.locator('xpath=ancestor::*[contains(@class, "docx-hf-band")]') }).count(), 'every picture in the text is shown').toBe(pictures);
    });

    test('downloading without edits changes nothing', async ({ page }) => {
      const original = await JSZip.loadAsync(buffer);
      await openInEditor(page, buffer, name);
      const main = await mainPart(original);
      const { zip } = await download(page);
      expect(await canonicalize(page, await zip.file(main).async('string'))).toEqual(await canonicalize(page, await original.file(main).async('string')));
      await samePartsBut(original, zip, [main]);
      await expectSoundPackage(page, zip);
    });

    test('a second save of a saved file changes nothing', async ({ page }) => {
      await openInEditor(page, buffer, name);
      const first = await download(page);
      await page.getByRole('button', { name: 'Close' }).click();
      await openInEditor(page, await first.zip.generateAsync({ type: 'nodebuffer' }), name);
      const second = await download(page);
      expect(await canonicalize(page, second.xml)).toEqual(await canonicalize(page, first.xml));
    });

    for (const [kind, locate] of Object.entries(KINDS)) {
      test(`typing in ${kind} changes only it`, async ({ page }) => {
        await openInEditor(page, buffer, name);
        const target = locate(page).filter({ hasText: /\S/, hasNot: page.locator('.textbox') }).last();
        test.skip(!(await target.count()), `${name} has no ${kind.slice(2)}`);
        const text = (await target.textContent()).trim();
        await typeAtEnd(page, target);
        const original = await JSZip.loadAsync(buffer);
        const { zip } = await download(page);
        await expectOnlyEdit(page, original, zip, await mainPart(original), text);
      });
    }

    test('typing and then undoing it gives back the file as it was, byte for byte', async ({ page }) => {
      await openInEditor(page, buffer, name);
      const target = KINDS['a paragraph'](page).filter({ hasText: /\S/, hasNot: page.locator('.textbox') }).last();
      await typeAtEnd(page, target);
      for (let i = 0; i < EDIT.length && (await target.textContent()).includes('[edited'); i++) await page.keyboard.press('ControlOrMeta+z');
      await expect(target).not.toContainText('[edited');
      const original = await JSZip.loadAsync(buffer);
      const { zip } = await download(page);
      await samePartsBut(original, zip, []);
    });

    test('Enter at the end of a paragraph adds one after it, and changes nothing else', async ({ page }) => {
      await openInEditor(page, buffer, name);
      const target = KINDS['a paragraph'](page).filter({ hasText: /\S/, hasNot: page.locator('.textbox, .raw-inline') }).last();
      test.skip(!(await target.count()), `${name} has no plain paragraph`);
      await select(page, target);
      await page.keyboard.press('Enter');
      await page.keyboard.type('New paragraph');
      const original = await JSZip.loadAsync(buffer);
      const main = await mainPart(original);
      const { zip } = await download(page);
      const originalXml = await original.file(main).async('string');
      let savedXml = await zip.file(main).async('string');
      if (await tracking(original)) savedXml = (await untracked(page, originalXml, savedXml)).xml;
      const before = await canonicalize(page, originalXml);
      const after = await canonicalize(page, savedXml);
      expect(after.length, 'one block more').toBe(before.length + 1);
      const at = after.findIndex((b) => JSON.stringify(b).includes('New paragraph'));
      expect([...after.slice(0, at), ...after.slice(at + 1)].map((b) => JSON.stringify(b)), 'the others as they were').toEqual(before.map((b) => JSON.stringify(b)));
      await samePartsBut(original, zip, [main, 'docProps/core.xml']);
      await expectSoundPackage(page, zip);
    });

    test('making a word bold changes only its paragraph', async ({ page }) => {
      await openInEditor(page, buffer, name);
      const target = KINDS['a paragraph'](page).filter({ hasText: /[A-Za-z]{4,}/, hasNot: page.locator('.textbox, strong') }).last();
      test.skip(!(await target.count()), `${name} has no plain paragraph`);
      const word = (await target.textContent()).match(/[A-Za-z]{4,}/)[0];
      await select(page, target, word);
      await page.keyboard.press('ControlOrMeta+b');
      const original = await JSZip.loadAsync(buffer);
      const main = await mainPart(original);
      const { zip } = await download(page);
      const changed = await changedBlocks(page, await original.file(main).async('string'), await zip.file(main).async('string'));
      expect(changed, 'one block changed').toHaveLength(1);
      expect(changed[0].after).toContain('w:b');
      await samePartsBut(original, zip, [main, 'docProps/core.xml']);
      await expectSoundPackage(page, zip);
    });

    for (const kind of ['header', 'footer']) {
      test(`typing in a ${kind} changes only its part`, async ({ page }) => {
        await openInEditor(page, buffer, name);
        const target = page.locator(`.docx-hf-band.${kind} .docx-hf-part:visible > p`).filter({ hasText: /\S/, hasNot: page.locator('.textbox') }).first();
        test.skip(!(await target.count()), `${name} shows no ${kind} with text`);
        const text = (await target.textContent()).trim();
        const host = target.locator('xpath=ancestor::*[@data-part][1]');
        const partName = await host.getAttribute('data-part');
        await typeAtEnd(page, target);
        const original = await JSZip.loadAsync(buffer);
        const { zip } = await download(page);
        await expectOnlyEdit(page, original, zip, partName, text);
      });
    }

    test('typing in a text box changes only its part', async ({ page }) => {
      await openInEditor(page, buffer, name);
      const target = page.locator('.docx-page .textbox p').filter({ hasText: /\S/ }).first();
      test.skip(!(await target.count()), `${name} has no text box`);
      const text = (await target.textContent()).trim();
      const original = await JSZip.loadAsync(buffer);
      const partName = await target.evaluate((el) => el.closest('[data-part]')?.dataset.part ?? null) ?? await mainPart(original);
      await typeAtEnd(page, target);
      const { zip } = await download(page);
      await expectOnlyEdit(page, original, zip, partName, text);
    });

    test('typing in a note changes only that note', async ({ page }) => {
      await openInEditor(page, buffer, name);
      const target = page.locator('.docx-notes .docx-note p').filter({ hasText: /\S/ }).first();
      test.skip(!(await target.count()), `${name} has no notes`);
      const kind = await target.locator('xpath=ancestor::*[@data-notes][1]').getAttribute('data-notes');
      const text = (await target.textContent()).trim();
      await typeAtEnd(page, target);
      const original = await JSZip.loadAsync(buffer);
      const { zip } = await download(page);
      await expectOnlyEdit(page, original, zip, `word/${kind}s.xml`, text);
    });
  });
}
