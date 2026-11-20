// Read Aloud, as Word's Review > Read Aloud (Ctrl+Alt+Space): from the caret's paragraph, or the
// selected text, paragraph by paragraph, with the paragraph and word highlighted. Only voices on
// this device are used (never online ones, which send the text away); hidden text (a test's
// answers) isn't read; Esc stops. The browser's speech is replaced by a fake here, which records
// what it is asked to say.
import { test, expect } from '@playwright/test';
import { buildDocx, openInEditor } from './helpers.js';

const P = (t, rPr = '') => `<w:p><w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const BODY = P('Read me first.') + P('Then this one.')
  + '<w:p><w:r><w:t xml:space="preserve">The answer is </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>secret</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>';

/** A fake speech synthesis with `voices`; what it is asked to say is in window.spoken. */
function fakeSpeech(page, voices) {
  return page.addInitScript((list) => {
    window.spoken = [];
    window.SpeechSynthesisUtterance = class { constructor(text) { this.text = text; } };
    Object.defineProperty(window, 'speechSynthesis', {
      configurable: true,
      value: {
        getVoices: () => list,
        speak(u) { window.spoken.push({ text: u.text, voice: u.voice?.name, lang: u.lang }); window.utterance = u; },
        cancel() { window.cancelled = (window.cancelled || 0) + 1; },
        pause() { window.paused = true; },
        resume() { window.paused = false; },
        addEventListener() {},
      },
    });
  }, voices);
}

const ON_DEVICE = { name: 'Local English', lang: 'en-GB', localService: true, default: false };
const ONLINE = { name: 'Google UK English', lang: 'en-GB', localService: false, default: true };
const spoken = (page) => page.evaluate(() => window.spoken);
const status = (page) => page.locator('.statusbar [role=status]');

test('reads from the caret with a voice on this device, paragraph by paragraph, highlighted; Esc stops', async ({ page }) => {
  await fakeSpeech(page, [ONLINE, ON_DEVICE]);
  await openInEditor(page, await buildDocx(BODY));
  await page.locator('.docx-page p', { hasText: 'Then this one.' }).click();
  await page.getByRole('button', { name: 'Read aloud' }).click();
  await expect.poll(() => spoken(page)).toEqual([{ text: 'Then this one.', voice: 'Local English', lang: 'en-GB' }]);
  await expect(status(page)).toHaveText('Reading aloud from the caret: Esc or Stop to stop');
  await expect(page.getByRole('button', { name: 'Stop reading aloud' })).toHaveAttribute('aria-pressed', 'true');
  expect(await page.evaluate(() => CSS.highlights.get('read-aloud')?.size)).toBe(1);
  // A word as it is spoken.
  await page.evaluate(() => window.utterance.onboundary({ name: 'word', charIndex: 5, charLength: 4 }));
  expect(await page.evaluate(() => [...CSS.highlights.get('read-aloud-word')][0].toString())).toBe('this');

  // The next paragraph when it ends; hidden text isn't read.
  await page.evaluate(() => window.utterance.onend());
  await expect.poll(async () => (await spoken(page)).map((s) => s.text)).toEqual(['Then this one.', 'The answer is .']);

  // Pause and resume.
  await page.getByRole('button', { name: 'Pause reading' }).click();
  expect(await page.evaluate(() => window.paused)).toBe(true);
  await page.getByRole('button', { name: 'Resume reading' }).click();
  expect(await page.evaluate(() => window.paused)).toBe(false);

  // Esc stops: the highlights go, and the page is as it was.
  await page.locator('.docx-page').focus();
  await page.keyboard.press('Escape');
  await expect(page.getByRole('button', { name: 'Read aloud' })).toHaveAttribute('aria-pressed', 'false');
  expect(await page.evaluate(() => CSS.highlights.has('read-aloud'))).toBe(false);
  await expect(page.locator('.docx-page p').first()).toHaveText('Read me first.');
});

test('the selected text only, and Ctrl+Alt+Space', async ({ page }) => {
  await fakeSpeech(page, [ON_DEVICE]);
  await openInEditor(page, await buildDocx(BODY));
  await page.evaluate(() => {
    const ps = [...document.querySelectorAll('.docx-page > p')];
    document.querySelector('.docx-page').focus();
    getSelection().setBaseAndExtent(ps[0].firstChild, 0, ps[1].firstChild, 4);
  });
  await page.keyboard.press('Control+Alt+Space');
  await expect.poll(() => spoken(page)).toHaveLength(1);
  await page.evaluate(() => window.utterance.onend());
  await page.evaluate(() => window.utterance.onend());
  await expect.poll(async () => (await spoken(page)).map((s) => s.text)).toEqual(['Read me first.', 'Then this one.']);
  await expect(page.getByRole('button', { name: 'Read aloud' })).toHaveAttribute('aria-pressed', 'false');
});

test('only online voices: nothing is read, and it says why', async ({ page }) => {
  await fakeSpeech(page, [ONLINE]);
  await openInEditor(page, await buildDocx(BODY));
  await page.locator('.docx-page p').first().click();
  await page.getByRole('button', { name: 'Read aloud' }).click();
  await expect(status(page)).toHaveText('No voice on this device to read with: add one in the system’s speech settings. Intact reads only with voices on this device, so the text isn’t sent anywhere');
  expect(await spoken(page)).toEqual([]);
});

test('text in a language no voice on this device reads: it says which', async ({ page }) => {
  await fakeSpeech(page, [ON_DEVICE]);
  await openInEditor(page, await buildDocx(P('Selamat pagi.', '<w:lang w:val="id-ID"/>')));
  await page.locator('.docx-page p').first().click();
  await page.getByRole('button', { name: 'Read aloud' }).click();
  await expect(status(page)).toHaveText('No voice on this device reads Indonesian: add one in the system’s speech settings. Intact reads only with voices on this device, so the text isn’t sent anywhere');
  expect(await spoken(page)).toEqual([]);
});
