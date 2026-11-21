// AutoCorrect entries, as Word's AutoCorrect Options › "Replace text as you type": a word typed then
// a space or punctuation is replaced by its entry ("cllr" → "Councillor", a long name from its
// initials), in the case it was typed in ("Cllr" → "Councillor", "CLLR" → "COUNCILLOR"). Your own
// entries are kept in this browser, as your colours are; common English typos ("teh" → "the") are
// corrected too, in English text only (Indonesian's "teh" is tea). Nothing leaves the device.

const KEY = 'intact.autocorrect';

// Common English typos Word corrects, those that are never a word themselves.
export const TYPOS = [
  ['teh', 'the'], ['hte', 'the'], ['adn', 'and'], ['nad', 'and'], ['taht', 'that'], ['thier', 'their'], ['recieve', 'receive'],
  ['recieved', 'received'], ['beleive', 'believe'], ['acheive', 'achieve'], ['seperate', 'separate'], ['definately', 'definitely'],
  ['occured', 'occurred'], ['occurence', 'occurrence'], ['untill', 'until'], ['wich', 'which'], ['becuase', 'because'],
  ['goverment', 'government'], ['accomodate', 'accommodate'], ['tommorow', 'tomorrow'], ['tomorow', 'tomorrow'], ['freind', 'friend'],
  ['wierd', 'weird'], ['alot', 'a lot'], ['dont', 'don’t'], ['doesnt', 'doesn’t'], ['didnt', 'didn’t'], ['isnt', 'isn’t'],
  ['wasnt', 'wasn’t'], ['couldnt', 'couldn’t'], ['wouldnt', 'wouldn’t'], ['shouldnt', 'shouldn’t'], ['thats', 'that’s'],
  ['im', 'I’m'], ['ive', 'I’ve'], ['i', 'I'], ['independant', 'independent'], ['neccessary', 'necessary'], ['publically', 'publicly'],
  ['concensus', 'consensus'], ['comittee', 'committee'], ['commitee', 'committee'], ['parliment', 'parliament'], ['buisness', 'business'],
  ['enviroment', 'environment'], ['existance', 'existence'], ['foriegn', 'foreign'], ['greatful', 'grateful'], ['harrass', 'harass'],
  ['millenium', 'millennium'], ['mispell', 'misspell'], ['noticable', 'noticeable'], ['posession', 'possession'], ['realy', 'really'],
  ['refered', 'referred'], ['relevent', 'relevant'], ['succesful', 'successful'], ['truely', 'truly'], ['wether', 'whether'],
];
const TYPO_MAP = new Map(TYPOS);

/** The settings, as kept in this browser: { on, typos, entries: [[from, to]] }. */
export function loadAutoCorrect() {
  const fallback = { on: true, typos: true, entries: [] };
  try {
    const v = JSON.parse(localStorage.getItem(KEY) || 'null');
    if (!v || !Array.isArray(v.entries)) return fallback;
    return { on: v.on !== false, typos: v.typos !== false, entries: v.entries.filter((e) => Array.isArray(e) && typeof e[0] === 'string' && typeof e[1] === 'string') };
  } catch {
    return fallback;
  }
}

/** Keeps the settings in this browser; false where it refuses storage (this session only, then). */
export function saveAutoCorrect(settings) {
  try {
    localStorage.setItem(KEY, JSON.stringify(settings));
    return true;
  } catch {
    return false;
  }
}

/** Why `from` can't be an entry's text to replace, or '' if it can. */
export function entryProblem(from, to) {
  if (!from) return 'Type the text to replace';
  if (/\s/.test(from)) return 'The text to replace is one word, without spaces';
  if (!to.trim()) return 'Type what to replace it with';
  return '';
}

/** `text` in the case `typed` was typed in: as typed in lower case, a capital first, or all capitals. */
export function inCase(text, typed) {
  const letters = typed.replace(/[^\p{L}]/gu, '');
  if (letters.length > 1 && letters === letters.toUpperCase() && letters !== letters.toLowerCase()) return text.toUpperCase();
  if (letters && letters[0] === letters[0].toUpperCase() && letters[0] !== letters[0].toLowerCase()) return text[0].toUpperCase() + text.slice(1);
  return text;
}

/**
 * The replacement for the word just typed: `before` is the paragraph's text up to `typed`, the space
 * or punctuation being typed ('\n' for Enter), `english` whether the text there is English. Returns
 * { from (as typed), to } or null. Your entries are matched whatever their case; typos only as typed
 * in lower case or with a capital first ("Teh"), never in capitals (an abbreviation), and only a
 * word on its own: not after a bracket or a quote, and "i" only before a space or an apostrophe, so
 * "(i)" (a list's numeral) and "i.e." stay as typed.
 */
export function autoCorrectWord(before, settings, english, typed = ' ') {
  if (!settings.on) return null;
  const m = /(?:^|[\s([{“‘«"'])([^\s([{“‘«"']+)$/u.exec(before);
  if (!m) return null;
  const word = m[1];
  const key = word.toLowerCase();
  const own = settings.entries.find(([from]) => from.toLowerCase() === key);
  if (own) return { from: word, to: own[0] === own[0].toLowerCase() ? inCase(own[1], word) : own[1] };
  const alone = before.length === word.length || /\s/u.test(before[before.length - word.length - 1]);
  if (!settings.typos || !english || !TYPO_MAP.has(key) || word === word.toUpperCase() || !alone) return null;
  if (key === 'i') return /^[\s'’]$/u.test(typed) ? { from: word, to: 'I' } : null;
  return { from: word, to: inCase(TYPO_MAP.get(key), word) };
}

// ---------------------------------------------------------------- moving entries between browsers

/** Most entries a list may bring in, and the largest file read for one. */
export const MAX_ENTRIES = 5000;
export const MAX_FILE = 1024 * 1024;

/**
 * Your entries as a CSV file's text, to take to another browser or device (or keep, or edit in a
 * spreadsheet): a header row "Replace,With", then an entry a row, in UTF-8 with a byte order mark,
 * as Excel's "CSV UTF-8" saves it, so Excel opens it with its accents.
 */
export function entriesToCsv(entries) {
  const field = (s) => (/[",;\t\r\n]|^\s|\s$/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  return `﻿${['Replace,With', ...entries.map(([f, w]) => `${field(f)},${field(w)}`)].join('\r\n')}\r\n`;
}

/**
 * The entries in a list's text (see entriesToCsv; also a list saved from a spreadsheet as CSV, with
 * commas, semicolons or tabs, with or without its header row): { entries: [[from, to]], skipped (rows
 * that aren't an entry: a word to replace that has spaces, or nothing to replace it with) }.
 * `parseCsv` is mail merge's (docx/merge.js), passed in to keep this file small.
 */
export function entriesFromCsv(text, parseCsv) {
  text = text.replace(/^﻿/, '');
  const first = text.split(/\r\n|\n|\r/, 1)[0];
  // A list without the header row: one is put on, so its first row is read as an entry.
  if (!/^\s*"?\s*(replace|ganti|from|text)\b/i.test(first)) {
    const separator = first.includes('\t') ? '\t' : first.split(';').length > first.split(',').length ? ';' : ',';
    text = `Replace${separator}With\n${text}`;
  }
  const entries = [];
  let skipped = 0;
  for (const [from = '', to = ''] of parseCsv(text).rows) {
    const f = from.trim();
    if (entryProblem(f, to)) skipped++;
    else entries.push([f, to.trim()]);
  }
  return { entries, skipped };
}

/**
 * `incoming` entries put into `current`: an entry for a word already there replaces it (the list
 * brought in wins, as its last row does for a word twice in it). { entries (sorted), added, changed }.
 */
export function mergeEntries(current, incoming) {
  const byKey = new Map(current.map((e) => [e[0].toLowerCase(), [...e]]));
  const before = new Map(current.map(([f, w]) => [f.toLowerCase(), w]));
  for (const [f, w] of incoming) byKey.set(f.toLowerCase(), [f, w]);
  let added = 0;
  let changed = 0;
  for (const [key, [, w]] of byKey) {
    if (!before.has(key)) added++;
    else if (before.get(key) !== w) changed++;
  }
  return { entries: [...byKey.values()].sort((a, b) => a[0].localeCompare(b[0])), added, changed };
}
