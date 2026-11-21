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
