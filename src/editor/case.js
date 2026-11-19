// Change Case, as Word's Aa button and Shift+F3: Sentence case., lowercase, UPPERCASE, Capitalize Each
// Word and tOGGLE cASE. Text is changed character by character, so a selection that spans several
// runs keeps each run's formatting: each piece is changed in place and keeps its length in pieces.

export const CASES = [
  { id: 'sentence', label: 'Sentence case.' },
  { id: 'lower', label: 'lowercase' },
  { id: 'upper', label: 'UPPERCASE' },
  { id: 'title', label: 'Capitalize Each Word' },
  { id: 'toggle', label: 'tOGGLE cASE' },
];

const LETTER = /[\p{L}\p{N}]/u;
const IN_WORD = /[\p{L}\p{N}'’­]/u; // an apostrophe doesn't start a new word: "don't", "O’Brien" stay one
const SENTENCE_END = /[.!?…]/u;

/**
 * Shift+F3's next case for `text`, as Word cycles it: UPPERCASE, then lowercase, then Capitalize Each
 * Word, then UPPERCASE again. Text with no letters stays as it is (null).
 */
export function nextCase(text, lang) {
  if (!/\p{L}/u.test(text)) return null;
  if (text === text.toLocaleUpperCase(lang)) return 'lower';
  if (text === text.toLocaleLowerCase(lang)) return 'title';
  return 'upper';
}

/**
 * `pieces` (strings: the selected text's runs, in order) in case `kind`. `before`: the text before
 * them in their paragraph, which says whether the first is at the start of a word or sentence.
 * Returns the changed pieces, one for each.
 */
export function changeCase(pieces, kind, before = '', lang = undefined) {
  const upper = (ch) => ch.toLocaleUpperCase(lang);
  const lower = (ch) => ch.toLocaleLowerCase(lang);
  const tail = [...before];
  let prev = tail[tail.length - 1] || '';
  // A sentence starts at the start of the paragraph, or after . ! ? (and spaces, quotes or brackets).
  let sentenceStart = !before.trim() || SENTENCE_END.test(before.replace(/[\s"'“”‘’)\]»]+$/u, '').slice(-1));
  return pieces.map((piece) => {
    let out = '';
    for (const ch of piece) {
      const wordStart = !IN_WORD.test(prev);
      if (kind === 'upper') out += upper(ch);
      else if (kind === 'lower') out += lower(ch);
      else if (kind === 'toggle') out += ch === upper(ch) ? lower(ch) : upper(ch);
      else if (kind === 'title') out += wordStart ? upper(ch) : lower(ch);
      else if (kind === 'sentence') out += sentenceStart && LETTER.test(ch) ? upper(ch) : lower(ch);
      else out += ch;
      if (LETTER.test(ch)) sentenceStart = false;
      else if (SENTENCE_END.test(ch)) sentenceStart = true;
      prev = ch;
    }
    return out;
  });
}

/** The label of a case, for messages. */
export const caseLabel = (id) => CASES.find((c) => c.id === id)?.label || id;
