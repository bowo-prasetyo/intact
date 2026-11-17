// Typographic quotes, dashes and ellipses as you type, as Word's AutoFormat As You Type: " and ' become
// curly quotes in the text's language, "--" between words an em dash, " - " (or " -- ") an en dash,
// and "..." an ellipsis. Pure functions of the text before the caret and the character typed.

// [opening double, closing double, opening single, closing single (and apostrophe)], by language.
const QUOTES = {
  en: ['“', '”', '‘', '’'],
  low: ['„', '“', '‚', '‘'], // German, Czech, Slovak, Slovenian, Icelandic, Estonian, Lithuanian, Latvian, Bulgarian
  lowRight: ['„', '”', '‚', '’'], // Polish, Romanian, Hungarian, Croatian, Serbian (Latin)
  fr: ['« ', ' »', '‘', '’'], // with no-break spaces inside, as Word types them in French
  guillemets: ['«', '»', '‘', '’'], // Spanish, Italian, Portuguese, Russian, Ukrainian, Norwegian, Greek
  sv: ['”', '”', '’', '’'], // Swedish, Finnish
};
const STYLE_OF = {
  de: 'low', cs: 'low', sk: 'low', sl: 'low', is: 'low', et: 'low', lt: 'low', lv: 'low', bg: 'low',
  pl: 'lowRight', ro: 'lowRight', hu: 'lowRight', hr: 'lowRight', sr: 'lowRight',
  fr: 'fr', es: 'guillemets', it: 'guillemets', pt: 'guillemets', ru: 'guillemets', uk: 'guillemets', nb: 'guillemets', no: 'guillemets', el: 'guillemets',
  sv: 'sv', fi: 'sv',
};

/** The quote marks for a language tag ("en-GB", "de"). */
export function quotesFor(lang) {
  const base = (lang || 'en').toLowerCase().split('-')[0];
  // Swiss and Liechtenstein German use guillemets; English and anything else, “curly” quotes.
  if (/^de-(ch|li)$/i.test(lang || '')) return QUOTES.guillemets;
  return QUOTES[STYLE_OF[base]] || QUOTES.en;
}

/** True if a quote typed after `prev` (the character before the caret, '' at the start) opens. */
function opens(prev) {
  return !prev || /[\s([{<“‘«„‚—–/-]/u.test(prev);
}

/** The quote to type for `typed` (" or ') after the text `before`, in language `lang`; null if not a quote. */
export function smartQuote(before, typed, lang) {
  if (typed !== '"' && typed !== "'") return null;
  const [od, cd, os, cs] = quotesFor(lang);
  const open = opens([...before].pop() || '');
  if (typed === '"') return open ? od : cd;
  return open ? os : cs;
}

/**
 * A replacement for the end of `before` when `typed` is typed: { remove (characters before the
 * caret to replace), text (what to type instead, with `typed`) }, or null.
 *  - "word--" then a letter: an em dash between the words ("word—w");
 *  - "word -" or "word --" then a space: an en dash between spaces ("word – ");
 *  - ".." then ".": an ellipsis.
 */
export function smartReplacement(before, typed) {
  if (typed === '.' && before.endsWith('..') && !before.endsWith('...')) return { remove: 2, text: '…' };
  if (typed === ' ') {
    const m = /(\S) (--?)$/.exec(before);
    if (m) return { remove: m[2].length, text: '– ' };
    return null;
  }
  if (/[\p{L}\p{N}]/u.test(typed) && /[\p{L}\p{N}.,;:!?)\]'"’”]--$/u.test(before)) return { remove: 2, text: `—${typed}` };
  return null;
}
