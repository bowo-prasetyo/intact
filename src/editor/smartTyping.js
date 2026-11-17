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

const ROMAN = { i: 1, ii: 2, iii: 3, iv: 4, v: 5, vi: 6, vii: 7, viii: 8, ix: 9, x: 10 };

/**
 * The list a paragraph starts when a space is typed after `before` (all of its text before the
 * caret), as Word's AutoFormat "Automatic bulleted lists" and "Automatic numbered lists" do, or null:
 *  - "*" or "•" a bulleted list, "-" a list with dash bullets;
 *  - "1." "1)" "(1)", "a." "a)" "(a)", "A." "A)", "i." "i)" "(i)", "I." "I)": a numbered list in that
 *    style, starting from the number typed ("3." starts at 3, "c)" at c; i, ii, iv… are Roman).
 * Returns { kind: 'bullet' | 'number', first } with `first` as numbering.addList takes it (null for
 * Word's default bullets).
 */
export function listTrigger(before) {
  if (before === '*' || before === '•') return { kind: 'bullet', first: null };
  if (before === '-') return { kind: 'bullet', first: { char: '-' } };
  const m = /^(\()?([0-9]{1,3}|[a-zA-Z]|[ivx]{1,4}|[IVX]{1,4})([.)])$/.exec(before);
  if (!m) return null;
  const [, open, mark, close] = m;
  if (open && close !== ')') return null; // "(1." isn't a list
  const text = open ? '(%1)' : `%1${close}`;
  if (/^[0-9]+$/.test(mark)) return { kind: 'number', first: { fmt: 'decimal', text, start: Number(mark) } };
  // One letter is a lettered list, except i (Roman, as in Word); ii, iv, ix… are Roman numerals.
  const lower = mark.toLowerCase();
  if (mark.length === 1 && lower !== 'i') return letterList(mark, text);
  const roman = ROMAN[lower];
  if (!roman) return null;
  return { kind: 'number', first: { fmt: mark === lower ? 'lowerRoman' : 'upperRoman', text, start: roman } };
}

function letterList(mark, text) {
  const lower = mark.toLowerCase();
  return { kind: 'number', first: { fmt: mark === lower ? 'lowerLetter' : 'upperLetter', text, start: lower.charCodeAt(0) - 96 } };
}

// Word's AutoCorrect entries for symbols, replaced as the last character is typed.
const AUTOCORRECT = [
  ['(c)', '©'], ['(r)', '®'], ['(tm)', '™'], ['(e)', '€'],
  ['-->', '→'], ['<--', '←'], ['==>', '⇒'], ['<==', '⇐'], ['<=>', '⇔'],
];

/**
 * Word's AutoCorrect for symbols: { from (the text typed, ending with `typed`), to }, when typing
 * `typed` after `before` (the paragraph's text before the caret) completes one, else null. "(c)",
 * "(r)" and "(e)", also in capitals, aren't replaced where they are an item's letter: at the start of
 * the paragraph, or after the letter before them, as in "(a) one, (b) two and (c) three".
 */
export function autoCorrect(before, typed) {
  const text = (before + typed).toLowerCase();
  const hit = AUTOCORRECT.find(([from]) => text.endsWith(from));
  if (!hit) return null;
  const [from, to] = hit;
  if (/^\([a-z]\)$/.test(from)) {
    // A letter in brackets at the start of the paragraph, or after the letter before it ("(b)"
    // before "(c)", "(d)" before "(e)"), is an item's letter, not a symbol.
    const lead = before.slice(0, before.length - (from.length - 1));
    const previous = String.fromCharCode(from.charCodeAt(1) - 1);
    if (!lead.trim() || lead.toLowerCase().includes(`(${previous})`)) return null;
  }
  return { from: (before + typed).slice(-from.length), to };
}
