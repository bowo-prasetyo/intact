// Insert Symbol, as Word's Ω Symbol: the characters people look for in letters, minutes, invoices and
// forms, by group, with their names to search by, and the ones used last (remembered in this
// browser, as Word's "recently used symbols"). Any character can also be inserted by its Unicode
// code (Word's "Character code", or Alt+X).

const group = (label, list) => ({ label, symbols: list.map(([char, name]) => ({ char, name })) });

export const SYMBOL_GROUPS = [
  group('Common', [
    ['©', 'Copyright'], ['®', 'Registered'], ['™', 'Trade mark'], ['°', 'Degree'], ['±', 'Plus-minus'],
    ['×', 'Multiplication'], ['÷', 'Division'], ['€', 'Euro'], ['£', 'Pound'], ['¥', 'Yen'],
    ['§', 'Section'], ['¶', 'Pilcrow (paragraph)'], ['–', 'En dash'], ['—', 'Em dash'], ['…', 'Ellipsis'],
    ['•', 'Bullet'], ['✓', 'Check mark'], ['✗', 'Ballot X'], ['☐', 'Ballot box'], ['☒', 'Ballot box with X'],
    ['→', 'Right arrow'], ['½', 'One half'], ['¼', 'One quarter'], ['¾', 'Three quarters'],
  ]),
  group('Punctuation', [
    ['–', 'En dash'], ['—', 'Em dash'], ['‑', 'Non-breaking hyphen'], ['…', 'Ellipsis'], ['·', 'Middle dot'],
    ['“', 'Left double quotation mark'], ['”', 'Right double quotation mark'], ['‘', 'Left single quotation mark'], ['’', 'Right single quotation mark (apostrophe)'],
    ['„', 'Double low quotation mark'], ['«', 'Left guillemet'], ['»', 'Right guillemet'], ['‹', 'Single left guillemet'], ['›', 'Single right guillemet'],
    ['§', 'Section'], ['¶', 'Pilcrow (paragraph)'], ['†', 'Dagger'], ['‡', 'Double dagger'], ['※', 'Reference mark'],
    ['¡', 'Inverted exclamation mark'], ['¿', 'Inverted question mark'], ['′', 'Prime (feet, minutes)'], ['″', 'Double prime (inches, seconds)'], ['№', 'Numero sign'],
  ]),
  group('Currency', [
    ['€', 'Euro'], ['£', 'Pound'], ['¥', 'Yen, yuan'], ['¢', 'Cent'], ['$', 'Dollar'],
    ['₹', 'Indian rupee'], ['₩', 'Won'], ['₽', 'Rouble'], ['₺', 'Turkish lira'], ['₫', 'Dong'],
    ['₱', 'Peso'], ['₦', 'Naira'], ['฿', 'Baht'], ['₪', 'Shekel'], ['₴', 'Hryvnia'],
    ['₿', 'Bitcoin'], ['¤', 'Currency sign'], ['₣', 'Franc'], ['₲', 'Guarani'], ['₸', 'Tenge'],
  ]),
  group('Arrows', [
    ['→', 'Right arrow'], ['←', 'Left arrow'], ['↑', 'Up arrow'], ['↓', 'Down arrow'], ['↔', 'Left right arrow'],
    ['↕', 'Up down arrow'], ['⇒', 'Rightwards double arrow'], ['⇐', 'Leftwards double arrow'], ['⇔', 'Left right double arrow'], ['↗', 'North east arrow'],
    ['↘', 'South east arrow'], ['↩', 'Return arrow'], ['➔', 'Heavy right arrow'], ['►', 'Right pointer'], ['◄', 'Left pointer'],
    ['▲', 'Up triangle'], ['▼', 'Down triangle'], ['▶', 'Right triangle'], ['◀', 'Left triangle'], ['⟶', 'Long right arrow'],
  ]),
  group('Maths', [
    ['±', 'Plus-minus'], ['×', 'Multiplication'], ['÷', 'Division'], ['−', 'Minus'], ['≈', 'Almost equal'],
    ['≠', 'Not equal'], ['≤', 'Less than or equal'], ['≥', 'Greater than or equal'], ['∞', 'Infinity'], ['√', 'Square root'],
    ['∑', 'Sum'], ['∏', 'Product'], ['∫', 'Integral'], ['∆', 'Increment (delta)'], ['∂', 'Partial differential'],
    ['‰', 'Per mille'], ['°', 'Degree'], ['µ', 'Micro'], ['²', 'Superscript two (squared)'], ['³', 'Superscript three (cubed)'],
    ['¹', 'Superscript one'], ['½', 'One half'], ['⅓', 'One third'], ['¼', 'One quarter'], ['¾', 'Three quarters'],
    ['∈', 'Element of'], ['∩', 'Intersection'], ['∪', 'Union'], ['∴', 'Therefore'], ['∠', 'Angle'],
  ]),
  group('Ticks and bullets', [
    ['✓', 'Check mark'], ['✔', 'Heavy check mark'], ['✗', 'Ballot X'], ['✘', 'Heavy ballot X'], ['☐', 'Ballot box'],
    ['☑', 'Ballot box with check'], ['☒', 'Ballot box with X'], ['•', 'Bullet'], ['◦', 'White bullet'], ['▪', 'Small black square'],
    ['▫', 'Small white square'], ['■', 'Black square'], ['□', 'White square'], ['●', 'Black circle'], ['○', 'White circle'],
    ['◆', 'Black diamond'], ['◇', 'White diamond'], ['★', 'Black star'], ['☆', 'White star'], ['❖', 'Diamond minus white X'],
    ['☎', 'Telephone'], ['✉', 'Envelope'], ['✎', 'Pencil'], ['⚠', 'Warning'], ['☺', 'Smiling face'],
  ]),
  group('Letters', [
    ['À', 'A grave'], ['Á', 'A acute'], ['Â', 'A circumflex'], ['Ä', 'A diaeresis'], ['Å', 'A ring'],
    ['Æ', 'AE'], ['Ç', 'C cedilla'], ['É', 'E acute'], ['È', 'E grave'], ['Ë', 'E diaeresis'],
    ['Ñ', 'N tilde'], ['Ö', 'O diaeresis'], ['Ø', 'O stroke'], ['Ü', 'U diaeresis'], ['ß', 'Sharp s'],
    ['à', 'a grave'], ['á', 'a acute'], ['â', 'a circumflex'], ['ä', 'a diaeresis'], ['å', 'a ring'],
    ['æ', 'ae'], ['ç', 'c cedilla'], ['é', 'e acute'], ['è', 'e grave'], ['ê', 'e circumflex'],
    ['ë', 'e diaeresis'], ['í', 'i acute'], ['ï', 'i diaeresis'], ['ñ', 'n tilde'], ['ó', 'o acute'],
    ['ö', 'o diaeresis'], ['ø', 'o stroke'], ['œ', 'oe'], ['ú', 'u acute'], ['ü', 'u diaeresis'],
  ]),
  group('Greek', [
    ['α', 'Alpha'], ['β', 'Beta'], ['γ', 'Gamma'], ['δ', 'Delta'], ['ε', 'Epsilon'],
    ['θ', 'Theta'], ['λ', 'Lambda'], ['μ', 'Mu'], ['π', 'Pi'], ['ρ', 'Rho'],
    ['σ', 'Sigma'], ['τ', 'Tau'], ['φ', 'Phi'], ['χ', 'Chi'], ['ψ', 'Psi'],
    ['ω', 'Omega'], ['Γ', 'Capital gamma'], ['Δ', 'Capital delta'], ['Σ', 'Capital sigma'], ['Ω', 'Capital omega (ohm)'],
  ]),
  group('Spaces', [
    [' ', 'Non-breaking space'], [' ', 'Thin space'], [' ', 'Narrow non-breaking space'], [' ', 'En space'], [' ', 'Em space'],
    ['­', 'Optional hyphen'], ['​', 'Zero-width space (optional break)'],
  ]),
];

/** Every symbol once, with the names it is known by, for searching. */
const ALL = (() => {
  const by = new Map();
  for (const g of SYMBOL_GROUPS) for (const s of g.symbols) if (!by.has(s.char)) by.set(s.char, s);
  return [...by.values()];
})();

/** U+00A9 for "©": a character's code, as Word's Symbol dialog shows it. */
export const codeOf = (char) => `U+${char.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')}`;

/** A symbol's name for a character (from the list, else its code). */
export const nameOf = (char) => ALL.find((s) => s.char === char)?.name || codeOf(char);

/** Symbols whose name (or code, or the character itself) matches `query`. */
export function searchSymbols(query) {
  const q = query.trim().toLowerCase();
  if (!q) return [];
  const code = /^(?:u\+)?([0-9a-f]{2,6})$/i.exec(q);
  return ALL.filter((s) => s.char === query.trim() || s.name.toLowerCase().includes(q) || (code && codeOf(s.char).slice(2) === code[1].toUpperCase().padStart(4, '0')));
}

/**
 * The character a Unicode code stands for, as typed in Word's "Character code" box ("00A9", "U+20AC",
 * "2713"), or null. Control characters and unassigned surrogates aren't characters to insert.
 */
export function charFromCode(text) {
  const m = /^\s*(?:u\+|0x)?([0-9a-f]{1,6})\s*$/i.exec(text || '');
  if (!m) return null;
  const cp = parseInt(m[1], 16);
  if (cp < 0x20 || (cp >= 0x7f && cp < 0xa0) || (cp >= 0xd800 && cp <= 0xdfff) || cp > 0x10ffff) return null;
  return String.fromCodePoint(cp);
}

// The symbols used last, remembered in this browser (as Word remembers its recently used symbols).
const RECENT_KEY = 'intact.recentSymbols';
const RECENT_MAX = 20;

export function recentSymbols() {
  try {
    const list = JSON.parse(localStorage.getItem(RECENT_KEY) || '[]');
    return Array.isArray(list) ? list.filter((c) => typeof c === 'string' && c.length && c.length <= 2).slice(0, RECENT_MAX) : [];
  } catch {
    return [];
  }
}

export function rememberSymbol(char) {
  const list = [char, ...recentSymbols().filter((c) => c !== char)].slice(0, RECENT_MAX);
  try { localStorage.setItem(RECENT_KEY, JSON.stringify(list)); } catch { /* this session only */ }
  return list;
}
