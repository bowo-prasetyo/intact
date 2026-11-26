// Word count and readability statistics, as Word's Word Count dialog and its Readability Statistics
// show them: the counts, the averages (sentences per paragraph, words per sentence, characters per
// word), the share of passive sentences and the Flesch Reading Ease and Flesch-Kincaid Grade Level.
// The scores and the passive voice are worked out for English, as Word's are; other languages get
// the counts and averages. Long sentences, long paragraphs and a word typed twice in a row ("the
// the") are listed with where they are, to show them one by one.
// The headline (the first heading of the text) is found here too, for a headline's character limit.
// Everything is read from the page as shown: nothing is sent anywhere, and nothing is changed.

/** Words, as the status bar counts them (as Word does): letters and digits, with ' and - inside. */
const WORD = /[\p{L}\p{N}]+(?:['’-][\p{L}\p{N}]+)*/gu;
export const wordCount = (text) => (text.match(WORD) || []).length;

/** Sentences longer than this many words are listed as long (plain-language guides say 20 to 30). */
export const LONG_SENTENCE = 30;
/** Paragraphs longer than this many words are listed as long (a screen's worth, read on the web). */
export const LONG_PARAGRAPH = 150;

// Words that are often right twice in a row ("had had", "that that", "what it is is"), and so aren't
// listed as repeated, as Word's spelling check leaves them.
const DOUBLED_ON_PURPOSE = new Set(['had', 'that', 'is', 'do', 'bye', 'very', 'so', 'no', 'yeah', 'ha', 'la', 'pom', 'tut', 'knock', 'blah', 'chop', 'night', 'cha']);

/**
 * A word typed twice in a row in `text` ("the the", "of of", across a line break too): [{ start, end }]
 * over both, the case of each ignored. Numbers and words doubled on purpose aren't listed.
 */
export function repeatedWords(text) {
  const out = [];
  let prev = null;
  for (const m of text.matchAll(WORD)) {
    const word = m[0].toLocaleLowerCase();
    // Only spaces between them: "Well, well" and "it. It" are written that way on purpose.
    if (prev && prev.word === word && /^\s+$/u.test(text.slice(prev.end, m.index)) && !/^\p{N}+$/u.test(word) && !DOUBLED_ON_PURPOSE.has(word)) {
      out.push({ start: prev.start, end: m.index + m[0].length });
    }
    prev = { word, start: m.index, end: m.index + m[0].length };
  }
  return out;
}

const BLOCKS = 'p, h1, h2, h3, h4, h5, h6';
// Text the reader doesn't read: deleted (tracked) text, hidden text, preserved content kept out of
// view, note numbers and the editor's own marks.
const SKIP = '[data-tc="del"], .raw-inline.deleted, [data-hidden], .raw-hidden, .note-ref, .comment-ref';

const inBody = (el) => !el.closest('.docx-notes, .docx-hf-band');

/** The text of `block` as read, with where each piece of it is: { text, nodes: [{ node, start }] }. */
export function blockText(block) {
  const nodes = [];
  let text = '';
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const parent = n.parentElement;
    if (!n.data || parent.closest(SKIP) || parent.closest(BLOCKS) !== block) continue;
    nodes.push({ node: n, start: text.length });
    text += n.data;
  }
  return { text, nodes };
}

/** A DOM range over characters `start` to `end` of a block's text (see blockText). */
export function textRange({ nodes }, start, end) {
  const at = (offset, atEnd) => {
    for (let i = 0; i < nodes.length; i++) {
      const { node, start: s } = nodes[i];
      if (atEnd ? offset <= s + node.length : offset < s + node.length || i === nodes.length - 1) return [node, Math.min(Math.max(offset - s, 0), node.length)];
    }
    return [nodes[0].node, 0];
  };
  const range = document.createRange();
  range.setStart(...at(start, false));
  range.setEnd(...at(end, true));
  return range;
}

// Abbreviations a full stop doesn't end a sentence after (titles, months, common Latin and the like).
const ABBREVIATIONS = new Set(['mr', 'mrs', 'ms', 'dr', 'prof', 'sr', 'jr', 'st', 'mt', 'gen', 'gov', 'sen', 'rep', 'rev', 'col', 'capt', 'lt', 'sgt',
  'cllr', 'hon', 'jan', 'feb', 'mar', 'apr', 'jun', 'jul', 'aug', 'sep', 'sept', 'oct', 'nov', 'dec', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun',
  'no', 'vol', 'p', 'pp', 'fig', 'eg', 'e.g', 'ie', 'i.e', 'etc', 'vs', 'v', 'cf', 'approx', 'inc', 'ltd', 'co', 'corp', 'dept', 'est', 'u.s', 'u.k', 'a.m', 'p.m',
  'bpk', 'ibu', 'sdr', 'yth', 'dll', 'dsb', 'tsb', 'jl', 'kab', 'kec', 'prov', 'drs', 'ir', 'h']);

/** The sentences of a paragraph's text: [{ start, end }] (a paragraph without an ending is one). */
export function sentences(text) {
  const out = [];
  let start = 0;
  const end = /[.!?…]+["'”’)\]]*(?=\s+|$)/gu;
  for (const m of text.matchAll(end)) {
    const stop = m.index + m[0].length;
    // '"Is it over?" she asked.': what goes on in lower case is the same sentence.
    if (/^\p{Ll}/u.test(text.slice(stop).trimStart())) continue;
    if (m[0][0] === '.' && m[0].length === 1) {
      // "Mr. Smith", "Nov. 20", "J. Smith": a full stop that doesn't end the sentence.
      const word = /([\p{L}.]+)\.$/u.exec(text.slice(start, stop))?.[1] || '';
      if (ABBREVIATIONS.has(word.toLowerCase()) || /^\p{L}$/u.test(word) || /^\p{N}/u.test(text.slice(stop).trimStart())) continue;
    }
    if (wordCount(text.slice(start, stop))) out.push({ start, end: stop });
    start = stop;
  }
  if (wordCount(text.slice(start))) out.push({ start, end: text.length });
  // Leading spaces aren't the sentence's.
  for (const s of out) while (s.start < s.end && /\s/.test(text[s.start])) s.start++;
  return out;
}

/** An English word's syllables, as readability formulas count them (vowel groups, less a silent e). */
export function syllables(word) {
  let w = word.toLowerCase().replace(/[^a-z]/g, '');
  if (!w) return 0;
  if (w.length <= 3) return 1;
  w = w.replace(/(?:[^laeiouy]es|[^laeiouy]ed|[^laeiouy]e)$/, '').replace(/^y/, '');
  const groups = w.match(/[aeiouy]{1,2}/g);
  return Math.max(1, groups ? groups.length : 1);
}

// The passive voice: a form of "be" (or "get"), perhaps an adverb, then a past participle.
const IRREGULAR = 'arisen|awoken|been|borne|beaten|become|begun|bent|bet|bound|bitten|bled|blown|broken|bred|brought|built|burnt|bought|caught|chosen|come|cost|cut|dealt|dug|done|drawn|driven|eaten|fallen|fed|felt|fought|found|fled|flown|forbidden|forgotten|forgiven|frozen|given|gone|ground|grown|hung|heard|hidden|hit|held|hurt|kept|known|laid|led|left|lent|let|lost|made|meant|met|paid|put|quit|read|ridden|rung|risen|run|said|seen|sought|sold|sent|set|shaken|shed|shot|shown|shut|sung|sunk|slain|slid|spoken|spent|spun|split|spread|stolen|stuck|stung|struck|sworn|swept|swum|taken|taught|torn|told|thought|thrown|understood|undertaken|upset|woken|worn|won|withdrawn|written';
const PASSIVE = new RegExp(`\\b(?:am|is|are|was|were|be|been|being|get|gets|got|gotten|getting)\\s+(?:\\w+ly\\s+)?(?:\\w{2,}ed|${IRREGULAR})\\b`, 'i');

/** The text's first paragraph, as a story's headline, if it is a heading (Title, Heading 1–6): its block, or null. */
export function headlineOf(page) {
  for (const el of page.querySelectorAll('h1, h2, h3, h4, h5, h6, p')) {
    if (!inBody(el) || el.closest('table, .textbox') || !blockText(el).text.trim()) continue;
    return /^H\d$/.test(el.tagName) ? el : null; // Title is drawn as h1, as Heading 1 is
  }
  return null;
}

/**
 * The statistics of the text in `page`: counts, averages and (for `english`) the passive share and
 * the Flesch scores, the long sentences as [{ block, start, end, words, text }], the long paragraphs
 * as [{ block, start, end, words }] and the words typed twice in a row as [{ block, start, end, text }].
 * Headers, footers and notes aren't read, as in Word, nor deleted or hidden text.
 */
export function readability(page, { english = true } = {}) {
  let paragraphs = 0;
  let sentenceCount = 0;
  let words = 0;
  let letters = 0;
  let syllableCount = 0;
  let passive = 0;
  const long = [];
  const longParagraphs = [];
  const repeated = [];
  for (const block of page.querySelectorAll(BLOCKS)) {
    if (!inBody(block)) continue;
    const bt = blockText(block);
    const blockWords = wordCount(bt.text);
    if (!blockWords) continue;
    paragraphs++;
    if (blockWords > LONG_PARAGRAPH) longParagraphs.push({ block, start: 0, end: bt.text.length, words: blockWords });
    for (const r of repeatedWords(bt.text)) repeated.push({ block, ...r, text: bt.text.slice(r.start, r.end) });
    for (const s of sentences(bt.text)) {
      const text = bt.text.slice(s.start, s.end);
      const list = text.match(WORD) || [];
      sentenceCount++;
      words += list.length;
      for (const w of list) {
        letters += [...w.replace(/['’-]/g, '')].length;
        if (english) syllableCount += syllables(w);
      }
      if (english && PASSIVE.test(text)) passive++;
      if (list.length > LONG_SENTENCE) long.push({ block, start: s.start, end: s.end, words: list.length, text });
    }
  }
  const round = (n, d = 1) => Math.round(n * 10 ** d) / 10 ** d;
  const perSentence = sentenceCount ? words / sentenceCount : 0;
  const perWord = words ? syllableCount / words : 0;
  const scores = english && words ? {
    ease: round(Math.min(100, Math.max(0, 206.835 - 1.015 * perSentence - 84.6 * perWord))),
    grade: round(Math.max(0, 0.39 * perSentence + 11.8 * perWord - 15.59)),
    passive: Math.round((passive / sentenceCount) * 100),
  } : null;
  return {
    paragraphs, sentences: sentenceCount, words,
    sentencesPerParagraph: paragraphs ? round(sentenceCount / paragraphs) : 0,
    wordsPerSentence: round(perSentence),
    charactersPerWord: words ? round(letters / words) : 0,
    english, scores, long, longParagraphs, repeated,
  };
}

/** What a Flesch Reading Ease score means, as the formula's own bands put it. */
export function easeLabel(score) {
  if (score >= 90) return 'very easy';
  if (score >= 80) return 'easy';
  if (score >= 70) return 'fairly easy';
  if (score >= 60) return 'plain English';
  if (score >= 50) return 'fairly difficult';
  if (score >= 30) return 'difficult';
  return 'very difficult';
}
