// A thesaurus, as Word's Review › Thesaurus (Shift+F7): a word's meanings, each with the words that
// say it, and similar and opposite meanings. The words are open wordnets, built into the app by
// scripts/thesaurus.mjs (src/data/, with their licences): English from Princeton WordNet, Indonesian
// from Wordnet Bahasa (only its hand-checked words, so it is small). A language's data is loaded the
// first time it is looked up in, from the app's own files, which work offline: nothing is sent anywhere.
// English words are looked up by their base and offered in the form looked up (see editor/inflect.js).

const LOADERS = {
  en: () => import('../data/thesaurus-en.txt?raw'),
  id: () => import('../data/thesaurus-id.txt?raw'),
};
const loaded = new Map(); // language -> Promise of a Thesaurus

/** The thesaurus's language for a language tag ("en-GB", "id-ID"): 'en' or 'id' (Malay too). */
export function thesaurusLanguage(tag) {
  return /^(id|in|ms)\b/i.test(tag || '') ? 'id' : 'en';
}

/** The thesaurus of `language` ('en' or 'id'), loaded once. */
export function loadThesaurus(language) {
  if (!loaded.has(language)) {
    loaded.set(language, LOADERS[language]().then((m) => new Thesaurus(m.default, language)).catch((e) => {
      loaded.delete(language); // tried again next time
      throw e;
    }));
  }
  return loaded.get(language);
}

import { formsOf, inflect } from './inflect.js';

const POS_LABEL = { n: 'noun', v: 'verb', a: 'adjective', r: 'adverb' };

export class Thesaurus {
  /** `language`: 'en' looks up English word forms by their base (see editor/inflect.js). */
  constructor(text, language = 'en') {
    this.english = language === 'en';
    this.lines = text.split('\n').filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'));
    this.index = new Map();
    this.lines.forEach(([, words], i) => {
      for (const w of words.split('|')) {
        const key = w.toLowerCase();
        if (!this.index.has(key)) this.index.set(key, []);
        this.index.get(key).push(i);
      }
    });
  }

  /**
   * `word`'s meanings: { word (the form found first, or null), inflected (whether words are offered
   * in another form than listed), meanings: [{ form, kind, pos, label, words, similar, opposite }] }:
   * the word's own meanings, then its base form's ("approves": "approve"), each the most used first.
   * A base's meanings are those of the part of speech its form is ("approves": a verb's, not a
   * noun's), and their words are in that form ("sanctions"); `participle`: an -ed form is a past
   * participle (after "has", "was"…), so "approved" offers "given" rather than "gave". `words` say the
   * same, without the form looked up.
   */
  lookup(word, { participle = false } = {}) {
    const w = word.toLowerCase().trim().replace(/\s+/g, ' ');
    const candidates = (this.english ? formsOf(w) : [{ base: w, pos: null, kind: 'base' }]).filter((f) => this.index.has(f.base));
    const refs = (s) => (s ? s.split(',').map((n) => parseInt(n, 36)) : []);
    const seen = new Set();
    const meanings = [];
    for (const f of candidates) {
      for (const i of this.index.get(f.base)) {
        const [pos, words, similar, opposite] = this.lines[i];
        if (seen.has(i) || (f.pos && f.pos !== pos)) continue;
        seen.add(i);
        const not = (x) => x.toLowerCase() !== f.base;
        const formed = (list) => [...new Set(list.filter(not).map((x) => inflect(x, f.kind, pos, participle)).filter(Boolean))];
        const related = (list) => formed(list.flatMap((j) => this.lines[j][1].split('|')));
        const m = { form: f.base, kind: f.kind, pos, label: POS_LABEL[pos] || pos, words: formed(words.split('|')), similar: related(refs(similar)), opposite: related(refs(opposite)) };
        if (m.words.length || m.similar.length || m.opposite.length) meanings.push(m);
      }
    }
    return { word: candidates[0]?.base || null, inflected: meanings.some((m) => m.kind !== 'base'), meanings };
  }
}

/** `word` in the case of `like`, the word it replaces: "Big" → "Large", "BIG" → "LARGE". */
export function matchCase(word, like) {
  if (like.length > 1 && like === like.toUpperCase() && like !== like.toLowerCase()) return word.toUpperCase();
  if (like[0] && like[0] === like[0].toUpperCase() && like[0] !== like[0].toLowerCase()) return word[0].toUpperCase() + word.slice(1);
  return word;
}
