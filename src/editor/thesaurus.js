// A thesaurus, as Word's Review › Thesaurus (Shift+F7): a word's meanings, each with the words that
// say it, and similar and opposite meanings. The words are open wordnets, built into the app by
// scripts/thesaurus.mjs (src/data/, with their licences): English from Princeton WordNet, Indonesian
// from Wordnet Bahasa (only its hand-checked words, so it is small). A language's data is loaded the
// first time it is looked up in, from the app's own files, which work offline: nothing is sent anywhere.

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
    loaded.set(language, LOADERS[language]().then((m) => new Thesaurus(m.default)).catch((e) => {
      loaded.delete(language); // tried again next time
      throw e;
    }));
  }
  return loaded.get(language);
}

const POS_LABEL = { n: 'noun', v: 'verb', a: 'adjective', r: 'adverb' };

// English word forms back to the form the thesaurus lists, as WordNet's own "morphy" does: endings
// by part of speech, and the common irregular forms.
const ENDINGS = {
  n: [['s', ''], ['ses', 's'], ['xes', 'x'], ['zes', 'z'], ['ches', 'ch'], ['shes', 'sh'], ['men', 'man'], ['ies', 'y']],
  v: [['s', ''], ['ies', 'y'], ['es', 'e'], ['es', ''], ['ed', 'e'], ['ed', ''], ['ing', 'e'], ['ing', ''], ['ied', 'y'], ['ying', 'ie']],
  a: [['er', ''], ['est', ''], ['er', 'e'], ['est', 'e'], ['ier', 'y'], ['iest', 'y']],
};
const IRREGULAR = {
  was: 'be', were: 'be', been: 'be', is: 'be', are: 'be', am: 'be', had: 'have', has: 'have', did: 'do', does: 'do', done: 'do',
  said: 'say', made: 'make', went: 'go', gone: 'go', took: 'take', taken: 'take', came: 'come', saw: 'see', seen: 'see',
  knew: 'know', known: 'know', got: 'get', gotten: 'get', gave: 'give', given: 'give', found: 'find', thought: 'think',
  told: 'tell', became: 'become', left: 'leave', felt: 'feel', brought: 'bring', began: 'begin', begun: 'begin', kept: 'keep',
  held: 'hold', wrote: 'write', written: 'write', stood: 'stand', heard: 'hear', meant: 'mean', met: 'meet', ran: 'run',
  paid: 'pay', sat: 'sit', spoke: 'speak', spoken: 'speak', led: 'lead', grew: 'grow', grown: 'grow', lost: 'lose',
  fell: 'fall', fallen: 'fall', sent: 'send', built: 'build', understood: 'understand', drew: 'draw', drawn: 'draw',
  broke: 'break', broken: 'break', spent: 'spend', rose: 'rise', risen: 'rise', drove: 'drive', driven: 'drive',
  bought: 'buy', wore: 'wear', worn: 'wear', chose: 'choose', chosen: 'choose', sought: 'seek', threw: 'throw', thrown: 'throw',
  caught: 'catch', fought: 'fight', taught: 'teach', sold: 'sell', won: 'win', ate: 'eat', eaten: 'eat', flew: 'fly', flown: 'fly',
  better: 'good', best: 'good', worse: 'bad', worst: 'bad', more: 'much', most: 'much', less: 'little', least: 'little',
  children: 'child', people: 'person', men: 'man', women: 'woman', feet: 'foot', teeth: 'tooth', mice: 'mouse', geese: 'goose',
};

export class Thesaurus {
  constructor(text) {
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

  /** The forms of `word` the thesaurus lists, the word itself first (English endings undone). */
  forms(word) {
    const w = word.toLowerCase().trim().replace(/\s+/g, ' ');
    const out = [w];
    if (IRREGULAR[w]) out.push(IRREGULAR[w]);
    for (const rules of Object.values(ENDINGS)) {
      for (const [end, base] of rules) if (w.endsWith(end) && w.length > end.length + 1) out.push(w.slice(0, -end.length) + base);
    }
    // A doubled consonant before -ed, -ing, -er, -est: "stopped", "running", "bigger".
    const doubled = /^(.*([bdfgklmnprtvz]))\2(?:ed|ing|er|est)$/.exec(w);
    if (doubled) out.push(doubled[1]);
    return [...new Set(out)].filter((f) => this.index.has(f));
  }

  /**
   * `word`'s meanings: { word (the form found first, or null), meanings: [{ form, pos, label, words,
   * similar, opposite }] }: the word's own meanings, then its base form's ("said": "say"), each the
   * most used first; `words` say the same, without the form looked up.
   */
  lookup(word) {
    const forms = this.forms(word);
    const refs = (s) => (s ? s.split(',').map((n) => parseInt(n, 36)) : []);
    const seen = new Set();
    const meanings = [];
    for (const form of forms) {
      const not = (w) => w.toLowerCase() !== form;
      for (const i of this.index.get(form)) {
        if (seen.has(i)) continue;
        seen.add(i);
        const [pos, words, similar, opposite] = this.lines[i];
        const related = (list) => [...new Set(list.flatMap((j) => this.lines[j][1].split('|')))].filter(not);
        const m = { form, pos, label: POS_LABEL[pos] || pos, words: words.split('|').filter(not), similar: related(refs(similar)), opposite: related(refs(opposite)) };
        if (m.words.length || m.similar.length || m.opposite.length) meanings.push(m);
      }
    }
    return { word: forms[0] || null, meanings };
  }
}

/** `word` in the case of `like`, the word it replaces: "Big" → "Large", "BIG" → "LARGE". */
export function matchCase(word, like) {
  if (like.length > 1 && like === like.toUpperCase() && like !== like.toLowerCase()) return word.toUpperCase();
  if (like[0] && like[0] === like[0].toUpperCase() && like[0] !== like[0].toLowerCase()) return word[0].toUpperCase() + word.slice(1);
  return word;
}
