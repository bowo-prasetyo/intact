// English word forms, for the thesaurus: a word looked up in a form ("approves", "took", "bigger") is
// found by its base ("approve", "take", "big") and the form it is in, and the words offered for it
// are put in that same form ("sanctions", "got", "larger"), so the one chosen fits where it goes.
// The rules are English spelling's (stop → stopped, carry → carried, big → bigger), with the common
// irregular verbs, nouns and adjectives. Pure functions: nothing here touches the page.

// Irregular verbs: base → [past, past participle].
const VERBS = {
  arise: ['arose', 'arisen'], awake: ['awoke', 'awoken'], be: ['was', 'been'], bear: ['bore', 'borne'], beat: ['beat', 'beaten'],
  become: ['became', 'become'], begin: ['began', 'begun'], bend: ['bent', 'bent'], bet: ['bet', 'bet'], bid: ['bid', 'bid'],
  bind: ['bound', 'bound'], bite: ['bit', 'bitten'], bleed: ['bled', 'bled'], blow: ['blew', 'blown'], break: ['broke', 'broken'],
  breed: ['bred', 'bred'], bring: ['brought', 'brought'], broadcast: ['broadcast', 'broadcast'], build: ['built', 'built'],
  burn: ['burnt', 'burnt'], burst: ['burst', 'burst'], buy: ['bought', 'bought'], cast: ['cast', 'cast'], catch: ['caught', 'caught'],
  choose: ['chose', 'chosen'], cling: ['clung', 'clung'], come: ['came', 'come'], cost: ['cost', 'cost'], creep: ['crept', 'crept'],
  cut: ['cut', 'cut'], deal: ['dealt', 'dealt'], dig: ['dug', 'dug'], do: ['did', 'done'], draw: ['drew', 'drawn'],
  dream: ['dreamt', 'dreamt'], drink: ['drank', 'drunk'], drive: ['drove', 'driven'], eat: ['ate', 'eaten'], fall: ['fell', 'fallen'],
  feed: ['fed', 'fed'], feel: ['felt', 'felt'], fight: ['fought', 'fought'], find: ['found', 'found'], flee: ['fled', 'fled'],
  fling: ['flung', 'flung'], fly: ['flew', 'flown'], forbid: ['forbade', 'forbidden'], forecast: ['forecast', 'forecast'],
  foresee: ['foresaw', 'foreseen'], forget: ['forgot', 'forgotten'], forgive: ['forgave', 'forgiven'], freeze: ['froze', 'frozen'],
  get: ['got', 'gotten'], give: ['gave', 'given'], go: ['went', 'gone'], grind: ['ground', 'ground'], grow: ['grew', 'grown'],
  hang: ['hung', 'hung'], have: ['had', 'had'], hear: ['heard', 'heard'], hide: ['hid', 'hidden'], hit: ['hit', 'hit'],
  hold: ['held', 'held'], hurt: ['hurt', 'hurt'], keep: ['kept', 'kept'], kneel: ['knelt', 'knelt'], know: ['knew', 'known'],
  lay: ['laid', 'laid'], lead: ['led', 'led'], lean: ['leant', 'leant'], leap: ['leapt', 'leapt'], learn: ['learnt', 'learnt'],
  leave: ['left', 'left'], lend: ['lent', 'lent'], let: ['let', 'let'], lie: ['lay', 'lain'], light: ['lit', 'lit'],
  lose: ['lost', 'lost'], make: ['made', 'made'], mean: ['meant', 'meant'], meet: ['met', 'met'], mislead: ['misled', 'misled'],
  mistake: ['mistook', 'mistaken'], overcome: ['overcame', 'overcome'], overtake: ['overtook', 'overtaken'], pay: ['paid', 'paid'],
  prove: ['proved', 'proven'], put: ['put', 'put'], quit: ['quit', 'quit'], read: ['read', 'read'], rid: ['rid', 'rid'],
  ride: ['rode', 'ridden'], ring: ['rang', 'rung'], rise: ['rose', 'risen'], run: ['ran', 'run'], say: ['said', 'said'],
  see: ['saw', 'seen'], seek: ['sought', 'sought'], sell: ['sold', 'sold'], send: ['sent', 'sent'], set: ['set', 'set'],
  sew: ['sewed', 'sewn'], shake: ['shook', 'shaken'], shed: ['shed', 'shed'], shine: ['shone', 'shone'], shoot: ['shot', 'shot'],
  show: ['showed', 'shown'], shrink: ['shrank', 'shrunk'], shut: ['shut', 'shut'], sing: ['sang', 'sung'], sink: ['sank', 'sunk'],
  sit: ['sat', 'sat'], sleep: ['slept', 'slept'], slide: ['slid', 'slid'], sling: ['slung', 'slung'], slit: ['slit', 'slit'],
  speak: ['spoke', 'spoken'], speed: ['sped', 'sped'], spend: ['spent', 'spent'], spill: ['spilt', 'spilt'], spin: ['spun', 'spun'],
  spit: ['spat', 'spat'], split: ['split', 'split'], spread: ['spread', 'spread'], spring: ['sprang', 'sprung'], stand: ['stood', 'stood'],
  steal: ['stole', 'stolen'], stick: ['stuck', 'stuck'], sting: ['stung', 'stung'], stink: ['stank', 'stunk'], stride: ['strode', 'stridden'],
  strike: ['struck', 'struck'], string: ['strung', 'strung'], strive: ['strove', 'striven'], swear: ['swore', 'sworn'],
  sweep: ['swept', 'swept'], swell: ['swelled', 'swollen'], swim: ['swam', 'swum'], swing: ['swung', 'swung'], take: ['took', 'taken'],
  teach: ['taught', 'taught'], tear: ['tore', 'torn'], tell: ['told', 'told'], think: ['thought', 'thought'], throw: ['threw', 'thrown'],
  thrust: ['thrust', 'thrust'], tread: ['trod', 'trodden'], undergo: ['underwent', 'undergone'], understand: ['understood', 'understood'],
  undertake: ['undertook', 'undertaken'], upset: ['upset', 'upset'], wake: ['woke', 'woken'], wear: ['wore', 'worn'],
  weave: ['wove', 'woven'], weep: ['wept', 'wept'], win: ['won', 'won'], wind: ['wound', 'wound'], withdraw: ['withdrew', 'withdrawn'],
  withhold: ['withheld', 'withheld'], withstand: ['withstood', 'withstood'], wring: ['wrung', 'wrung'], write: ['wrote', 'written'],
};
// Irregular plurals and the -s forms that aren't base + s.
const NOUNS = {
  man: 'men', woman: 'women', child: 'children', person: 'people', foot: 'feet', tooth: 'teeth', mouse: 'mice', goose: 'geese',
  ox: 'oxen', life: 'lives', wife: 'wives', knife: 'knives', leaf: 'leaves', half: 'halves', wolf: 'wolves', shelf: 'shelves',
  thief: 'thieves', loaf: 'loaves', calf: 'calves', analysis: 'analyses', crisis: 'crises', thesis: 'theses', basis: 'bases',
  criterion: 'criteria', phenomenon: 'phenomena', sheep: 'sheep', fish: 'fish', deer: 'deer', series: 'series', species: 'species',
};
const THIRD = { be: 'is', have: 'has', do: 'does', go: 'goes' };
// Irregular comparatives: base → [comparative, superlative].
const ADJECTIVES = {
  good: ['better', 'best'], bad: ['worse', 'worst'], far: ['farther', 'farthest'], little: ['less', 'least'],
  many: ['more', 'most'], much: ['more', 'most'], ill: ['worse', 'worst'],
};

// The other way: an irregular form → [{ base, pos, kind }].
const IRREGULAR = new Map();
const addIrregular = (form, entry) => {
  if (!IRREGULAR.has(form)) IRREGULAR.set(form, []);
  IRREGULAR.get(form).push(entry);
};
for (const [base, [past, part]] of Object.entries(VERBS)) {
  if (past === part) addIrregular(past, { base, pos: 'v', kind: 'past' }); // past or participle: the words before it say
  else {
    addIrregular(past, { base, pos: 'v', kind: 'simplePast' });
    addIrregular(part, { base, pos: 'v', kind: 'participle' });
  }
}
for (const [base, third] of Object.entries(THIRD)) addIrregular(third, { base, pos: 'v', kind: 'third' });
addIrregular('were', { base: 'be', pos: 'v', kind: 'simplePast' });
addIrregular('are', { base: 'be', pos: 'v', kind: 'base' });
addIrregular('am', { base: 'be', pos: 'v', kind: 'base' });
for (const [base, plural] of Object.entries(NOUNS)) if (plural !== base) addIrregular(plural, { base, pos: 'n', kind: 'plural' });
for (const [base, [er, est]] of Object.entries(ADJECTIVES)) {
  addIrregular(er, { base, pos: 'a', kind: 'er' });
  addIrregular(est, { base, pos: 'a', kind: 'est' });
}

// Regular endings, undone: [ending, what replaces it, part of speech, form]. As WordNet's "morphy".
const ENDINGS = [
  ['s', '', 'n', 'plural'], ['ses', 's', 'n', 'plural'], ['xes', 'x', 'n', 'plural'], ['zes', 'z', 'n', 'plural'],
  ['ches', 'ch', 'n', 'plural'], ['shes', 'sh', 'n', 'plural'], ['ies', 'y', 'n', 'plural'],
  ['s', '', 'v', 'third'], ['ies', 'y', 'v', 'third'], ['es', 'e', 'v', 'third'], ['es', '', 'v', 'third'],
  ['ed', 'e', 'v', 'past'], ['ed', '', 'v', 'past'], ['ied', 'y', 'v', 'past'],
  ['ing', 'e', 'v', 'ing'], ['ing', '', 'v', 'ing'], ['ying', 'ie', 'v', 'ing'],
  ['er', '', 'a', 'er'], ['est', '', 'a', 'est'], ['er', 'e', 'a', 'er'], ['est', 'e', 'a', 'est'], ['ier', 'y', 'a', 'er'], ['iest', 'y', 'a', 'est'],
];
const DOUBLED = { ed: ['v', 'past'], ing: ['v', 'ing'], er: ['a', 'er'], est: ['a', 'est'] };

/**
 * What `word` (lower case) may be a form of: [{ base, pos, kind }], the word itself first (kind
 * 'base', any part of speech). kind: 'plural' (a noun's), 'third' (approves), 'past' (approved: past
 * or participle, as the words before it say), 'simplePast' (took), 'participle' (taken), 'ing',
 * 'er', 'est'. Not every one is a word: the thesaurus keeps those it has.
 */
export function formsOf(word) {
  const out = [{ base: word, pos: null, kind: 'base' }];
  out.push(...(IRREGULAR.get(word) || []));
  for (const [end, add, pos, kind] of ENDINGS) {
    if (word.endsWith(end) && word.length > end.length + 1) out.push({ base: word.slice(0, -end.length) + add, pos, kind });
  }
  // A doubled consonant before -ed, -ing, -er, -est: "stopped", "running", "bigger".
  const doubled = /^(.*([bdfgklmnprtvz]))\2(ed|ing|er|est)$/.exec(word);
  if (doubled) out.push({ base: doubled[1], pos: DOUBLED[doubled[3]][0], kind: DOUBLED[doubled[3]][1] });
  return out;
}

/** Syllables, roughly: groups of vowels, a silent e at the end left out. */
function syllables(word) {
  const groups = word.match(/[aeiouy]+/g)?.length || 1;
  return word.length > 2 && /[^aeiouy]e$/.test(word) && !/le$/.test(word) && groups > 1 ? groups - 1 : groups;
}
/** Whether the last consonant doubles before an ending (stop → stopped, big → bigger): one syllable, consonant-vowel-consonant. */
const doubles = (w) => (syllables(w) === 1 && /(?:^|[^aeiou])[aeiou][bdfgklmnprtvz]$/.test(w)) || STRESSED_LAST.has(w);
// Longer verbs stressed at the end double it too (permitted, preferred), unlike visited or offered.
const STRESSED_LAST = new Set(['admit', 'commit', 'permit', 'submit', 'omit', 'emit', 'transmit', 'remit', 'acquit', 'refer', 'prefer',
  'defer', 'confer', 'transfer', 'infer', 'deter', 'occur', 'recur', 'incur', 'concur', 'compel', 'expel', 'propel', 'repel', 'dispel',
  'rebel', 'regret', 'control', 'patrol', 'equip', 'allot', 'abet', 'embed', 'outwit', 'abhor', 'befit', 'forbid', 'outbid', 'aver', 'demur', 'annul']);

/** A verb's -s form: approve → approves, carry → carries, push → pushes. */
function third(w) {
  if (THIRD[w]) return THIRD[w];
  if (/(s|x|z|ch|sh)$/.test(w)) return `${w}es`;
  if (/[^aeiou]y$/.test(w)) return `${w.slice(0, -1)}ies`;
  if (/[^aeiou]o$/.test(w)) return `${w}es`;
  return `${w}s`;
}
/** A regular verb's -ed form: approve → approved, carry → carried, stop → stopped. */
function ed(w) {
  if (w.endsWith('e')) return `${w}d`;
  if (/[^aeiou]y$/.test(w)) return `${w.slice(0, -1)}ied`;
  if (doubles(w)) return `${w}${w.at(-1)}ed`;
  return `${w}ed`;
}
/** A verb's -ing form: approve → approving, die → dying, stop → stopping, see → seeing. */
function ing(w) {
  if (w.endsWith('ie')) return `${w.slice(0, -2)}ying`;
  if (/[^aeioy]e$/.test(w) && w.length > 2) return `${w.slice(0, -1)}ing`;
  if (doubles(w)) return `${w}${w.at(-1)}ing`;
  return `${w}ing`;
}
/** A noun's plural: vote → votes, levy → levies, tax → taxes, man → men. */
function plural(w) {
  if (NOUNS[w]) return NOUNS[w];
  if (/(s|x|z|ch|sh)$/.test(w)) return `${w}es`;
  if (/[^aeiou]y$/.test(w)) return `${w.slice(0, -1)}ies`;
  return `${w}s`;
}
/** An adjective's -er or -est form (`i` 0 or 1): big → bigger, happy → happiest, important → more important. */
function compare(w, i) {
  if (ADJECTIVES[w]) return ADJECTIVES[w][i];
  const end = i ? 'est' : 'er';
  const n = syllables(w);
  if (n === 1) {
    if (w.endsWith('e')) return `${w}${end.slice(1)}`;
    if (/[^aeiou]y$/.test(w)) return `${w.slice(0, -1)}i${end}`;
    if (doubles(w)) return `${w}${w.at(-1)}${end}`;
    return `${w}${end}`;
  }
  if (n === 2 && /[^aeiou]y$/.test(w)) return `${w.slice(0, -1)}i${end}`;
  return `${i ? 'most' : 'more'} ${w}`;
}

/**
 * `word` (as the thesaurus lists it: a base form, perhaps of several words) in form `kind` (see
 * formsOf), as a `pos` ('n', 'v', 'a', 'r'); `participle`: an ambiguous -ed form is a past participle
 * (after "has", "was"…). Null for a name, an abbreviation or a word with digits, which can't be put
 * in another form ("O.K." for "approves").
 */
export function inflect(word, kind, pos, participle = false) {
  if (!kind || kind === 'base') return word;
  if (/[A-Z0-9.]/.test(word)) return null;
  const parts = word.split(' ');
  if (pos === 'n' && kind === 'plural') {
    // The last word of a noun: "flood walls"; "...of..." the first: "rights of way".
    const at = parts.indexOf('of') > 0 ? parts.indexOf('of') - 1 : parts.length - 1;
    parts[at] = plural(parts[at]);
    return parts.join(' ');
  }
  if (pos === 'v') {
    // The verb is the first word: "turns over", "turned down".
    const w = parts[0];
    const irregular = VERBS[w];
    const form = {
      third: () => third(w),
      ing: () => ing(w),
      simplePast: () => (irregular ? irregular[0] : ed(w)),
      participle: () => (irregular ? irregular[1] : ed(w)),
      past: () => (irregular ? irregular[participle ? 1 : 0] : ed(w)),
    }[kind];
    if (!form) return word;
    parts[0] = form();
    return parts.join(' ');
  }
  if (pos === 'a' && (kind === 'er' || kind === 'est')) {
    return parts.length > 1 ? `${kind === 'est' ? 'most' : 'more'} ${word}` : compare(word, kind === 'est' ? 1 : 0);
  }
  return word;
}

/** Whether a verb after `before` (the text before it) is a past participle: after have, be or get. */
export function participleAfter(before) {
  return /\b(?:have|has|had|having|['’]ve|['’]d|be|been|being|is|are|was|were|am|['’]s|['’]re|get|gets|got|gotten|getting)\s+(?:\w+ly\s+|not\s+|never\s+|just\s+|already\s+)*$/i.test(before);
}

