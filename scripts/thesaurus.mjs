// Builds the thesaurus data in src/data/ from open wordnets, for Review › Thesaurus (Shift+F7):
//
//   node scripts/thesaurus.mjs <wordnet-db>/dict <wn-msa-all.tab>
//
// English: Princeton WordNet 3.1 (the npm package wordnet-db: `npm pack wordnet-db@3.1.14`).
// Indonesian: Wordnet Bahasa 1.2, wn-msa-all.tab, as the package id-wordnet keeps it
// (https://github.com/open-language/id-wordnet, database/1.2/), with each word's quality: only the
// words checked by hand (Y) are kept, as the ones made automatically (O, M, L) often aren't the same
// meaning ("mematikan", to kill, among the words for "to extend"). It is small, but right.
// Both licences allow redistribution with their notices, which head each file written here.
//
// The format, one meaning (synset) per line, so the browser indexes it in a moment:
//   <part of speech: n, v, a, r>\t<words, |-separated>\t<similar meanings>\t<opposite meanings>
// where the last two are line numbers (base 36, ,-separated), counted after the notice's lines
// (which start with #). Only meanings that offer something are kept: two words or more, or a link
// to a similar or opposite meaning.
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const [dict, indonesian] = process.argv.slice(2);
if (!dict || !indonesian) {
  console.error('usage: node scripts/thesaurus.mjs <wordnet-db>/dict <wn-msa-all.tab>');
  process.exit(1);
}

const WORDNET_NOTICE = `# English: WordNet 3.1, https://wordnet.princeton.edu/ (meanings, similar and opposite meanings)
# WordNet Release 3.0 This software and database is being provided to you, the LICENSEE, by Princeton
# University under the following license. By obtaining, using and/or copying this software and database,
# you agree that you have read, understood, and will comply with these terms and conditions.: Permission
# to use, copy, modify and distribute this software and database and its documentation for any purpose
# and without fee or royalty is hereby granted, provided that you agree to comply with the following
# copyright notice and statements, including the disclaimer, and that the same appear on ALL copies of
# the software, database and documentation, including modifications that you make for internal use or
# for distribution. WordNet 3.0 Copyright 2006 by Princeton University. All rights reserved. THIS
# SOFTWARE AND DATABASE IS PROVIDED "AS IS" AND PRINCETON UNIVERSITY MAKES NO REPRESENTATIONS OR
# WARRANTIES, EXPRESS OR IMPLIED. BY WAY OF EXAMPLE, BUT NOT LIMITATION, PRINCETON UNIVERSITY MAKES NO
# REPRESENTATIONS OR WARRANTIES OF MERCHANT- ABILITY OR FITNESS FOR ANY PARTICULAR PURPOSE OR THAT THE
# USE OF THE LICENSED SOFTWARE, DATABASE OR DOCUMENTATION WILL NOT INFRINGE ANY THIRD PARTY PATENTS,
# COPYRIGHTS, TRADEMARKS OR OTHER RIGHTS. The name of Princeton University or Princeton may not be used
# in advertising or publicity pertaining to distribution of the software and/or database. Title to
# copyright in this software, database and any associated documentation shall at all times remain with
# Princeton University and LICENSEE agrees to preserve same.
`;

const BAHASA_NOTICE = `# Indonesian: Wordnet Bahasa 1.2, http://wn-msa.sourceforge.net/ (meanings; the words checked by hand)
# The Wordnet Bahasa is released under the MIT License (MIT)
# Copyright (c) 2014-2015 Francis Bond, David Moeljadi
# Copyright (c) 2013 Francis Bond, Muhammad Zulhelmy Mohamed Rosman
# Copyright (c) 2012 Francis Bond, Nurril Hirfana Mohamed Noor, Suerya Sapuan, Muhammad Zulhelmy Mohamed Rosman, Hammam Riza, Tan Enya Kong
# Copyright (c) 2011 Francis Bond, Nurril Hirfana Mohamed Noor, Suerya Sapuan
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
# associated documentation files (the "Software"), to deal in the Software without restriction,
# including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
# and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
# subject to the following conditions: The above copyright notice and this permission notice shall be
# included in all copies or substantial portions of the Software. THE SOFTWARE IS PROVIDED "AS IS",
# WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
# COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF
# CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
# OTHER DEALINGS IN THE SOFTWARE. (The Wordnet Bahasa consists mainly of data rather than code; for the
# purposes of the license the data is referred to as software.)
`;

const word = (w) => w.replace(/\(\w+\)$/, '').replace(/_/g, ' ');

/** Writes `synsets` ([{ pos, words, similar: [keys], opposite: [keys] }], keyed by .key) as lines. */
function write(file, notice, synsets) {
  const offers = (s) => s.words.length > 1 || s.similar.length || s.opposite.length;
  const kept = synsets.filter(offers);
  const line = new Map(kept.map((s, i) => [s.key, i]));
  const refs = (keys) => [...new Set(keys.map((k) => line.get(k)).filter((i) => i != null))].map((i) => i.toString(36)).join(',');
  const body = kept.map((s) => `${s.pos}\t${s.words.join('|')}\t${refs(s.similar)}\t${refs(s.opposite)}`.replace(/\t+$/, ''));
  writeFileSync(file, `${notice}${body.join('\n')}\n`);
  console.log(`${file}: ${kept.length} meanings, ${(Buffer.byteLength(body.join('\n')) / 1e6).toFixed(1)} MB`);
}

// ---------------------------------------------------------------- English, from WordNet's data files

const english = [];
for (const [file, pos] of [['data.noun', 'n'], ['data.verb', 'v'], ['data.adj', 'a'], ['data.adv', 'r']]) {
  for (const raw of readFileSync(join(dict, file), 'latin1').split('\n')) {
    if (!raw || raw.startsWith(' ')) continue; // the licence at the top of each file
    const f = raw.split(' | ')[0].trim().split(/\s+/);
    const offset = f[0];
    const count = parseInt(f[3], 16);
    const words = [];
    for (let i = 0; i < count; i++) words.push(word(f[4 + i * 2]));
    let at = 4 + count * 2;
    const pointers = Number(f[at++]);
    const similar = [];
    const opposite = [];
    for (let i = 0; i < pointers; i++, at += 4) {
      const [symbol, target, targetPos] = f.slice(at, at + 3);
      const key = `${targetPos === 's' ? 'a' : targetPos}${target}`;
      // Similar meanings: an adjective's "similar to" and a verb's group; opposites: antonyms.
      if (symbol === '&' || symbol === '$') similar.push(key);
      else if (symbol === '!') opposite.push(key);
    }
    english.push({ key: `${pos}${offset}`, pos, words: [...new Set(words)], similar, opposite });
  }
}
// The meanings most used first: by how often their words were found in that meaning in WordNet's
// tagged texts (index.sense's counts), so "big" offers "large" before "pregnant".
const used = new Map();
const POS = { 1: 'n', 2: 'v', 3: 'a', 4: 'r', 5: 'a' };
for (const raw of readFileSync(join(dict, 'index.sense'), 'latin1').split('\n')) {
  const [key, offset, , count] = raw.split(' ');
  if (!key) continue;
  const synset = `${POS[key.split('%')[1][0]]}${offset}`;
  used.set(synset, (used.get(synset) || 0) + Number(count || 0));
}
english.sort((x, y) => (used.get(y.key) || 0) - (used.get(x.key) || 0));
write('src/data/thesaurus-en.txt', WORDNET_NOTICE, english);

// ---------------------------------------------------------------- Indonesian, from Wordnet Bahasa

const groups = new Map();
for (const raw of readFileSync(indonesian, 'utf8').split('\n')) {
  const [synset, language, quality, lemma] = raw.split('\t');
  // Indonesian (I) or both Indonesian and Malaysian (B), checked by hand (Y).
  if (!lemma || !['I', 'B'].includes(language) || quality !== 'Y') continue;
  const key = synset.replace(/-s$/, '-a');
  if (!groups.has(key)) groups.set(key, { key, pos: key.slice(-1), words: [], similar: [], opposite: [] });
  const w = word(lemma.trim());
  if (!groups.get(key).words.includes(w)) groups.get(key).words.push(w);
}
write('src/data/thesaurus-id.txt', BAHASA_NOTICE, [...groups.values()]);
