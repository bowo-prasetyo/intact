// Placeholders typed into a document for whoever fills it in: "[Client name]", "{{date}}",
// "«Amount»" or "<<Amount>>", as templates for quotes, letters and contracts write them. Each is
// found wherever it appears (across formatting, as Find finds text) and listed once, so it can be
// filled in once for every place it is used. Word's own fields (content controls) aren't these:
// they are filled in where they are.

import { textSegments } from './find.js';

// [Text], {{Text}}, «Text» and <<Text>>: on one line, at most 80 characters, with no brackets of
// the same kind inside.
const PATTERN = /\[([^[\]\n]{1,80})\]|\{\{([^{}\n]{1,80})\}\}|«([^«»\n]{1,80})»|<<([^<>\n]{1,80})>>/gu;

/**
 * Whether the text between the brackets names something to fill in: a word of at least two letters
 * ("Name", "Nama Organisasi", "date"), not a citation or reference ("[1]", "[2–4]", "[Smith 2020]"),
 * a check box ("[ ]", "[x]") or an editorial note ("[sic]", "[…]").
 */
function names(inner) {
  const text = inner.trim();
  if (!/\p{L}.*\p{L}/su.test(text)) return false; // at least two letters
  if (/^(sic|sic!|emphasis added|citation needed|ibid\.?|et al\.?)$/iu.test(text)) return false;
  if (/[\s,](1[5-9]|20)\d\d[a-z]?$/u.test(text)) return false; // a reference: "[Smith 2020]", "[Lee et al., 2019a]"
  return true;
}

/**
 * The placeholders in `roots` (an element or a list, in order), each once, in the order they first
 * appear: { text, label, matches, editable }. `text` is the placeholder as written ("[Client name]"),
 * `label` what it asks for ("Client name"), `matches` every place it is (as Find's matches, in
 * document order) and `editable` how many of them can be filled in here.
 */
export function findPlaceholders(roots) {
  const byText = new Map();
  let order = 0;
  for (const segment of [roots].flat().flatMap(textSegments)) {
    for (const m of segment.text.matchAll(PATTERN)) {
      const inner = m[1] ?? m[2] ?? m[3] ?? m[4];
      if (!names(inner)) continue;
      let entry = byText.get(m[0]);
      if (!entry) {
        entry = { text: m[0], label: inner.trim(), matches: [], editable: 0 };
        byText.set(m[0], entry);
      }
      entry.matches.push({ segment, start: m.index, end: m.index + m[0].length, editable: segment.editable, order: order++ });
      if (segment.editable) entry.editable++;
    }
  }
  return [...byText.values()];
}
