// Versioned file names, for downloads: where the browser can't save back to the file opened, each
// download is a new file, and browsers name a second one "Report (1).docx". People who send a
// document back and forth name its versions instead: "Report v2", "Report_v4_JD" (with the
// initials of whoever changed it), as their Word habits have it. This works out the next one.

/** A version at the end of a name: its separator, its word (v, V, ver, rev…), its number, and initials after it. */
const VERSION = /([ _-])((?:v|ver|version|rev|r)\.?[ _-]?)(\d+(?:\.\d+)*)(?:([ _-])([A-Z]{2,4}))?$/i;

/** The separator a name uses between its words: an underscore or a hyphen where it has no spaces. */
function separatorOf(base) {
  if (/ /.test(base)) return ' ';
  if (/_/.test(base)) return '_';
  if (/-/.test(base)) return '-';
  return ' ';
}

/**
 * Initials from a name, for a version's: "Dana Executive" → "DE", "Ana María de la Cruz" → "AMC"
 * (only the capitalised words, as people write their initials), at most four letters; '' if none.
 */
export function initialsOf(name = '') {
  const words = name.trim().split(/\s+/).filter(Boolean);
  const capital = words.filter((w) => /^\p{Lu}/u.test(w));
  const letters = (capital.length ? capital : words).map((w) => [...w][0]).join('').toUpperCase();
  return /^\p{L}{1,4}$/u.test(letters) ? letters : letters.replace(/[^\p{L}]/gu, '').slice(0, 4);
}

/**
 * The next version of `base` (a file name without its extension): its version number one up,
 * written as it is written ("v2" → "v3", "_V09" → "_V10", "rev 3" → "rev 4", "v1.2" → "v1.3"), or "v2" added (the
 * file opened was the first) with the separator the name uses; with `initials`, those after it
 * (replacing any there were), else the initials it had are left as they were.
 */
export function nextVersionName(base, initials = '') {
  const name = base.trim() || 'document';
  const m = VERSION.exec(name);
  if (m) {
    const [whole, sep, word, digits, initialsSep, oldInitials] = m;
    // The last part goes up: v2 → v3, v1.2 → v1.3, v09 → v10.
    const parts = digits.split('.');
    const last = parts.pop();
    const next = [...parts, String(Number(last) + 1).padStart(last.length, '0')].join('.');
    const tail = initials ? `${initialsSep || sep}${initials}` : oldInitials ? `${initialsSep}${oldInitials}` : '';
    return `${name.slice(0, name.length - whole.length)}${sep}${word}${next}${tail}`;
  }
  const sep = separatorOf(name);
  return `${name}${sep}v2${initials ? `${sep}${initials}` : ''}`;
}
