// Sources from a reference manager's export, read on this device: BibTeX (.bib, from Zotero,
// Mendeley, JabRef, Google Scholar or a LaTeX project) and RIS (.ris, from EndNote, Zotero,
// Mendeley and most databases). Each entry becomes a source as the Citation dialog makes one
// ({ type, authors: [{ last, first }], corporate, title, year, publisher, city, journal, volume,
// issue, pages, site, url, accessed }), of the three kinds Word's sources here are: a book, a
// journal article or a website. A conference paper or a book chapter is kept as an article in its
// proceedings or book, and a thesis or report as a book from its university or institution.

const SOURCE = () => ({
  type: 'Book', authors: [], corporate: '', title: '', year: '', publisher: '', city: '', journal: '',
  volume: '', issue: '', pages: '', site: '', url: '', accessed: '',
});

// ------------------------------------------------------------------ BibTeX

// LaTeX accents and symbols as the characters they make ({\"o} -> ö).
const ACCENTS = { '"': '̈', "'": '́', '`': '̀', '^': '̂', '~': '̃', '=': '̄', '.': '̇', u: '̆', v: '̌', H: '̋', c: '̧', k: '̨', r: '̊' };
const SYMBOLS = { ss: 'ß', ae: 'æ', AE: 'Æ', oe: 'œ', OE: 'Œ', o: 'ø', O: 'Ø', aa: 'å', AA: 'Å', l: 'ł', L: 'Ł', i: 'ı', j: 'ȷ', '&': '&', '%': '%', $: '$', '#': '#', _: '_' };

/** A BibTeX value's text as it reads: accents made, braces and commands taken off, dashes made. */
export function latexText(value) {
  let t = String(value);
  // \"{o}, {\"o}, \"o, \c{c}, \v s
  t = t.replace(/\{?\\([`'^"~=.]|[uvHckr](?![a-zA-Z]))\s*\{?\s*(\\[ij]|[a-zA-Z])\s*\}?\}?/g, (m, acc, letter) => {
    const base = letter === '\\i' ? 'i' : letter === '\\j' ? 'j' : letter;
    return `${base}${ACCENTS[acc] || ''}`.normalize('NFC');
  });
  t = t.replace(/\\(ss|ae|AE|oe|OE|aa|AA|o|O|l|L|i|j)(?![a-zA-Z])\s?/g, (m, name) => SYMBOLS[name]);
  t = t.replace(/\\([&%$#_])/g, '$1');
  t = t.replace(/\\(?:textit|textbf|emph|textrm|textsc|textsf|texttt|mathrm|url|href)\s*/g, '');
  t = t.replace(/\\[a-zA-Z]+\s*/g, '');
  t = t.replace(/[{}]/g, '').replace(/---/g, '—').replace(/--/g, '–').replace(/~/g, ' ');
  return t.replace(/\s+/g, ' ').trim();
}

/** Splits `text` at top-level occurrences of `sep` (a regex source), outside braces. */
function splitTop(text, sep) {
  const out = [];
  let depth = 0;
  let from = 0;
  const re = new RegExp(sep, 'gi');
  for (let i = 0; i < text.length; i++) {
    if (text[i] === '{') depth++;
    else if (text[i] === '}') depth--;
    else if (depth === 0) {
      re.lastIndex = i;
      const m = re.exec(text);
      if (m && m.index === i) {
        out.push(text.slice(from, i));
        from = i + m[0].length;
        i = from - 1;
      }
    }
  }
  out.push(text.slice(from));
  return out.map((x) => x.trim()).filter(Boolean);
}

/** True if the first brace of `text` closes at its very end ("{World Health Organization}"). */
function wholeBraced(text) {
  if (!text.startsWith('{') || !text.endsWith('}')) return false;
  let depth = 0;
  for (let i = 0; i < text.length; i++) {
    if (text[i] === '{') depth++;
    else if (text[i] === '}' && --depth === 0) return i === text.length - 1;
  }
  return false;
}

/**
 * BibTeX names ("Smith, Jane and Lee, Kim and {World Health Organization}") as { authors, corporate }:
 * a name wholly in braces is an organisation's.
 */
function bibNames(raw) {
  const authors = [];
  let corporate = '';
  for (const name of splitTop(raw, '\\s+and\\s+')) {
    if (wholeBraced(name)) {
      if (!authors.length && !corporate) corporate = latexText(name);
      continue;
    }
    if (/^others$/i.test(name)) continue;
    const parts = splitTop(name, ',');
    if (parts.length >= 2) {
      // "von Last, Jr, First" or "Last, First"
      authors.push({ last: latexText(parts[0]), first: latexText(parts[parts.length - 1]) });
    } else {
      const words = splitTop(name, '\\s+');
      const last = words.pop() || '';
      authors.push({ last: latexText(last), first: latexText(words.join(' ')) });
    }
  }
  return { authors, corporate };
}

/** The entries of a BibTeX file: [{ type ('article'), key, fields: { title: raw value } }]. */
function bibEntries(text) {
  const entries = [];
  const strings = {};
  const re = /@\s*([a-zA-Z]+)\s*([{(])/g;
  let m;
  while ((m = re.exec(text))) {
    const type = m[1].toLowerCase();
    const close = m[2] === '{' ? '}' : ')';
    // The entry's body, to its matching close.
    let depth = 1;
    let i = re.lastIndex;
    for (; i < text.length && depth > 0; i++) {
      const c = text[i];
      if (c === '{' || (c === '(' && close === ')')) depth++;
      else if (c === '}' || (c === ')' && close === ')')) depth--;
    }
    const body = text.slice(re.lastIndex, i - 1);
    re.lastIndex = i;
    if (type === 'comment' || type === 'preamble') continue;
    const fields = {};
    let rest = body;
    let key = '';
    if (type !== 'string') {
      const comma = rest.indexOf(',');
      key = (comma < 0 ? rest : rest.slice(0, comma)).trim();
      rest = comma < 0 ? '' : rest.slice(comma + 1);
    }
    for (const part of splitTop(rest, ',')) {
      const eq = part.indexOf('=');
      if (eq < 0) continue;
      const name = part.slice(0, eq).trim().toLowerCase();
      // value: {…}, "…" or a number or @string name, joined with #
      const value = splitTop(part.slice(eq + 1).trim(), '\\s*#\\s*').map((v) => {
        if (/^\{[\s\S]*\}$/.test(v)) return v.slice(1, -1);
        if (/^"[\s\S]*"$/.test(v)) return v.slice(1, -1);
        return strings[v.toLowerCase()] ?? v;
      }).join('');
      fields[name] = value;
    }
    if (type === 'string') Object.assign(strings, fields);
    else entries.push({ type, key, fields });
  }
  return entries;
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

/** An access date ("2024-03-05") as Word's sources keep one: year, month's name, day ("2024 March 5"). */
function accessedOf(date) {
  const m = /^(\d{4})(?:[-/](\d{1,2})(?:[-/](\d{1,2}))?)?/.exec(String(date || '').trim());
  if (!m) return '';
  return [m[1], m[2] && MONTHS[Number(m[2]) - 1], m[3] && String(Number(m[3]))].filter(Boolean).join(' ');
}

const doiUrl = (doi) => (doi ? (/^https?:/i.test(doi) ? doi : `https://doi.org/${doi.replace(/^doi:\s*/i, '')}`) : '');

/** The sources in a BibTeX file's text. */
export function parseBibtex(text) {
  const out = [];
  for (const { type, fields: f } of bibEntries(text)) {
    const s = SOURCE();
    const get = (...names) => latexText(names.map((n) => f[n]).find((v) => v != null && String(v).trim()) ?? '');
    const names = bibNames(f.author || f.editor || '');
    s.authors = names.authors;
    s.corporate = names.corporate || (!names.authors.length ? get('organization', 'institution') : '');
    s.title = get('title');
    s.year = get('year') || (get('date').match(/\d{4}/)?.[0] ?? '');
    s.url = get('url') || doiUrl(get('doi'));
    if (type === 'article') {
      s.type = 'JournalArticle';
      s.journal = get('journal', 'journaltitle');
    } else if (['inproceedings', 'conference', 'incollection', 'inbook'].includes(type)) {
      s.type = 'JournalArticle';
      s.journal = get('booktitle');
    } else if (['online', 'electronic', 'www', 'webpage'].includes(type) || (type === 'misc' && s.url && !f.publisher)) {
      s.type = 'InternetSite';
      s.site = get('organization', 'publisher', 'howpublished', 'journal').replace(/^\\?url.*$/i, '');
      s.accessed = accessedOf(get('urldate', 'accessed'));
    } else {
      s.type = 'Book';
      s.publisher = get('publisher', 'school', 'institution', 'organization');
      s.city = get('address', 'location');
    }
    if (s.type === 'JournalArticle') {
      s.volume = get('volume');
      s.issue = get('number', 'issue');
      s.pages = get('pages').replace(/–/g, '-');
    } else s.url = s.type === 'Book' ? '' : s.url;
    if (s.title || s.authors.length) out.push(s);
  }
  return out;
}

// ------------------------------------------------------------------ RIS

/** The sources in a RIS file's text. */
export function parseRis(text) {
  const out = [];
  let tags = null;
  for (const line of text.replace(/^﻿/, '').split(/\r?\n/)) {
    const m = /^([A-Z][A-Z0-9])  -\s?(.*)$/.exec(line);
    if (!m) continue;
    const [, tag, value] = m;
    if (tag === 'TY') tags = { TY: [value.trim()] };
    else if (tag === 'ER') {
      if (tags) out.push(risSource(tags));
      tags = null;
    } else if (tags) (tags[tag] ||= []).push(value.trim());
  }
  return out.filter((s) => s.title || s.authors.length);
}

function risSource(t) {
  const s = SOURCE();
  const get = (...names) => names.map((n) => t[n]?.[0]).find(Boolean) || '';
  const kind = get('TY').toUpperCase();
  for (const name of [...(t.AU || []), ...(t.A1 || [])]) {
    const [last, ...first] = name.split(',');
    if (first.length) s.authors.push({ last: last.trim(), first: first.join(',').trim() });
    else if (!s.authors.length && !s.corporate) s.corporate = name.trim();
  }
  s.title = get('TI', 'T1', 'CT');
  s.year = (get('PY', 'Y1', 'DA').match(/\d{4}/) || [''])[0];
  s.url = get('UR', 'L2') || doiUrl(get('DO'));
  const pages = [get('SP'), get('EP')].filter(Boolean).join('-');
  if (['JOUR', 'JFULL', 'MGZN', 'NEWS', 'EJOUR', 'CONF', 'CPAPER', 'CHAP', 'ABST', 'INPR'].includes(kind)) {
    s.type = 'JournalArticle';
    s.journal = get('JO', 'JF', 'T2', 'JA', 'J2', 'BT');
    s.volume = get('VL');
    s.issue = get('IS');
    s.pages = /-/.test(get('SP')) ? get('SP') : pages;
  } else if (['ELEC', 'WEB', 'BLOG'].includes(kind)) {
    s.type = 'InternetSite';
    s.site = get('T2', 'PB', 'DB');
    s.accessed = accessedOf(get('Y2'));
  } else {
    s.type = 'Book';
    s.publisher = get('PB');
    s.city = get('CY', 'PP');
    s.url = '';
  }
  return s;
}

/**
 * The sources in a reference manager's export, by what its text is: BibTeX (entries that start
 * with @) or RIS (lines such as "TY  - JOUR"). [] if it is neither.
 */
export function parseReferences(text) {
  if (/^\s*TY  -/m.test(text)) return parseRis(text);
  if (/@\s*[a-zA-Z]+\s*[{(]/.test(text)) return parseBibtex(text);
  return [];
}
