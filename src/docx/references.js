// References, as on Word's References tab: a table of contents, captions, cross-references,
// citations and a bibliography, written as Word writes them (fields: TOC, SEQ, REF, CITATION and
// BIBLIOGRAPHY), and the document's sources as Word keeps them (its Source Manager's list, a
// b:Sources part in customXml). Citations and the bibliography are formatted here in APA (7th
// edition), MLA (9th edition), Chicago (17th edition, author-date), Harvard (Cite Them Right) or
// IEEE style, as those styles' guides lay them out; Word formats them again with its own styles of
// those names when it updates the fields. A citation may cite several sources, as Word's CITATION
// field does with \m.
import { NS } from './xml.js';

export const BIB_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography';
export const CUSTOM_XML_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
export const CUSTOM_XML_PROPS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps';
export const CUSTOM_XML_PROPS_CT = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml';
export const CUSTOM_XML_DS = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml';

// Word's names for its citation styles, as b:Sources records the one chosen.
// `name` is how the Citation dialog offers each.
export const CITATION_STYLES = {
  APA: { label: 'APA', selected: '\\APASixthEditionOfficeOnline.xsl', version: '6', heading: 'References', name: 'APA (7th edition)' },
  MLA: { label: 'MLA', selected: '\\MLASeventhEditionOfficeOnline.xsl', version: '7', heading: 'Works Cited', name: 'MLA (9th edition)' },
  Chicago: { label: 'Chicago', selected: '\\CHICAGO.XSL', version: '16', heading: 'Bibliography', name: 'Chicago (17th edition, author-date)' },
  Harvard: { label: 'Harvard - Anglia', selected: '\\HarvardAnglia2008OfficeOnline.xsl', version: '2008', heading: 'References', name: 'Harvard' },
  IEEE: { label: 'IEEE', selected: '\\IEEE2006OfficeOnline.xsl', version: '2006', heading: 'References', name: 'IEEE (numbered)' },
};

/** The style citations are formatted in here: the document's own if it is one of these, else APA. */
export const citationStyle = (style) => (CITATION_STYLES[style] ? style : 'APA');

// The kinds of source offered, as Word's Create Source dialog names them (b:SourceType).
export const SOURCE_TYPES = {
  Book: 'Book',
  JournalArticle: 'Journal article',
  InternetSite: 'Website',
};

const esc = (t) => String(t ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const W_ATTR = `xmlns:w="${NS.w}"`;

/** A run of text (`italic`, `bold` optional), as XML. */
export function runXml(text, { italic = false, bold = false, rPr = '' } = {}) {
  const props = `${bold ? '<w:b/><w:bCs/>' : ''}${italic ? '<w:i/><w:iCs/>' : ''}${rPr}`;
  return `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

/** A complex field as runs: begin, instruction, separate, `resultRuns` (XML), end. */
export function fieldRuns(instr, resultRuns, rPr = '') {
  const props = rPr ? `<w:rPr>${rPr}</w:rPr>` : '';
  return `<w:r>${props}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${props}<w:instrText xml:space="preserve">${esc(instr)}</w:instrText></w:r>`
    + `<w:r>${props}<w:fldChar w:fldCharType="separate"/></w:r>${resultRuns}<w:r>${props}<w:fldChar w:fldCharType="end"/></w:r>`;
}

/** Paragraphs (XML strings) wrapped to parse together, in the w namespace. */
export const bodyXml = (paras) => `<w:body ${W_ATTR}>${paras.join('')}</w:body>`;

// ------------------------------------------------------------------ table of contents

/**
 * A table of contents as Word's Automatic Table gallery inserts one: its heading ("Contents", in
 * Word's TOC Heading style), then the TOC field over headings 1 to 3, linked (\h), its entries
 * filled in by updating it. `headingStyle` is the TOC Heading style's id.
 */
export function tocXml(headingStyle) {
  return [
    `<w:p><w:pPr><w:pStyle w:val="${esc(headingStyle)}"/></w:pPr>${runXml('Contents')}</w:p>`,
    `<w:p>${fieldRuns(' TOC \\o "1-3" \\h \\z \\u ', runXml('Update the table of contents to list the headings.'))}</w:p>`,
  ];
}

/** The heading a table of figures is given here, by its caption label. */
export const FIGURES_HEADINGS = { Figure: 'List of Figures', Table: 'List of Tables', Equation: 'List of Equations' };

/**
 * A table of figures (a list of figures or of tables), as Word's Insert Table of Figures writes one:
 * the TOC field with \c and the caption label, linked (\h), its entries the captions of that label,
 * filled in by updating it. Word adds no heading; a thesis's front matter has one, so it is put
 * before it ("List of Figures"), in Word's TOC Heading style (`headingStyle`), which isn't listed in
 * a table of contents.
 */
export function tableOfFiguresXml(headingStyle, label) {
  return [
    `<w:p><w:pPr><w:pStyle w:val="${esc(headingStyle)}"/></w:pPr>${runXml(FIGURES_HEADINGS[label] || `List of ${label}s`)}</w:p>`,
    `<w:p>${fieldRuns(` TOC \\h \\z \\c "${label}" `, runXml(`Update the list to show the ${label.toLowerCase()} captions.`))}</w:p>`,
  ];
}

// ------------------------------------------------------------------ captions

// Word's caption labels (Insert Caption's Label list).
export const CAPTION_LABELS = ['Figure', 'Table', 'Equation'];

/**
 * A caption, as Word's Insert Caption writes one: "Figure " and its number (a SEQ field), then the
 * text after it (": Sales by region"), in Word's Caption style (`style`). `n` is the number shown.
 */
export function captionXml(style, label, n, text) {
  const after = text ? `: ${text}` : '';
  return `<w:p><w:pPr><w:pStyle w:val="${esc(style)}"/></w:pPr>${runXml(`${label} `)}`
    + `${fieldRuns(` SEQ ${label} \\* ARABIC `, `<w:r><w:rPr><w:noProof/></w:rPr><w:t>${n}</w:t></w:r>`)}`
    + `${after ? runXml(after) : ''}</w:p>`;
}

// ------------------------------------------------------------------ sources

/** The text of b:`name` in `el`, or ''. */
function bText(el, name) {
  const c = [...(el?.children || [])].find((x) => x.localName === name && x.namespaceURI === BIB_NS);
  return c ? c.textContent.trim() : '';
}
const bChild = (el, name) => [...(el?.children || [])].find((x) => x.localName === name && x.namespaceURI === BIB_NS) || null;

/** True if `xmlDoc` is a Word sources part (b:Sources). */
export const isSourcesPart = (xmlDoc) => xmlDoc?.documentElement?.namespaceURI === BIB_NS && xmlDoc.documentElement.localName === 'Sources';

/**
 * The sources in a b:Sources part: { style ('APA', 'MLA' or the part's own name), sources: [{ tag,
 * type, authors: [{ last, first }], corporate, title, year, publisher, city, journal, volume, issue,
 * pages, site, url, accessed }] }.
 */
export function readSources(xmlDoc) {
  const root = xmlDoc.documentElement;
  const name = (root.getAttribute('StyleName') || '').toUpperCase();
  const selected = (root.getAttribute('SelectedStyle') || '').toUpperCase();
  const known = Object.keys(CITATION_STYLES).find((k) => name.startsWith(k.toUpperCase()) || (!name && selected.includes(k.toUpperCase())));
  const style = known || (!name ? 'APA' : root.getAttribute('StyleName'));
  const sources = [];
  for (const s of root.children) {
    if (s.localName !== 'Source' || s.namespaceURI !== BIB_NS) continue;
    const authorEl = bChild(bChild(bChild(s, 'Author'), 'Author'), 'NameList');
    const corporate = bText(bChild(bChild(s, 'Author'), 'Author'), 'Corporate');
    const authors = authorEl ? [...authorEl.children].filter((p) => p.localName === 'Person')
      .map((p) => ({ last: bText(p, 'Last'), first: [bText(p, 'First'), bText(p, 'Middle')].filter(Boolean).join(' ') })) : [];
    const y = bText(s, 'YearAccessed');
    const m = bText(s, 'MonthAccessed');
    const d = bText(s, 'DayAccessed');
    sources.push({
      tag: bText(s, 'Tag'), type: bText(s, 'SourceType') || 'Book', authors, corporate,
      title: bText(s, 'Title'), year: bText(s, 'Year'), publisher: bText(s, 'Publisher'), city: bText(s, 'City'),
      journal: bText(s, 'JournalName'), volume: bText(s, 'Volume'), issue: bText(s, 'Issue'), pages: bText(s, 'Pages'),
      site: bText(s, 'InternetSiteTitle') || bText(s, 'ProductionCompany'), url: bText(s, 'URL'),
      accessed: y ? [y, m, d].filter(Boolean).join(' ') : '',
    });
  }
  return { style, sources };
}

/** A tag for a new source, as Word makes one: the author's first three letters and the year's last two ("Smi20"). */
export function newTag(source, taken) {
  const who = (source.authors[0]?.last || source.corporate || source.title || 'Src').replace(/[^\p{L}\p{N}]/gu, '').slice(0, 3) || 'Src';
  const base = `${who[0].toUpperCase()}${who.slice(1).toLowerCase()}${(source.year || '').slice(-2)}`;
  let tag = base;
  for (let i = 0; taken.has(tag); i++) tag = `${base}${String.fromCharCode(97 + (i % 26))}${i >= 26 ? Math.floor(i / 26) : ''}`;
  return tag;
}

/** The b:Source element for `source` (and its order in the list, `order`), as XML. */
function sourceXml(source, order) {
  const el = (name, value) => (value ? `<b:${name}>${esc(value)}</b:${name}>` : '');
  const people = source.authors.filter((a) => a.last || a.first)
    .map((a) => `<b:Person>${el('Last', a.last)}${el('First', a.first)}</b:Person>`).join('');
  const author = people ? `<b:Author><b:Author><b:NameList>${people}</b:NameList></b:Author></b:Author>`
    : source.corporate ? `<b:Author><b:Author>${el('Corporate', source.corporate)}</b:Author></b:Author>` : '';
  const [ay, am, ad] = (source.accessed || '').split(/[ -]/);
  return `<b:Source>${el('Tag', source.tag)}${el('SourceType', source.type)}${el('Guid', source.guid || '')}${author}`
    + `${el('Title', source.title)}${el('Year', source.year)}${el('City', source.city)}${el('Publisher', source.publisher)}`
    + `${el('JournalName', source.journal)}${el('Volume', source.volume)}${el('Issue', source.issue)}${el('Pages', source.pages)}`
    + `${el('InternetSiteTitle', source.site)}${el('URL', source.url)}${el('YearAccessed', ay)}${el('MonthAccessed', am)}${el('DayAccessed', ad)}`
    + `<b:RefOrder>${order}</b:RefOrder></b:Source>`;
}

/** A whole b:Sources part for `sources`, in citation style `style`, as Word writes it. */
export function sourcesXml(style, sources) {
  const s = CITATION_STYLES[style] || CITATION_STYLES.APA;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<b:Sources xmlns:b="${BIB_NS}" xmlns="${BIB_NS}" SelectedStyle="${esc(s.selected)}" StyleName="${s.label}" Version="${s.version}">`
    + `${sources.map((src, i) => sourceXml(src, i + 1)).join('')}</b:Sources>`;
}

/** The properties part Word writes beside a sources part (its id and schema). */
export function sourcesPropsXml(guid) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="${guid}" xmlns:ds="${CUSTOM_XML_DS}"><ds:schemaRefs><ds:schemaRef ds:uri="${BIB_NS}"/></ds:schemaRefs></ds:datastoreItem>`;
}

/** A GUID in braces, as Word writes them. */
export function newGuid() {
  const hex = [...crypto.getRandomValues(new Uint8Array(16))].map((b) => b.toString(16).padStart(2, '0')).join('').toUpperCase();
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}}`;
}

// ------------------------------------------------------------------ formatting citations

const initials = (first) => first.split(/\s+/).filter(Boolean).map((n) => n.split('-').map((p) => `${p[0].toUpperCase()}.`).join('-')).join(' ');
const endsSentence = (t) => /[.?!]$/.test(t.trim());
/** `text` and a full stop, unless it ends in one already (or in ? or !). */
const stop = (text) => (endsSentence(text) ? text : `${text}.`);

/** The authors as APA lists them: "Smith, J. A., Lee, K., & Park, S.". */
function apaAuthors(s) {
  if (!s.authors.length) return s.corporate || '';
  const names = s.authors.map((a) => [a.last, initials(a.first || '')].filter(Boolean).join(', '));
  if (names.length === 1) return names[0];
  if (names.length <= 20) return `${names.slice(0, -1).join(', ')}, & ${names[names.length - 1]}`;
  return `${names.slice(0, 19).join(', ')}, … ${names[names.length - 1]}`;
}

/** The authors as MLA lists them: "Smith, Jane", "Smith, Jane, and Kim Lee", "Smith, Jane, et al.". */
function mlaAuthors(s) {
  if (!s.authors.length) return s.corporate || '';
  const [a, b] = s.authors;
  const firstOne = [a.last, a.first].filter(Boolean).join(', ');
  if (s.authors.length === 1) return firstOne;
  if (s.authors.length === 2) return `${firstOne}, and ${[b.first, b.last].filter(Boolean).join(' ')}`;
  return `${firstOne}, et al.`;
}

/** The names a citation shows, as each style names one, two, three or more authors. */
function citedNames(s, style) {
  if (!s.authors.length) return s.corporate || (style === 'MLA' || style === 'Chicago' ? `“${s.title}”` : s.title);
  const l = s.authors.map((a) => a.last || a.first);
  if (l.length === 1) return l[0];
  if (style === 'APA') return l.length === 2 ? `${l[0]} & ${l[1]}` : `${l[0]} et al.`;
  if (style === 'MLA') return l.length === 2 ? `${l[0]} and ${l[1]}` : `${l[0]} et al.`;
  // Chicago and Harvard name up to three authors, and four or more as the first et al.
  if (l.length >= 4) return `${l[0]} et al.`;
  if (l.length === 2) return `${l[0]} and ${l[1]}`;
  return style === 'Chicago' ? `${l[0]}, ${l[1]}, and ${l[2]}` : `${l[0]}, ${l[1]} and ${l[2]}`;
}

const pagesLabel = (page) => `${/[-–,]/.test(page) ? 'pp.' : 'p.'} ${page}`;

/** One source's part of an in-text citation, with the page cited (`page`) and its number in IEEE (`n`). */
function citedPart(s, style, page, n) {
  const who = citedNames(s, style);
  const year = s.year || 'n.d.';
  if (style === 'MLA') return [who, page].filter(Boolean).join(' ');
  if (style === 'Chicago') return `${who} ${year}${page ? `, ${page}` : ''}`;
  if (style === 'IEEE') return `${n ?? '#'}${page ? `, ${pagesLabel(page)}` : ''}`;
  return [who, year, page && pagesLabel(page)].filter(Boolean).join(', '); // APA, Harvard
}

/**
 * An in-text citation for `sources` (one, or several cited together), the page cited (`page`) going
 * with the first: "(Smith, 2020, p. 4; Lee, 2021)" in APA and Harvard, "(Smith 4; Lee)" in MLA,
 * "(Smith 2020, 4)" in Chicago, "[1, p. 4], [3]" in IEEE, numbered as `numbers` (tag -> number)
 * says; a source without one shows #.
 */
export function formatCitation(sources, style, page = '', numbers = null) {
  const list = Array.isArray(sources) ? sources : [sources];
  if (!list.length || list.some((s) => !s)) return '(Source not found)';
  const parts = list.map((s, i) => citedPart(s, style, i === 0 ? page : '', numbers?.get(s.tag)));
  return style === 'IEEE' ? parts.map((p) => `[${p}]`).join(', ') : `(${parts.join('; ')})`;
}

const natural = (a) => [a.first, a.last].filter(Boolean).join(' ');
const inverted = (a) => [a.last, a.first].filter(Boolean).join(', ');
/** "A, B, and C" (with `serial` false, "A, B and C"). */
const listed = (names, serial = true) => (names.length <= 2 ? names.join(' and ') : `${names.slice(0, -1).join(', ')}${serial ? ',' : ''} and ${names[names.length - 1]}`);

/** The authors as Chicago lists them: "Smith, Jane", "Smith, Jane, and Kim Lee"; more than ten, the first seven et al. */
function chicagoAuthors(s) {
  if (!s.authors.length) return s.corporate || '';
  const many = s.authors.length > 10;
  const names = (many ? s.authors.slice(0, 7) : s.authors).map((a, i) => (i ? natural(a) : inverted(a)));
  if (many) return `${names.join(', ')}, et al.`;
  return names.length === 2 ? `${names[0]}, and ${names[1]}` : listed(names);
}

/** The authors as Harvard (Cite Them Right) lists them: "Smith, J.A.", "Smith, J., Lee, K. and Park, S."; four or more, the first et al. */
function harvardAuthors(s) {
  if (!s.authors.length) return s.corporate || '';
  const names = s.authors.map((a) => [a.last, initials(a.first || '').replace(/ /g, '')].filter(Boolean).join(', '));
  return names.length >= 4 ? `${names[0]} et al.` : listed(names, false);
}

/** The authors as IEEE lists them: "J. A. Smith", "J. Smith and K. Lee", "J. Smith, K. Lee, and S. Park"; seven or more, the first et al. */
function ieeeAuthors(s) {
  if (!s.authors.length) return s.corporate || '';
  const names = s.authors.map((a) => [initials(a.first || ''), a.last].filter(Boolean).join(' '));
  return names.length >= 7 ? `${names[0]} et al.` : listed(names);
}

/**
 * A source's entry in the bibliography, as pieces of text, some italic: [{ text, italic }]. APA:
 * "Smith, J. (2020). *A book*. Press." MLA: "Smith, Jane. *A Book*. Press, 2020." Chicago: "Smith,
 * Jane. 2020. *A Book*. City: Press." Harvard: "Smith, J. (2020) *A book*. City: Press." IEEE:
 * "J. Smith, *A Book*. City: Press, 2020." (its number is put before it by the bibliography).
 */
export function formatEntry(s, style) {
  const out = [];
  const add = (text, italic = false) => { if (text) out.push({ text, italic }); };
  const year = s.year || 'n.d.';
  const pages = s.pages ? s.pages.replace(/-/g, '–') : '';
  const place = [s.city, s.publisher].filter(Boolean).join(': ');
  const done = () => {
    if (out.length) out[out.length - 1].text = out[out.length - 1].text.trimEnd();
    return out;
  };
  if (style === 'MLA') {
    const who = mlaAuthors(s);
    if (who) add(`${stop(who)} `);
    if (s.type === 'Book') {
      add(stop(s.title), true);
      add([' ', [s.publisher, s.year].filter(Boolean).join(', ')].join(''));
      if (s.publisher || s.year) add('.');
    } else if (s.type === 'JournalArticle') {
      add(`“${stop(s.title)}” `);
      add(s.journal, true);
      const rest = [s.volume && `vol. ${s.volume}`, s.issue && `no. ${s.issue}`, s.year, pages && `pp. ${pages}`].filter(Boolean).join(', ');
      add(rest ? `, ${rest}.` : '.');
      if (s.url) add(` ${s.url}.`);
    } else {
      add(`“${stop(s.title)}” `);
      add(s.site, true);
      const rest = [s.year, s.url].filter(Boolean).join(', ');
      add(s.site ? (rest ? `, ${rest}.` : '.') : (rest ? `${rest}.` : ''));
      if (s.accessed) add(` Accessed ${s.accessed}.`);
    }
    return done();
  }
  if (style === 'Chicago') {
    const who = chicagoAuthors(s);
    if (who) add(`${stop(who)} ${year}. `);
    if (s.type === 'Book') {
      add(stop(s.title), true);
      if (!who) add(` ${year}.`);
      if (place) add(` ${stop(place)}`);
    } else if (s.type === 'JournalArticle') {
      add(`“${stop(s.title)}”`);
      if (!who) add(` ${year}.`);
      add(' ');
      add(s.journal, true);
      if (s.volume) add(` ${s.volume}`);
      if (s.issue) add(` (${s.issue})`);
      if (pages) add(`: ${pages}`);
      add('.');
      if (s.url) add(` ${s.url}.`);
    } else {
      add(`“${stop(s.title)}”`);
      if (!who) add(` ${year}.`);
      if (s.site) add(` ${stop(s.site)}`);
      if (s.accessed) add(` Accessed ${s.accessed}.`);
      if (s.url) add(` ${s.url}.`);
    }
    return done();
  }
  if (style === 'Harvard') {
    const who = harvardAuthors(s);
    if (who) add(`${who} (${year}) `);
    const titleYear = () => { if (!who) add(` (${year})`); };
    if (s.type === 'Book') {
      add(s.title, true);
      titleYear();
      add('. ');
      if (place) add(stop(place));
    } else if (s.type === 'JournalArticle') {
      add(`‘${s.title}’`);
      titleYear();
      add(', ');
      add(s.journal, true);
      const rest = [s.volume && `${s.volume}${s.issue ? `(${s.issue})` : ''}`, pages && `pp. ${pages}`].filter(Boolean).join(', ');
      add(rest ? `, ${rest}.` : '.');
      if (s.url) add(` Available at: ${s.url}.`);
    } else {
      add(s.title, true);
      titleYear();
      add('. ');
      if (s.site && s.site !== who) add(`${stop(s.site)} `);
      if (s.url) add(`Available at: ${s.url}${s.accessed ? ` (Accessed: ${s.accessed})` : ''}.`);
    }
    return done();
  }
  if (style === 'IEEE') {
    const who = ieeeAuthors(s);
    if (s.type === 'Book') {
      if (who) add(`${who}, `);
      add(s.title, true);
      const rest = [place, s.year].filter(Boolean).join(', ');
      add(rest ? `. ${rest}.` : '.');
    } else if (s.type === 'JournalArticle') {
      add(`${who ? `${who}, ` : ''}“${s.title},” `);
      add(s.journal, true);
      const rest = [s.volume && `vol. ${s.volume}`, s.issue && `no. ${s.issue}`, pages && `pp. ${pages}`, s.year].filter(Boolean).join(', ');
      add(rest ? `, ${rest}.` : '.');
      if (s.url) add(` ${s.url}`);
    } else {
      add(`${who ? `${who}, ` : ''}“${stop(s.title)}” `);
      if (s.site) add(`${stop(s.site)} `);
      if (s.accessed) add(`Accessed: ${s.accessed}. `);
      if (s.url) add(`[Online]. Available: ${s.url}`);
    }
    return done();
  }
  // APA
  const who = apaAuthors(s);
  if (who) add(`${stop(who)} (${year}). `);
  if (s.type === 'Book') {
    add(stop(s.title), true);
    if (!who) add(` (${year}).`);
    if (s.publisher) add(` ${stop(s.publisher)}`);
  } else if (s.type === 'JournalArticle') {
    add(stop(s.title));
    if (!who) add(` (${year}).`);
    add(' ');
    add(s.journal, true);
    if (s.volume) {
      add(', ');
      add(s.volume, true);
    }
    if (s.issue) add(`(${s.issue})`);
    if (pages) add(`, ${pages}`);
    add('.');
    if (s.url) add(` ${s.url}`);
  } else {
    add(stop(s.title), true);
    if (!who) add(` (${year}).`);
    if (s.site && s.site !== who) add(` ${stop(s.site)}`);
    if (s.url) add(` ${s.url}`);
  }
  return done();
}

/**
 * Numbers for IEEE's citations: each source's tag numbered in the order it is first cited, from
 * `cited` (the tags each citation cites, in the order of the text). Returns tag -> number.
 */
export function numberCitations(cited) {
  const numbers = new Map();
  for (const tags of cited) for (const tag of tags) if (!numbers.has(tag)) numbers.set(tag, numbers.size + 1);
  return numbers;
}

/**
 * The sources in the order a bibliography lists them: by first author (or title), then year; in
 * IEEE, by their numbers (`numbers`), those not cited after them in the order they were added.
 */
export function sortedSources(sources, style = 'APA', numbers = null) {
  if (style === 'IEEE') {
    const n = (s) => numbers?.get(s.tag) ?? Infinity;
    return sources.map((s, i) => [s, i]).sort((a, b) => n(a[0]) - n(b[0]) || a[1] - b[1]).map(([s]) => s);
  }
  const key = (s) => `${(s.authors[0]?.last || s.corporate || s.title || '').toLowerCase()}\u0000${s.year || ''}`;
  return [...sources].sort((a, b) => key(a).localeCompare(key(b)));
}

/**
 * A bibliography as Word's Bibliography gallery inserts one: a heading (the style's: "References"
 * in APA, Harvard and IEEE, "Works Cited" in MLA, "Bibliography" in Chicago, in `headingStyle`),
 * then the BIBLIOGRAPHY field, one paragraph a source in Word's Bibliography style (`entryStyle`),
 * with a hanging indent of half an inch; in IEEE each starts with its number ("[1]") and a tab.
 * Returns the paragraphs as XML: [heading, ...field paragraphs]; `withHeading` false leaves the
 * heading out. `numbers` numbers IEEE's sources (see numberCitations).
 */
export function bibliographyXml({ sources, style, headingStyle, entryStyle, withHeading = true, numbers = null }) {
  const list = sortedSources(sources, style, numbers);
  const pPr = `<w:pPr><w:pStyle w:val="${esc(entryStyle)}"/><w:ind w:left="720" w:hanging="720"/></w:pPr>`;
  let next = Math.max(0, ...(numbers ? numbers.values() : [])) + 1;
  const entryRuns = (s) => {
    const prefix = style === 'IEEE' ? `${runXml(`[${numbers?.get(s.tag) ?? next++}]`, { rPr: '<w:noProof/>' })}<w:r><w:rPr><w:noProof/></w:rPr><w:tab/></w:r>` : '';
    return prefix + formatEntry(s, style).map((p) => runXml(p.text, { italic: p.italic, rPr: '<w:noProof/>' })).join('');
  };
  const entries = list.length ? list.map(entryRuns) : [runXml('There are no sources in the current document.', { rPr: '<w:noProof/>' })];
  const begin = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> BIBLIOGRAPHY </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
  const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const paras = entries.map((runs, i) => `<w:p>${pPr}${i === 0 ? begin : ''}${runs}${i === entries.length - 1 ? end : ''}</w:p>`);
  const heading = `<w:p><w:pPr><w:pStyle w:val="${esc(headingStyle)}"/></w:pPr>${runXml(CITATION_STYLES[citationStyle(style)].heading)}</w:p>`;
  return withHeading ? [heading, ...paras] : paras;
}

/**
 * A citation field's sources and page, from its instruction (" CITATION Smi20 \\p 4 \\l 1033 \\m
 * Lee21 "): { tags: ['Smi20', 'Lee21'], page: '4' }, or null if it isn't one.
 */
export function parseCitation(instr) {
  const tokens = [...(instr || '').matchAll(/"([^"]*)"|(\S+)/g)].map((m) => m[1] ?? m[2]);
  if ((tokens.shift() || '').toUpperCase() !== 'CITATION') return null;
  const tags = [];
  let page = '';
  if (tokens[0] && !tokens[0].startsWith('\\')) tags.push(tokens.shift());
  for (let i = 0; i < tokens.length; i++) {
    const sw = tokens[i].toLowerCase();
    if (sw === '\\m' && tokens[i + 1]) tags.push(tokens[++i]);
    else if (sw === '\\p' && tokens[i + 1]) {
      if (!page) page = tokens[i + 1];
      i++;
    } else if (['\\l', '\\s', '\\v', '\\f'].includes(sw)) i++;
  }
  return { tags, page };
}

/**
 * A citation field, as Word's Insert Citation writes one: CITATION with the source's tag, a page
 * with \\p, and each further source cited with it after \\m. `sources` is one source or several.
 */
export function citationXml(sources, style, page = '', numbers = null) {
  const list = Array.isArray(sources) ? sources : [sources];
  const more = list.slice(1).map((s) => ` \\m ${s.tag}`).join('');
  const instr = ` CITATION ${list[0].tag}${page ? ` \\p ${page.replace(/\s+/g, '')}` : ''} \\l 1033${more} `;
  return fieldRuns(instr, runXml(formatCitation(list, style, page, numbers), { rPr: '<w:noProof/>' }));
}
