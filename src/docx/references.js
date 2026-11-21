// References, as on Word's References tab: a table of contents, captions, cross-references,
// citations and a bibliography, written as Word writes them (fields: TOC, SEQ, REF, CITATION and
// BIBLIOGRAPHY), and the document's sources as Word keeps them (its Source Manager's list, a
// b:Sources part in customXml). Citations and the bibliography are formatted here in APA (7th
// edition) or MLA (9th edition) style, as Word's styles of those names lay them out; Word formats
// them again with its own when it updates the fields.
import { NS } from './xml.js';

export const BIB_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography';
export const CUSTOM_XML_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
export const CUSTOM_XML_PROPS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps';
export const CUSTOM_XML_PROPS_CT = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml';
export const CUSTOM_XML_DS = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml';

// Word's names for its citation styles, as b:Sources records the one chosen.
export const CITATION_STYLES = {
  APA: { label: 'APA', selected: '\\APASixthEditionOfficeOnline.xsl', version: '6', heading: 'References' },
  MLA: { label: 'MLA', selected: '\\MLASeventhEditionOfficeOnline.xsl', version: '7', heading: 'Works Cited' },
};

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
  const style = name.startsWith('MLA') ? 'MLA' : name.startsWith('APA') || !name ? 'APA' : root.getAttribute('StyleName');
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

/** The names a citation shows: "Smith", "Smith & Lee" (APA) or "Smith and Lee" (MLA), "Smith et al.". */
function citedNames(s, style) {
  if (!s.authors.length) return s.corporate || (style === 'MLA' ? `“${s.title}”` : s.title);
  const lasts = s.authors.map((a) => a.last || a.first);
  if (lasts.length === 1) return lasts[0];
  if (lasts.length === 2) return `${lasts[0]} ${style === 'MLA' ? 'and' : '&'} ${lasts[1]}`;
  return `${lasts[0]} et al.`;
}

/** An in-text citation for `source`: "(Smith, 2020, p. 4)" in APA, "(Smith 4)" in MLA. */
export function formatCitation(source, style, page = '') {
  if (!source) return '(Source not found)';
  const who = citedNames(source, style);
  if (style === 'MLA') return `(${[who, page].filter(Boolean).join(' ')})`;
  const pages = page ? `${/[-–,]/.test(page) ? 'pp.' : 'p.'} ${page}` : '';
  return `(${[who, source.year || 'n.d.', pages].filter(Boolean).join(', ')})`;
}

/**
 * A source's entry in the bibliography, as pieces of text, some italic: [{ text, italic }]. APA:
 * "Smith, J. (2020). *A book*. Press." MLA: "Smith, Jane. *A Book*. Press, 2020."
 */
export function formatEntry(s, style) {
  const out = [];
  const add = (text, italic = false) => { if (text) out.push({ text, italic }); };
  const year = s.year || 'n.d.';
  const pages = s.pages ? s.pages.replace(/-/g, '–') : '';
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
    return out;
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
  return out;
}

/** The sources in the order a bibliography lists them: by first author (or title), then year. */
export function sortedSources(sources) {
  const key = (s) => `${(s.authors[0]?.last || s.corporate || s.title || '').toLowerCase()}\u0000${s.year || ''}`;
  return [...sources].sort((a, b) => key(a).localeCompare(key(b)));
}

/**
 * A bibliography as Word's Bibliography gallery inserts one: a heading ("References" in APA,
 * "Works Cited" in MLA, in `headingStyle`), then the BIBLIOGRAPHY field, one paragraph a source in
 * Word's Bibliography style (`entryStyle`), with a hanging indent of half an inch. Returns the
 * paragraphs as XML: [heading, ...field paragraphs]; `withHeading` false leaves the heading out.
 */
export function bibliographyXml({ sources, style, headingStyle, entryStyle, withHeading = true }) {
  const list = sortedSources(sources);
  const pPr = `<w:pPr><w:pStyle w:val="${esc(entryStyle)}"/><w:ind w:left="720" w:hanging="720"/></w:pPr>`;
  const entryRuns = (s) => formatEntry(s, style).map((p) => runXml(p.text, { italic: p.italic, rPr: '<w:noProof/>' })).join('');
  const entries = list.length ? list.map(entryRuns) : [runXml('There are no sources in the current document.', { rPr: '<w:noProof/>' })];
  const begin = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> BIBLIOGRAPHY </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
  const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const paras = entries.map((runs, i) => `<w:p>${pPr}${i === 0 ? begin : ''}${runs}${i === entries.length - 1 ? end : ''}</w:p>`);
  const heading = `<w:p><w:pPr><w:pStyle w:val="${esc(headingStyle)}"/></w:pPr>${runXml((CITATION_STYLES[style] || CITATION_STYLES.APA).heading)}</w:p>`;
  return withHeading ? [heading, ...paras] : paras;
}

/** A citation field, as Word's Insert Citation writes one (CITATION with the source's tag, a page with \p). */
export function citationXml(source, style, page = '') {
  const instr = ` CITATION ${source.tag}${page ? ` \\p ${page.replace(/\s+/g, '')}` : ''} \\l 1033 `;
  return fieldRuns(instr, runXml(formatCitation(source, style, page), { rPr: '<w:noProof/>' }));
}
