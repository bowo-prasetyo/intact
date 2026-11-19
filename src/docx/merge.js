// Mail merge from a CSV, as Word's Mailings tab does it with a list: each row of the list fills the
// document's placeholders ("[Recipient name]", see editor/placeholders.js), giving either one
// document with every row's copy on a page of its own (Word's "Edit individual documents") or a
// separate document for each row, in a .zip. Everything happens here, on the document as it would
// be downloaded: nothing is uploaded, and the document itself isn't changed.
import JSZip from 'jszip';
import { NS, parseXml, serializeXml, createW, wChild, ensureOrdered, PPR_ORDER } from './xml.js';
import { PATTERN, names } from '../editor/placeholders.js';

const W = NS.w;
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const MAX_ROWS = 1000;

// ---------------------------------------------------------------- reading the list

/**
 * The text of a CSV file: UTF-8 (with or without a byte order mark, as Excel's "CSV UTF-8" saves
 * it), or, when it isn't valid UTF-8, Windows-1252, as Excel's plain "CSV" saves it on Windows.
 */
export function decodeCsv(bytes) {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, '');
  } catch {
    return new TextDecoder('windows-1252').decode(bytes);
  }
}

/**
 * A CSV's columns (its first row, named) and rows, as RFC 4180 reads them: fields in double quotes
 * may hold the separator, line breaks and doubled quotes. The separator is the one the header row
 * uses most of: a comma, a semicolon (Excel where the decimal mark is a comma) or a tab. Empty
 * rows are left out. { columns: [string], rows: [[string]], separator }.
 */
export function parseCsv(text) {
  text = text.replace(/^﻿/, '');
  const firstLine = text.split(/\r\n|\n|\r/, 1)[0].replace(/"[^"]*"/g, '');
  const separator = [',', ';', '\t'].map((c) => [c, firstLine.split(c).length - 1]).sort((a, b) => b[1] - a[1])[0][0];
  const records = [];
  let field = '';
  let record = [];
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"' && text[i + 1] === '"') { field += '"'; i++; } else if (c === '"') quoted = false;
      else field += c;
    } else if (c === '"' && field === '') quoted = true;
    else if (c === separator) { record.push(field); field = ''; } else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      record.push(field);
      records.push(record);
      record = [];
      field = '';
    } else field += c;
  }
  if (field !== '' || record.length) {
    record.push(field);
    records.push(record);
  }
  const nonEmpty = records.filter((r) => r.some((f) => f.trim()));
  const header = nonEmpty.shift() || [];
  const seen = new Map();
  const columns = header.map((h, i) => {
    const name = h.trim() || `Column ${i + 1}`;
    const n = (seen.get(name) || 0) + 1;
    seen.set(name, n);
    return n > 1 ? `${name} (${n})` : name;
  });
  const width = Math.max(columns.length, ...nonEmpty.map((r) => r.length));
  for (let i = columns.length; i < width; i++) columns.push(`Column ${i + 1}`);
  return { columns, rows: nonEmpty.map((r) => columns.map((_, i) => (r[i] ?? '').replace(/\r\n?/g, '\n'))), separator };
}

/** What a name says, for matching a placeholder to a column: letters and digits, lower case, no accents. */
const key = (s) => s.normalize('NFKD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^\p{L}\p{N}]+/gu, '');

/**
 * Which column fills each placeholder: the one of the same name, whatever its case, spaces,
 * underscores or accents ("Recipient name" for "recipient_name"); else none (-1).
 * `labels`: what the placeholders ask for. Returns a column index for each.
 */
export function matchColumns(labels, columns) {
  const keys = columns.map(key);
  return labels.map((label) => keys.indexOf(key(label)));
}

// ---------------------------------------------------------------- filling the XML

const INSIDE_DELETION = (t) => {
  for (let n = t.parentNode; n && n.localName !== 'p'; n = n.parentNode) if (n.namespaceURI === W && (n.localName === 'del' || n.localName === 'moveFrom')) return true;
  return false;
};

/**
 * Fills the placeholders in the paragraphs under `root` (an XML element or document) with
 * `values` (a Map from the placeholder as written, "[Recipient name]", to what goes in it), as
 * Fill in does in the page: the text goes into the run where the placeholder starts, so it takes
 * that run's formatting, also where the placeholder runs across several runs. A value with line
 * breaks (an address) gets Word's line breaks. Returns how many placeholders were filled.
 */
export function fillXml(root, values) {
  const doc = root.ownerDocument || root;
  let filled = 0;
  for (const p of [...root.getElementsByTagNameNS(W, 'p')]) {
    const nodes = [];
    let text = '';
    for (const t of p.getElementsByTagNameNS(W, 't')) {
      // This paragraph's own text: not a text box's paragraph inside it, nor deleted text.
      let owner = t.parentNode;
      while (owner && !(owner.namespaceURI === W && owner.localName === 'p')) owner = owner.parentNode;
      if (owner !== p || INSIDE_DELETION(t)) continue;
      nodes.push({ t, start: text.length });
      text += t.textContent;
    }
    if (!text) continue;
    const hits = [...text.matchAll(PATTERN)].filter((m) => values.has(m[0]) && names(m[1] ?? m[2] ?? m[3] ?? m[4]));
    for (const m of hits.reverse()) {
      const start = m.index;
      const end = start + m[0].length;
      let first = true;
      for (const { t, start: s } of nodes) {
        const e = s + t.textContent.length;
        if (e <= start || s >= end) continue;
        const from = Math.max(start, s) - s;
        const to = Math.min(end, e) - s;
        const data = t.textContent;
        if (first) {
          writeText(doc, t, data.slice(0, from), values.get(m[0]), data.slice(to));
          first = false;
        } else {
          t.textContent = data.slice(0, from) + data.slice(to);
        }
      }
      filled++;
    }
  }
  return filled;
}

/** `before` + `value` + `after` into text element `t`, the value's line breaks as Word's (w:br). */
function writeText(doc, t, before, value, after) {
  const lines = String(value).split('\n');
  const preserve = (el) => el.setAttributeNS(NS.xml, 'xml:space', 'preserve');
  t.textContent = before + lines[0] + (lines.length === 1 ? after : '');
  preserve(t);
  let at = t;
  for (let i = 1; i < lines.length; i++) {
    const br = createW(doc, 'br');
    const next = createW(doc, 't');
    next.textContent = lines[i] + (i === lines.length - 1 ? after : '');
    preserve(next);
    at.parentNode.insertBefore(br, at.nextSibling);
    br.parentNode.insertBefore(next, br.nextSibling);
    at = next;
  }
}

/** The values for one row: placeholder as written -> its column's cell (those with a column). */
export function rowValues(placeholders, mapping, row) {
  const values = new Map();
  placeholders.forEach((p, i) => {
    if (mapping[i] >= 0) values.set(p.text, row[mapping[i]] ?? '');
  });
  return values;
}

// ---------------------------------------------------------------- one document, a copy per row

/** Why the rows can't go into one document (its notes would be referred to by every copy), or null. */
export function singleBlocker(documentXml) {
  if (/<w:(footnote|endnote)Reference\b/.test(documentXml)) return 'it has footnotes or endnotes, which each copy would refer to';
  return null;
}

/**
 * A copy of the body's content made fit to follow others in one document: the ids that must be
 * unique in a document (bookmarks, tracked changes, drawings) moved by `offset`, bookmarks' names
 * made its own, Word 2010's paragraph ids taken off (Word makes new ones), and comments' marks
 * taken out (the comment belongs to the first copy).
 */
function makeCopy(nodes, n, offset) {
  for (const root of nodes) {
    if (root.nodeType !== 1) continue;
    const all = [root, ...root.getElementsByTagName('*')];
    for (const el of all) {
      if (el.namespaceURI === W && (el.localName === 'commentRangeStart' || el.localName === 'commentRangeEnd')) {
        el.parentNode?.removeChild(el);
        continue;
      }
      if (el.namespaceURI === W && el.localName === 'commentReference') {
        const run = el.parentNode;
        run?.parentNode?.removeChild(run);
        continue;
      }
      const id = el.getAttributeNS(W, 'id');
      if (id && /^\d+$/.test(id) && !/Reference$/.test(el.localName)) el.setAttributeNS(W, 'w:id', String(Number(id) + offset));
      if (el.namespaceURI === W && el.localName === 'bookmarkStart') {
        const name = el.getAttributeNS(W, 'name');
        if (name) el.setAttributeNS(W, 'w:name', `${name}_${n}`.slice(0, 40));
      }
      if (el.localName === 'docPr' && /^\d+$/.test(el.getAttribute('id') || '')) el.setAttribute('id', String(Number(el.getAttribute('id')) + offset));
      el.removeAttributeNS(W14, 'paraId');
      el.removeAttributeNS(W14, 'textId');
    }
  }
}

/** Starts a copy on a new page: "page break before" on its first paragraph, or a page break before its first table. */
function startOnNewPage(doc, nodes, body, before) {
  const first = nodes.find((n) => n.nodeType === 1);
  if (first?.localName === 'p') {
    const pPr = wChild(first, 'pPr') || first.insertBefore(createW(doc, 'pPr'), first.firstChild);
    ensureOrdered(pPr, 'pageBreakBefore', PPR_ORDER);
    return;
  }
  const p = createW(doc, 'p');
  const r = p.appendChild(createW(doc, 'r'));
  r.appendChild(createW(doc, 'br', { type: 'page' }));
  body.insertBefore(p, before);
}

/** The highest number used as an id in `doc` (w:id, docPr's id), for copies' ids to start above. */
function highestId(doc) {
  let max = 0;
  for (const el of doc.getElementsByTagName('*')) {
    for (const v of [el.getAttributeNS(W, 'id'), el.localName === 'docPr' ? el.getAttribute('id') : null]) {
      if (v && /^\d+$/.test(v)) max = Math.max(max, Number(v));
    }
  }
  return max;
}

/**
 * Merges `rows` into the document `bytes` (a .docx as downloaded). `placeholders`: [{ text }] in
 * the order of `mapping`, each one's column (-1: left as it is). `parts`: { main, others } — the
 * main part's path and the header, footer and note parts' paths.
 * mode 'single': one .docx, a copy for each row, each on a new page; the headers and footers
 * shared, so only what every row fills in the same way is filled in them.
 * mode 'separate': a .zip of a .docx for each row, named by column `nameColumn` (or numbered).
 * Returns { blob, name, count, keptInHeaders } (keptInHeaders: placeholders left in headers and
 * footers because the rows fill them differently).
 */
export async function mailMerge(bytes, { rows, placeholders, mapping, parts, mode = 'single', nameColumn = -1, baseName = 'document', mime }) {
  rows = rows.slice(0, MAX_ROWS);
  const base = await JSZip.loadAsync(bytes);
  const fillPart = async (zip, path, values) => {
    const text = await zip.file(path)?.async('string');
    if (text == null) return;
    const xml = parseXml(text);
    if (fillXml(xml, values)) zip.file(path, serializeXml(xml));
  };

  if (mode === 'separate') {
    const out = new JSZip();
    const used = new Set();
    for (let i = 0; i < rows.length; i++) {
      const zip = await JSZip.loadAsync(bytes);
      const values = rowValues(placeholders, mapping, rows[i]);
      for (const path of [parts.main, ...parts.others]) await fillPart(zip, path, values);
      let name = fileName(nameColumn >= 0 ? rows[i][nameColumn] : '') || `${baseName} ${i + 1}`;
      for (let k = 2; used.has(name.toLowerCase()); k++) name = `${fileName(nameColumn >= 0 ? rows[i][nameColumn] : '') || baseName} (${k})`;
      used.add(name.toLowerCase());
      out.file(`${name}.docx`, await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
    }
    const blob = await out.generateAsync({ type: 'blob', mimeType: 'application/zip', compression: 'STORE' });
    return { blob, name: `${baseName} (merged).zip`, count: rows.length, keptInHeaders: [] };
  }

  // One document: the body's content once for each row.
  const xml = parseXml(await base.file(parts.main).async('string'));
  const body = xml.getElementsByTagNameNS(W, 'body')[0];
  const sectPr = [...body.childNodes].filter((n) => n.nodeType === 1).pop();
  const finalSect = sectPr?.localName === 'sectPr' ? sectPr : null;
  // Each copy's ids go above every id in the document (taken before its content is taken out).
  const stride = highestId(xml) + 1;
  const content = [...body.childNodes].filter((n) => n !== finalSect);
  for (const n of content) body.removeChild(n);
  rows.forEach((row, i) => {
    const copy = content.map((n) => n.cloneNode(true));
    if (i) makeCopy(copy, i + 1, stride * i);
    const holder = xml.createElementNS(W, 'w:body');
    for (const n of copy) holder.appendChild(n);
    fillXml(holder, rowValues(placeholders, mapping, row));
    const nodes = [...holder.childNodes];
    if (i) startOnNewPage(xml, nodes, body, finalSect);
    for (const n of nodes) body.insertBefore(n, finalSect);
  });
  base.file(parts.main, serializeXml(xml));
  // Headers and footers are shared by every copy: what every row fills the same way is filled.
  const same = new Map();
  const differ = [];
  placeholders.forEach((p, i) => {
    if (mapping[i] < 0) return;
    const first = rows[0]?.[mapping[i]] ?? '';
    if (rows.every((r) => (r[mapping[i]] ?? '') === first)) same.set(p.text, first);
    else differ.push(p.text);
  });
  const keptInHeaders = [];
  for (const path of parts.others) {
    const text = await base.file(path)?.async('string');
    if (text == null) continue;
    const part = parseXml(text);
    if (fillXml(part, same)) base.file(path, serializeXml(part));
    for (const p of differ) if (!keptInHeaders.includes(p) && plainIncludes(part, p)) keptInHeaders.push(p);
  }
  const blob = await base.generateAsync({ type: 'blob', mimeType: mime, compression: 'DEFLATE' });
  return { blob, name: `${baseName} (merged).docx`, count: rows.length, keptInHeaders };
}

/** Whether the text of `doc`'s paragraphs holds `s` (across runs). */
function plainIncludes(doc, s) {
  for (const p of doc.getElementsByTagNameNS(W, 'p')) {
    if ([...p.getElementsByTagNameNS(W, 't')].map((t) => t.textContent).join('').includes(s)) return true;
  }
  return false;
}

/** A cell's text as a file name: on one line, without the characters file systems refuse. */
function fileName(s) {
  return String(s || '').replace(/[\r\n]+/g, ' ').replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '').replace(/\s+/g, ' ').trim().slice(0, 80);
}
