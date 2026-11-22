// Document properties (docProps/core.xml, app.xml, custom.xml), the sensitivity label or
// classification marking a document carries, and removing personal information before release, as
// Word's Document Inspector does ("Document Properties and Personal Information").
import { NS, parseXml, serializeXml } from './xml.js';

const W = NS.w;

const CORE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const APP_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
const CUSTOM_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';
const CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const DC = 'http://purl.org/dc/elements/1.1/';
const EP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const LABEL_NS = 'http://schemas.microsoft.com/office/2020/mipLabelMetadata';

/** The core properties shown before release, by element: [namespace, local name, label]. */
const CORE = [
  [DC, 'creator', 'Author'], [CP, 'lastModifiedBy', 'Last saved by'], [DC, 'title', 'Title'], [DC, 'subject', 'Subject'],
  [CP, 'keywords', 'Keywords'], [DC, 'description', 'Comments'], [CP, 'category', 'Category'], [CP, 'contentStatus', 'Status'],
];
const APP = [['Company', 'Company'], ['Manager', 'Manager'], ['Template', 'Template'], ['HyperlinkBase', 'Hyperlink base']];
/** The properties that name people or organisations: removed with personal information. */
export const PERSONAL = new Set(['Author', 'Last saved by', 'Company', 'Manager']);

/** Custom properties that carry a classification marking, besides Purview's labels. */
const MARKING = /(^|[_ ])(security ?classification|classification|protective ?marking|sensitivity|dissemination|caveat)/i;

/** The package's property parts: { core, app, custom } paths (null where missing). */
export function propertyPaths(pkgRels) {
  const of = (type) => [...pkgRels.values()].find((r) => r.type === type && !r.external)?.path || null;
  return { core: of(CORE_REL), app: of(APP_REL), custom: of(CUSTOM_REL) };
}

const text = async (zip, path) => (path && zip.file(path) ? zip.file(path).async('string') : null);

/**
 * The document's properties: { core: [{ label, value }], app: [...], custom: [{ name, value }],
 * label, marking }. `label` is the Microsoft Purview sensitivity label, `marking` a classification
 * from another tool's custom property (see readLabel).
 */
export async function readProperties(zip, pkgRels) {
  const paths = propertyPaths(pkgRels);
  const out = { core: [], app: [], custom: [], label: null, marking: null };
  const core = await text(zip, paths.core);
  if (core) {
    const root = parseXml(core).documentElement;
    for (const [ns, local, label] of CORE) {
      const value = root.getElementsByTagNameNS(ns, local)[0]?.textContent.trim();
      if (value) out.core.push({ label, value });
    }
  }
  const app = await text(zip, paths.app);
  if (app) {
    const root = parseXml(app).documentElement;
    for (const [local, label] of APP) {
      const value = root.getElementsByTagNameNS(EP, local)[0]?.textContent.trim();
      // Word writes Normal.dotm as the template of every document: only another one says anything.
      if (value && !(local === 'Template' && /^normal\.dotm?$/i.test(value))) out.app.push({ label, value });
    }
  }
  const custom = await text(zip, paths.custom);
  if (custom) {
    for (const p of parseXml(custom).documentElement.children) {
      const name = p.getAttribute('name');
      if (name) out.custom.push({ name, value: p.textContent.trim() });
    }
  }
  out.label = readLabel(out.custom, await text(zip, 'docMetadata/LabelInfo.xml'));
  const marking = out.custom.find((p) => !p.name.startsWith('MSIP_Label_') && MARKING.test(p.name) && p.value);
  if (marking) out.marking = marking;
  return out;
}

const guid = (s) => (s || '').replace(/[{}]/g, '').toLowerCase();

/**
 * The sensitivity label applied with Microsoft Purview (Azure Information Protection), or null:
 * { name, id, method, date, markings }. Office writes it as custom properties named
 * MSIP_Label_<id>_<field> and, since 2021, in docMetadata/LabelInfo.xml (without the name, which
 * then isn't in the file). `method` is Privileged (chosen by a person) or Standard (by default or
 * automatically); `markings` the content marking the label adds: header, footer, watermark, and
 * whether it encrypts.
 */
export function readLabel(custom, labelInfo) {
  const labels = new Map();
  const at = (id) => labels.get(guid(id)) || labels.set(guid(id), { id: guid(id) }).get(guid(id));
  for (const { name, value } of custom) {
    const m = /^MSIP_Label_([0-9a-f-]{36})_(\w+)$/i.exec(name);
    if (m) at(m[1])[m[2].toLowerCase()] = value;
  }
  if (labelInfo) {
    for (const el of parseXml(labelInfo).getElementsByTagNameNS(LABEL_NS, 'label')) {
      const l = at(el.getAttribute('id'));
      l.enabled = el.getAttribute('enabled');
      l.method ||= el.getAttribute('method');
      l.contentbits ??= el.getAttribute('contentBits');
      l.removed = el.getAttribute('removed');
    }
  }
  const on = (v) => /^(1|true)$/i.test(v || '');
  const label = [...labels.values()].find((l) => on(l.enabled) && !on(l.removed));
  if (!label) return null;
  const bits = Number(label.contentbits) || 0;
  return {
    id: label.id,
    name: label.name || null,
    method: label.method || null,
    date: label.setdate || null,
    markings: [bits & 1 && 'header', bits & 2 && 'footer', bits & 4 && 'watermark', bits & 8 && 'encryption'].filter(Boolean),
  };
}

// ------------------------------------------------------------------ removing personal information

/** The name Word puts on comments and tracked changes when personal information is removed. */
export const ANONYMOUS = 'Author';

/**
 * The parts that can hold people's names: the main document and the parts beside it (comments,
 * notes, headers, footers, people.xml) and the property parts.
 */
export function personalParts(zip, mainDir, props) {
  const dir = mainDir ? `${mainDir}/` : '';
  const inMain = Object.keys(zip.files).filter((p) => p.startsWith(dir) && !p.slice(dir.length).includes('/') && /\.xml$/i.test(p));
  return [...inMain, props.core, props.app].filter(Boolean);
}

/**
 * `xml` with personal information taken out, or null if it has none: in the document's parts, the
 * names (and initials) on comments and tracked changes become "Author", and people.xml loses the
 * accounts behind them; in the properties, the author, the last to save it, company and manager go.
 */
export function anonymize(xml) {
  const doc = parseXml(xml);
  let changed = false;
  for (const el of doc.getElementsByTagName('*')) {
    for (const [ns, name, value] of [[W, 'author', ANONYMOUS], [W, 'initials', 'A'], [W15, 'author', ANONYMOUS]]) {
      const v = el.getAttributeNS(ns, name);
      if (v && v !== value) {
        el.setAttributeNS(ns, el.getAttributeNodeNS(ns, name).name, value);
        changed = true;
      }
    }
  }
  const remove = [
    ...doc.getElementsByTagNameNS(W15, 'presenceInfo'),
    ...doc.getElementsByTagNameNS(DC, 'creator'), ...doc.getElementsByTagNameNS(CP, 'lastModifiedBy'),
    ...doc.getElementsByTagNameNS(EP, 'Company'), ...doc.getElementsByTagNameNS(EP, 'Manager'),
  ];
  for (const el of remove) el.remove();
  if (remove.length) changed = true;
  // people.xml: one person per name, all "Author" now; Word keeps one.
  const seen = new Set();
  for (const p of [...doc.getElementsByTagNameNS(W15, 'person')]) {
    const name = p.getAttributeNS(W15, 'author');
    if (seen.has(name)) {
      p.remove();
      changed = true;
    }
    seen.add(name);
  }
  return changed ? serializeXml(doc) : null;
}

// ------------------------------------------------------------------ recording an edit

const DCTERMS = 'http://purl.org/dc/terms/';
const XSI = 'http://www.w3.org/2001/XMLSchema-instance';

/**
 * core.xml (`xml`) with an edit recorded, as Word records one on saving: `modified` (a Date) as
 * dcterms:modified, `by` as cp:lastModifiedBy (left as it is when empty), the cp:revision number one
 * up when `revision`; and `title`, when given, as dc:title.
 */
export function updateCore(xml, { modified, by, revision } = {}, title = null) {
  const doc = parseXml(xml);
  const root = doc.documentElement;
  const set = (ns, qname, value, attrs = []) => {
    let el = root.getElementsByTagNameNS(ns, qname.split(':')[1])[0];
    if (!el) {
      el = doc.createElementNS(ns, qname);
      for (const [ans, aname, avalue] of attrs) el.setAttributeNS(ans, aname, avalue);
      root.append(el);
    }
    el.textContent = value;
  };
  if (title != null) set(DC, 'dc:title', title);
  if (by) set(CP, 'cp:lastModifiedBy', by);
  if (revision) {
    const n = Number(root.getElementsByTagNameNS(CP, 'revision')[0]?.textContent);
    if (Number.isInteger(n) && n > 0) set(CP, 'cp:revision', String(n + 1));
  }
  // W3CDTF to the second, in UTC, as Word writes it.
  if (modified) set(DCTERMS, 'dcterms:modified', modified.toISOString().replace(/\.\d+Z$/, 'Z'), [[XSI, 'xsi:type', 'dcterms:W3CDTF']]);
  return serializeXml(doc);
}

// ------------------------------------------------------------------ editing properties

/**
 * The core properties edited here, as File > Info > Properties shows them: [key, namespace, qualified
 * name, label, the DOCPROPERTY field's name (null: not offered as one)].
 */
export const EDITABLE_CORE = [
  ['title', DC, 'dc:title', 'Title', 'Title'],
  ['subject', DC, 'dc:subject', 'Subject', 'Subject'],
  ['keywords', CP, 'cp:keywords', 'Keywords', 'Keywords'],
  ['category', CP, 'cp:category', 'Category', 'Category'],
  ['contentStatus', CP, 'cp:contentStatus', 'Status', null],
  ['description', DC, 'dc:description', 'Comments', 'Comments'],
];

/** The editable core properties' values in core.xml `xml` (or none): { title, subject, … }, '' where unset. */
export function readCore(xml) {
  const root = xml ? parseXml(xml).documentElement : null;
  return Object.fromEntries(EDITABLE_CORE.map(([key, ns, qname]) => [key, root?.getElementsByTagNameNS(ns, qname.split(':')[1])[0]?.textContent || '']));
}

/** core.xml `xml` with the editable properties in `edits` ({ title: '…' }) set; an empty one taken out. */
export function setCore(xml, edits) {
  const doc = parseXml(xml);
  const root = doc.documentElement;
  for (const [key, ns, qname] of EDITABLE_CORE) {
    if (!(key in edits)) continue;
    let el = root.getElementsByTagNameNS(ns, qname.split(':')[1])[0];
    if (!edits[key]) {
      el?.remove();
      continue;
    }
    if (!el) root.append((el = doc.createElementNS(ns, qname)));
    el.textContent = edits[key];
  }
  return serializeXml(doc);
}

/** A new, empty core.xml, as Word writes one. */
export const EMPTY_CORE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="${CP}" xmlns:dc="${DC}" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>`;

const CUSTOM_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties';
const VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
// The format id every custom property has, as Word writes it.
const FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';
export const CUSTOM_PROPS_REL = CUSTOM_REL;
export const CUSTOM_PROPS_CT = 'application/vnd.openxmlformats-officedocument.custom-properties+xml';
export const CORE_PROPS_REL = CORE_REL;
export const CORE_PROPS_CT = 'application/vnd.openxmlformats-package.core-properties+xml';

/** True for custom properties kept as they are (a sensitivity label's, which Office manages). */
export const managedProperty = (name) => /^MSIP_Label_/i.test(name);

/**
 * custom.xml (`xml`, or none for a new part) holding the custom properties `list` ([{ name, value }],
 * text): a property already there keeps its type and id with the value given (a number or a date
 * that no longer reads as one becomes text, vt:lpwstr), one not in `list` is taken out (but for those
 * Office manages, see managedProperty), and a new one is added as text, as Word's Advanced
 * Properties > Custom adds one.
 */
export function setCustom(xml, list) {
  const doc = parseXml(xml || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="${CUSTOM_NS}" xmlns:vt="${VT}"/>`);
  const root = doc.documentElement;
  const wanted = new Map(list.map((p) => [p.name.toLowerCase(), p]));
  const seen = new Set();
  let pid = 1;
  for (const p of [...root.children]) {
    pid = Math.max(pid, Number(p.getAttribute('pid')) || 0);
    const name = p.getAttribute('name') || '';
    if (managedProperty(name)) continue;
    const want = wanted.get(name.toLowerCase());
    if (!want) {
      p.remove();
      continue;
    }
    seen.add(name.toLowerCase());
    const v = p.firstElementChild;
    if (v && v.textContent === want.value) continue;
    const fits = v && ({ i4: /^-?\d+$/, r8: /^-?\d+(\.\d+)?$/, bool: /^(true|false)$/, filetime: /^\d{4}-\d\d-\d\dT/ }[v.localName]?.test(want.value) ?? v.localName === 'lpwstr');
    if (fits) v.textContent = want.value;
    else {
      const t = doc.createElementNS(VT, 'vt:lpwstr');
      t.textContent = want.value;
      if (v) v.replaceWith(t);
      else p.append(t);
    }
  }
  for (const want of list) {
    if (seen.has(want.name.toLowerCase())) continue;
    const p = doc.createElementNS(CUSTOM_NS, 'property');
    p.setAttribute('fmtid', FMTID);
    p.setAttribute('pid', String(Math.max(2, ++pid)));
    p.setAttribute('name', want.name);
    const t = doc.createElementNS(VT, 'vt:lpwstr');
    t.textContent = want.value;
    p.append(t);
    root.append(p);
    pid = Number(p.getAttribute('pid'));
  }
  return serializeXml(doc);
}
