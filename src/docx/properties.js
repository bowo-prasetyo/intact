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
