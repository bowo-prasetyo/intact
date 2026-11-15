// Small helpers for working with WordprocessingML (OOXML) via the browser DOM.

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  v: 'urn:schemas-microsoft-com:vml',
  pr: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  xml: 'http://www.w3.org/XML/1998/namespace',
};
const W = NS.w;

export const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  if (err) throw new Error('Invalid XML: ' + err.textContent.slice(0, 200));
  return doc;
}

export function serializeXml(doc) {
  let s = new XMLSerializer().serializeToString(doc);
  s = s.replace(/^<\?xml[^>]*\?>\s*/, '');
  return XML_HEADER + s;
}

/** Element children in the w: namespace, optionally filtered by local name. */
export function wChildren(el, local) {
  if (!el) return [];
  const out = [];
  for (const n of el.childNodes) {
    if (n.nodeType === 1 && n.namespaceURI === W && (!local || n.localName === local)) out.push(n);
  }
  return out;
}

export function wChild(el, local) {
  if (!el) return null;
  for (const n of el.childNodes) {
    if (n.nodeType === 1 && n.namespaceURI === W && n.localName === local) return n;
  }
  return null;
}

export function elementChildren(el) {
  return el ? [...el.childNodes].filter((n) => n.nodeType === 1) : [];
}

export function wAttr(el, name) {
  if (!el) return null;
  return el.getAttributeNS(W, name) ?? el.getAttribute('w:' + name);
}

/** Interprets an OOXML on/off property (e.g. <w:b/>, <w:b w:val="0"/>). Returns undefined when absent. */
export function onOff(el) {
  if (!el) return undefined;
  const v = wAttr(el, 'val');
  return !(v === '0' || v === 'false' || v === 'off');
}

export function createW(doc, local, attrs = {}) {
  const el = doc.createElementNS(W, 'w:' + local);
  for (const [k, v] of Object.entries(attrs)) el.setAttributeNS(W, 'w:' + k, String(v));
  return el;
}

// Schema order of children (CT_RPr / CT_PPr). Word rejects out-of-order elements.
export const RPR_ORDER = [
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
  'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
  'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
  'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange',
];
export const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
  'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
  'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
  'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
  'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
  'sectPr', 'pPrChange',
];

/** Returns the existing child `local` of `parent`, or inserts a new one at its schema position. */
export function ensureOrdered(parent, local, order) {
  const existing = wChild(parent, local);
  if (existing) return existing;
  const el = createW(parent.ownerDocument, local);
  const idx = order.indexOf(local);
  const before = wChildren(parent).find((c) => {
    const j = order.indexOf(c.localName);
    return j > idx || j === -1;
  });
  parent.insertBefore(el, before || null);
  return el;
}

export function removeW(parent, ...locals) {
  for (const local of locals) for (const c of wChildren(parent, local)) parent.removeChild(c);
}

export function isEmptyElement(el) {
  return el.childNodes.length === 0 && el.attributes.length === 0;
}

/** Concatenated visible text of an OOXML fragment (for previews of unsupported content). */
export function plainText(el) {
  let s = '';
  const walk = (n) => {
    if (n.nodeType !== 1) return;
    if (n.namespaceURI === W) {
      if (n.localName === 't' || n.localName === 'delText') { s += n.textContent; return; }
      if (n.localName === 'tab') { s += '\t'; return; }
      if (n.localName === 'br' || n.localName === 'cr') { s += '\n'; return; }
      if (n.localName === 'p' && s && !s.endsWith('\n')) s += '\n';
      if (n.localName === 'instrText') return;
    }
    for (const c of n.childNodes) walk(c);
  };
  walk(el);
  return s;
}

/** Resolves a relationship target relative to the part that owns the .rels file. */
export function resolvePath(baseDir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (baseDir ? baseDir.split('/') : []).filter(Boolean);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg !== '') parts.push(seg);
  }
  return parts.join('/');
}

export function parseRels(xmlText, baseDir) {
  const map = new Map();
  if (!xmlText) return map;
  const doc = parseXml(xmlText);
  for (const rel of doc.getElementsByTagNameNS(NS.pr, 'Relationship')) {
    const target = rel.getAttribute('Target');
    const external = rel.getAttribute('TargetMode') === 'External';
    map.set(rel.getAttribute('Id'), {
      type: rel.getAttribute('Type'),
      target,
      external,
      path: external ? null : resolvePath(baseDir, target),
    });
  }
  return map;
}
