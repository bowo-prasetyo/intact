// DocxDocument: loads a .docx package, renders its body into editable HTML, and writes the
// edited HTML back into word/document.xml while keeping everything else in the package intact.
//
// Strategy (lossless where it matters):
//  * Paragraph properties (<w:pPr>) and run properties (<w:rPr>) are kept as original XML in a
//    side table ("raw store") and referenced from the HTML via data-ppr / data-rpr attributes.
//  * Bold / italic / underline / strike are expressed as HTML tags so the browser can edit them;
//    on save they are merged back into the stored rPr.
//  * Anything the editor does not understand (images, fields, content controls, equations,
//    bookmarks, ...) becomes a non-editable placeholder that is written back verbatim.
//  * Tracked changes (insertions and deletions) are shown read-only and written back verbatim, so
//    their author and date survive and the reviewer can still accept or reject them in Word.
//  * Tables keep their structure (read-only) while each cell's content is editable.
import JSZip from 'jszip';
import {
  NS, parseXml, serializeXml, wChild, wChildren, wAttr, onOff, createW, ensureOrdered, removeW,
  RPR_ORDER, PPR_ORDER, isEmptyElement, plainText, parseRels, elementChildren,
} from './xml.js';
import { StyleSheet, jcToCss } from './styles.js';
import { Numbering } from './numbering.js';
import { blankDocxFiles } from './blank.js';

const W = NS.w;
const MAIN_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const STYLES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const STYLES_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
export const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

const SIMPLE_RUN_CHILDREN = new Set(['rPr', 't', 'tab', 'br', 'cr', 'noBreakHyphen', 'softHyphen', 'lastRenderedPageBreak']);
const HIDDEN_INLINE = new Set(['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'permStart', 'permEnd', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']);
const FLATTEN_INLINE = new Set(['smartTag', 'customXml', 'dir', 'bdo']);
const TRACKED_INLINE = { ins: 'Inserted', moveTo: 'Moved here', del: 'Deleted', moveFrom: 'Moved away' };
const BLOCK_TAGS = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'DIV', 'TABLE', 'UL', 'OL', 'LI', 'BLOCKQUOTE', 'PRE', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER']);
const HEADING_TAG = { title: 'h1', h1: 'h1', h2: 'h2', h3: 'h3', h4: 'h4', h5: 'h5', h6: 'h6' };
const IMAGE_TYPES = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff' };

export class DocxDocument {
  /** @param {ArrayBuffer|Blob} data */
  static async open(data, name = 'document.docx') {
    const zip = await JSZip.loadAsync(data);
    const doc = new DocxDocument(zip, name);
    await doc.init();
    return doc;
  }

  static async createBlank(name = 'Untitled.docx') {
    const zip = new JSZip();
    for (const [path, content] of Object.entries(blankDocxFiles())) zip.file(path, content);
    return DocxDocument.open(await zip.generateAsync({ type: 'arraybuffer' }), name);
  }

  constructor(zip, name) {
    this.zip = zip;
    this.name = name;
    this.mediaUrls = new Map();
  }

  async text(path) {
    const f = this.zip.file(path);
    return f ? f.async('string') : null;
  }

  async init() {
    // Locate the main document part via the package relationships.
    const pkgRels = parseRels(await this.text('_rels/.rels'), '');
    const main = [...pkgRels.values()].find((r) => r.type === MAIN_TYPE || r.type.endsWith('/officeDocument'));
    this.mainPath = main?.path || 'word/document.xml';
    const docXml = await this.text(this.mainPath);
    if (!docXml) throw new Error('This file does not look like a Word document (no main document part).');
    this.xml = parseXml(docXml);
    this.body = this.xml.getElementsByTagNameNS(W, 'body')[0];
    if (!this.body) throw new Error('The document has no body.');

    const slash = this.mainPath.lastIndexOf('/');
    this.mainDir = slash >= 0 ? this.mainPath.slice(0, slash) : '';
    this.relsPath = `${this.mainDir ? this.mainDir + '/' : ''}_rels/${this.mainPath.slice(slash + 1)}.rels`;
    this.rels = parseRels(await this.text(this.relsPath), this.mainDir);

    const relOf = (suffix) => [...this.rels.values()].find((r) => r.type.endsWith(suffix) && !r.external);
    this.stylesPath = relOf('/styles')?.path || null;
    const stylesXml = this.stylesPath && (await this.text(this.stylesPath));
    const themePath = relOf('/theme')?.path;
    const themeXml = themePath && (await this.text(themePath));
    const numPath = relOf('/numbering')?.path;
    const numXml = numPath && (await this.text(numPath));

    this.styles = new StyleSheet(
      parseXml(stylesXml || `<w:styles xmlns:w="${W}"/>`),
      themeXml ? parseXml(themeXml) : null,
    );
    this.numberingXml = numXml ? parseXml(numXml) : null;

    // Pre-create blob URLs for all images so rendering can stay synchronous.
    const mediaFiles = Object.values(this.zip.files).filter((f) => !f.dir && /\.(png|jpe?g|gif|bmp|svg|webp|tiff?)$/i.test(f.name));
    await Promise.all(mediaFiles.map(async (f) => {
      const ext = f.name.split('.').pop().toLowerCase();
      const blob = new Blob([await f.async('uint8array')], { type: IMAGE_TYPES[ext] || 'application/octet-stream' });
      this.mediaUrls.set(f.name, URL.createObjectURL(blob));
    }));
  }

  dispose() {
    for (const url of this.mediaUrls.values()) URL.revokeObjectURL(url);
    this.mediaUrls.clear();
  }

  /** Page geometry in CSS px from the final section properties. */
  get page() {
    const sectPr = wChild(this.body, 'sectPr');
    const pgSz = wChild(sectPr, 'pgSz');
    const pgMar = wChild(sectPr, 'pgMar');
    const px = (twips, dflt) => (twips != null ? Number(twips) / 15 : dflt);
    return {
      width: px(wAttr(pgSz, 'w'), 816),
      minHeight: px(wAttr(pgSz, 'h'), 1056),
      top: px(wAttr(pgMar, 'top'), 96),
      bottom: px(wAttr(pgMar, 'bottom'), 96),
      left: px(wAttr(pgMar, 'left'), 96),
      right: px(wAttr(pgMar, 'right'), 96),
    };
  }

  css(scope) {
    return this.styles.css(scope);
  }

  /** Style id to use for a builtin block kind ('normal', 'title', 'h1'..). Adds missing styles. */
  styleIdFor(kind) {
    if (kind === 'normal') return null;
    return this.styles.ensureBuiltin(kind);
  }

  blockKind(styleId) {
    return this.styles.kindOf(styleId) || (styleId && styleId !== this.styles.defaultParagraph ? 'other' : 'normal');
  }

  // ---------------------------------------------------------------- XML -> HTML

  /** Renders the document body into a DocumentFragment of editable HTML. */
  render() {
    this.raw = [];
    this.numbering = new Numbering(this.numberingXml);
    const frag = document.createDocumentFragment();
    this.renderBlocks(this.body, frag);
    if (!frag.childNodes.length) frag.append(this.emptyParagraph());
    return frag;
  }

  keep(node) {
    this.raw.push(node);
    return String(this.raw.length - 1);
  }

  emptyParagraph() {
    const p = document.createElement('p');
    p.append(document.createElement('br'));
    return p;
  }

  renderBlocks(container, parent) {
    for (const child of elementChildren(container)) {
      const local = child.namespaceURI === W ? child.localName : null;
      if (local === 'p') parent.append(this.renderParagraph(child));
      else if (local === 'tbl') parent.append(this.renderTable(child));
      else if (local === 'sectPr' || local === 'tblPr' || local === 'tcPr' || local === 'trPr') continue;
      else parent.append(this.rawBlock(child));
    }
  }

  rawBlock(node) {
    const div = document.createElement('div');
    div.contentEditable = 'false';
    div.dataset.raw = this.keep(node);
    const local = node.localName;
    if (HIDDEN_INLINE.has(local)) {
      div.className = 'raw-hidden';
      return div;
    }
    div.className = 'raw-block';
    div.title = `Preserved content (<${node.nodeName}>) – not editable here`;
    div.textContent = plainText(node).trim() || `[${node.nodeName}]`;
    return div;
  }

  renderParagraph(p) {
    const pPr = wChild(p, 'pPr');
    const styleId = wAttr(wChild(pPr, 'pStyle'), 'val');
    const kind = this.styles.kindOf(styleId);
    const el = document.createElement(HEADING_TAG[kind] || 'p');
    if (styleId) el.dataset.style = styleId;
    if (pPr) {
      el.dataset.ppr = this.keep(pPr);
      Object.assign(el.style, cssToStyle(this.styles.pPrCss(pPr)));
    }
    const numPr = wChild(pPr, 'numPr') || this.styles.styleNumPr(styleId);
    const num = numPr && this.numbering.next(numPr);
    if (num) {
      el.classList.add('list');
      el.dataset.marker = num.marker;
      if (num.pPr && !wChild(pPr, 'ind')) {
        const { 'padding-left': pl } = this.styles.pPrCss(num.pPr);
        if (pl) el.style.paddingLeft = pl;
      }
      if (!el.style.paddingLeft) el.style.paddingLeft = '36pt';
      el.style.textIndent = '';
    }
    this.renderInlines(p, el);
    if (!el.textContent && !el.querySelector('img,.raw-inline:not(.raw-hidden)')) el.append(document.createElement('br'));
    return el;
  }

  renderInlines(container, parent) {
    for (const child of elementChildren(container)) {
      const local = child.namespaceURI === W ? child.localName : null;
      if (local === 'pPr' || local === 'proofErr') continue;
      if (local === 'r') this.renderRun(child, parent);
      else if (local === 'hyperlink') {
        const a = document.createElement('a');
        const shell = child.cloneNode(false);
        a.dataset.hl = this.keep(shell);
        const rid = child.getAttributeNS(NS.r, 'id');
        const anchor = wAttr(child, 'anchor');
        const rel = rid && this.rels.get(rid);
        a.href = rel ? rel.target : anchor ? '#' + anchor : '#';
        a.title = (rel ? rel.target : anchor || '') + ' (Ctrl+click to open)';
        this.renderInlines(child, a);
        parent.append(a);
      } else if (local && FLATTEN_INLINE.has(local)) {
        // Unwrap smart tags and similar containers: render their runs as normal text.
        this.renderInlines(child, parent);
      } else {
        parent.append(this.rawInline(child));
      }
    }
  }

  rawInline(node) {
    const span = document.createElement('span');
    span.contentEditable = 'false';
    span.dataset.raw = this.keep(node);
    const local = node.localName;
    if (HIDDEN_INLINE.has(local)) {
      span.className = 'raw-inline raw-hidden';
      return span;
    }
    span.className = 'raw-inline';
    // Images (DrawingML or legacy VML)
    const blips = [...node.getElementsByTagNameNS(NS.a, 'blip'), ...node.getElementsByTagNameNS(NS.v, 'imagedata')];
    if (blips.length) {
      const extent = node.getElementsByTagNameNS(NS.wp, 'extent')[0];
      for (const blip of blips) {
        const rid = blip.getAttributeNS(NS.r, 'embed') || blip.getAttributeNS(NS.r, 'id');
        const rel = rid && this.rels.get(rid);
        const url = rel && this.mediaUrls.get(rel.path);
        if (!url) continue;
        const img = document.createElement('img');
        img.src = url;
        img.alt = '';
        if (extent) {
          img.style.width = `${Number(extent.getAttribute('cx')) / 9525}px`;
          img.style.height = `${Number(extent.getAttribute('cy')) / 9525}px`;
        }
        span.append(img);
        break;
      }
      if (span.childNodes.length) return span;
    }
    if (local === 'r' && wChildren(node, 'br').some((b) => wAttr(b, 'type') === 'page')) {
      span.classList.add('page-break');
      span.title = 'Page break';
      return span;
    }
    if (local === 'r' && wChild(node, 'footnoteReference')) {
      span.classList.add('note-ref');
      span.textContent = '*';
      span.title = 'Footnote';
      return span;
    }
    if (TRACKED_INLINE[local]) {
      span.classList.add(local === 'ins' || local === 'moveTo' ? 'inserted' : 'deleted');
      const author = wAttr(node, 'author');
      const date = wAttr(node, 'date');
      span.title = `${TRACKED_INLINE[local]}${author ? ` by ${author}` : ''}${date ? ` on ${date.slice(0, 10)}` : ''} – tracked change, accept or reject it in Word`;
      span.textContent = plainText(node) || '⧉';
      return span;
    }
    span.title = `Preserved content (<${node.nodeName}>)`;
    span.textContent = plainText(node) || '⧉';
    return span;
  }

  renderRun(r, parent) {
    const kids = elementChildren(r);
    const simple = kids.every((k) => k.namespaceURI === W && SIMPLE_RUN_CHILDREN.has(k.localName)
      && !(k.localName === 'br' && wAttr(k, 'type') && wAttr(k, 'type') !== 'textWrapping'));
    if (!simple) {
      parent.append(this.rawInline(r));
      return;
    }
    const rPr = wChild(r, 'rPr');
    let target = parent;
    if (rPr) {
      const span = document.createElement('span');
      span.dataset.rpr = this.keep(rPr);
      const rStyle = wAttr(wChild(rPr, 'rStyle'), 'val');
      if (rStyle) span.dataset.rstyle = rStyle;
      Object.assign(span.style, cssToStyle(this.styles.rPrCss(rPr, false)));
      // Explicitly switched-off toggles (e.g. <w:b w:val="0"/>) are kept as inline CSS.
      if (onOff(wChild(rPr, 'b')) === false) span.style.fontWeight = 'normal';
      if (onOff(wChild(rPr, 'i')) === false) span.style.fontStyle = 'normal';
      parent.append(span);
      target = span;
      const wrap = (tag) => { const e = document.createElement(tag); target.append(e); target = e; };
      if (onOff(wChild(rPr, 'b'))) wrap('strong');
      if (onOff(wChild(rPr, 'i'))) wrap('em');
      const u = wChild(rPr, 'u');
      if (u && wAttr(u, 'val') !== 'none') wrap('u');
      if (onOff(wChild(rPr, 'strike')) || onOff(wChild(rPr, 'dstrike'))) wrap('s');
    }
    let text = '';
    const flush = () => { if (text) { target.append(document.createTextNode(text)); text = ''; } };
    for (const k of kids) {
      switch (k.localName) {
        case 't': text += k.textContent; break;
        case 'tab': text += '\t'; break;
        case 'noBreakHyphen': text += '‑'; break;
        case 'softHyphen': text += '­'; break;
        case 'br': case 'cr': flush(); target.append(document.createElement('br')); break;
        default: break;
      }
    }
    flush();
  }

  renderTable(tbl) {
    const table = document.createElement('table');
    table.className = 'docx-table';
    table.contentEditable = 'false';
    table.dataset.raw = this.keep(tbl);
    const tbody = document.createElement('tbody');
    table.append(tbody);
    for (const tr of wChildren(tbl, 'tr')) {
      const row = document.createElement('tr');
      for (const tc of wChildren(tr, 'tc')) {
        const td = document.createElement('td');
        td.contentEditable = 'true';
        const tcPr = wChild(tc, 'tcPr');
        const span = Number(wAttr(wChild(tcPr, 'gridSpan'), 'val') || 1);
        if (span > 1) td.colSpan = span;
        const vMerge = wChild(tcPr, 'vMerge');
        if (vMerge && wAttr(vMerge, 'val') !== 'restart') td.classList.add('vmerge');
        const w = wChild(tcPr, 'tcW');
        if (w && wAttr(w, 'type') === 'dxa') td.style.width = `${Number(wAttr(w, 'w')) / 15}px`;
        const shd = wAttr(wChild(tcPr, 'shd'), 'fill');
        if (shd && /^[0-9a-f]{6}$/i.test(shd)) td.style.backgroundColor = '#' + shd;
        this.renderBlocks(tc, td);
        if (!td.childNodes.length) td.append(this.emptyParagraph());
        row.append(td);
      }
      tbody.append(row);
    }
    return table;
  }

  /**
   * Recomputes list markers for the editor content in document order, so that adding, removing or
   * restyling paragraphs keeps "1. 2. 3." correct while editing.
   */
  renumber(root) {
    if (!this.numberingXml && !root.querySelector('.list')) return;
    const numbering = new Numbering(this.numberingXml);
    for (const el of root.querySelectorAll('p,h1,h2,h3,h4,h5,h6')) {
      const pPr = el.dataset.ppr != null ? this.raw[Number(el.dataset.ppr)] : null;
      const numPr = wChild(pPr, 'numPr') || this.styles.styleNumPr(el.dataset.style);
      const num = numPr && numbering.next(numPr);
      if (num) {
        el.classList.add('list');
        if (el.dataset.marker !== num.marker) el.dataset.marker = num.marker;
      } else if (el.classList.contains('list')) {
        el.classList.remove('list');
        delete el.dataset.marker;
      }
    }
  }

  // ---------------------------------------------------------------- HTML -> XML

  /** Writes the editor content back into document.xml (in memory). */
  update(root) {
    const sectPr = wChild(this.body, 'sectPr');
    const nodes = this.serializeBlocks(root);
    while (this.body.firstChild) this.body.removeChild(this.body.firstChild);
    for (const n of nodes) this.body.appendChild(n);
    if (!nodes.some((n) => n.localName === 'p')) this.body.appendChild(createW(this.xml, 'p'));
    if (sectPr) this.body.appendChild(sectPr);
  }

  /** Produces a .docx Blob from the current editor content. */
  async save(root) {
    this.update(root);
    this.zip.file(this.mainPath, serializeXml(this.xml));
    await this.writeStyles();
    return this.zip.generateAsync({ type: 'blob', mimeType: DOCX_MIME, compression: 'DEFLATE' });
  }

  async writeStyles() {
    const xml = serializeXml(this.styles.doc);
    if (this.stylesPath) {
      this.zip.file(this.stylesPath, xml);
      return;
    }
    // The package had no styles part: add one plus its relationship and content type.
    this.stylesPath = `${this.mainDir ? this.mainDir + '/' : ''}styles.xml`;
    this.zip.file(this.stylesPath, xml);
    const relsXml = await this.text(this.relsPath);
    const rels = parseXml(relsXml || `<Relationships xmlns="${NS.pr}"/>`);
    const rel = rels.createElementNS(NS.pr, 'Relationship');
    let n = 1;
    while ([...rels.documentElement.children].some((r) => r.getAttribute('Id') === `rIdStyles${n}`)) n++;
    rel.setAttribute('Id', `rIdStyles${n}`);
    rel.setAttribute('Type', STYLES_REL);
    rel.setAttribute('Target', 'styles.xml');
    rels.documentElement.appendChild(rel);
    this.zip.file(this.relsPath, serializeXml(rels));
    const ct = parseXml(await this.text('[Content_Types].xml'));
    const ov = ct.createElementNS(NS.ct, 'Override');
    ov.setAttribute('PartName', '/' + this.stylesPath);
    ov.setAttribute('ContentType', STYLES_CT);
    ct.documentElement.appendChild(ov);
    this.zip.file('[Content_Types].xml', serializeXml(ct));
  }

  rawClone(index) {
    return this.raw[Number(index)].cloneNode(true);
  }

  serializeBlocks(container) {
    const out = [];
    let pending = [];
    const flush = () => {
      if (pending.some((n) => n.nodeType === 1 || n.data.trim())) out.push(this.serializeParagraph(null, pending));
      pending = [];
    };
    for (const node of container.childNodes) {
      if (node.nodeType === 1 && BLOCK_TAGS.has(node.tagName)) {
        flush();
        if (node.tagName === 'TABLE' && node.dataset.raw != null) out.push(this.serializeTable(node));
        else if (node.tagName === 'TABLE') out.push(...this.serializeBlocks(node.querySelector('tbody') || node));
        else if (node.dataset.raw != null) out.push(this.rawClone(node.dataset.raw));
        else if ([...node.children].some((c) => BLOCK_TAGS.has(c.tagName))) out.push(...this.serializeBlocks(node));
        else out.push(this.serializeParagraph(node, [...node.childNodes]));
      } else if (node.nodeType === 1 || node.nodeType === 3) {
        pending.push(node);
      }
    }
    flush();
    return out;
  }

  serializeTable(table) {
    const tbl = this.rawClone(table.dataset.raw);
    const rows = [...table.querySelectorAll(':scope > tbody > tr, :scope > tr')];
    wChildren(tbl, 'tr').forEach((tr, i) => {
      const cells = rows[i] ? [...rows[i].children].filter((c) => c.tagName === 'TD') : [];
      wChildren(tr, 'tc').forEach((tc, j) => {
        const td = cells[j];
        if (!td) return; // structure changed unexpectedly: keep original cell content
        for (const c of elementChildren(tc)) if (c.localName !== 'tcPr') tc.removeChild(c);
        const blocks = this.serializeBlocks(td);
        for (const b of blocks) tc.appendChild(b);
        // A cell must end with a paragraph.
        if (!blocks.length || blocks[blocks.length - 1].localName !== 'p') tc.appendChild(createW(this.xml, 'p'));
      });
    });
    return tbl;
  }

  serializeParagraph(el, nodes) {
    const doc = this.xml;
    const p = createW(doc, 'p');
    let pPr = el?.dataset.ppr != null ? this.rawClone(el.dataset.ppr) : createW(doc, 'pPr');
    const styleId = el?.dataset.style;
    if (styleId) ensureOrdered(pPr, 'pStyle', PPR_ORDER).setAttributeNS(W, 'w:val', styleId);
    else removeW(pPr, 'pStyle');
    const align = el?.style.textAlign || '';
    const jc = wChild(pPr, 'jc');
    if (!align || align === 'start') {
      removeW(pPr, 'jc');
    } else if (!jc || jcToCss(wAttr(jc, 'val')) !== align) {
      const val = { left: 'left', center: 'center', right: 'right', justify: 'both', end: 'right' }[align] || 'left';
      ensureOrdered(pPr, 'jc', PPR_ORDER).setAttributeNS(W, 'w:val', val);
    }
    if (!isEmptyElement(pPr)) p.appendChild(pPr);
    const state = { last: null };
    for (const n of nodes) this.serializeInline(n, p, {}, el, state);
    return p;
  }

  serializeInline(node, out, ctx, block, state) {
    if (node.nodeType === 3) {
      // Text is written exactly as it is: non-breaking and zero-width spaces are meaningful in Word,
      // and the editor uses white-space: pre-wrap, so typed spaces are ordinary spaces.
      const text = node.data;
      if (text) this.emitText(text, out, ctx, state);
      return;
    }
    if (node.nodeType !== 1) return;
    const el = node;
    if (el.tagName === 'BR') {
      if (!isTrailing(el, block)) this.emitRunChild(createW(this.xml, 'br'), out, ctx, state);
      return;
    }
    if (el.dataset.raw != null) {
      out.appendChild(this.rawClone(el.dataset.raw));
      state.last = null;
      return;
    }
    if (el.tagName === 'A' && el.dataset.hl != null) {
      const hl = this.rawClone(el.dataset.hl);
      const inner = { last: null };
      for (const c of el.childNodes) this.serializeInline(c, hl, ctx, block, inner);
      if (hl.childNodes.length) out.appendChild(hl);
      state.last = null;
      return;
    }
    if (el.tagName === 'IMG' || el.tagName === 'SCRIPT' || el.tagName === 'STYLE') return;
    const next = { ...ctx };
    if (el.dataset.rpr != null) {
      next.rpr = el.dataset.rpr;
      next.b = next.i = next.u = next.s = undefined; // the run's own tags/styles define these
    }
    switch (el.tagName) {
      case 'B': case 'STRONG': next.b = true; break;
      case 'I': case 'EM': next.i = true; break;
      case 'U': case 'INS': next.u = true; break;
      case 'S': case 'STRIKE': case 'DEL': next.s = true; break;
      default: break;
    }
    const st = el.style;
    if (st) {
      if (st.fontWeight) next.b = st.fontWeight === 'bold' || st.fontWeight === 'bolder' || Number(st.fontWeight) >= 600;
      if (st.fontStyle) next.i = st.fontStyle === 'italic' || st.fontStyle === 'oblique';
      const deco = st.textDecorationLine || st.textDecoration || '';
      if (deco.includes('underline')) next.u = true;
      if (deco.includes('line-through')) next.s = true;
    }
    for (const c of el.childNodes) this.serializeInline(c, out, next, block, state);
  }

  emitText(text, out, ctx, state) {
    const parts = text.split(/(\t|\n)/);
    for (const part of parts) {
      if (!part) continue;
      if (part === '\t') this.emitRunChild(createW(this.xml, 'tab'), out, ctx, state);
      else if (part === '\n') this.emitRunChild(createW(this.xml, 'br'), out, ctx, state);
      else {
        const t = createW(this.xml, 't');
        t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
        t.textContent = part;
        this.emitRunChild(t, out, ctx, state);
      }
    }
  }

  /** Appends a run child, merging into the previous run when formatting is identical. */
  emitRunChild(child, out, ctx, state) {
    const key = `${ctx.rpr ?? ''}|${ctx.b}|${ctx.i}|${ctx.u}|${ctx.s}`;
    if (state.last && state.last.key === key && state.last.run.parentNode === out && out.lastChild === state.last.run) {
      state.last.run.appendChild(child);
      return;
    }
    const run = createW(this.xml, 'r');
    const rPr = this.buildRPr(ctx);
    if (rPr) run.appendChild(rPr);
    run.appendChild(child);
    out.appendChild(run);
    state.last = { key, run };
  }

  buildRPr(ctx) {
    const rPr = ctx.rpr != null ? this.rawClone(ctx.rpr) : createW(this.xml, 'rPr');
    const toggle = (local, value) => {
      if (value === undefined) { removeW(rPr, local); return; }
      const e = ensureOrdered(rPr, local, RPR_ORDER);
      if (value) e.removeAttributeNS(W, 'val');
      else e.setAttributeNS(W, 'w:val', '0');
    };
    toggle('b', ctx.b);
    toggle('i', ctx.i);
    const u = wChild(rPr, 'u');
    if (ctx.u) {
      if (!u || wAttr(u, 'val') === 'none') ensureOrdered(rPr, 'u', RPR_ORDER).setAttributeNS(W, 'w:val', 'single');
    } else removeW(rPr, 'u');
    if (ctx.s) {
      if (!onOff(wChild(rPr, 'strike')) && !onOff(wChild(rPr, 'dstrike'))) {
        removeW(rPr, 'dstrike');
        toggle('strike', true);
      }
    } else removeW(rPr, 'strike', 'dstrike');
    return isEmptyElement(rPr) ? null : rPr;
  }
}

/** True if `br` is the browser's placeholder <br> at the very end of a block. */
function isTrailing(br, block) {
  let n = br;
  while (n && n !== block) {
    let s = n.nextSibling;
    while (s && s.nodeType === 3 && !s.data) s = s.nextSibling;
    if (s) return false;
    n = n.parentNode;
  }
  return true;
}

function cssToStyle(css) {
  const out = {};
  for (const [k, v] of Object.entries(css)) out[k.replace(/-([a-z])/g, (_, c) => c.toUpperCase())] = v;
  return out;
}
