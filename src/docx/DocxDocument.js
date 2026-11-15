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
  RPR_ORDER, PPR_ORDER, isEmptyElement, plainText, parseRels, elementChildren, XML_HEADER,
} from './xml.js';
import { StyleSheet, jcToCss } from './styles.js';
import { Numbering, addList, MAX_LEVEL } from './numbering.js';
import { blankDocxFiles } from './blank.js';
import {
  loadPart, parseComments, parseNotes, headerFooterVariants, fieldName, describe,
} from './annotations.js';
import {
  controlInfo, controlTitle, controlChoices, controlDate, formatWordDate, updateControlProperties,
} from './contentControls.js';
import { trackedChangeLabel, paragraphMarkLabel, paragraphMarkRevision, inlineRevisionKind } from './revisions.js';
import { reviewParagraphs } from './review.js';
import {
  COMMENT_PART_RELS, commentIdsIn, pairedCommentRanges, removedComments, pruneCommentMarkers, pruneCommentParts,
} from './commentRemoval.js';

const W = NS.w;
const MAIN_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const STYLES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const STYLES_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const NUMBERING_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const NUMBERING_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const NUMPR_ORDER = ['ilvl', 'numId', 'numberingChange', 'ins'];
const SETTINGS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';
const SETTINGS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';
// The settings that come before <w:trackRevisions> in CT_Settings.
const BEFORE_TRACK_REVISIONS = new Set([
  'writeProtection', 'view', 'zoom', 'removePersonalInformation', 'removeDateAndTime', 'doNotDisplayPageBoundaries',
  'displayBackgroundShape', 'printPostScriptOverText', 'printFractionalCharacterWidth', 'printFormsData',
  'embedTrueTypeFonts', 'embedSystemFonts', 'saveSubsetFonts', 'saveFormsData', 'mirrorMargins', 'alignBordersAndEdges',
  'bordersDoNotSurroundHeader', 'bordersDoNotSurroundFooter', 'gutterAtTop', 'hideSpellingErrors', 'hideGrammaticalErrors',
  'activeWritingStyle', 'proofState', 'formsDesign', 'attachedTemplate', 'linkStyles', 'stylePaneFormatFilter',
  'stylePaneSortMethod', 'documentType', 'mailMerge', 'revisionView',
]);
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
    this.originalBody = this.body.cloneNode(true); // what "Review my changes" compares with

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

    // Read-only companions of the body: comments, notes, and the first page's header and footer.
    const partOf = (suffix) => loadPart(this, relOf(suffix));
    const [comments, commentsExt, footnotes, endnotes] = await Promise.all(
      ['/comments', '/commentsExtended', '/footnotes', '/endnotes'].map(partOf),
    );
    this.comments = parseComments(comments, commentsExt);
    await this.initCommentRemoval(relOf);
    this.notes = { footnote: parseNotes(footnotes, 'footnote'), endnote: parseNotes(endnotes, 'endnote') };
    const settings = await partOf('/settings');
    const evenAndOdd = onOff(wChild(settings?.xml.documentElement, 'evenAndOddHeaders')) === true;
    // Track Changes: on when the document says so, as in Word. Saved back only if it was switched.
    this.settingsPath = settings?.path || null;
    this.settingsText = settings ? await this.text(settings.path) : null;
    this.trackRevisions = onOff(wChild(settings?.xml.documentElement, 'trackRevisions')) === true;
    this.tracking = this.trackRevisions;
    // New revisions get ids above every id in the document.
    let maxId = 0;
    for (const el of this.xml.getElementsByTagName('*')) {
      const id = Number(el.getAttributeNS(W, 'id'));
      if (Number.isInteger(id) && id > maxId) maxId = id;
    }
    this.firstRevisionId = maxId + 1;
    const loadVariants = async (local) => {
      const variants = await Promise.all(headerFooterVariants(this.body, local, evenAndOdd).map(async (v) => {
        const part = await loadPart(this, this.rels.get(v.rid));
        return part && { ...v, ...part };
      }));
      return variants.filter(Boolean);
    };
    this.headers = await loadVariants('headerReference');
    this.footers = await loadVariants('footerReference');
    this.numberingPath = numPath || null;
    this.numberingChanged = false;
    this.listDefs = {}; // list definitions added this session: kind -> abstractNumId, plus the bullet numId

    // Pre-create blob URLs for all images so rendering can stay synchronous.
    const mediaFiles = Object.values(this.zip.files).filter((f) => !f.dir && /\.(png|jpe?g|gif|bmp|svg|webp|tiff?)$/i.test(f.name));
    await Promise.all(mediaFiles.map(async (f) => {
      const ext = f.name.split('.').pop().toLowerCase();
      const blob = new Blob([await f.async('uint8array')], { type: IMAGE_TYPES[ext] || 'application/octet-stream' });
      this.mediaUrls.set(f.name, URL.createObjectURL(blob));
    }));
  }

  /**
   * What deleting a comment's text needs: the original comment parts, which comments the body
   * references (and which with a complete range), and which are referenced from other parts.
   */
  async initCommentRemoval(relOf) {
    this.commentParts = {};
    for (const [key, suffix] of Object.entries(COMMENT_PART_RELS)) {
      const path = relOf(suffix)?.path;
      const text = path && (await this.text(path));
      if (text) this.commentParts[key] = { path, text };
    }
    this.commentsInBody = commentIdsIn(this.body);
    this.commentRangesInBody = pairedCommentRanges(this.body);
    this.commentsElsewhere = new Set();
    if (!this.commentsInBody.size) return;
    const stories = Object.values(this.zip.files).filter((f) => !f.dir && /(^|\/)(header|footer|footnotes|endnotes)\d*\.xml$/.test(f.name));
    for (const f of stories) {
      const text = await f.async('string');
      if (text.includes('commentReference')) for (const id of commentIdsIn(parseXml(text))) this.commentsElsewhere.add(id);
    }
  }

  /**
   * The comments that would be removed if the editor content `root` were saved now (a Set of ids):
   * their reference was deleted with the text they were attached to.
   */
  removedComments(root) {
    const present = new Set();
    for (const el of root.querySelectorAll('[data-raw]')) {
      for (const id of this.commentRefs.get(el.dataset.raw) || []) present.add(id);
    }
    return removedComments(this.comments, { wasInBody: this.commentsInBody, present, elsewhere: this.commentsElsewhere });
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
      header: px(wAttr(pgMar, 'header'), 48),
      footer: px(wAttr(pgMar, 'footer'), 48),
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
    this.commentOrder = new Map(); // comment id -> marker number, in document order
    this.noteCount = { footnote: 0, endnote: 0 };
    const frag = document.createDocumentFragment();
    this.renderBlocks(this.body, frag);
    if (!frag.childNodes.length) frag.append(this.emptyParagraph());
    this.applyContextualSpacing(frag);
    this.markTracked(frag);
    return frag;
  }

  /**
   * Labels everything that is or holds a tracked change (data-tracked), so the editor can refuse
   * edits that would remove it: deleting a revision would silently accept or reject it.
   */
  markTracked(root, reset = true) {
    if (reset) this.commentRefs = new Map(); // raw index -> ids of the comments it references (see removedComments)
    for (const el of root.querySelectorAll('[data-raw]')) {
      const node = this.raw[Number(el.dataset.raw)];
      const label = trackedChangeLabel(node);
      if (label) el.dataset.tracked = label;
      // A table's cells are rendered (and saved) on their own, so its references are counted there.
      const refs = node.localName === 'tbl' ? [] : [...commentIdsIn(node)];
      if (refs.length) this.commentRefs.set(el.dataset.raw, refs);
    }
    for (const el of root.querySelectorAll('[data-ppr]')) {
      const pPr = this.raw[Number(el.dataset.ppr)];
      const label = paragraphMarkLabel(pPr);
      if (label) el.dataset.tracked = label;
      const rev = paragraphMarkRevision(pPr);
      if (rev) {
        el.dataset.revMark = rev.kind; // shown as ¶ (see EditorView), and accepted or rejected there
        el.dataset.revTitle = `${rev.kind === 'ins' ? 'Paragraph break inserted' : 'Paragraph break deleted'}${rev.author ? ` by ${rev.author}` : ''}${rev.date ? ` on ${rev.date.slice(0, 10)}` : ''}`;
      }
    }
  }

  // ---------------------------------------------------------------- accepting and rejecting

  /**
   * The tracked change an element shows, or null: { kind: 'ins' | 'del', mark (a paragraph break),
   * own (made this session), author, date }. Inline revisions are spans; a paragraph break is its
   * paragraph.
   */
  revisionOf(el) {
    if (el?.nodeType !== 1) return null;
    const d = el.dataset;
    if (d.tc) return { kind: d.tc, mark: false, own: true, author: d.tcAuthor, date: d.tcDate };
    if (d.tcMark) return { kind: d.tcMark, mark: true, own: true, author: d.tcMarkAuthor, date: d.tcMarkDate };
    if (d.revMark && d.ppr != null) {
      const rev = paragraphMarkRevision(this.raw[Number(d.ppr)]);
      return rev && { kind: rev.kind, mark: true, own: false, author: rev.author, date: rev.date };
    }
    if (d.raw != null && el.classList.contains('raw-inline')) {
      const node = this.raw[Number(d.raw)];
      const kind = inlineRevisionKind(node);
      return kind && { kind, mark: false, own: false, author: wAttr(node, 'author'), date: wAttr(node, 'date') };
    }
    return null;
  }

  /**
   * Accepts (or rejects) the tracked change shown by `el`, in the editor: an insertion accepted (or
   * a deletion rejected) becomes ordinary, editable text; the other way round, its text goes. A
   * paragraph break kept loses its revision; one removed joins its paragraph with the next, which
   * keeps its own properties (its break is the one that stays), as in Word.
   */
  resolveRevision(el, accept) {
    const rev = this.revisionOf(el);
    if (!rev) return;
    const keep = (rev.kind === 'ins') === accept;
    if (rev.mark) {
      this.resolveMark(el, rev, keep);
    } else if (!keep) {
      el.remove();
    } else if (rev.own) {
      el.replaceWith(...el.childNodes);
    } else {
      // Someone else's revision: render what it holds as ordinary text (deleted text undeleted).
      const node = this.raw[Number(el.dataset.raw)].cloneNode(true);
      if (rev.kind === 'del') {
        for (const [from, to] of [['delText', 't'], ['delInstrText', 'instrText']]) {
          for (const t of [...node.getElementsByTagNameNS(W, from)]) {
            const r = createW(this.xml, to);
            for (const a of [...t.attributes]) r.setAttributeNS(a.namespaceURI, a.name, a.value);
            r.textContent = t.textContent;
            t.replaceWith(r);
          }
        }
      }
      const holder = document.createElement('span');
      this.renderInlines(node, holder);
      this.markTracked(holder, false);
      el.replaceWith(...holder.childNodes);
    }
  }

  resolveMark(el, rev, keep) {
    let next = el.nextElementSibling;
    while (next?.classList.contains('raw-hidden')) next = next.nextElementSibling;
    // Removing the last paragraph's break (or one before a table) isn't possible: it stays.
    if (keep || !next?.matches('p, h1, h2, h3, h4, h5, h6')) {
      if (rev.own) {
        delete el.dataset.tcMark;
        delete el.dataset.tcMarkAuthor;
        delete el.dataset.tcMarkDate;
        return;
      }
      const pPr = this.raw[Number(el.dataset.ppr)].cloneNode(true);
      const rPr = wChild(pPr, 'rPr');
      removeW(rPr, 'ins', 'del', 'moveTo', 'moveFrom');
      if (rPr && !elementChildren(rPr).length) pPr.removeChild(rPr);
      el.dataset.ppr = this.keep(pPr);
      delete el.dataset.revMark;
      delete el.dataset.revTitle;
      delete el.dataset.tracked;
      return;
    }
    const kids = [...el.childNodes].filter((n) => !(n.tagName === 'BR' && n === el.lastChild && !el.textContent));
    if (next.lastChild?.tagName === 'BR' && !next.textContent && kids.length) next.lastChild.remove();
    next.prepend(...kids);
    el.remove();
  }

  /** How many tracked formatting changes the editor content holds (these are accepted in Word). */
  formattingChanges(root) {
    const seen = new Set();
    for (const el of root.querySelectorAll('[data-rpr], [data-ppr]')) {
      const node = this.raw[Number(el.dataset.rpr ?? el.dataset.ppr)];
      if (wChild(node, 'rPrChange') || wChild(node, 'pPrChange')) seen.add(el.dataset.rpr ?? `p${el.dataset.ppr}`);
    }
    return seen.size;
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
      else if (local === 'sdt') parent.append(this.renderControl(child, true));
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
      markCommentRange(div, node);
      return div;
    }
    div.className = 'raw-block';
    div.title = `${describe(node)} – kept as it is; edit it in Word`;
    div.textContent = plainText(node).trim() || `[${node.nodeName}]`;
    return div;
  }

  /**
   * A content control: a frame (not editable itself) around its rendered content. Text controls hold
   * an editable `.sdt-content`; checkbox, drop-down and date controls change through clicks. Read-only
   * controls are kept verbatim (data-raw), the others are rebuilt on save (data-sdt).
   */
  renderControl(sdt, block) {
    const info = controlInfo(sdt);
    const content = wChild(sdt, 'sdtContent');
    const el = document.createElement(block ? 'div' : 'span');
    el.contentEditable = 'false';
    el.className = `sdt sdt-${info.kind}`;
    el.title = controlTitle(info);
    const render = (into) => {
      if (content) {
        if (block) this.renderBlocks(content, into);
        else this.renderInlines(content, into);
      }
      if (block && !into.firstChild) into.append(this.emptyParagraph());
    };
    if (info.readOnly) {
      el.dataset.raw = this.keep(sdt);
      el.classList.add('sdt-locked');
      render(el);
      for (const e of el.querySelectorAll('[contenteditable="true"]')) e.contentEditable = 'false';
      return el;
    }
    el.dataset.sdt = this.keep(sdt);
    if (info.placeholder) el.classList.add('showing-placeholder');
    if (info.kind === 'checkbox') el.dataset.checked = info.checked ? '1' : '0';
    if (info.kind === 'text' || info.kind === 'combo') {
      const box = document.createElement(block ? 'div' : 'span');
      box.className = 'sdt-content';
      box.contentEditable = 'true';
      if (info.plain) box.dataset.plain = info.multiLine ? 'multi' : 'single';
      render(box);
      el.append(box);
    } else {
      render(el);
    }
    return el;
  }

  /** The control settings behind a rendered `.sdt` element. */
  control(el) {
    const sdt = this.raw[Number(el.dataset.sdt ?? el.dataset.raw)];
    const info = controlInfo(sdt);
    return {
      ...info,
      choices: controlChoices(sdt),
      date: info.kind === 'date' ? controlDate(sdt) : null,
    };
  }

  formatDate(el, iso) {
    const { format, locale } = this.control(el).date;
    return formatWordDate(iso, format, locale);
  }

  isPlaceholderStyle(styleId) {
    return this.styles.byId.get(styleId)?.name.toLowerCase() === 'placeholder text';
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
    const direct = wChild(pPr, 'numPr');
    if (direct) {
      // The paragraph's own list membership is editable state: see listOf() and listEdited().
      const list = this.originalList(el, pPr);
      el.dataset.numId = list?.numId ?? '0';
      el.dataset.ilvl = String(list?.ilvl ?? 0);
    }
    const list = this.listOf(el);
    const num = list && this.numbering.next(list.numId, list.ilvl);
    if (num) {
      el.classList.add('list');
      el.dataset.marker = num.marker;
      if (num.pPr && !wChild(pPr, 'ind')) this.indentForList(el, num, list.ilvl);
      if (!el.style.paddingLeft) el.style.paddingLeft = '36pt';
      el.style.textIndent = '';
    }
    this.renderInlines(p, el);
    if (!el.textContent && !el.querySelector('img,.raw-inline:not(.raw-hidden)')) el.append(document.createElement('br'));
    return el;
  }

  renderInlines(container, outer) {
    // Complex fields are runs: begin, instruction, separate, result…, end. The markers are kept
    // hidden and the result (still editable, as in Word) is wrapped to say what the field is.
    const fields = [];
    let parent = outer;
    for (const child of elementChildren(container)) {
      const local = child.namespaceURI === W ? child.localName : null;
      if (local === 'pPr' || local === 'proofErr') continue;
      const mark = local === 'r' && fieldMarker(child);
      if (mark) {
        parent.append(this.hiddenRaw(child));
        if (mark.type === 'begin') fields.push({ instr: '', outer: parent });
        else if (mark.type === 'instr' && fields.length) fields[fields.length - 1].instr += mark.text;
        else if (mark.type === 'separate' && fields.length) {
          const wrap = document.createElement('span');
          wrap.className = 'field';
          wrap.title = `${fieldName(fields[fields.length - 1].instr)} – Word updates it`;
          parent.append(wrap);
          parent = wrap;
        } else if (mark.type === 'end' && fields.length) {
          parent = fields.pop().outer;
        }
        continue;
      }
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
      } else if (local === 'sdt') {
        parent.append(this.renderControl(child, false));
      } else if (local && FLATTEN_INLINE.has(local)) {
        // Unwrap smart tags and similar containers: render their runs as normal text.
        this.renderInlines(child, parent);
      } else {
        parent.append(this.rawInline(child));
      }
    }
  }

  hiddenRaw(node) {
    const span = document.createElement('span');
    span.contentEditable = 'false';
    span.dataset.raw = this.keep(node);
    span.className = 'raw-inline raw-hidden';
    markCommentRange(span, node);
    return span;
  }

  rawInline(node) {
    const local = node.localName;
    if (HIDDEN_INLINE.has(local)) return this.hiddenRaw(node);
    const commentRef = local === 'r' && onlyChild(node, 'commentReference');
    if (commentRef) {
      const id = wAttr(commentRef, 'id');
      const comment = this.comments.find((c) => c.id === id);
      const span = this.hiddenRaw(node);
      span.dataset.commentRef = id; // the comment's mark: deleting it deletes the comment
      if (comment?.parentId) return span; // replies are shown with their thread
      span.className = 'raw-inline comment-ref';
      span.dataset.commentId = id;
      if (!this.commentOrder.has(id)) this.commentOrder.set(id, this.commentOrder.size + 1);
      span.dataset.n = String(this.commentOrder.get(id));
      span.title = comment ? `Comment by ${comment.author}: ${comment.paragraphs.join(' ')}` : 'Comment';
      return span;
    }
    const span = document.createElement('span');
    span.contentEditable = 'false';
    span.dataset.raw = this.keep(node);
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
      if (span.childNodes.length) {
        span.title = `${describe(node)} – kept as it is; edit it in Word`;
        return span;
      }
    }
    if (local === 'r' && wChildren(node, 'br').some((b) => wAttr(b, 'type') === 'page')) {
      span.classList.add('page-break');
      span.title = 'Page break';
      return span;
    }
    const noteRef = local === 'r' && (wChild(node, 'footnoteReference') || wChild(node, 'endnoteReference'));
    if (noteRef) {
      const kind = noteRef.localName === 'footnoteReference' ? 'footnote' : 'endnote';
      const n = ++this.noteCount[kind];
      span.classList.add('note-ref');
      span.dataset.n = kind === 'footnote' ? String(n) : romanLower(n);
      const text = this.notes[kind].get(wAttr(noteRef, 'id'));
      span.title = `${kind === 'footnote' ? 'Footnote' : 'Endnote'} ${span.dataset.n}${text ? `: ${text}` : ''} – edit it in Word`;
      return span;
    }
    const sym = local === 'r' && wChild(node, 'sym');
    if (sym) {
      const code = parseInt(wAttr(sym, 'char') || '', 16);
      span.classList.add('symbol');
      span.textContent = Number.isFinite(code) ? String.fromCodePoint(code >= 0xf000 ? code - 0xf000 : code) : '?';
      const font = wAttr(sym, 'font');
      if (font) span.style.fontFamily = `"${font}"`;
      span.title = `Symbol${font ? ` (${font})` : ''}`;
      return span;
    }
    if (TRACKED_INLINE[local]) {
      span.classList.add(local === 'ins' || local === 'moveTo' ? 'inserted' : 'deleted');
      const author = wAttr(node, 'author');
      const date = wAttr(node, 'date');
      span.title = `${TRACKED_INLINE[local]}${author ? ` by ${author}` : ''}${date ? ` on ${date.slice(0, 10)}` : ''} – tracked change: Accept or Reject on the toolbar`;
      span.textContent = plainText(node) || '⧉';
      return span;
    }
    span.title = `${describe(node)} – kept as it is; edit it in Word`;
    span.textContent = plainText(node) || '⧉';
    return span;
  }

  /** A new page break (<w:r><w:br w:type="page"/></w:r>), rendered as rawInline renders one. */
  newPageBreak() {
    const r = createW(this.xml, 'r');
    r.appendChild(createW(this.xml, 'br')).setAttributeNS(W, 'w:type', 'page');
    const span = document.createElement('span');
    span.contentEditable = 'false';
    span.dataset.raw = this.keep(r);
    span.className = 'raw-inline page-break';
    span.title = 'Page break';
    return span;
  }

  /** A header or footer (an entry of this.headers / this.footers), rendered read-only, or null if empty. */
  renderHeaderFooter(part) {
    if (!part) return null;
    const [rels, numbering, order, notes] = [this.rels, this.numbering, this.commentOrder, this.noteCount];
    this.rels = part.rels;
    this.numbering = new Numbering(this.numberingXml);
    this.commentOrder = new Map();
    this.noteCount = { footnote: 0, endnote: 0 };
    const frag = document.createDocumentFragment();
    try {
      this.renderBlocks(part.xml.documentElement, frag);
    } finally {
      [this.rels, this.numbering, this.commentOrder, this.noteCount] = [rels, numbering, order, notes];
    }
    for (const e of frag.querySelectorAll('[contenteditable]')) e.removeAttribute('contenteditable');
    const holder = document.createElement('div');
    holder.append(frag);
    if (!holder.textContent.trim() && !holder.querySelector('img')) return null;
    return holder;
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
      const va = wAttr(wChild(rPr, 'vertAlign'), 'val');
      if (va === 'superscript') wrap('sup');
      if (va === 'subscript') wrap('sub');
      const hl = wAttr(wChild(rPr, 'highlight'), 'val');
      if (hl && hl !== 'none') {
        wrap('mark');
        target.dataset.color = hl;
      }
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
    const tblStyle = wAttr(wChild(wChild(tbl, 'tblPr'), 'tblStyle'), 'val');
    if (tblStyle) table.dataset.tstyle = tblStyle;
    const tbody = document.createElement('tbody');
    table.append(tbody);
    for (const { el: tr, wrappers: rowWrappers } of tableParts(tbl, 'tr')) {
      const row = document.createElement('tr');
      if (rowWrappers.length) {
        row.classList.add('sdt-row');
        row.title = controlTitle(controlInfo(rowWrappers[rowWrappers.length - 1]));
      }
      for (const { el: tc, wrappers: cellWrappers } of tableParts(tr, 'tc')) {
        const td = document.createElement('td');
        td.contentEditable = 'true';
        // Content controls around whole rows or cells: the cell is the field. Anything but a text
        // field (and anything locked or bound) stays read-only.
        const wrappers = [...rowWrappers, ...cellWrappers];
        if (wrappers.length) {
          const infos = wrappers.map(controlInfo);
          td.classList.add('sdt-cell');
          if (cellWrappers.length) td.title = controlTitle(infos[infos.length - 1]);
          if (infos.some((i) => i.readOnly || i.kind !== 'text')) {
            td.contentEditable = 'false';
            td.classList.add('sdt-locked');
          } else if (infos[infos.length - 1].placeholder) {
            td.classList.add('showing-placeholder');
          }
        }
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
    this.applyContextualSpacing(root);
    if (!this.numberingXml && !root.querySelector('.list')) return;
    const numbering = new Numbering(this.numberingXml);
    this.numbering = numbering;
    for (const el of root.querySelectorAll('p,h1,h2,h3,h4,h5,h6')) {
      const list = this.listOf(el);
      const num = list && numbering.next(list.numId, list.ilvl);
      // Indentation follows the list level, unless the paragraph keeps its own (as when rendered).
      const edited = this.listEdited(el);
      if (num) {
        el.classList.add('list');
        if (el.dataset.marker !== num.marker) el.dataset.marker = num.marker;
        const pPr = el.dataset.ppr != null ? this.raw[Number(el.dataset.ppr)] : null;
        if (edited || !wChild(pPr, 'ind')) this.indentForList(el, num, list.ilvl);
        else el.style.paddingLeft = this.styles.pPrCss(pPr)['padding-left'] || '36pt';
      } else if (el.classList.contains('list')) {
        // Out of the list: back to the paragraph's own indentation (none once its list was edited).
        el.classList.remove('list');
        delete el.dataset.marker;
        const pPr = el.dataset.ppr != null && !edited ? this.raw[Number(el.dataset.ppr)] : null;
        const css = this.styles.pPrCss(pPr);
        el.style.paddingLeft = css['padding-left'] || '';
        el.style.textIndent = css['text-indent'] || '';
      }
    }
  }

  /**
   * Word's "Don't add space between paragraphs of the same style" (w:contextualSpacing): such a
   * paragraph drops its space before / after when the paragraph above / below has the same style.
   * This is what keeps list items tight. Shown with classes; the document itself is unchanged.
   */
  applyContextualSpacing(root) {
    const blocks = root.querySelectorAll('p, h1, h2, h3, h4, h5, h6');
    const styleOf = (el) => el.dataset.style || this.styles.defaultParagraph || '';
    const sibling = (el, dir) => {
      let s = dir < 0 ? el.previousElementSibling : el.nextElementSibling;
      while (s?.classList.contains('raw-hidden')) s = dir < 0 ? s.previousElementSibling : s.nextElementSibling;
      return s?.matches('p, h1, h2, h3, h4, h5, h6') ? s : null;
    };
    for (const el of blocks) {
      const pPr = el.dataset.ppr != null ? this.raw[Number(el.dataset.ppr)] : null;
      const own = onOff(wChild(pPr, 'contextualSpacing'));
      const on = own !== undefined ? own : this.styles.contextualSpacing(el.dataset.style);
      const sameStyle = (other) => !!other && styleOf(other) === styleOf(el);
      el.classList.toggle('cs-before', !!on && sameStyle(sibling(el, -1)));
      el.classList.toggle('cs-after', !!on && sameStyle(sibling(el, 1)));
    }
  }

  indentForList(el, num, ilvl) {
    const { 'padding-left': pl } = num.pPr ? this.styles.pPrCss(num.pPr) : {};
    el.style.paddingLeft = pl || `${36 * (ilvl + 1)}pt`;
    el.style.textIndent = '';
  }

  // ---------------------------------------------------------------- lists

  /** The list a block is in, as { numId, ilvl }, or null. Direct numbering wins over the style's. */
  listOf(el) {
    if (el.dataset.numId != null) {
      return el.dataset.numId === '0' ? null : { numId: el.dataset.numId, ilvl: Number(el.dataset.ilvl || 0) };
    }
    return this.styleList(el.dataset.style);
  }

  styleList(styleId) {
    const numPr = this.styles.styleNumPr(styleId);
    const numId = wAttr(wChild(numPr, 'numId'), 'val');
    if (!numId || numId === '0') return null;
    return { numId, ilvl: Number(wAttr(wChild(numPr, 'ilvl'), 'val') ?? 0) };
  }

  /** The list given by a paragraph's stored numPr (falling back to its style for missing parts). */
  originalList(el, pPr) {
    const numPr = wChild(pPr, 'numPr');
    if (!numPr) return undefined;
    const styleNumPr = this.styles.styleNumPr(el.dataset.style);
    const numId = wAttr(wChild(numPr, 'numId'), 'val') ?? wAttr(wChild(styleNumPr, 'numId'), 'val');
    const ilvl = wAttr(wChild(numPr, 'ilvl'), 'val') ?? wAttr(wChild(styleNumPr, 'ilvl'), 'val') ?? '0';
    return !numId || numId === '0' ? null : { numId, ilvl: Number(ilvl) };
  }

  /** True if the user changed which list, or which level, the paragraph is in. */
  listEdited(el) {
    const pPr = el.dataset.ppr != null ? this.raw[Number(el.dataset.ppr)] : null;
    const orig = this.originalList(el, pPr);
    if (el.dataset.numId == null) return orig !== undefined; // direct numbering was removed
    if (orig === undefined) return true; // direct numbering was added
    const now = this.listOf(el);
    return (orig?.numId ?? '0') !== (now?.numId ?? '0') || (orig?.ilvl ?? 0) !== (now?.ilvl ?? 0);
  }

  listKind(list) {
    if (!list) return null;
    this.kinds ||= new Numbering(this.numberingXml);
    return this.kinds.kind(list.numId, list.ilvl);
  }

  /** True if list `numId` defines level `ilvl` (Word's "List Bullet" style, for one, has only one). */
  hasListLevel(numId, ilvl) {
    this.kinds ||= new Numbering(this.numberingXml);
    return ilvl >= 0 && ilvl <= MAX_LEVEL && this.kinds.hasLevel(numId, ilvl);
  }

  /** Puts a block in list `numId` at level `ilvl`. */
  setList(el, numId, ilvl = 0) {
    el.dataset.numId = numId;
    el.dataset.ilvl = String(Math.max(0, Math.min(MAX_LEVEL, ilvl)));
  }

  /**
   * Takes a block out of its list, as Word does when you press Enter on an empty item: a paragraph
   * whose list came from a list style ("List Bullet", "List Paragraph") goes back to Normal.
   */
  removeList(el) {
    const styleId = el.dataset.style;
    const isHeading = !!this.styles.kindOf(styleId);
    if (styleId && !isHeading && (this.styleList(styleId) || this.styles.isListParagraph(styleId))) delete el.dataset.style;
    if (this.styleList(el.dataset.style)) {
      el.dataset.numId = '0'; // e.g. a numbered heading style: switch its numbering off
      el.dataset.ilvl = '0';
    } else {
      delete el.dataset.numId;
      delete el.dataset.ilvl;
    }
  }

  /** A numId for a new list of `kind` ('bullet' or 'number'); bullets share one list. */
  newList(kind) {
    if (kind === 'bullet' && this.listDefs.bulletNumId) return this.listDefs.bulletNumId;
    if (!this.numberingXml) this.numberingXml = parseXml(`<w:numbering xmlns:w="${W}"/>`);
    const numId = addList(this.numberingXml, kind, this.listDefs);
    if (kind === 'bullet') this.listDefs.bulletNumId = numId;
    this.numberingChanged = true;
    this.kinds = null;
    return numId;
  }

  // ---------------------------------------------------------------- HTML -> XML

  /** Writes the editor content back into document.xml (in memory). */
  update(root) {
    const sectPr = wChild(this.body, 'sectPr');
    this.revisions = []; // <w:ins>/<w:del> written for this session's tracked changes
    const nodes = this.serializeBlocks(root);
    this.revisions.forEach((rev, i) => rev.setAttributeNS(W, 'w:id', String(this.firstRevisionId + i)));
    while (this.body.firstChild) this.body.removeChild(this.body.firstChild);
    for (const n of nodes) this.body.appendChild(n);
    if (!nodes.some((n) => n.localName === 'p')) this.body.appendChild(createW(this.xml, 'p'));
    if (sectPr) this.body.appendChild(sectPr);
    this.commentsRemoved = removedComments(this.comments, {
      wasInBody: this.commentsInBody, present: commentIdsIn(this.body), elsewhere: this.commentsElsewhere,
    });
    pruneCommentMarkers(this.body, this.commentsRemoved, this.commentRangesInBody);
  }

  /**
   * What saving the editor content now would change, for "Review my changes": the paragraphs that
   * differ, comments removed with their text, this session's tracked changes, and the other parts
   * that would change.
   */
  review(root) {
    this.update(root);
    const trackedParagraphs = new Set();
    for (const rev of this.revisions) {
      for (let n = rev.parentNode; n?.nodeType === 1; n = n.parentNode) {
        if (n.namespaceURI === W && n.localName === 'p') { trackedParagraphs.add(n); break; }
      }
    }
    const styleName = (id) => {
      const name = (id ? this.styles.byId.get(id)?.name || id : this.styles.byId.get(this.styles.defaultParagraph)?.name) || 'Normal';
      return name[0].toUpperCase() + name.slice(1); // built-in names are stored as "heading 1"
    };
    return {
      paragraphs: reviewParagraphs(this.originalBody, this.body, { tracked: (p) => trackedParagraphs.has(p), styleName }),
      commentsRemoved: this.comments.filter((c) => this.commentsRemoved.has(c.id)),
      insertions: this.revisions.filter((r) => r.localName === 'ins').length,
      deletions: this.revisions.filter((r) => r.localName === 'del').length,
      tracking: this.tracking,
      trackingSwitched: this.tracking !== this.trackRevisions,
      stylesAdded: this.styles.changed,
      listsAdded: this.numberingChanged,
    };
  }

  /** Produces a .docx Blob from the current editor content. */
  async save(root) {
    this.update(root);
    this.zip.file(this.mainPath, serializeXml(this.xml));
    await this.writeStyles();
    if (this.numberingChanged) await this.writeNumbering();
    this.writeComments();
    await this.writeSettings();
    return this.zip.generateAsync({ type: 'blob', mimeType: DOCX_MIME, compression: 'DEFLATE' });
  }

  async writeStyles() {
    // Only when a style was added: re-serializing an untouched styles.xml would still reorder its
    // namespace declarations and the like, and nothing the user didn't touch should change.
    if (!this.styles.changed) return;
    const xml = serializeXml(this.styles.doc);
    if (this.stylesPath) this.zip.file(this.stylesPath, xml);
    else this.stylesPath = await this.addPart('styles.xml', xml, STYLES_REL, STYLES_CT);
  }

  /**
   * Saves whether Track Changes is on (w:trackRevisions), so Word keeps tracking for the recipient.
   * settings.xml is only rewritten when it was switched; switched back, the original goes back.
   */
  async writeSettings() {
    if (this.tracking === this.trackRevisions) {
      if (this.settingsChanged) this.zip.file(this.settingsPath, this.settingsText);
      this.settingsChanged = false;
      return;
    }
    this.settingsChanged = true;
    if (!this.settingsPath) {
      this.settingsPath = await this.addPart('settings.xml', `${XML_HEADER}<w:settings xmlns:w="${W}"><w:trackRevisions/></w:settings>`, SETTINGS_REL, SETTINGS_CT);
      this.settingsText = `${XML_HEADER}<w:settings xmlns:w="${W}"/>`;
      return;
    }
    const doc = parseXml(this.settingsText);
    const root = doc.documentElement;
    if (this.tracking) {
      const before = wChildren(root).filter((c) => BEFORE_TRACK_REVISIONS.has(c.localName)).pop();
      root.insertBefore(createW(doc, 'trackRevisions'), before ? before.nextSibling : root.firstChild);
    } else {
      removeW(root, 'trackRevisions');
    }
    this.zip.file(this.settingsPath, serializeXml(doc));
  }

  /** Takes removed comments out of the comment parts (or puts back the original parts if none). */
  writeComments() {
    if (!this.commentsRemoved.size && !this.commentPartsChanged) return;
    for (const [path, xml] of Object.entries(pruneCommentParts(this.commentParts, this.comments, this.commentsRemoved))) {
      this.zip.file(path, xml);
    }
    this.commentPartsChanged = this.commentsRemoved.size > 0;
  }

  async writeNumbering() {
    const xml = serializeXml(this.numberingXml);
    if (this.numberingPath) this.zip.file(this.numberingPath, xml);
    else this.numberingPath = await this.addPart('numbering.xml', xml, NUMBERING_REL, NUMBERING_CT);
  }

  /** Adds a part next to the main document, with its relationship and content type. */
  async addPart(name, xml, relType, contentType) {
    const path = `${this.mainDir ? this.mainDir + '/' : ''}${name}`;
    this.zip.file(path, xml);
    const relsXml = await this.text(this.relsPath);
    const rels = parseXml(relsXml || `<Relationships xmlns="${NS.pr}"/>`);
    const rel = rels.createElementNS(NS.pr, 'Relationship');
    const prefix = `rId${name[0].toUpperCase()}${name.slice(1, name.indexOf('.'))}`;
    let n = 1;
    while ([...rels.documentElement.children].some((r) => r.getAttribute('Id') === `${prefix}${n}`)) n++;
    rel.setAttribute('Id', `${prefix}${n}`);
    rel.setAttribute('Type', relType);
    rel.setAttribute('Target', name);
    rels.documentElement.appendChild(rel);
    this.zip.file(this.relsPath, serializeXml(rels));
    const ct = parseXml(await this.text('[Content_Types].xml'));
    const ov = ct.createElementNS(NS.ct, 'Override');
    ov.setAttribute('PartName', '/' + path);
    ov.setAttribute('ContentType', contentType);
    ct.documentElement.appendChild(ov);
    this.zip.file('[Content_Types].xml', serializeXml(ct));
    return path;
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
        else if (node.dataset.sdt != null) out.push(this.serializeControl(node));
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
    const original = this.raw[Number(table.dataset.raw)];
    const tbl = original.cloneNode(true);
    const rows = [...table.querySelectorAll(':scope > tbody > tr, :scope > tr')];
    // Rows and cells are matched by the same walk as rendering, through row/cell-level wrappers.
    tableParts(tbl, 'tr').forEach(({ el: tr }, i) => {
      const cells = rows[i] ? [...rows[i].children].filter((c) => c.tagName === 'TD') : [];
      tableParts(tr, 'tc').forEach(({ el: tc }, j) => {
        const td = cells[j];
        if (!td || td.contentEditable === 'false') return; // read-only, or structure changed: keep it
        for (const c of elementChildren(tc)) if (c.localName !== 'tcPr') tc.removeChild(c);
        const blocks = this.serializeBlocks(td);
        for (const b of blocks) tc.appendChild(b);
        // A cell must end with a paragraph.
        if (!blocks.length || blocks[blocks.length - 1].localName !== 'p') tc.appendChild(createW(this.xml, 'p'));
      });
    });
    // Placeholders of row/cell fields: replaced text drops the placeholder state; untouched or
    // cleared placeholder text is written back as it was.
    const before = structuralControls(original);
    structuralControls(tbl).forEach((sdt, i) => {
      const orig = before[i];
      if (!orig || !controlInfo(orig).placeholder) return;
      const was = plainText(wChild(orig, 'sdtContent'));
      const now = plainText(wChild(sdt, 'sdtContent'));
      if (now === was || !now.trim()) sdt.replaceChild(wChild(orig, 'sdtContent').cloneNode(true), wChild(sdt, 'sdtContent'));
      else updateControlProperties(sdt, { dataset: {} }, { placeholderReplaced: true, isPlaceholderStyle: (id) => this.isPlaceholderStyle(id) });
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
    if (el && this.listEdited(el)) {
      // Write the new list membership; the list level's indentation replaces direct indentation.
      removeW(pPr, 'ind');
      if (el.dataset.numId == null) {
        removeW(pPr, 'numPr');
      } else {
        const numPr = ensureOrdered(pPr, 'numPr', PPR_ORDER);
        ensureOrdered(numPr, 'ilvl', NUMPR_ORDER).setAttributeNS(W, 'w:val', el.dataset.ilvl || '0');
        ensureOrdered(numPr, 'numId', NUMPR_ORDER).setAttributeNS(W, 'w:val', el.dataset.numId);
      }
    }
    const mark = el?.dataset.tcMark;
    if (mark === 'ins' || mark === 'del') {
      // The paragraph's break is a tracked change of this session: <w:pPr><w:rPr><w:ins|del/>.
      const rPr = ensureOrdered(pPr, 'rPr', PPR_ORDER);
      const rev = this.newRevision(mark, el.dataset.tcMarkAuthor, el.dataset.tcMarkDate);
      const ins = wChild(rPr, 'ins');
      rPr.insertBefore(rev, mark === 'del' && ins ? ins.nextSibling : rPr.firstChild);
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
    if (el.dataset.tc === 'ins' || el.dataset.tc === 'del') {
      this.serializeRevision(el, out, ctx, block, state);
      return;
    }
    if (el.dataset.raw != null) {
      out.appendChild(this.rawClone(el.dataset.raw));
      state.last = null;
      return;
    }
    if (el.dataset.sdt != null) {
      out.appendChild(this.serializeControl(el, ctx, block));
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
      next.b = next.i = next.u = next.s = next.va = next.hl = undefined; // the run's own tags/styles define these
    }
    switch (el.tagName) {
      case 'B': case 'STRONG': next.b = true; break;
      case 'I': case 'EM': next.i = true; break;
      case 'U': case 'INS': next.u = true; break;
      case 'S': case 'STRIKE': case 'DEL': next.s = true; break;
      case 'SUP': next.va = 'superscript'; break;
      case 'SUB': next.va = 'subscript'; break;
      case 'MARK': next.hl = el.dataset.color || 'yellow'; break;
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

  /** A revision of this session, for <w:ins>, <w:del> or a paragraph mark (its id is set in update). */
  newRevision(kind, author, date) {
    const rev = createW(this.xml, kind);
    rev.setAttributeNS(W, 'w:id', '0');
    rev.setAttributeNS(W, 'w:author', author || 'Author');
    if (date) rev.setAttributeNS(W, 'w:date', date);
    this.revisions.push(rev);
    return rev;
  }

  /**
   * A tracked insertion or deletion of this session (see editor/track.js) as <w:ins> / <w:del>,
   * joined with the one just written by the same author. Deleted text is <w:delText>.
   */
  serializeRevision(el, out, ctx, block, state) {
    const kind = el.dataset.tc;
    const author = el.dataset.tcAuthor || 'Author';
    const prev = state.revision;
    const wrap = prev && prev.kind === kind && prev.author === author && out.lastChild === prev.el
      ? prev.el : this.newRevision(kind, author, el.dataset.tcDate);
    const inner = { last: null };
    for (const c of el.childNodes) this.serializeInline(c, wrap, ctx, block, inner);
    if (kind === 'del') {
      for (const [from, to] of [['t', 'delText'], ['instrText', 'delInstrText']]) {
        for (const t of [...wrap.getElementsByTagNameNS(W, from)]) {
          const d = createW(this.xml, to);
          for (const a of [...t.attributes]) d.setAttributeNS(a.namespaceURI, a.name, a.value);
          d.textContent = t.textContent;
          t.replaceWith(d);
        }
      }
    }
    if (wrap.childNodes.length && wrap.parentNode !== out) out.appendChild(wrap);
    else if (!wrap.childNodes.length) this.revisions.splice(this.revisions.indexOf(wrap), 1);
    state.last = null;
    state.revision = { kind, author, el: wrap };
  }

  /** Rebuilds a content control's <w:sdtContent> from the editor, keeping its properties. */
  serializeControl(el, ctx = {}, block = null) {
    const original = this.raw[Number(el.dataset.sdt)];
    const sdt = original.cloneNode(true);
    const info = controlInfo(original);
    const box = el.querySelector(':scope > .sdt-content') || el;
    const text = box.textContent;
    const before = plainText(wChild(original, 'sdtContent') || original);
    // Placeholder text left alone, or cleared, is written back as it was so Word shows it again.
    if (info.placeholder && (text === before || !text.trim())) return sdt;

    let content = wChild(sdt, 'sdtContent');
    if (!content) content = sdt.appendChild(createW(this.xml, 'sdtContent'));
    while (content.firstChild) content.removeChild(content.firstChild);
    if (el.tagName === 'DIV') {
      const blocks = this.serializeBlocks(box);
      for (const b of blocks) content.appendChild(b);
      if (!blocks.some((b) => b.localName === 'p' || b.localName === 'tbl')) content.appendChild(createW(this.xml, 'p'));
    } else if (text || box.querySelector('img, [data-raw], [data-sdt]')) {
      const state = { last: null };
      for (const c of box.childNodes) this.serializeInline(c, content, ctx, block, state);
    }
    updateControlProperties(sdt, el, {
      placeholderReplaced: info.placeholder,
      isPlaceholderStyle: (id) => this.isPlaceholderStyle(id),
    });
    return sdt;
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
    const key = `${ctx.rpr ?? ''}|${ctx.b}|${ctx.i}|${ctx.u}|${ctx.s}|${ctx.va}|${ctx.hl}`;
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
    // Superscript/subscript and highlight are tags (<sup>, <sub>, <mark>): without one, they are off,
    // but an explicit "baseline" or "none" (overriding the style) is kept.
    const va = wChild(rPr, 'vertAlign');
    if (ctx.va) ensureOrdered(rPr, 'vertAlign', RPR_ORDER).setAttributeNS(W, 'w:val', ctx.va);
    else if (va && wAttr(va, 'val') !== 'baseline') removeW(rPr, 'vertAlign');
    const hl = wChild(rPr, 'highlight');
    if (ctx.hl) ensureOrdered(rPr, 'highlight', RPR_ORDER).setAttributeNS(W, 'w:val', ctx.hl);
    else if (hl && wAttr(hl, 'val') !== 'none') removeW(rPr, 'highlight');
    return isEmptyElement(rPr) ? null : rPr;
  }
}

/**
 * The rows of a table (local 'tr') or cells of a row ('tc') in order, looking through content
 * controls and customXml wrapped around them: [{ el, wrappers: [<w:sdt>…] }].
 */
function tableParts(parent, local, wrappers = []) {
  const out = [];
  for (const c of elementChildren(parent)) {
    if (c.namespaceURI !== W) continue;
    if (c.localName === local) out.push({ el: c, wrappers });
    else if (c.localName === 'sdt') {
      const content = wChild(c, 'sdtContent');
      if (content) out.push(...tableParts(content, local, [...wrappers, c]));
    } else if (c.localName === 'customXml') out.push(...tableParts(c, local, wrappers));
  }
  return out;
}

/** Content controls wrapped around rows or cells of a table, in document order. */
function structuralControls(tbl) {
  const out = [];
  for (const { el: tr, wrappers } of tableParts(tbl, 'tr')) {
    for (const w of wrappers) if (!out.includes(w)) out.push(w);
    for (const { wrappers: cw } of tableParts(tr, 'tc')) for (const w of cw) if (!out.includes(w)) out.push(w);
  }
  return out;
}

/** If run `r` is only a field marker, its kind: begin / separate / end, or instr (with its text). */
function fieldMarker(r) {
  const kids = elementChildren(r).filter((k) => !(k.namespaceURI === W && (k.localName === 'rPr' || k.localName === 'lastRenderedPageBreak')));
  if (kids.length !== 1 || kids[0].namespaceURI !== W) return null;
  const k = kids[0];
  if (k.localName === 'fldChar') return { type: wAttr(k, 'fldCharType') };
  if (k.localName === 'instrText') return { type: 'instr', text: k.textContent };
  return null;
}

/** The run's only content element if it is `local` (ignoring run properties). */
function onlyChild(r, local) {
  const kids = elementChildren(r).filter((k) => !(k.namespaceURI === W && k.localName === 'rPr'));
  return kids.length === 1 && kids[0].namespaceURI === W && kids[0].localName === local ? kids[0] : null;
}

function markCommentRange(el, node) {
  if (node.namespaceURI !== W) return;
  if (node.localName === 'commentRangeStart') el.dataset.commentStart = wAttr(node, 'id');
  if (node.localName === 'commentRangeEnd') el.dataset.commentEnd = wAttr(node, 'id');
}

function romanLower(n) {
  const map = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let s = '';
  for (const [v, r] of map) while (n >= v) { s += r; n -= v; }
  return s;
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
