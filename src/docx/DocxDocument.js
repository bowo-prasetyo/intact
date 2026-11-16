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
import {
  trackedChangeLabel, paragraphMarkLabel, paragraphMarkRevision, inlineRevisionKind, formattingChange, rowRevision,
  cellRevision, propertyChange,
} from './revisions.js';
import { reviewParagraphs, tableSummary } from './review.js';
import {
  pictureInfo, resizePicture, setWrap, setPosition, setDescription, EMU_PER_PX, pictureLook, setCrop, setRotation, setEffects, wrapPolygon, setWrapPolygon,
} from './pictures.js';
import {
  rowsOf, cellsOf, layout, cellPosition, gridWidths, lockedParts, isWrappedCell, insertRows, deleteRows, insertColumns, deleteColumns,
  fixMerges, fitWidths, totalWidth, deleteRowTracked, undeleteRow, newTableXml, mergeArea, areaCells, mergeCells, splitCell,
  markCell, deleteCellTracked, undeleteCell, acceptInsertedCell, removeCells, removeRow, resolveMerge, restoreGrid, dropGridChange,
} from './tables.js';
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
const COMMENTS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const COMMENTS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
const COMMENTS_EX_REL = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
const COMMENTS_EX_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
/** A new paragraph id, as Word writes them (eight hex digits, below 0x80000000). */
const newParaId = () => (Math.floor(Math.random() * 0x7ffffffe) + 1).toString(16).toUpperCase().padStart(8, '0');
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
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
// Pictures that can be inserted: the formats every version of Word shows.
const PICTURE_TYPES = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp' };
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
    this.newMedia = new Map(); // pictures inserted this session: relationship id -> { path, target, data, ext, type }
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
    for (const c of this.comments) c.hadParaId = !!c.paraId; // a reply or resolving may need one added
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
    // Annotations (comments, revisions, bookmarks) share one id space: new ones come from here.
    for (const c of this.comments) if (Number(c.id) > maxId) maxId = Number(c.id);
    this.nextId = maxId + 1;
    this.firstNewId = this.nextId; // ids from here on are this session's
    this.newComments = []; // comments added this session (see addComment)
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
    for (const el of root.querySelectorAll('[data-rpr]')) {
      const change = formattingChange(this.raw[Number(el.dataset.rpr)]);
      if (!change) continue;
      el.dataset.fmt = 'run'; // a tracked formatting change, accepted or rejected on the toolbar
      el.title = `Formatting changed${byWhom(change)} – tracked change: Accept or Reject on the toolbar`;
    }
    for (const el of root.querySelectorAll('[data-ppr]')) {
      const pPr = this.raw[Number(el.dataset.ppr)];
      if (formattingChange(pPr)) el.dataset.fmtPara = '1';
      const label = paragraphMarkLabel(pPr);
      if (label) el.dataset.tracked = label;
      const rev = paragraphMarkRevision(pPr);
      if (rev) {
        el.dataset.revMark = rev.kind; // shown as ¶ (see EditorView), and accepted or rejected there
        el.dataset.revTitle = `${rev.kind === 'ins' ? 'Paragraph break inserted' : 'Paragraph break deleted'}${rev.author ? ` by ${rev.author}` : ''}${rev.date ? ` on ${rev.date.slice(0, 10)}` : ''}`;
      }
    }
  }

  // ---------------------------------------------------------------- adding comments

  /**
   * Adds a comment of this session on the editor range `range` (a point if collapsed): its range
   * markers and its mark go into the text, as Word writes them, and it joins this.comments.
   */
  addComment(range, { author, text, date }) {
    const id = String(this.nextId++);
    const initials = author.split(/\s+/).filter(Boolean).map((w) => w[0]).join('').slice(0, 3).toUpperCase();
    const comment = {
      id, author, initials, date, text, paraId: newParaId(), parentId: null, done: false, isNew: true,
      paragraphs: text.split('\n').map((s) => s.trim()).filter(Boolean),
    };
    this.comments.push(comment);
    this.newComments.push(comment);
    // Like comments that were there: deleting its mark deletes it, and its range stays whole.
    this.commentsInBody.add(id);
    this.commentRangesInBody.add(id);
    const marker = (local) => {
      const el = createW(this.xml, local);
      el.setAttributeNS(W, 'w:id', id);
      return el;
    };
    const run = createW(this.xml, 'r');
    run.appendChild(marker('commentReference'));
    const mark = this.rawInline(run);
    this.commentRefs.set(mark.dataset.raw, [id]);
    const start = range.cloneRange();
    start.collapse(true);
    const end = range.cloneRange();
    end.collapse(false);
    end.insertNode(mark);
    end.insertNode(this.hiddenRaw(marker('commentRangeEnd')));
    start.insertNode(this.hiddenRaw(marker('commentRangeStart')));
    return comment;
  }

  /**
   * Adds a reply of this session to comment `parentId`, as Word writes one: its own range markers
   * inside the parent's range and its mark after the parent's, linked to the parent in
   * commentsExtended.xml on save. Null if the parent's mark isn't in the editor content `root`.
   */
  addReply(parentId, { author, text, date }, root) {
    const q = (attr) => root.querySelector(`[${attr}="${CSS.escape(parentId)}"]`);
    const parentRef = q('data-comment-ref');
    const parent = this.comments.find((c) => c.id === parentId);
    if (!parentRef || !parent) return null;
    parent.paraId ||= newParaId(); // written into comments.xml on save if the reply is kept
    const id = String(this.nextId++);
    const initials = author.split(/\s+/).filter(Boolean).map((w) => w[0]).join('').slice(0, 3).toUpperCase();
    const reply = {
      id, author, initials, date, text, paraId: newParaId(), parentId, done: false, isNew: true,
      paragraphs: text.split('\n').map((s) => s.trim()).filter(Boolean),
    };
    this.comments.push(reply);
    this.newComments.push(reply);
    this.commentsInBody.add(id);
    this.commentRangesInBody.add(id);
    const marker = (local) => {
      const el = createW(this.xml, local);
      el.setAttributeNS(W, 'w:id', id);
      return el;
    };
    const run = createW(this.xml, 'r');
    run.appendChild(marker('commentReference'));
    const mark = this.rawInline(run); // a reply's mark is hidden: replies are shown with their thread
    this.commentRefs.set(mark.dataset.raw, [id]);
    // Inside the parent's range (or, for a comment on a point, just before its mark).
    const start = q('data-comment-start');
    const end = q('data-comment-end');
    const replyStart = this.hiddenRaw(marker('commentRangeStart'));
    const replyEnd = this.hiddenRaw(marker('commentRangeEnd'));
    if (start) start.after(replyStart);
    else parentRef.before(replyStart);
    if (end) end.after(replyEnd);
    else parentRef.before(replyEnd);
    parentRef.after(mark);
    return reply;
  }

  /** A comment's text as the panel shows it: its paragraphs, one per line. */
  commentText(c) {
    return c.text ?? c.paragraphs.join('\n');
  }

  /**
   * Changes to comments kept on their marks in the editor content (so Undo covers them): an edited
   * text (data-edit) and a thread resolved or reopened (data-done). Map id -> { text?, done? }, only
   * where they differ from the comment as it was.
   */
  commentOverrides(root) {
    const out = new Map();
    for (const el of root.querySelectorAll('[data-comment-ref]')) {
      const c = this.comments.find((x) => x.id === el.dataset.commentRef);
      if (!c) continue;
      const o = {};
      if (el.dataset.edit != null && el.dataset.edit !== this.commentText(c)) o.text = el.dataset.edit;
      if (el.dataset.done != null && (el.dataset.done === '1') !== !!c.done) o.done = el.dataset.done === '1';
      if (Object.keys(o).length) out.set(c.id, o);
    }
    return out;
  }

  /** Numbers the comment marks under `root` in document order, as the panel lists them. */
  renumberComments(root) {
    this.commentOrder = new Map();
    for (const el of root.querySelectorAll('.comment-ref')) {
      const id = el.dataset.commentId;
      if (!this.commentOrder.has(id)) this.commentOrder.set(id, this.commentOrder.size + 1);
      el.dataset.n = String(this.commentOrder.get(id));
    }
  }

  // ---------------------------------------------------------------- accepting and rejecting

  /**
   * The tracked change an element shows, or null: { kind: 'ins' | 'del', mark (a paragraph break),
   * own (made this session), author, date }. Inline revisions are spans; a paragraph break is its
   * paragraph.
   */
  revisionOf(el, kind = null) {
    if (el?.nodeType !== 1) return null;
    const d = el.dataset;
    if (kind === 'run' || kind === 'para') {
      const change = formattingChange(this.raw[Number(kind === 'run' ? d.rpr : d.ppr)]);
      return change && { kind: 'format', mark: false, own: false, para: kind === 'para', ...change };
    }
    if (kind === 'tprops' || kind === 'rprops' || kind === 'cprops') {
      const node = kind === 'tprops' ? this.raw[Number(el.dataset.raw)] : kind === 'rprops' ? this.rowXml(el) : this.cellXml(el);
      const change = propertyChange(node);
      return change && { kind: 'format', mark: false, own: false, ...change };
    }
    if (kind === 'merge') {
      const merge = wChild(wChild(this.cellXml(el), 'tcPr'), 'cellMerge');
      return merge && { kind: 'format', mark: false, own: false, author: wAttr(merge, 'author'), date: wAttr(merge, 'date') };
    }
    if (kind === 'cells') {
      const tc = this.cellXml(el);
      const rev = tc && cellRevision(tc);
      return rev && { kind: rev.kind, mark: false, own: false, cells: true, author: rev.author, date: rev.date };
    }
    if (kind === 'row') {
      const rev = this.rowXml(el) && rowRevision(this.rowXml(el));
      return rev && { kind: rev.kind, mark: false, own: false, row: true, author: rev.author, date: rev.date };
    }
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
  resolveRevision(item, accept) {
    const { el } = item;
    if (item.kind === 'run' || item.kind === 'para') return this.resolveFormatting(item, accept);
    if (item.kind === 'row') return this.resolveRow(el, accept);
    if (item.kind === 'cells') return this.resolveCells(el, accept);
    if (item.kind === 'merge') return this.resolveMergeCells(el, accept);
    if (item.kind === 'tprops' || item.kind === 'rprops' || item.kind === 'cprops') return this.resolveTableProps(item, accept);
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

  /**
   * A tracked formatting change: accepted, the record of the old formatting goes; rejected, the old
   * formatting comes back (the run or paragraph is restyled; text formatting applied inside the run
   * since goes with it, as the run's formatting is what is rejected).
   */
  resolveFormatting({ el, kind }, accept) {
    if (kind === 'run') {
      const rPr = this.raw[Number(el.dataset.rpr)];
      const change = wChild(rPr, 'rPrChange');
      if (!change) return;
      let next;
      if (accept) {
        next = rPr.cloneNode(true);
        removeW(next, 'rPrChange');
      } else {
        next = (wChild(change, 'rPr') || createW(this.xml, 'rPr')).cloneNode(true);
      }
      // Restyle: the run's tags (bold, italic…) come from its properties, so they are rebuilt.
      for (const tag of [...el.querySelectorAll('strong, b, em, i, u, s, strike, sup, sub, mark')]) tag.replaceWith(...tag.childNodes);
      const { span, target } = this.runSpan(next);
      target.append(...el.childNodes);
      el.replaceWith(span);
      this.markTracked(span.parentElement || span, false);
      return;
    }
    const pPr = this.raw[Number(el.dataset.ppr)];
    let next;
    if (accept) {
      next = pPr.cloneNode(true);
      removeW(next, 'pPrChange');
      removeW(wChild(next, 'rPr'), 'rPrChange');
    } else {
      const change = wChild(pPr, 'pPrChange');
      next = change ? (wChild(change, 'pPr') || createW(this.xml, 'pPr')).cloneNode(true) : pPr.cloneNode(true);
      // The mark's own properties and section break stay with the paragraph, as in Word.
      const rPr = wChild(pPr, 'rPr')?.cloneNode(true);
      if (rPr) {
        const old = wChild(rPr, 'rPrChange');
        const restored = old ? (wChild(old, 'rPr') || createW(this.xml, 'rPr')).cloneNode(true) : rPr;
        for (const rev of ['ins', 'del', 'moveTo', 'moveFrom']) {
          const mark = wChild(rPr, rev);
          if (mark && restored !== rPr) restored.insertBefore(mark.cloneNode(true), restored.firstChild);
        }
        removeW(next, 'rPr');
        ensureOrdered(next, 'rPr', PPR_ORDER).replaceWith(restored);
      }
      const sectPr = wChild(pPr, 'sectPr');
      if (sectPr && !wChild(next, 'sectPr')) next.appendChild(sectPr.cloneNode(true));
    }
    const fresh = this.paragraphElement(next);
    for (const a of ['tcMark', 'tcMarkAuthor', 'tcMarkDate', 'revMark', 'revTitle', 'tracked']) {
      if (el.dataset[a] != null) fresh.dataset[a] = el.dataset[a];
    }
    fresh.append(...el.childNodes);
    el.replaceWith(fresh);
  }

  /** The XML row (<w:tr>) a rendered table row stands for (rows match in order). */
  rowXml(tr) {
    const table = tr.closest('table');
    if (!table || table.dataset.raw == null) return null;
    const rows = [...table.querySelectorAll(':scope > tbody > tr, :scope > tr')];
    const parts = tableParts(this.raw[Number(table.dataset.raw)], 'tr');
    return parts[rows.indexOf(tr)]?.el || null;
  }

  /**
   * A tracked row: an inserted row accepted (or a deleted one rejected) stays, without the revision;
   * otherwise the row goes. The table gets a new copy of its XML, so Undo stays exact; a table left
   * with no rows goes too.
   */
  resolveRow(tr, accept) {
    const table = tr.closest('table');
    const rows = [...table.querySelectorAll(':scope > tbody > tr, :scope > tr')];
    const index = rows.indexOf(tr);
    const current = tableParts(this.raw[Number(table.dataset.raw)], 'tr')[index]?.el;
    const was = current && rowRevision(current);
    if (!was) return;
    // A deleted row rejected is rendered again: from the table with the edits made in it.
    const tbl = was.kind === 'del' && !accept ? this.tableWithEdits(table) : this.raw[Number(table.dataset.raw)].cloneNode(true);
    const row = tableParts(tbl, 'tr')[index]?.el;
    const rev = row && rowRevision(row);
    if (!rev) return;
    if ((rev.kind === 'ins') === accept) {
      const trPr = wChild(row, 'trPr');
      removeW(trPr, 'ins', 'del');
      if (!elementChildren(trPr).length) row.removeChild(trPr);
      if (rev.kind === 'del') {
        // A deleted row rejected: its text comes back, and its cells can be edited again.
        undeleteRow(row);
        const fresh = this.renderRow(row);
        this.markTracked(fresh, false);
        tr.replaceWith(fresh);
      } else {
        delete tr.dataset.revRow;
        tr.removeAttribute('title');
      }
    } else {
      removeRow(row);
      tr.remove();
      if (!tableParts(tbl, 'tr').length) {
        table.remove();
        return;
      }
    }
    table.dataset.raw = this.keep(tbl);
    const label = trackedChangeLabel(tbl);
    if (label) table.dataset.tracked = label;
    else delete table.dataset.tracked;
  }

  /**
   * How many tracked changes under `root` can only be accepted or rejected in Word: those inside
   * content saved verbatim (locked fields and cells, preserved blocks and items) and tracked cell
   * merges.
   */
  wordOnlyChanges(root) {
    let n = 0;
    for (const el of root.querySelectorAll('[data-raw]')) {
      const node = this.raw[Number(el.dataset.raw)];
      if (node.localName === 'tbl') {
      } else if (!el.classList.contains('inserted') && !el.classList.contains('deleted') && el.dataset.tracked) {
        n++;
      }
    }
    return n;
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

  /**
   * An empty paragraph element for paragraph properties `pPr`: its tag (heading or not), style,
   * formatting as CSS and list membership, as rendered.
   */
  paragraphElement(pPr) {
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
    return el;
  }

  renderParagraph(p) {
    const el = this.paragraphElement(wChild(p, 'pPr'));
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
        img.draggable = false; // a selected floating picture is moved by Intact itself
        const docPr = node.getElementsByTagNameNS(NS.wp, 'docPr')[0];
        img.alt = docPr?.getAttribute('descr') || '';
        if (extent) {
          img.style.width = `${Number(extent.getAttribute('cx')) / 9525}px`;
          img.style.height = `${Number(extent.getAttribute('cy')) / 9525}px`;
        }
        span.append(img);
        break;
      }
      if (span.childNodes.length) {
        const info = local === 'r' && pictureInfo(node);
        if (info) {
          // A picture laid out here: selected by a click, then resized, wrapped and positioned.
          span.classList.add('picture');
          span.title = 'Picture – click to select it: drag a corner to resize, or use Layout Options';
          this.stylePicture(span, info);
        } else {
          span.title = `${describe(node)} – kept as it is; edit it in Word`;
        }
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
      const run = this.runSpan(rPr);
      parent.append(run.span);
      target = run.target;
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

  /**
   * The span for a run with properties `rPr`: its formatting as CSS, with bold, italic, underline,
   * strikethrough, superscript, subscript and highlight as tags inside it. Returns the span and the
   * innermost element, which holds the text.
   */
  runSpan(rPr) {
    const span = document.createElement('span');
    span.dataset.rpr = this.keep(rPr);
    const rStyle = wAttr(wChild(rPr, 'rStyle'), 'val');
    if (rStyle) span.dataset.rstyle = rStyle;
    Object.assign(span.style, cssToStyle(this.styles.rPrCss(rPr, false)));
    // Explicitly switched-off toggles (e.g. <w:b w:val="0"/>) are kept as inline CSS.
    if (onOff(wChild(rPr, 'b')) === false) span.style.fontWeight = 'normal';
    if (onOff(wChild(rPr, 'i')) === false) span.style.fontStyle = 'normal';
    let target = span;
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
    return { span, target };
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
    for (const { el: tr, wrappers: rowWrappers } of tableParts(tbl, 'tr')) tbody.append(this.renderRow(tr, rowWrappers));
    this.applyTableProps(table, tbl);
    return table;
  }

  /** A table row (`rowWrappers`: content controls around it), its cells editable unless locked. */
  renderRow(tr, rowWrappers = []) {
    const row = document.createElement('tr');
    if (rowWrappers.length) {
      row.classList.add('sdt-row');
      row.title = controlTitle(controlInfo(rowWrappers[rowWrappers.length - 1]));
    }
    for (const { el: tc, wrappers: cellWrappers } of tableParts(tr, 'tc')) row.append(this.renderCell(tc, rowWrappers, cellWrappers));
    this.applyRowProps(row, tr, rowWrappers);
    return row;
  }

  /** A tracked change to a table's properties (or its grid), shown on the table. */
  applyTableProps(table, tbl) {
    let change = propertyChange(tbl);
    // A grid change made with tracked cells (a split) is part of that change, resolved with it.
    const grid = wChild(wChild(tbl, 'tblGrid'), 'tblGridChange');
    if (change && grid && !wChild(wChild(tbl, 'tblPr'), 'tblPrChange')) {
      const key = `${wAttr(grid, 'author') || ''}|${wAttr(grid, 'date') || ''}`;
      if ([...tbl.getElementsByTagNameNS(W, 'cellIns')].some((c) => `${wAttr(c, 'author') || ''}|${wAttr(c, 'date') || ''}` === key)) change = null;
    }
    if (change) {
      table.dataset.fmtTable = '1';
      table.title = `Table formatting changed${byWhom(change)} – tracked change: Accept or Reject on the toolbar`;
    } else if (table.dataset.fmtTable) {
      delete table.dataset.fmtTable;
      table.removeAttribute('title');
    }
  }

  /** A row's tracked insertion or deletion, shown on the row. */
  applyRowProps(row, tr, rowWrappers = []) {
    const rev = rowRevision(tr);
    if (rev) {
      row.dataset.revRow = rev.kind; // a tracked row, accepted or rejected on the toolbar
      row.title = `Row ${rev.kind === 'ins' ? 'inserted' : 'deleted'}${byWhom(rev)} – tracked change: Accept or Reject on the toolbar`;
    } else {
      delete row.dataset.revRow;
      if (rowWrappers.length) row.title = controlTitle(controlInfo(rowWrappers[rowWrappers.length - 1]));
      else row.removeAttribute('title');
    }
    const change = propertyChange(tr);
    if (change) {
      row.dataset.fmtRow = '1'; // a tracked change to the row's properties (height, header repeat…)
      if (!rev) row.title = `Row formatting changed${byWhom(change)} – tracked change: Accept or Reject on the toolbar`;
    } else {
      delete row.dataset.fmtRow;
    }
  }

  /** A table cell (`rowWrappers`, `cellWrappers`: content controls around its row or itself). */
  renderCell(tc, rowWrappers = [], cellWrappers = []) {
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
    this.applyCellProps(td, tc);
    this.renderBlocks(tc, td);
    if (!td.childNodes.length) td.append(this.emptyParagraph());
    return td;
  }

  /** What a cell's properties show: columns spanned, merged with the cell above, width, shading. */
  applyCellProps(td, tc) {
    const tcPr = wChild(tc, 'tcPr');
    const span = Number(wAttr(wChild(tcPr, 'gridSpan'), 'val') || 1);
    if (span > 1) td.colSpan = span;
    else td.removeAttribute('colspan');
    const vMerge = wChild(tcPr, 'vMerge');
    td.classList.toggle('vmerge', !!vMerge && wAttr(vMerge, 'val') !== 'restart');
    const w = wChild(tcPr, 'tcW');
    td.style.width = w && wAttr(w, 'type') === 'dxa' ? `${Number(wAttr(w, 'w')) / 15}px` : '';
    const shd = wAttr(wChild(tcPr, 'shd'), 'fill');
    td.style.backgroundColor = shd && /^[0-9a-f]{6}$/i.test(shd) ? '#' + shd : '';
    // A cell inserted or deleted with tracking (a tracked column): cells of one change share a key.
    const rev = cellRevision(tc);
    if (rev) {
      td.dataset.revCell = rev.kind;
      td.dataset.revKey = `${rev.kind}|${rev.author || ''}|${rev.date || ''}`;
      td.title = `Cell ${rev.kind === 'ins' ? 'inserted' : 'deleted'}${byWhom(rev)} – tracked change: Accept or Reject on the toolbar`;
    } else if (td.dataset.revCell) {
      delete td.dataset.revCell;
      delete td.dataset.revKey;
      td.removeAttribute('title');
    }
    const merge = wChild(tcPr, 'cellMerge');
    if (merge) {
      // A tracked merge or split down a column: its cells share a key (author and time).
      td.dataset.revMerge = wAttr(merge, 'vMerge') === 'cont' ? 'merge' : 'split';
      td.dataset.revMergeKey = `${wAttr(merge, 'author') || ''}|${wAttr(merge, 'date') || ''}`;
      if (!rev) td.title = `Cells ${td.dataset.revMerge === 'merge' ? 'merged' : 'split'}${byWhom({ author: wAttr(merge, 'author'), date: wAttr(merge, 'date') })} – tracked change: Accept or Reject on the toolbar`;
    } else if (td.dataset.revMerge) {
      delete td.dataset.revMerge;
      delete td.dataset.revMergeKey;
      if (!rev) td.removeAttribute('title');
    }
    const change = propertyChange(tc);
    if (change) {
      td.dataset.fmtCell = '1'; // a tracked change to the cell's properties (width, shading…)
      if (!rev) td.title = `Cell formatting changed${byWhom(change)} – tracked change: Accept or Reject on the toolbar`;
    } else if (td.dataset.fmtCell) {
      delete td.dataset.fmtCell;
      if (!rev) td.removeAttribute('title');
    }
  }

  /**
   * A tracked change to a table's, row's or cell's properties: accepted, its record goes; rejected,
   * the old properties come back (a row's or cell's own tracked insertion or deletion stays).
   */
  resolveTableProps({ el, kind }, accept) {
    const table = kind === 'tprops' ? el : el.closest('table');
    const rows = this.tableRows(table);
    const r = kind === 'tprops' ? -1 : rows.indexOf(kind === 'rprops' ? el : el.parentElement);
    const c = kind === 'cprops' ? [...el.parentElement.cells].indexOf(el) : -1;
    const restore = (props, changeLocal, innerLocal, keep = []) => {
      const change = wChild(props, changeLocal);
      if (!change) return;
      props.removeChild(change);
      if (accept) return;
      const kept = keep.map((local) => wChild(props, local)).filter(Boolean);
      const old = wChild(change, innerLocal);
      while (props.firstChild) props.removeChild(props.firstChild);
      for (const n of [...(old?.childNodes || [])]) props.appendChild(n);
      for (const n of kept) props.appendChild(n);
    };
    this.editTable(table, (tbl) => {
      if (kind === 'tprops') {
        restore(wChild(tbl, 'tblPr'), 'tblPrChange', 'tblPr');
        restore(wChild(tbl, 'tblGrid'), 'tblGridChange', 'tblGrid');
        return;
      }
      const tr = rowsOf(tbl)[r];
      if (kind === 'rprops') restore(wChild(tr, 'trPr'), 'trPrChange', 'trPr', ['ins', 'del']);
      else restore(wChild(cellsOf(tr)[c], 'tcPr'), 'tcPrChange', 'tcPr', ['cellIns', 'cellDel', 'cellMerge']);
    });
  }

  /** The XML cell (<w:tc>) a rendered cell stands for (rows and cells match in order). */
  cellXml(td) {
    const tr = this.rowXml(td.parentElement);
    return tr ? cellsOf(tr)[[...td.parentElement.cells].indexOf(td)] || null : null;
  }

  /** The rendered cells of the same tracked cell change as `td` (one tracked column), in order. */
  cellGroup(td) {
    const table = td.closest('table');
    return this.tableRows(table).flatMap((r) => [...r.cells]).filter((c) => c.dataset.revKey && c.dataset.revKey === td.dataset.revKey);
  }

  /**
   * A tracked cell change (a column inserted or deleted with tracking), all its cells at once: an
   * insertion accepted or a deletion rejected keeps the cells (a deleted cell's text comes back);
   * otherwise the cells go, with their grid columns.
   */
  resolveCells(td, accept) {
    const table = td.closest('table');
    const group = this.cellGroup(td);
    const kind = td.dataset.revCell;
    const rows = this.tableRows(table);
    const at = group.map((c) => [rows.indexOf(c.parentElement), [...c.parentElement.cells].indexOf(c)]);
    const pick = (tbl) => at.map(([r, i]) => cellsOf(rowsOf(tbl)[r])[i]).filter(Boolean);
    const key = td.dataset.revKey.split('|').slice(1).join('|'); // author|date
    if ((kind === 'ins') === accept) {
      if (kind === 'ins') {
        this.editTable(table, (tbl) => {
          pick(tbl).forEach(acceptInsertedCell);
          dropGridChange(tbl, key);
        });
      }
      else {
        this.editTable(table, (tbl) => pick(tbl).forEach((tc) => {
          const copy = tc.cloneNode(true); // rendered again, its text no longer deleted
          undeleteCell(copy);
          tc.replaceWith(copy);
        }), { withEdits: true, keepInsertions: true });
      }
    } else {
      this.editTable(table, (tbl) => {
        removeCells(tbl, pick(tbl));
        if (kind === 'ins') restoreGrid(tbl, key); // cells split off: the widths they took come back
        fitWidths(tbl, { before: Infinity, textWidth: 0 });
      });
    }
  }

  // ---------------------------------------------------------------- inserting pictures

  /**
   * A picture from a file (PNG, JPEG, GIF or BMP) as Word inserts one: in line with the text, at its
   * own size or narrowed to fit between the margins. Its file goes into word/media when saved (if
   * it is still in the document). Returns { span } (the rendered picture) or { reason }.
   */
  async addPicture(file) {
    const ext = PICTURE_TYPES[file.type];
    if (!ext) return { reason: `${file.name ? `“${file.name}”` : 'This file'} isn’t a picture Word shows everywhere: insert a PNG, JPEG, GIF or BMP picture` };
    const data = new Uint8Array(await file.arrayBuffer());
    const url = URL.createObjectURL(new Blob([data], { type: file.type }));
    const size = await new Promise((resolve) => {
      const img = new Image();
      img.onload = () => resolve(img.naturalWidth ? { w: img.naturalWidth, h: img.naturalHeight } : null);
      img.onerror = () => resolve(null);
      img.src = url;
    });
    if (!size) {
      URL.revokeObjectURL(url);
      return { reason: `${file.name ? `“${file.name}”` : 'This picture'} couldn’t be read` };
    }
    const maxWidth = this.textWidth / 15;
    let { w, h } = size;
    if (w > maxWidth) {
      h = Math.max(1, Math.round((h * maxWidth) / w));
      w = maxWidth;
    }
    let n = 1;
    const taken = (path) => this.zip.file(path) || [...this.newMedia.values()].some((m) => m.path === path);
    while (taken(`${this.mainDir ? this.mainDir + '/' : ''}media/image${n}.${ext}`)) n++;
    const target = `media/image${n}.${ext}`;
    const path = `${this.mainDir ? this.mainDir + '/' : ''}${target}`;
    let k = 1;
    while (this.rels.has(`rIdImage${k}`)) k++;
    const rid = `rIdImage${k}`;
    this.rels.set(rid, { type: IMAGE_REL, target, external: false, path });
    this.mediaUrls.set(path, url);
    this.newMedia.set(rid, { path, target, data, ext, type: file.type });
    const id = await this.nextDocPrId();
    const cx = Math.round(w * 9525);
    const cy = Math.round(h * 9525);
    const name = (file.name || `image${n}.${ext}`).replace(/[<>&"]/g, '');
    const xml = `<w:r xmlns:w="${W}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${PIC}" xmlns:r="${NS.r}"><w:drawing>`
      + `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
      + `<wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>`
      + `<a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>`
      + `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
      + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
      + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const r = this.xml.importNode(parseXml(xml).documentElement, true);
    for (const a of [...r.attributes]) if (a.name.startsWith('xmlns')) r.removeAttributeNode(a); // declared where needed when saved
    return { span: this.rawInline(r) };
  }

  /** A drawing id no other drawing in the package has (Word wants them unique). */
  async nextDocPrId() {
    if (this.docPrId == null) {
      let max = 0;
      for (const f of Object.values(this.zip.files)) {
        if (f.dir || !/\.xml$/.test(f.name)) continue;
        const text = await f.async('string');
        for (const m of text.matchAll(/<(?:\w+:)?docPr\b[^>]*?\bid="(\d+)"/g)) max = Math.max(max, Number(m[1]));
      }
      this.docPrId = max;
    }
    return ++this.docPrId;
  }

  /**
   * Writes the pictures inserted this session that are in the document (file, relationship, content
   * type), and takes out any written by an earlier save that no longer are.
   */
  async writeMedia() {
    if (!this.newMedia.size) return;
    const used = new Set([...this.body.getElementsByTagNameNS(NS.a, 'blip')].map((b) => b.getAttributeNS(NS.r, 'embed')));
    const rels = parseXml((await this.text(this.relsPath)) || `<Relationships xmlns="${NS.pr}"/>`);
    const ct = parseXml(await this.text('[Content_Types].xml'));
    let changed = false;
    for (const [rid, m] of this.newMedia) {
      const rel = [...rels.documentElement.children].find((r) => r.getAttribute('Id') === rid);
      if (used.has(rid)) {
        if (rel) continue;
        this.zip.file(m.path, m.data);
        const el = rels.createElementNS(NS.pr, 'Relationship');
        el.setAttribute('Id', rid);
        el.setAttribute('Type', IMAGE_REL);
        el.setAttribute('Target', m.target);
        rels.documentElement.appendChild(el);
        const defaults = [...ct.documentElement.children].filter((c) => c.localName === 'Default');
        if (!defaults.some((d) => d.getAttribute('Extension')?.toLowerCase() === m.ext)) {
          const d = ct.createElementNS(NS.ct, 'Default');
          d.setAttribute('Extension', m.ext);
          d.setAttribute('ContentType', m.type);
          ct.documentElement.insertBefore(d, ct.documentElement.firstElementChild);
        }
        changed = true;
      } else if (rel) {
        this.zip.remove(m.path);
        rel.remove();
        changed = true;
      }
    }
    if (!changed) return;
    this.zip.file(this.relsPath, serializeXml(rels));
    this.zip.file('[Content_Types].xml', serializeXml(ct));
  }

  // ---------------------------------------------------------------- picture layout

  /** The layout of a rendered picture (a span.picture): see pictures.js pictureInfo. */
  pictureLayout(span) {
    return pictureInfo(this.raw[Number(span.dataset.raw)]);
  }

  /**
   * Shows a picture as it is laid out, as near as a web page can: in line; floating left or right
   * (or centred) with the text around it; on a line of its own; or behind or in front of the text at
   * its position. Offsets from the paragraph are measured from where the picture's run is.
   */
  stylePicture(span, info = this.pictureLayout(span), look = pictureLook(this.raw[Number(span.dataset.raw)])) {
    const img = span.querySelector('img');
    const w = info.cx / EMU_PER_PX;
    const h = info.cy / EMU_PER_PX;
    span.dataset.wrap = info.wrap;
    span.removeAttribute('style');
    const st = span.style;
    this.placePicture(span, info, w, h);
    // What shows of it: its frame, the picture cropped inside (as Word's srcRect cuts it).
    st.width = `${w}px`;
    st.height = `${h}px`;
    st.overflow = 'hidden';
    const c = look.crop;
    const fullW = w / (1 - c.l - c.r);
    const fullH = h / (1 - c.t - c.b);
    if (img) {
      img.alt = info.descr;
      Object.assign(img.style, {
        width: `${fullW}px`, height: `${fullH}px`, maxWidth: 'none', marginLeft: `${-c.l * fullW}px`, marginTop: `${-c.t * fullH}px`,
        filter: [look.greyscale && 'grayscale(1)', look.brightness && `brightness(${1 + look.brightness})`, look.contrast && `contrast(${1 + look.contrast})`].filter(Boolean).join(' '),
      });
    }
    // Rotation and flips, border, shadow, glow, soft edges and reflection: as near as CSS shows them.
    const turns = [look.rot && `rotate(${look.rot}deg)`, look.flipH && 'scaleX(-1)', look.flipV && 'scaleY(-1)'].filter(Boolean).join(' ');
    if (turns) st.transform = turns;
    span.dataset.turn = turns;
    if (look.border) st.outline = `${Math.max(1, look.border.width * 1.333)}px solid #${look.border.color}`;
    const drops = [look.shadow && 'drop-shadow(3px 3px 3px rgb(0 0 0 / 40%))', look.glow && `drop-shadow(0 0 5px #${look.glow})`].filter(Boolean);
    if (drops.length) st.filter = drops.join(' ');
    if (look.softEdge) {
      const fade = 'linear-gradient(to right, transparent, #000 8px, #000 calc(100% - 8px), transparent), linear-gradient(to bottom, transparent, #000 8px, #000 calc(100% - 8px), transparent)';
      st.maskImage = fade;
      st.webkitMaskImage = fade;
      st.maskComposite = 'intersect';
      st.webkitMaskComposite = 'source-in';
    }
    if (look.reflection) st.webkitBoxReflect = 'below 2px linear-gradient(transparent 55%, rgb(255 255 255 / 45%))';
  }

  /** Where a picture goes: in line, floating left or right, on a line of its own, or behind or in front of the text. */
  placePicture(span, info, w, h) {
    if (info.wrap === 'inline') return;
    const st = span.style;
    const page = this.page;
    const content = page.width - page.left - page.right;
    const px = (emu) => (emu || 0) / EMU_PER_PX;
    // Horizontal place within the text column: from the left margin.
    let left = 0;
    const { h: hp, v: vp } = info;
    if (hp.align === 'center') left = (content - w) / 2;
    else if (hp.align === 'right' || hp.align === 'outside') left = content - w;
    else if (hp.align == null) left = px(hp.offset) - (hp.relativeFrom === 'page' ? page.left : 0);
    const fromParagraph = !vp.relativeFrom || ['paragraph', 'line'].includes(vp.relativeFrom);
    const top = vp.align == null ? px(vp.offset) : 0;
    if (info.wrap === 'behind' || info.wrap === 'front') {
      st.position = 'absolute';
      st.left = `${page.left + left}px`;
      st.zIndex = info.wrap === 'behind' ? '-1' : '2';
      if (fromParagraph) st.marginTop = `${top}px`;
      else {
        const area = vp.relativeFrom === 'page' ? page.minHeight : page.minHeight - page.top - page.bottom;
        const base = vp.relativeFrom === 'page' ? 0 : page.top;
        st.top = `${base + (vp.align === 'center' ? (area - h) / 2 : vp.align === 'bottom' ? area - h : vp.align === 'top' ? 0 : top)}px`;
      }
      return;
    }
    if (fromParagraph) st.marginTop = `${top}px`;
    if (info.wrap === 'topAndBottom') {
      st.display = 'block';
      st.marginLeft = hp.align === 'center' || hp.align === 'right' ? 'auto' : `${Math.max(0, left)}px`;
      if (hp.align === 'center') st.marginRight = 'auto';
      return;
    }
    // Square, tight and through: the text flows around it; tight and through, around its outline.
    const outline = (info.wrap === 'tight' || info.wrap === 'through') && wrapPolygon(this.raw[Number(span.dataset.raw)]);
    if (outline?.length > 2) {
      st.shapeOutside = `polygon(${outline.map(([x, y]) => `${(x / 216).toFixed(2)}% ${(y / 216).toFixed(2)}%`).join(', ')}) content-box`;
      st.shapeMargin = '8px';
    }
    if (hp.align === 'right' || (hp.align == null && left + w / 2 > content / 2 + 1)) {
      st.float = 'right';
      st.marginLeft = '10px';
      if (hp.align == null) st.marginRight = `${Math.max(0, content - left - w)}px`;
    } else {
      st.float = 'left';
      st.marginRight = '10px';
      st.marginLeft = `${Math.max(0, left)}px`;
    }
  }

  /**
   * Changes a picture's layout: `change(run)` changes a copy of its run (see pictures.js), which
   * replaces it (so Undo stays exact), and the picture is shown again.
   */
  changePicture(span, change) {
    const run = this.raw[Number(span.dataset.raw)].cloneNode(true);
    change(run);
    span.dataset.raw = this.keep(run);
    this.stylePicture(span, pictureInfo(run), pictureLook(run));
  }

  setPictureSize(span, cx, cy) {
    this.changePicture(span, (run) => resizePicture(run, cx, cy));
  }

  /** Changes a picture's text wrapping; for Tight and Through, `outline` (see editor pictureOutline) is what the text wraps to. */
  setPictureWrap(span, wrap, outline = null) {
    this.zOrder = (this.zOrder || 251659264) + 1024;
    this.changePicture(span, (run) => {
      setWrap(run, wrap, this.zOrder);
      if (outline) setWrapPolygon(run, outline);
    });
  }

  setPicturePosition(span, position) {
    this.changePicture(span, (run) => setPosition(run, position));
  }

  /** A picture's crop, rotation and effects (see pictures.js pictureLook). */
  pictureLookOf(span) {
    return pictureLook(this.raw[Number(span.dataset.raw)]);
  }

  setPictureCrop(span, crop) {
    this.changePicture(span, (run) => setCrop(run, crop));
  }

  setPictureRotation(span, rotation) {
    this.changePicture(span, (run) => setRotation(run, rotation));
  }

  setPictureEffects(span, look) {
    this.changePicture(span, (run) => setEffects(run, look));
  }

  setPictureDescription(span, text) {
    this.changePicture(span, (run) => setDescription(run, text));
  }

  // ---------------------------------------------------------------- table structure

  /** The rendered rows of a table, in order (they stand for its <w:tr> elements, in order). */
  tableRows(table) {
    return [...table.querySelectorAll(':scope > tbody > tr, :scope > tr')];
  }

  /** Why a table's structure can't be changed here, or null. */
  tableLocked(table) {
    if (table?.dataset.raw == null) return 'This table can’t be changed here: change it in Word';
    return lockedParts(this.raw[Number(table.dataset.raw)]);
  }

  /** Text width between the page margins, in twips. */
  get textWidth() {
    const p = this.page;
    return Math.round((p.width - p.left - p.right) * 15);
  }

  /** The grid columns and row of a rendered cell: { row, start, span }. */
  cellPosition(td) {
    const table = td.closest('table');
    const row = this.tableRows(table).indexOf(td.parentElement);
    const index = [...td.parentElement.cells].indexOf(td);
    return cellPosition(this.raw[Number(table.dataset.raw)], row, index);
  }

  /**
   * Changes a table's structure: `edit(tbl)` changes a copy of its XML (which then replaces it, so
   * Undo stays exact), and the rendered table follows: rows and cells still there are kept as they
   * are (with any edits in them), new ones are rendered. Cells' edits are kept in the copy first
   * when `withEdits` (for rows rebuilt from XML). Returns false if no rows are left: the table goes.
   */
  editTable(table, edit, { withEdits = false, keepInsertions = false } = {}) {
    const tbl = withEdits ? this.tableWithEdits(table, !keepInsertions) : this.raw[Number(table.dataset.raw)].cloneNode(true);
    const domRows = this.tableRows(table);
    const rowOf = new Map();
    const cellOf = new Map();
    rowsOf(tbl).forEach((tr, i) => {
      if (!domRows[i]) return;
      rowOf.set(tr, domRows[i]);
      cellsOf(tr).forEach((tc, j) => { if (domRows[i].cells[j]) cellOf.set(tc, domRows[i].cells[j]); });
    });
    edit(tbl);
    fixMerges(tbl);
    const rows = rowsOf(tbl);
    if (!rows.length || !gridWidths(tbl).length) {
      table.remove();
      return false;
    }
    // Fields around rows and cells, for rendering new rows and cells as fields too.
    const rowWrappers = new Map(tableParts(tbl, 'tr').map(({ el, wrappers }) => [el, wrappers]));
    const rendered = rows.map((tr) => {
      const row = rowOf.get(tr);
      const around = rowWrappers.get(tr) || [];
      if (!row) {
        const fresh = this.renderRow(tr, around);
        this.markTracked(fresh, false);
        return fresh;
      }
      const cellWrappers = new Map(tableParts(tr, 'tc').map(({ el, wrappers }) => [el, wrappers]));
      row.replaceChildren(...cellsOf(tr).map((tc) => {
        const td = cellOf.get(tc);
        if (!td) {
          const fresh = this.renderCell(tc, around, cellWrappers.get(tc) || []);
          this.markTracked(fresh, false);
          return fresh;
        }
        this.applyCellProps(td, tc);
        return td;
      }));
      this.applyRowProps(row, tr, around);
      return row;
    });
    const tblStyle = wAttr(wChild(wChild(tbl, 'tblPr'), 'tblStyle'), 'val');
    if (tblStyle) table.dataset.tstyle = tblStyle;
    else delete table.dataset.tstyle;
    this.applyTableProps(table, tbl);
    (table.tBodies[0] || table).replaceChildren(...rendered);
    table.dataset.raw = this.keep(tbl);
    const label = trackedChangeLabel(tbl);
    if (label) table.dataset.tracked = label;
    else delete table.dataset.tracked;
    return true;
  }

  /**
   * A copy of a table's XML with the edits made in its cells. This session's tracked changes in it
   * get their ids now; with `dropInsertions` (its rows are being deleted) its insertions go, with
   * their content, as deleting your own insertion removes it in Word.
   */
  tableWithEdits(table, dropInsertions = false) {
    this.revisions = [];
    const tbl = this.serializeTable(table);
    for (const rev of this.revisions) {
      if (dropInsertions && rev.localName === 'ins' && rev.parentNode?.localName !== 'rPr') rev.parentNode?.removeChild(rev);
      else rev.setAttributeNS(W, 'w:id', String(this.nextId++));
    }
    this.revisions = [];
    return tbl;
  }

  /** A revision of this session for a table row, with its id now: <w:ins> or <w:del>. */
  rowRevisionMaker(tracked) {
    return (kind = 'ins') => {
      const rev = createW(this.xml, kind);
      rev.setAttributeNS(W, 'w:id', String(this.nextId++));
      rev.setAttributeNS(W, 'w:author', tracked.author || 'Author');
      if (tracked.date) rev.setAttributeNS(W, 'w:date', tracked.date);
      return rev;
    };
  }

  /**
   * Inserts `count` rows above or below rendered row `index` (copies of it, empty), tracked while
   * tracking (`tracked`: author and date). Returns { rows } (the new rendered rows) or { reason }.
   */
  insertTableRows(table, index, count, below, tracked = null) {
    const reason = this.tableLocked(table);
    if (reason) return { reason };
    let made = [];
    const rev = tracked && this.rowRevisionMaker(tracked);
    this.editTable(table, (tbl) => { made = insertRows(tbl, index, count, below, rev); });
    const rows = this.tableRows(table);
    const first = below ? index + 1 : index;
    return { rows: rows.slice(first, first + made.length) };
  }

  /** Tab in the last cell: a row like the last one, below it (see insertTableRows). */
  addTableRow(table, tracked = null) {
    const rows = this.tableRows(table);
    if (!rows.length) return { row: null, reason: 'This table has no rows to copy' };
    const res = this.insertTableRows(table, rows.length - 1, 1, true, tracked);
    if (res.reason) return { row: null, reason: res.reason.replace('change its layout', 'add rows') };
    return { row: res.rows[0], reason: null };
  }

  /**
   * Deletes rendered rows `from` to `to`. While tracking they are marked deleted (rows inserted
   * this session go); otherwise they go, and so does a table left with no rows. Returns
   * { removed } (the table went) or { reason }.
   */
  deleteTableRows(table, from, to, tracked = null) {
    const reason = this.tableLocked(table);
    if (reason) return { reason };
    const domRows = this.tableRows(table).slice(from, to + 1);
    if (!tracked) {
      const kept = domRows.find((r) => r.dataset.revRow || r.querySelector('[data-tracked], [data-rev-mark], [data-fmt], [data-fmt-para], td[data-rev-cell]'));
      if (kept) return { reason: 'The rows hold tracked changes: accept or reject them first, or turn on Track changes' };
      return { removed: !this.editTable(table, (tbl) => deleteRows(tbl, range(from, to))) };
    }
    if (domRows.some((r) => r.dataset.revRow && !this.ownRow(table, r))) {
      return { reason: 'A row is already a tracked change: accept or reject it first' };
    }
    const rev = this.rowRevisionMaker(tracked);
    const removed = !this.editTable(table, (tbl) => {
      const rows = rowsOf(tbl);
      const own = [];
      for (let i = from; i <= to; i++) {
        const tr = rows[i];
        if (!tr) continue;
        if (this.isOwnRow(tr)) { own.push(i); continue; }
        const copy = tr.cloneNode(true); // a new element, so the row is rendered again
        deleteRowTracked(copy, rev);
        tr.replaceWith(copy);
      }
      deleteRows(tbl, own);
    }, { withEdits: true });
    return { removed };
  }

  /** True for a row inserted with tracking in this session. */
  isOwnRow(tr) {
    const ins = wChild(wChild(tr, 'trPr'), 'ins');
    return !!ins && Number(wAttr(ins, 'id')) >= this.firstNewId;
  }

  ownRow(table, row) {
    const tr = this.rowXml(row);
    return !!tr && this.isOwnRow(tr);
  }

  /**
   * Inserts `count` columns left or right of the grid columns `start`..`start + span - 1` (a cell's),
   * as wide as the column next to them; a table that fitted the page narrows to fit. Column changes
   * aren't tracked. Returns { reason } or {}.
   */
  insertTableColumns(table, start, span, count, right, tracked = null) {
    const reason = this.tableLocked(table);
    if (reason) return { reason };
    const at = right ? start + span : start;
    const like = right ? start + span - 1 : start;
    const rev = tracked && this.rowRevisionMaker(tracked);
    this.editTable(table, (tbl) => {
      const before = totalWidth(tbl);
      const made = insertColumns(tbl, at, count, like);
      if (rev) for (const tc of made) markCell(tc, rev('cellIns')); // a tracked column: its cells inserted
      fitWidths(tbl, { before, textWidth: this.textWidth });
    });
    return {};
  }

  /**
   * Deletes grid columns `from`..`to`. While tracking, the cells wholly in them are marked deleted
   * (<w:cellDel>, their text in <w:del>), as Word tracks a deleted column; cells spanning into them
   * narrow when the deletion is accepted. A column you inserted while tracking just goes.
   * Returns { removed } (the table went) or { reason }.
   */
  deleteTableColumns(table, from, to, tracked = null) {
    const reason = this.tableLocked(table);
    if (reason) return { reason };
    const tbl = this.raw[Number(table.dataset.raw)];
    if (tracked) {
      const targets = this.tableRows(table).flatMap((row, i) => {
        const cells = layout(rowsOf(tbl)[i]);
        return [...row.cells].map((td, j) => ({ td, i, j, c: cells[j] }))
          .filter(({ c }) => c && c.start >= from && c.start + c.span - 1 <= to);
      });
      if (!targets.length) return { reason: 'No cell lies wholly in these columns: nothing to delete' };
      const own = (t) => t.td.dataset.revCell === 'ins' && Number(cellRevision(t.c.tc)?.id) >= this.firstNewId;
      if (targets.every(own)) return this.deleteTableColumns(table, from, to); // your own inserted column
      if (targets.some((t) => t.td.dataset.revCell)) return { reason: 'These columns hold a tracked column change: accept or reject it first' };
      const rev = this.rowRevisionMaker(tracked);
      this.editTable(table, (t) => {
        for (const { i, j } of targets) {
          const tc = cellsOf(rowsOf(t)[i])[j];
          const copy = tc.cloneNode(true); // rendered again, its text marked deleted
          deleteCellTracked(copy, rev);
          tc.replaceWith(copy);
        }
      }, { withEdits: true });
      return {};
    }
    const held = this.tableRows(table).some((row, i) => {
      if (row.dataset.revRow) return false;
      const cells = layout(rowsOf(tbl)[i]);
      // Cells wholly in the columns go (the others only narrow): none may hold someone's tracked
      // change (a column you inserted while tracking may go).
      return [...row.cells].some((td, j) => cells[j] && cells[j].start >= from && cells[j].start + cells[j].span - 1 <= to
        && (td.querySelector('[data-tracked], [data-rev-mark], [data-fmt], [data-fmt-para]')
          || (td.dataset.revCell && !(td.dataset.revCell === 'ins' && Number(cellRevision(cells[j].tc)?.id) >= this.firstNewId))));
    });
    if (held) return { reason: 'Cells in these columns hold tracked changes: accept or reject them first' };
    return { removed: !this.editTable(table, (tbl2) => {
      deleteColumns(tbl2, from, to);
      fitWidths(tbl2, { before: Infinity, textWidth: 0 });
    }) };
  }

  /**
   * Merges the cells of rendered rows `r0`..`r1` and grid columns `c0`..`c1` (grown to whole cells,
   * as Word does) into one: their content goes into the first, one cell after another (empty cells
   * add nothing). Returns { cell } (the merged cell) or { reason }.
   */
  mergeTableCells(table, [r0, r1], [c0, c1], tracked = null) {
    const reason = this.tableLocked(table);
    if (reason) return { reason };
    const tbl = this.raw[Number(table.dataset.raw)];
    const area = mergeArea(tbl, r0, r1, c0, c1);
    if (!area) return { reason: 'These cells can’t be merged: select a rectangle of cells' };
    const cells = areaCells(tbl, area);
    if (cells.flat().length < 2) return { reason: 'Select the cells to merge: more than one' };
    const domRows = this.tableRows(table);
    const tds = cells.map((row, i) => row.map((c) => domRows[area.r0 + i].cells[c.index]));
    if (cells.flat().some((c) => isWrappedCell(c.tc))) return { reason: 'Cells that are fields can’t be merged: merge them in Word' };
    if (tds.flat().some((td) => !td || td.dataset.revCell || td.dataset.revMerge) || domRows.slice(area.r0, area.r1 + 1).some((r) => r.dataset.revRow)) {
      return { reason: 'These cells are a tracked change: accept or reject it first' };
    }
    const down = area.r1 > area.r0 && !!tracked; // merging down the rows: tracked while tracking
    if (down && tds.slice(1).flat().some((td) => td.querySelector('[data-tracked], [data-tc], [data-tc-mark], [data-rev-mark]'))) {
      return { reason: 'These cells hold tracked changes: accept or reject them first' };
    }
    const target = tds[0][0];
    const empty = (td) => !td.textContent.trim() && !td.querySelector('img, .raw-inline, .sdt, table');
    // Across each row (never tracked: Word can't record it), the cells' content joins the row's first cell.
    tds.forEach((row) => {
      for (const td of row.slice(1)) {
        if (!empty(td)) {
          if (empty(row[0])) row[0].replaceChildren();
          row[0].append(...td.childNodes);
        }
        td.replaceChildren(this.emptyParagraph());
      }
    });
    // Down the rows: moved into the merged cell, or, tracked, copied in as an insertion and left
    // as a deletion where it was (so rejecting the merge puts it back).
    for (const row of tds.slice(1)) {
      const td = row[0];
      if (empty(td)) continue;
      if (!down) {
        if (empty(target)) target.replaceChildren();
        target.append(...td.childNodes);
        td.replaceChildren(this.emptyParagraph());
        continue;
      }
      for (const block of [...td.children]) {
        const copy = block.cloneNode(true);
        copy.dataset.tcMark = 'ins';
        copy.dataset.tcMarkAuthor = tracked.author;
        copy.dataset.tcMarkDate = tracked.date;
        copy.replaceChildren(this.trackedSpan('ins', tracked, [...copy.childNodes]));
        target.append(copy);
        block.replaceChildren(this.trackedSpan('del', tracked, [...block.childNodes]));
      }
    }
    const rev = down && this.rowRevisionMaker(tracked);
    this.editTable(table, (t) => {
      mergeCells(t, area, rev);
      fitWidths(t, { before: Infinity, textWidth: 0 });
    });
    return { cell: target, across: area.c1 > area.c0, down: area.r1 > area.r0 };
  }

  /** A tracked insertion or deletion of this session around `nodes`, as editor/track.js makes them. */
  trackedSpan(kind, { author, date }, nodes) {
    const span = document.createElement('span');
    span.className = kind === 'ins' ? 'tc-ins' : 'tc-del';
    span.dataset.tc = kind;
    span.dataset.tcAuthor = author;
    span.dataset.tcDate = date;
    span.title = `${kind === 'ins' ? 'Inserted' : 'Deleted'} by ${author} – tracked change: Accept or Reject on the toolbar`;
    if (kind === 'del') span.contentEditable = 'false';
    span.append(...nodes.filter((n) => !(n.nodeType === 1 && n.tagName === 'BR' && !n.nextSibling)));
    return span;
  }

  /** Splits a rendered cell into `cols` × `rows` cells (see tables.js splitCell), tracked while tracking. Returns {} or { reason }. */
  splitTableCell(table, td, cols, rows, tracked = null) {
    const reason = this.tableLocked(table);
    if (reason) return { reason };
    if (td.dataset.revCell || td.dataset.revMerge || td.parentElement.dataset.revRow) return { reason: 'This cell is a tracked change: accept or reject it first' };
    const rowIndex = this.tableRows(table).indexOf(td.parentElement);
    const cellIndex = [...td.parentElement.cells].indexOf(td);
    const trial = splitCell(this.raw[Number(table.dataset.raw)].cloneNode(true), rowIndex, cellIndex, cols, rows);
    if (trial.reason) return trial;
    const rev = tracked && this.rowRevisionMaker(tracked);
    this.editTable(table, (t) => {
      splitCell(t, rowIndex, cellIndex, cols, rows, rev);
      fitWidths(t, { before: Infinity, textWidth: 0 });
    });
    return {};
  }

  /**
   * A tracked merge (or split) of cells down a column, all its cells at once, with the text the
   * merge moved: accepted, the merge stays and the moved text is only in the merged cell; rejected,
   * the cells are as they were, their text back in them.
   */
  resolveMergeCells(td, accept) {
    const table = td.closest('table');
    const key = td.dataset.revMergeKey;
    const group = this.tableRows(table).flatMap((r) => [...r.cells]).filter((c) => c.dataset.revMergeKey === key);
    const [author, date] = key.split('|');
    for (const cell of group) {
      // This session's moved text: its insertions and deletions of the same author and time.
      for (const span of [...cell.querySelectorAll('[data-tc]')].filter((sp) => sp.dataset.tcAuthor === author && sp.dataset.tcDate === date)) {
        if ((span.dataset.tc === 'ins') === accept) span.replaceWith(...span.childNodes);
        else span.remove();
      }
      for (const block of [...cell.querySelectorAll('[data-tc-mark]')].filter((b) => b.dataset.tcMarkAuthor === author && b.dataset.tcMarkDate === date)) {
        if (accept) {
          delete block.dataset.tcMark;
          delete block.dataset.tcMarkAuthor;
          delete block.dataset.tcMarkDate;
        } else block.remove();
      }
      // Someone's (from the file): their revisions of the same author and time, one by one.
      for (const el of [...cell.querySelectorAll('.raw-inline.inserted, .raw-inline.deleted')]) {
        const rev = this.revisionOf(el);
        if (rev && rev.author === author && (rev.date || '') === date) this.resolveRevision({ el, kind: 'inline' }, accept);
      }
      if (!cell.firstElementChild) cell.append(this.emptyParagraph());
    }
    const rows = this.tableRows(table);
    const at = group.map((c) => [rows.indexOf(c.parentElement), [...c.parentElement.cells].indexOf(c)]);
    this.editTable(table, (tbl) => at.forEach(([r, i]) => {
      const tc = cellsOf(rowsOf(tbl)[r])[i];
      if (tc) resolveMerge(tc, accept);
    }));
  }

  /** Deletes a whole table: its rows marked deleted while tracking, else the table goes. */
  deleteTable(table, tracked = null) {
    const rows = this.tableRows(table);
    if (!tracked) {
      const reason = this.tableLocked(table);
      if (reason && !reason.startsWith('This table’s rows')) return { reason };
      if (table.querySelector('[data-tracked], [data-rev-mark], [data-fmt], [data-fmt-para], [data-rev-row], [data-rev-cell]')) {
        return { reason: 'The table holds tracked changes: accept or reject them first, or turn on Track changes' };
      }
      table.remove();
      return { removed: true };
    }
    return this.deleteTableRows(table, 0, rows.length - 1, tracked);
  }

  /**
   * A new table of `rows` × `cols` empty cells across the page, in the Table Grid style (added to
   * the file if need be); while tracking, its rows are tracked insertions. Returns the rendered table.
   */
  newTable(rows, cols, tracked = null) {
    const styleId = this.styles.ensureTableGrid();
    const tbl = newTableXml(this.xml, rows, cols, this.textWidth, styleId, tracked && this.rowRevisionMaker(tracked));
    const table = this.renderTable(tbl);
    const label = trackedChangeLabel(tbl);
    if (label) table.dataset.tracked = label;
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
    this.revisions.forEach((rev, i) => rev.setAttributeNS(W, 'w:id', String(this.nextId + i)));
    while (this.body.firstChild) this.body.removeChild(this.body.firstChild);
    for (const n of nodes) this.body.appendChild(n);
    if (!nodes.some((n) => n.localName === 'p')) this.body.appendChild(createW(this.xml, 'p'));
    if (sectPr) this.body.appendChild(sectPr);
    this.commentsRemoved = removedComments(this.comments, {
      wasInBody: this.commentsInBody, present: commentIdsIn(this.body), elsewhere: this.commentsElsewhere,
    });
    pruneCommentMarkers(this.body, this.commentsRemoved, this.commentRangesInBody);
    this.commentEdits = this.commentOverrides(root);
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
    // Rows inserted or deleted with tracking this session: their paragraphs are tracked too.
    for (const tr of this.body.getElementsByTagNameNS(W, 'tr')) {
      const trPr = wChild(tr, 'trPr');
      const rev = wChild(trPr, 'ins') || wChild(trPr, 'del');
      if (rev && Number(wAttr(rev, 'id')) >= this.firstNewId) for (const p of tr.getElementsByTagNameNS(W, 'p')) trackedParagraphs.add(p);
    }
    let cellsTracked = 0;
    for (const tc of this.body.getElementsByTagNameNS(W, 'tc')) {
      const rev = cellRevision(tc);
      if (!rev || Number(rev.id) < this.firstNewId) continue;
      cellsTracked++;
      for (const p of tc.getElementsByTagNameNS(W, 'p')) trackedParagraphs.add(p);
    }
    const styleName = (id) => {
      const name = (id ? this.styles.byId.get(id)?.name || id : this.styles.byId.get(this.styles.defaultParagraph)?.name) || 'Normal';
      return name[0].toUpperCase() + name.slice(1); // built-in names are stored as "heading 1"
    };
    return {
      paragraphs: reviewParagraphs(this.originalBody, this.body, { tracked: (p) => trackedParagraphs.has(p), styleName }),
      commentsRemoved: this.comments.filter((c) => this.commentsRemoved.has(c.id) && !c.isNew),
      commentsAdded: this.newComments.filter((c) => commentIdsIn(this.body).has(c.id) && !this.commentsRemoved.has(c.id))
        .map((c) => ({ ...c, text: this.commentEdits.get(c.id)?.text ?? c.text, parentAuthor: this.comments.find((p) => p.id === c.parentId)?.author })),
      ...(() => {
        const changed = [...this.commentEdits].filter(([id]) => !this.commentsRemoved.has(id)).map(([id, e]) => ({ c: this.comments.find((x) => x.id === id), e }));
        return {
          commentsEdited: changed.filter(({ c, e }) => e.text != null && !c.isNew).map(({ c }) => c),
          commentsResolved: changed.filter(({ e }) => e.done === true).map(({ c }) => c),
          commentsReopened: changed.filter(({ e }) => e.done === false).map(({ c }) => c),
        };
      })(),
      insertions: this.revisions.filter((r) => r.localName === 'ins').length,
      deletions: this.revisions.filter((r) => r.localName === 'del').length,
      tracking: this.tracking,
      trackingSwitched: this.tracking !== this.trackRevisions,
      stylesAdded: this.styles.changed,
      layout: tableSummary(this.originalBody, this.body),
      cellsTracked,
      rowsTracked: [...this.body.getElementsByTagNameNS(W, 'trPr')].reduce((n, trPr) => {
        const rev = wChild(trPr, 'ins') || wChild(trPr, 'del');
        return n + (rev && Number(wAttr(rev, 'id')) >= this.firstNewId ? 1 : 0);
      }, 0),
      listsAdded: this.numberingChanged,
    };
  }

  /** Produces a .docx Blob from the current editor content. */
  async save(root) {
    this.update(root);
    this.zip.file(this.mainPath, serializeXml(this.xml));
    await this.writeMedia();
    await this.writeStyles();
    if (this.numberingChanged) await this.writeNumbering();
    await this.writeComments();
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
  /**
   * Writes the comment parts when comments changed: removed ones taken out (see commentRemoval.js),
   * edited ones rewritten, new comments and replies appended, and resolved state and reply links in
   * commentsExtended.xml. A part is created if the document has none; untouched, they stay as read.
   */
  async writeComments() {
    const present = commentIdsIn(this.body);
    const removed = this.commentsRemoved;
    const added = this.newComments.filter((c) => present.has(c.id) && !removed.has(c.id));
    const edits = new Map([...this.commentEdits].filter(([id]) => present.has(id) && !removed.has(id)));
    if (!removed.size && !added.length && !edits.size && !this.commentPartsChanged) return;
    const out = pruneCommentParts(this.commentParts, this.comments, removed);
    const byId = (id) => this.comments.find((c) => c.id === id);
    const replies = added.filter((c) => c.parentId);
    const doneChanges = [...edits].filter(([, e]) => e.done != null);
    // Comments that commentsExtended.xml will name: they need a paragraph id, added if they had none.
    const named = new Set([...replies.map((r) => r.parentId), ...doneChanges.map(([id]) => id)]);
    for (const id of named) {
      const c = byId(id);
      if (c) c.paraId ||= newParaId();
    }
    const touched = this.comments.filter((c) => !c.isNew && present.has(c.id) && !removed.has(c.id)
      && (edits.get(c.id)?.text != null || (named.has(c.id) && !c.hadParaId)));
    if (added.length || touched.length) {
      if (!this.commentParts.comments) {
        const empty = `${XML_HEADER}<w:comments xmlns:w="${W}" xmlns:w14="${W14}"/>`;
        const path = await this.addPart('comments.xml', empty, COMMENTS_REL, COMMENTS_CT);
        this.commentParts.comments = { path, text: empty };
      }
      const { path, text } = this.commentParts.comments;
      const doc = parseXml(out[path] ?? text);
      for (const c of touched) {
        const el = [...doc.getElementsByTagNameNS(W, 'comment')].find((x) => wAttr(x, 'id') === c.id);
        if (!el) continue;
        const edit = edits.get(c.id)?.text;
        if (edit != null) setCommentText(doc, el, edit);
        const paras = wChildren(el, 'p');
        if (c.paraId && paras.length) paras[paras.length - 1].setAttributeNS(W14, 'w14:paraId', c.paraId);
      }
      for (const c of added) doc.documentElement.appendChild(commentXml(doc, { ...c, text: edits.get(c.id)?.text ?? c.text }));
      out[path] = serializeXml(doc);
    }
    if (replies.length || doneChanges.length) {
      if (!this.commentParts.extended) {
        const empty = `${XML_HEADER}<w15:commentsEx xmlns:w15="${W15}"/>`;
        const path = await this.addPart('commentsExtended.xml', empty, COMMENTS_EX_REL, COMMENTS_EX_CT);
        this.commentParts.extended = { path, text: empty };
      }
      const { path, text } = this.commentParts.extended;
      const doc = parseXml(out[path] ?? text);
      const entry = (paraId) => {
        let e = [...doc.getElementsByTagNameNS(W15, 'commentEx')].find((x) => x.getAttributeNS(W15, 'paraId') === paraId);
        if (!e) {
          e = doc.createElementNS(W15, 'w15:commentEx');
          e.setAttributeNS(W15, 'w15:paraId', paraId);
          e.setAttributeNS(W15, 'w15:done', '0');
          doc.documentElement.appendChild(e);
        }
        return e;
      };
      for (const r of replies) {
        const parent = byId(r.parentId);
        entry(parent.paraId);
        entry(r.paraId).setAttributeNS(W15, 'w15:paraIdParent', parent.paraId);
      }
      for (const [id, e] of doneChanges) entry(byId(id).paraId).setAttributeNS(W15, 'w15:done', e.done ? '1' : '0');
      out[path] = serializeXml(doc);
    }
    for (const [path, xml] of Object.entries(out)) this.zip.file(path, xml);
    this.commentPartsChanged = removed.size > 0 || added.length > 0 || edits.size > 0 || touched.length > 0;
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

/** The numbers from `a` to `b`. */
function range(a, b) {
  return Array.from({ length: b - a + 1 }, (_, i) => a + i);
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

/** A comment of this session as <w:comment>, its text one paragraph per line, as Word writes it. */
function commentXml(doc, c) {
  const el = createW(doc, 'comment');
  el.setAttributeNS(W, 'w:id', c.id);
  el.setAttributeNS(W, 'w:author', c.author);
  el.setAttributeNS(W, 'w:date', c.date);
  if (c.initials) el.setAttributeNS(W, 'w:initials', c.initials);
  const lines = c.text.split('\n');
  lines.forEach((line, i) => {
    const p = el.appendChild(createW(doc, 'p'));
    if (i === 0) p.appendChild(createW(doc, 'r')).appendChild(createW(doc, 'annotationRef'));
    if (line) {
      const t = p.appendChild(createW(doc, 'r')).appendChild(createW(doc, 't'));
      t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
      t.textContent = line;
    }
  });
  const paras = wChildren(el, 'p');
  if (c.paraId) paras[paras.length - 1].setAttributeNS(W14, 'w14:paraId', c.paraId);
  return el;
}

/**
 * Replaces a comment's text (in comments.xml): one paragraph per line, keeping its first
 * paragraph's properties and its reference mark. (Formatting within the old text goes.)
 */
function setCommentText(doc, comment, text) {
  const paras = wChildren(comment, 'p');
  const first = paras[0];
  const pPr = wChild(first, 'pPr');
  const mark = !!first?.getElementsByTagNameNS(W, 'annotationRef').length;
  for (const p of paras) comment.removeChild(p);
  text.split('\n').forEach((line, i) => {
    const p = comment.appendChild(createW(doc, 'p'));
    if (i === 0 && pPr) p.appendChild(pPr.cloneNode(true));
    if (i === 0 && mark) p.appendChild(createW(doc, 'r')).appendChild(createW(doc, 'annotationRef'));
    if (line) {
      const t = p.appendChild(createW(doc, 'r')).appendChild(createW(doc, 't'));
      t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
      t.textContent = line;
    }
  });
}

/** " by Ann on 2026-01-05", from a revision's author and date. */
function byWhom({ author, date }) {
  return `${author ? ` by ${author}` : ''}${date ? ` on ${date.slice(0, 10)}` : ''}`;
}
