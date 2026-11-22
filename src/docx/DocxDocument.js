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
import { layoutTabs } from '../editor/tabLayout.js';
import { ommlToLatex } from '../editor/mathOutput.js';
import {
  NS, parseXml, serializeXml, wChild, wChildren, wAttr, onOff, createW, ensureOrdered, removeW,
  RPR_ORDER, PPR_ORDER, isEmptyElement, plainText, parseRels, elementChildren, XML_HEADER,
} from './xml.js';
import { StyleSheet, jcToCss, borderCss } from './styles.js';
import { textboxInfo, setTextboxText, textboxContents, syncCopies } from './textboxes.js';
import { loadEmbeddedFonts } from './fonts.js';
import { Numbering, addList, addPresetList, addHeadingList, MAX_LEVEL } from './numbering.js';
import { blankDocxFiles } from './blank.js';
import { STARTERS, starterFiles } from './starters.js';
import {
  loadPart, parseComments, parseNotes, headerFooterVariants, labelVariants, fieldTitle, describe,
} from './annotations.js';
import {
  controlInfo, controlTitle, controlChoices, controlDate, formatWordDate, updateControlProperties,
} from './contentControls.js';
import {
  trackedChangeLabel, paragraphMarkLabel, paragraphMarkRevision, inlineRevisionKind, formattingChange, rowRevision,
  cellRevision, propertyChange,
} from './revisions.js';
import { reviewParagraphs, tableSummary } from './review.js';
import { resolveRevisionsIn, countRevisions, revisionAuthors } from './resolveAll.js';
import {
  pictureInfo, resizePicture, setWrap, setPosition, setDescription, EMU_PER_PX, pictureLook, setCrop, setRotation, setEffects, wrapPolygon, setWrapPolygon,
  setGeometry, set3D, applyPictureStyle, setArtistic, clearArtistic, blipOf, shapeOf, newShapeRun,
} from './pictures.js';
import { drawPicture, drawShape } from './pictureView.js';
import { ommlToMathml, isEquation, MATH_NS } from './omml.js';
import {
  bodyXml, isSourcesPart, readSources, sourcesXml, sourcesPropsXml, newGuid, newTag,
  CUSTOM_XML_REL, CUSTOM_XML_PROPS_REL, CUSTOM_XML_PROPS_CT,
} from './references.js';
import { fieldReview, referenceManager } from './fields.js';
import { relsPathOf } from './links.js';
import { formFieldInfo, formFieldTitle, setFormCheck, setFormDropdown } from './formFields.js';
import { sniffFile, readProtection, readWriteProtection, readSignatures, removeSignatures } from './package.js';
import { PasswordError, decrypt, encrypt, keyForPassword } from './encryption.js';
import {
  propertyPaths, readProperties, personalParts, anonymize, updateCore, readCore, setCore, setCustom, managedProperty, EDITABLE_CORE,
  EMPTY_CORE, CUSTOM_PROPS_REL, CUSTOM_PROPS_CT, CORE_PROPS_REL, CORE_PROPS_CT,
} from './properties.js';
import {
  rowsOf, cellsOf, layout, cellPosition, gridWidths, lockedParts, insertRows, deleteRows, insertColumns, deleteColumns,
  fixMerges, fitWidths, totalWidth, deleteRowTracked, undeleteRow, newTableXml, mergeArea, areaCells, mergeCells, splitCell,
  markCell, deleteCellTracked, undeleteCell, acceptInsertedCell, removeCells, removeRow, resolveMerge, restoreGrid, dropGridChange,
  hMerge, isTrackedHMerge, resolveCellProps, normalizeHMerge, formatCells, setShading, setBorders, cellBorders,
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
const FOOTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
const FOOTER_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
const HEADER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const HEADER_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const FOOTNOTES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes';
const FOOTNOTES_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml';
const ENDNOTES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes';
const ENDNOTES_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml';
// The final section's properties Page setup sets: kept as opened, and written afresh at each save.
const PAGE_SETUP_PARTS = ['type', 'pgSz', 'pgMar', 'pgNumType', 'cols'];
// The order of a section's properties (CT_SectPr), for adding page size and margins.
const SECTPR_ORDER = [
  'headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc',
  'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg',
  'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange',
];
const SETTINGS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';
const SETTINGS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';
// The settings that come before <w:trackRevisions> in CT_Settings.
// The settings after w:updateFields in CT_Settings, in order.
const AFTER_UPDATE_FIELDS = new Set(['hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat', 'docVars', 'rsids', 'attachedSchema',
  'themeFontLang', 'clrSchemeMapping', 'doNotIncludeSubdocsInStats', 'doNotAutoCompressPictures', 'forceUpgrade', 'captions',
  'readModeInkLockDown', 'smartTagType', 'shapeDefaults', 'doNotEmbedSmartTags', 'decimalSymbol', 'listSeparator']);
// The settings after evenAndOddHeaders in CT_Settings, to put it before them.
const AFTER_EVEN_AND_ODD = new Set(['bookFoldRevPrinting', 'bookFoldPrinting', 'bookFoldPrintingSheets', 'drawingGridHorizontalSpacing',
  'drawingGridVerticalSpacing', 'displayHorizontalDrawingGridEvery', 'displayVerticalDrawingGridEvery', 'doNotUseMarginsForDrawingGridOrigin',
  'drawingGridHorizontalOrigin', 'drawingGridVerticalOrigin', 'doNotShadeFormData', 'noPunctuationKerning', 'characterSpacingControl',
  'printTwoOnOne', 'strictFirstAndLastChars', 'noLineBreaksAfter', 'noLineBreaksBefore', 'savePreviewPicture', 'doNotValidateAgainstSchema',
  'saveInvalidXml', 'ignoreMixedContent', 'alwaysShowPlaceholderText', 'doNotDemarcateInvalidXml', 'saveXmlDataOnly', 'useXSLTWhenSaving',
  'saveThroughXslt', 'showXMLTags', 'alwaysMergeEmptyNamespace', 'updateFields', ...AFTER_UPDATE_FIELDS]);
const BEFORE_TRACK_REVISIONS = new Set([
  'writeProtection', 'view', 'zoom', 'removePersonalInformation', 'removeDateAndTime', 'doNotDisplayPageBoundaries',
  'displayBackgroundShape', 'printPostScriptOverText', 'printFractionalCharacterWidth', 'printFormsData',
  'embedTrueTypeFonts', 'embedSystemFonts', 'saveSubsetFonts', 'saveFormsData', 'mirrorMargins', 'alignBordersAndEdges',
  'bordersDoNotSurroundHeader', 'bordersDoNotSurroundFooter', 'gutterAtTop', 'hideSpellingErrors', 'hideGrammaticalErrors',
  'activeWritingStyle', 'proofState', 'formsDesign', 'attachedTemplate', 'linkStyles', 'stylePaneFormatFilter',
  'stylePaneSortMethod', 'documentType', 'mailMerge', 'revisionView',
]);
export const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
/** The type of each kind of Word file, as it is downloaded. */
const FORMAT_MIME = {
  docx: DOCX_MIME,
  dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
  docm: 'application/vnd.ms-word.document.macroEnabled.12',
  dotm: 'application/vnd.ms-word.template.macroEnabled.12',
};

/** The content type of the main part of each kind of Word file, in [Content_Types].xml. */
const MAIN_PART_TYPE = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
  docm: 'application/vnd.ms-word.document.macroEnabled.main+xml',
  dotm: 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml',
};
/** The kind of document a template makes, as Word makes one when a template is opened from Explorer. */
const DOCUMENT_OF = { dotx: 'docx', dotm: 'docm' };

// What a run holds that is drawn on its own: pictures, shapes and text boxes, legacy VML, objects.
const RUN_OBJECTS = new Set(['drawing', 'AlternateContent', 'pict', 'object']);
/**
 * A run's content (rPr aside) split so each object in it stands alone, the rest grouped as it comes:
 * one piece (no split) unless the run holds an object beside other content.
 */
function runPieces(kids, isSimple) {
  const content = kids.filter((k) => !(k.namespaceURI === W && k.localName === 'rPr'));
  if (!content.some((k) => RUN_OBJECTS.has(k.localName))) return [content];
  const pieces = [];
  for (const k of content) {
    const last = pieces[pieces.length - 1];
    if (RUN_OBJECTS.has(k.localName) || !last || RUN_OBJECTS.has(last[0].localName) || isSimple(k) !== isSimple(last[0])) pieces.push([k]);
    else last.push(k);
  }
  return pieces;
}
const SIMPLE_RUN_CHILDREN = new Set(['rPr', 't', 'tab', 'br', 'cr', 'noBreakHyphen', 'softHyphen', 'lastRenderedPageBreak']);
const HIDDEN_INLINE = new Set(['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'permStart', 'permEnd', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']);
const FLATTEN_INLINE = new Set(['smartTag', 'customXml', 'dir', 'bdo']);
const NOTE_KINDS = ['footnote', 'endnote'];
const NOTE_NAMES = { footnote: 'Footnote', endnote: 'Endnote' };
const TRACKED_INLINE = { ins: 'Inserted', moveTo: 'Moved here', del: 'Deleted', moveFrom: 'Moved away' };
const BLOCK_TAGS = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'DIV', 'TABLE', 'UL', 'OL', 'LI', 'BLOCKQUOTE', 'PRE', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER']);
const HEADING_TAG = { title: 'h1', h1: 'h1', h2: 'h2', h3: 'h3', h4: 'h4', h5: 'h5', h6: 'h6' };
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const HYPERLINK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
// Pictures that can be inserted: the formats every version of Word shows.
const PICTURE_TYPES = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/svg+xml': 'svg' };
// Office 2016's extension for an SVG picture (with a PNG for older readers), as Word writes it.
const SVG_EXT = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';
const ASVG = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';
/**
 * A PNG of the SVG picture `img` (loaded) at twice `size` ({ w, h } in pixels), or null if the
 * browser can't draw it. Drawn from the file in an image, which runs none of its scripts.
 */
async function svgToPng(img, { w, h }) {
  try {
    const scale = Math.min(2, 4096 / Math.max(w, h));
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(w * scale));
    canvas.height = Math.max(1, Math.round(h * scale));
    canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
    return blob ? new Uint8Array(await blob.arrayBuffer()) : null;
  } catch {
    return null;
  }
}

/** `el` alone in a new <w:pPr>, for comparing one property with sameProps. */
function wrapIn(xml, el) {
  const w = createW(xml, 'pPr');
  w.appendChild(el.cloneNode(true));
  return w;
}

// Word's Line and Page Breaks that the Paragraph dialog sets (on/off paragraph properties).
const BREAK_PROPS = ['keepNext', 'keepLines', 'pageBreakBefore'];
const IMAGE_TYPES = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff' };

/**
 * Writes paper size, orientation and margins into section properties `sect`: `set` has any of w,
 * h, orient, top, right, bottom, left (twips); what it doesn't name is kept.
 */
function applySetup(sect, set) {
  if (!sect) return;
  const put = (el, attr, v) => { if (v !== undefined) el.setAttributeNS(W, `w:${attr}`, String(Math.round(v))); };
  if (['w', 'h', 'orient'].some((k) => set[k] !== undefined)) {
    const sz = ensureOrdered(sect, 'pgSz', SECTPR_ORDER);
    put(sz, 'w', set.w);
    put(sz, 'h', set.h);
    if (set.orient === 'landscape') sz.setAttributeNS(W, 'w:orient', 'landscape');
    else if (set.orient === 'portrait') sz.removeAttributeNS(W, 'orient');
  }
  if (['top', 'right', 'bottom', 'left'].some((k) => set[k] !== undefined)) {
    let mar = wChild(sect, 'pgMar');
    if (!mar) {
      // Word requires all of them: its defaults for what isn't set.
      mar = ensureOrdered(sect, 'pgMar', SECTPR_ORDER);
      for (const [a, v] of [['top', 1440], ['right', 1440], ['bottom', 1440], ['left', 1440], ['header', 708], ['footer', 708], ['gutter', 0]]) put(mar, a, v);
    }
    for (const k of ['top', 'right', 'bottom', 'left']) put(mar, k, set[k]);
  }
  // How the section starts (w:type): on a new page, or continuous (on the same page, as a section
  // break made for columns is). Written as Word writes the kind of break.
  if (set.type !== undefined) {
    removeW(sect, 'type');
    if (set.type) ensureOrdered(sect, 'type', SECTPR_ORDER).setAttributeNS(W, 'w:val', set.type);
  }
  // Columns, as Word's Columns presets write them: a number of equal columns (w:num), the space
  // between them (w:space) and a line between them (w:sep). Columns of their own widths (w:col)
  // give way to equal ones; one column is w:cols with its space alone, as Word writes it.
  if (set.cols !== undefined || set.colSpace !== undefined || set.colSep !== undefined) {
    const cols = ensureOrdered(sect, 'cols', SECTPR_ORDER);
    const count = Math.max(1, Math.round(set.cols ?? columnCount(cols)));
    const space = set.colSpace ?? (Number(wAttr(cols, 'space')) || 720);
    while (cols.firstChild) cols.removeChild(cols.firstChild);
    for (const a of ['num', 'space', 'equalWidth', 'sep']) cols.removeAttributeNS(W, a);
    // In the order Word writes them.
    if (count > 1) cols.setAttributeNS(W, 'w:num', String(count));
    put(cols, 'space', space);
    if (count > 1 && set.colSep) cols.setAttributeNS(W, 'w:sep', '1');
  }
  // Page numbering, as Word's Format Page Numbers: the number format (w:fmt; Arabic numerals,
  // decimal, are Word's default and written as no format) and where it starts (w:start; none
  // continues from the section before). The element's other attributes (chapter numbers) are kept.
  if (set.numFmt !== undefined || set.numStart !== undefined) {
    const num = ensureOrdered(sect, 'pgNumType', SECTPR_ORDER);
    if (set.numFmt !== undefined) {
      if (!set.numFmt || set.numFmt === 'decimal') num.removeAttributeNS(W, 'fmt');
      else num.setAttributeNS(W, 'w:fmt', set.numFmt);
    }
    if (set.numStart !== undefined) {
      if (set.numStart == null) num.removeAttributeNS(W, 'start');
      else num.setAttributeNS(W, 'w:start', String(Math.max(0, Math.round(set.numStart))));
    }
    if (!num.attributes.length && !num.childNodes.length) sect.removeChild(num);
  }
}

/** A section's page numbering (w:pgNumType): { numFmt ('decimal', 'lowerRoman'…), numStart (a number, or null to continue) }. */
function numberingOf(pgNumType) {
  const start = wAttr(pgNumType, 'start');
  return { numFmt: wAttr(pgNumType, 'fmt') || 'decimal', numStart: start != null && /^\d+$/.test(start) ? Number(start) : null };
}

/** How many columns a section's w:cols sets: its number, or its own columns (w:col); 1 without it. */
function columnCount(cols) {
  return cols ? Number(wAttr(cols, 'num')) || wChildren(cols, 'col').length || 1 : 1;
}

/**
 * A section's columns and start, for Page setup: { cols (how many), colSpace (twips between them),
 * colSep (a line between them), colUnequal (columns of their own widths), type (how it starts) }.
 */
function columnsSetupOf(sectPr) {
  const cols = wChild(sectPr, 'cols');
  const space = wAttr(cols, 'space');
  return {
    cols: columnCount(cols),
    colSpace: space != null && /^\d+$/.test(space) ? Number(space) : 720,
    colSep: ['1', 'true', 'on'].includes(wAttr(cols, 'sep') || ''),
    colUnequal: ['0', 'false', 'off'].includes(wAttr(cols, 'equalWidth') || '') && wChildren(cols, 'col').length > 1,
    type: wAttr(wChild(sectPr, 'type'), 'val') || 'nextPage',
  };
}

/** "Header", "First page header", "Even pages footer": a header or footer host's name. */
function hostLabel(which, label) {
  const name = !label || label === 'All pages' ? which : `${label} ${which}`;
  return name[0].toUpperCase() + name.slice(1);
}

/** A different first page (w:titlePg) on or off in section properties `sect`. */
function setTitlePage(sect, on) {
  removeW(sect, 'titlePg');
  if (on) ensureOrdered(sect, 'titlePg', SECTPR_ORDER);
}

// The namespaces of Word's VML, for a watermark (declared on a header made here).
const VML_NS = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"';

/**
 * A run that is Word's text watermark (a VML text path whose shape Word names
 * PowerPlusWaterMarkObject…): { text, diagonal, color, font }, or null.
 */
export function watermarkInfo(run) {
  const shape = [...run.getElementsByTagNameNS(NS.v, 'shape')].find((s) => /^PowerPlusWaterMarkObject/.test(s.getAttribute('id') || ''));
  const path = shape?.getElementsByTagNameNS(NS.v, 'textpath')[0];
  if (!path) return null;
  const style = shape.getAttribute('style') || '';
  const rot = Number(/rotation:\s*(-?[\d.]+)/.exec(style)?.[1] || 0);
  const color = shape.getAttribute('fillcolor') || 'silver';
  const font = /font-family:\s*"?([^";]+)/.exec(path.getAttribute('style') || '')?.[1] || 'Calibri';
  return { text: path.getAttribute('string') || '', diagonal: Math.abs(((rot % 360) + 360) % 360 - 315) < 1 || Math.abs(rot + 45) < 1, color, font };
}

// Fonts Word offers that come with Windows and macOS (with Office), offered after a document's own.
const STANDARD_FONTS = [
  'Aptos', 'Arial', 'Arial Black', 'Arial Narrow', 'Book Antiqua', 'Bookman Old Style', 'Calibri', 'Calibri Light', 'Cambria',
  'Candara', 'Century Gothic', 'Comic Sans MS', 'Consolas', 'Constantia', 'Corbel', 'Courier New', 'Franklin Gothic Medium',
  'Garamond', 'Georgia', 'Gill Sans MT', 'Lucida Console', 'Lucida Sans Unicode', 'Palatino Linotype', 'Segoe UI',
  'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana',
];

export class DocxDocument {
  /** @param {ArrayBuffer|Blob} data */
  /**
   * Opens a .docx (its bytes). One encrypted with a password is decrypted here with `password`, on
   * this device (see encryption.js); without one, or with a wrong one, a PasswordError says so.
   */
  static async open(data, name = 'document.docx', { password = null } = {}) {
    let bytes = new Uint8Array(data instanceof Blob ? await data.arrayBuffer() : data);
    // Encrypted, rights-managed and legacy files aren't ZIP packages: say what they are instead.
    const unreadable = sniffFile(bytes);
    let encryption = null;
    if (unreadable?.kind === 'password') {
      if (password == null) throw new PasswordError('needed', `${name} is protected with a password. Enter it to open the document: it is checked on this device and never sent anywhere.`);
      const decrypted = await decrypt(bytes, password);
      // A document encrypted the way Office 2007 did is encrypted again as Word does now, with the same password.
      encryption = { key: decrypted.key || await keyForPassword(password), source: bytes, changed: !decrypted.key };
      bytes = decrypted.zip;
      const inner = sniffFile(bytes);
      if (inner) throw inner;
    } else if (unreadable) throw unreadable;
    const zip = await JSZip.loadAsync(bytes);
    const doc = new DocxDocument(zip, name);
    doc.sourceBytes = bytes; // what was opened: downloaded as it is when a signed file is unchanged
    // { key (to encrypt it again), source (the encrypted file opened), changed (the password since) }, or null.
    doc.encryption = encryption;
    await doc.init();
    return doc;
  }

  /**
   * Protects the document with `password` from now on (Word's Encrypt with Password), or, with null,
   * takes the protection off. Only the downloaded file changes: the text doesn't.
   */
  async setPassword(password) {
    this.encryption = password == null ? null : { key: await keyForPassword(password), source: null, changed: true };
    this.encryptionChanged = true;
  }

  /** The saved package, encrypted with the document's password if it has one. */
  async sealed(blob, changed) {
    if (!this.encryption) return blob;
    // Unchanged, with the same password: the file goes back exactly as it was opened.
    if (!changed && !this.encryption.changed && this.encryption.source) return new Blob([this.encryption.source], { type: this.format.mime });
    return new Blob([await encrypt(new Uint8Array(await blob.arrayBuffer()), this.encryption.key)], { type: this.format.mime });
  }

  static async createBlank(name = 'Untitled.docx', files = blankDocxFiles()) {
    const zip = new JSZip();
    for (const [path, content] of Object.entries(files)) zip.file(path, content);
    const doc = await DocxDocument.open(await zip.generateAsync({ type: 'arraybuffer' }), name);
    doc.startedHere = true; // a new document: nobody has seen it, so its edits needn't be tracked
    return doc;
  }

  /** A new document from starter `id` (a letter, a memo…, see starters.js), on the paper used here. */
  static async createStarter(id) {
    const starter = STARTERS.find((s) => s.id === id);
    const doc = await DocxDocument.createBlank(`${starter.file}.docx`, starterFiles(id));
    doc.starter = id; // the editor opens Fill in for its placeholders
    return doc;
  }

  constructor(zip, name) {
    this.zip = zip;
    this.name = name;
    this.mediaUrls = new Map();
    // Pictures inserted this session, in the body or a header or footer: `${relsPath}|${rid}` ->
    // { rid, relsPath, part (a header or footer's path, or null for the body), path, target, data, ext, type }.
    this.newMedia = new Map();
    // Links added this session (pasted, inserted or retargeted), per relationships part:
    // rels path -> Map(address -> relationship id). The ids stay the same from one save to the next.
    this.linkIds = new Map();
    // Links from the file taken out of the text by redaction: their addresses go from the package too
    // (see forgetLink, writeForgottenLinks). `${part path}|${relationship id}` -> { path, rid, rel }.
    this.forgottenLinks = new Map();
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
    // The notes themselves are edited in the page, below the text (see renderNotes, serializeNotes).
    this.noteParts = {};
    for (const [kind, part] of [['footnote', footnotes], ['endnote', endnotes]]) {
      if (part) this.noteParts[kind] = { ...part, text: await this.text(part.path) };
    }
    const settings = await partOf('/settings');
    const evenAndOdd = onOff(wChild(settings?.xml.documentElement, 'evenAndOddHeaders')) === true;
    // Track Changes: on when the document says so, as in Word. Saved back only if it was switched.
    this.settingsPath = settings?.path || null;
    this.settingsText = settings ? await this.text(settings.path) : null;
    this.trackRevisions = onOff(wChild(settings?.xml.documentElement, 'trackRevisions')) === true;
    // Restrictions on editing that Word enforces, and digital signatures (see package.js). They are
    // never changed here: the editor opens such a document locked, as Word does.
    this.protection = readProtection(settings?.xml.documentElement);
    this.writeProtection = readWriteProtection(settings?.xml.documentElement);
    this.signatures = await readSignatures(this.zip);
    await this.readFormat();
    // Document properties and the sensitivity label (see properties.js), shown, never changed, except
    // that personal information can be removed before release (removePersonal, see writePersonal).
    this.propertyPaths = propertyPaths(pkgRels);
    this.properties = await readProperties(this.zip, pkgRels);
    this.removePersonal = false;
    this.removePersonalSet = !!wChild(settings?.xml.documentElement, 'removePersonalInformation');
    this.anonymized = new Map(); // part -> { written, original }: what writePersonal changed
    // Saving an edit records it in the core properties, as Word does (see writeCore): when, by whom
    // (modifiedBy, set by the editor), the revision number; and a title, if one is given here.
    this.coreOriginal = this.propertyPaths.core ? await this.text(this.propertyPaths.core) : null;
    this.updateProperties = true;
    this.modifiedBy = '';
    this.newTitle = null;
    // Properties edited here (Document properties…): { core: { title, … } or null, custom: [{ name,
    // value }] or null }, written on saving (see writeCore, writeCustom).
    this.customOriginal = this.propertyPaths.custom ? await this.text(this.propertyPaths.custom) : null;
    this.propEdits = { core: null, custom: null };
    this.changed = false; // whether the last write changed the document (see write)
    this.restored = false; // opened from a draft (see drafts.js): it differs from the file first opened
    // The package's parts as opened, to tell whether a save changes anything (see packageChanged).
    this.openedEntries = new Map(Object.entries(this.zip.files));
    this.tracking = this.trackRevisions;
    // Word updates fields (tables of contents, cross-references…) on opening if the file asks it to.
    this.updateFieldsSet = onOff(wChild(settings?.xml.documentElement, 'updateFields')) === true;
    this.updateFields = false; // asked for in the review, this session
    // New revisions get ids above every id in the document.
    let maxId = 0;
    for (const xml of [this.xml, ...Object.values(this.noteParts).map((p) => p.xml)]) {
      for (const el of xml.getElementsByTagName('*')) {
        const id = Number(el.getAttributeNS(W, 'id'));
        if (Number.isInteger(id) && id > maxId) maxId = id;
      }
    }
    // Annotations (comments, revisions, bookmarks) share one id space: new ones come from here.
    for (const c of this.comments) if (Number(c.id) > maxId) maxId = Number(c.id);
    this.nextId = maxId + 1;
    this.firstNewId = this.nextId; // ids from here on are this session's
    this.newComments = []; // comments added this session (see addComment)
    const loadVariants = async (local) => {
      const variants = await Promise.all(headerFooterVariants(this.body, local, evenAndOdd, true).map(async (v) => {
        const part = await loadPart(this, this.rels.get(v.rid));
        return part && { ...v, ...part, text: await this.text(part.path) }; // edited in the page (see renderBand)
      }));
      return variants.filter(Boolean);
    };
    // A different first page and different odd and even pages, as set here (Page setup) for the
    // whole document: null for as the file has them.
    this.evenAndOddInFile = evenAndOdd;
    this.hfFlags = { titlePg: null, evenAndOdd: null };
    this.headers = await loadVariants('headerReference');
    this.footers = await loadVariants('footerReference');
    // The document's sources, as Word's Source Manager keeps them (a b:Sources part in customXml),
    // for citations and a bibliography; written back only when a source or the style is changed.
    this.sources = { path: null, style: 'APA', list: [], changed: false };
    for (const path of Object.keys(this.zip.files).filter((f) => /^customXml\/item\d+\.xml$/i.test(f)).sort()) {
      try {
        const xml = parseXml(await this.text(path));
        if (!isSourcesPart(xml)) continue;
        const { style, sources } = readSources(xml);
        this.sources = { path, style, list: sources, changed: false };
        break;
      } catch { /* not XML this editor reads: kept as it is */ }
    }
    this.newNotes = new Map(); // 'footnote:7' -> the <w:footnote> of a note added this session
    // Paper size, orientation and margins set here for the whole document (see setPageSetup), and
    // the final section's own, to apply them to afresh at each save.
    this.pageSetup = null;
    const finalSect = wChild(this.body, 'sectPr');
    this.originalPage = Object.fromEntries(PAGE_SETUP_PARTS.map((local) => [local, wChild(finalSect, local)?.cloneNode(true) || null]));
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
    // Fonts embedded in the file are shown with it, as in Word (read from the file, never fetched).
    this.fontFaces = await loadEmbeddedFonts(await partOf('/fontTable'), (path) => this.zip.file(path)?.async('uint8array'));
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
    // Comments anchored in a header, footer or note: which part each is in, and its original text.
    this.commentStory = new Map();
    this.storyTexts = {};
    this.story = { edits: {}, done: {}, removed: [], replies: [] }; // changes to them (see storyState)
    const stories = Object.values(this.zip.files).filter((f) => !f.dir && /(^|\/)(header|footer|footnotes|endnotes)\d*\.xml$/.test(f.name));
    for (const f of stories) {
      const text = await f.async('string');
      if (!text.includes('commentReference')) continue;
      this.storyTexts[f.name] = text;
      for (const id of commentIdsIn(parseXml(text))) {
        this.commentsElsewhere.add(id);
        this.commentStory.set(id, f.name);
      }
    }
  }

  // ---------------------------------------------------------------- comments in headers, footers and notes

  /** The part (header, footer, notes) a comment (or the comment it replies to) is anchored in, or null. */
  storyOf(id) {
    if (this.commentStory.has(id)) return this.commentStory.get(id);
    const c = this.comments.find((x) => x.id === id);
    return c?.parentId ? this.storyOf(c.parentId) : null;
  }

  /** Where a comment in a header, footer or note is, in words: "In a header", or null. */
  storyPlace(id) {
    const path = this.storyOf(id);
    if (!path) return null;
    const name = path.split('/').pop();
    return name.startsWith('header') ? 'In a header' : name.startsWith('footer') ? 'In a footer'
      : name.startsWith('footnotes') ? 'In a footnote' : 'In an endnote';
  }

  /** Changes to comments in headers, footers and notes, as a string, so the editor's Undo covers them. */
  storyState() {
    return JSON.stringify(this.story);
  }

  setStoryState(json) {
    this.story = JSON.parse(json);
  }

  /** The comments removed from headers, footers and notes, with their replies. */
  storyRemoved() {
    const gone = new Set(this.story.removed);
    for (let grew = gone.size > 0; grew;) {
      grew = false;
      for (const c of this.comments) {
        if (c.parentId && gone.has(c.parentId) && !gone.has(c.id)) {
          gone.add(c.id);
          grew = true;
        }
      }
    }
    return gone;
  }

  /** True for a comment kept: not removed, and (for a reply added here) not undone. */
  storyCommentLive(c) {
    return !this.storyRemoved().has(c.id) && (!c.storyReply || this.story.replies.includes(c.id));
  }

  /** A reply to a comment anchored in a header, footer or note, written next to it there on save. */
  replyInStory(parentId, { author, text, date }) {
    const parent = this.comments.find((c) => c.id === parentId);
    if (!parent || !this.storyOf(parentId)) return null;
    parent.paraId ||= newParaId();
    const id = String(this.nextId++);
    const initials = author.split(/\s+/).filter(Boolean).map((w) => w[0]).join('').slice(0, 3).toUpperCase();
    const reply = {
      id, author, initials, date, text, paraId: newParaId(), parentId, done: false, isNew: true, storyReply: true,
      paragraphs: text.split('\n').map((x) => x.trim()).filter(Boolean),
    };
    this.comments.push(reply);
    this.newComments.push(reply);
    this.story.replies.push(id);
    return reply;
  }

  editStoryComment(id, text) {
    this.story.edits[id] = text;
  }

  resolveStoryComment(id, done) {
    this.story.done[id] = done;
  }

  deleteStoryComment(id) {
    this.story.removed.push(id);
  }

  /** The ids of comments in headers, footers and notes that are kept (with replies added here). */
  liveStoryComments() {
    return this.comments.filter((c) => this.storyOf(c.id) && this.storyCommentLive(c)).map((c) => c.id);
  }

  /** Comments that will be in the saved file: referenced in the body, or kept in a header, footer or note. */
  presentComments() {
    const present = commentIdsIn(this.body);
    for (const id of this.liveStoryComments()) present.add(id);
    return present;
  }

  /**
   * Writes the headers, footers and notes whose comments changed: removed comments' marks go,
   * replies' marks go next to their comment's (inside its range, as Word writes them). A part
   * changed by an earlier save and no longer is gets its original text back.
   */
  writeStories() {
    const removed = this.commentsRemoved;
    const replies = this.comments.filter((c) => c.storyReply && this.storyCommentLive(c) && !removed.has(c.id));
    const docs = new Map();
    const docFor = (path) => {
      if (!docs.has(path)) docs.set(path, this.partOutAt(path) || parseXml(this.storyTexts[path]));
      return docs.get(path);
    };
    for (const id of removed) if (this.commentStory.has(id)) docFor(this.commentStory.get(id));
    for (const r of replies) docFor(this.storyOf(r.parentId));
    for (const [path, xml] of docs) {
      pruneCommentMarkers(xml.documentElement, removed, new Set());
      for (const r of replies.filter((x) => this.storyOf(x.parentId) === path)) addReplyMarkers(xml, r.parentId, r.id);
      this.zip.file(path, serializeXml(xml));
    }
    for (const path of this.storiesWritten || []) if (!docs.has(path) && !this.partOutAt(path)) this.zip.file(path, this.storyTexts[path]);
    this.storiesWritten = new Set(docs.keys());
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
    for (const face of this.fontFaces || []) document.fonts.delete(face);
    this.fontFaces = [];
  }

  /** Page geometry in CSS px from the final section properties. */
  get page() {
    // As opened (a save writes the page setup into the section), with what is set here.
    const sectPr = wChild(this.body, 'sectPr');
    const pgSz = this.originalPage ? this.originalPage.pgSz : wChild(sectPr, 'pgSz');
    const pgMar = this.originalPage ? this.originalPage.pgMar : wChild(sectPr, 'pgMar');
    const px = (twips, dflt) => (twips != null ? Number(twips) / 15 : dflt);
    const set = this.pageSetup || {};
    return {
      width: px(set.w ?? wAttr(pgSz, 'w'), 816),
      minHeight: px(set.h ?? wAttr(pgSz, 'h'), 1056),
      top: px(set.top ?? wAttr(pgMar, 'top'), 96),
      bottom: px(set.bottom ?? wAttr(pgMar, 'bottom'), 96),
      left: px(set.left ?? wAttr(pgMar, 'left'), 96),
      right: px(set.right ?? wAttr(pgMar, 'right'), 96),
      header: px(wAttr(pgMar, 'header'), 48),
      footer: px(wAttr(pgMar, 'footer'), 48),
    };
  }

  /**
   * The paper size, orientation and margins as they are now (the final section's, with any set
   * here), in twips: { w, h, orient ('portrait' or 'landscape'), top, right, bottom, left }, and its
   * page numbering: { numFmt, numStart } (see numberingOf).
   */
  pageSetupNow() {
    // As opened: a save writes the setup into the section.
    const sect = createW(this.xml, 'sectPr');
    for (const local of PAGE_SETUP_PARTS) if (this.originalPage[local]) sect.append(this.originalPage[local].cloneNode(true));
    return { ...this.sectionSetupOf(sect), ...this.pageSetup };
  }

  /**
   * Sets the paper size, orientation and margins of the final section (the whole document's, when
   * it has one section; the editor writes the others' into their section breaks, see
   * setSectionSetup): `setup` has any of w, h, orient, top, right, bottom, left (twips), numFmt
   * and numStart (page numbering: a format, and a start or null to continue); null sets
   * them back as they were. What it doesn't name is kept as the section has it (see applyPageSetup).
   */
  setPageSetup(setup) {
    this.pageSetup = setup && Object.keys(setup).length ? { ...setup } : null;
  }

  // ---------------------------------------------------------------- sections

  /**
   * The section setup a sectPr gives: { w, h, orient, top, right, bottom, left } in twips, { numFmt,
   * numStart }, and its columns and start ({ cols, colSpace, colSep, colUnequal, type }).
   */
  sectionSetupOf(sectPr) {
    const sz = wChild(sectPr, 'pgSz');
    const mar = wChild(sectPr, 'pgMar');
    const num = (el, a, d) => Number(wAttr(el, a) ?? d);
    return {
      w: num(sz, 'w', 12240), h: num(sz, 'h', 15840), orient: wAttr(sz, 'orient') || 'portrait',
      top: num(mar, 'top', 1440), right: num(mar, 'right', 1440), bottom: num(mar, 'bottom', 1440), left: num(mar, 'left', 1440),
      ...numberingOf(wChild(sectPr, 'pgNumType')),
      ...columnsSetupOf(sectPr),
    };
  }

  /**
   * The columns of the section that section-end paragraph `el` ends, or of the final section (el
   * null), as they are now, for showing them: { count, gap (px), rule } or null for one column.
   */
  sectionColumns(el) {
    return columnsOf(el ? this.sectionBreakOf(el) : this.finalSectPr());
  }

  /** The section break a section-end paragraph holds (its w:sectPr), as it is now. */
  sectionBreakOf(el) {
    return el?.dataset.ppr != null ? wChild(this.raw[Number(el.dataset.ppr)], 'sectPr') : null;
  }

  /**
   * The setup of the section that section-end paragraph `el` ends, or of the final section (el
   * null), as it is now: { w, h, orient, top, right, bottom, left }.
   */
  sectionSetup(el) {
    return el ? this.sectionSetupOf(this.sectionBreakOf(el)) : this.pageSetupNow();
  }

  /**
   * Sets the paper, orientation and margins of the section that section-end paragraph `el` ends
   * (in its own w:sectPr), as Word's Page Setup does for "This section": `set` as setPageSetup's.
   */
  setSectionSetup(el, set) {
    const pPr = this.rawClone(el.dataset.ppr);
    applySetup(wChild(pPr, 'sectPr'), set);
    el.dataset.ppr = this.keep(pPr);
  }

  /**
   * Makes paragraph `el` end a section, as Word's Insert > Section break does: it holds a copy of
   * `sectPr` (the section it was in: its paper, margins, columns, headers and footers), so the text
   * up to it keeps them, and what follows begins a new section. A section's w:type says how that
   * section starts, so the copy keeps the type the section had; the section after the break is
   * the one to start on a new page or, for `type` 'continuous' (a break made for columns), on the
   * same page: the caller sets that on it (setSectionSetup or setPageSetup with { type }). A new
   * page's break, made before there was a choice, says so on the copy too, as before.
   */
  addSectionBreak(el, sectPr, type = 'nextPage') {
    const pPr = el.dataset.ppr != null ? this.rawClone(el.dataset.ppr) : createW(this.xml, 'pPr');
    removeW(pPr, 'sectPr');
    const copy = this.xml.importNode(sectPr, true);
    removeW(copy, 'sectPrChange');
    if (type === 'nextPage') {
      removeW(copy, 'type');
      ensureOrdered(copy, 'type', SECTPR_ORDER).setAttributeNS(W, 'w:val', 'nextPage');
    }
    const before = elementChildren(pPr).find((c) => c.localName === 'pPrChange');
    pPr.insertBefore(copy, before || null);
    el.dataset.ppr = this.keep(pPr);
    el.dataset.sectionEnd = type;
  }

  /**
   * The document's sections as they print, in order, from the section breaks in `root` (the page) as
   * they are now: [{ end (the paragraph ending it, or null for the final one), setup (sectionSetup's,
   * in twips), headerDist and footerDist (px), start ('nextPage', 'continuous'… how it begins), titlePg,
   * header and footer: { default, first, even } (indexes into this.headers and this.footers, or
   * null) }]. A section without a header (footer) of a type has the one before it has, as in Word;
   * one made here is the first section's (see newPart) where that section has none of its own.
   */
  printSections(root) {
    const ends = root ? [...root.querySelectorAll('[data-section-end]')].filter((el) => !el.closest('.docx-hf-band, .docx-notes, .textbox')) : [];
    const sects = [...ends.map((el) => ({ end: el, sectPr: this.sectionBreakOf(el) })), { end: null, sectPr: this.finalSectPr() }];
    const inherited = { header: {}, footer: {} };
    for (const [which, list] of [['header', this.headers], ['footer', this.footers]]) {
      list.forEach((v, i) => { if (v.added && v.section === 1) inherited[which][v.type] = i; });
    }
    return sects.map(({ end, sectPr }, i) => {
      const bands = {};
      for (const [which, list, local] of [['header', this.headers, 'headerReference'], ['footer', this.footers, 'footerReference']]) {
        for (const ref of wChildren(sectPr, local)) {
          const at = list.findIndex((v) => v.rid === ref.getAttributeNS(NS.r, 'id'));
          if (at >= 0) inherited[which][wAttr(ref, 'type') || 'default'] = at;
        }
        bands[which] = { default: inherited[which].default ?? null, first: inherited[which].first ?? null, even: inherited[which].even ?? null };
      }
      const mar = wChild(sectPr, 'pgMar');
      const dist = (a) => (wAttr(mar, a) != null ? Number(wAttr(mar, a)) / 15 : 48);
      return {
        end,
        setup: end ? this.sectionSetup(end) : this.pageSetupNow(),
        headerDist: dist('header'),
        footerDist: dist('footer'),
        // How it begins: the break before it says (see renderBody), the first one on a new page.
        start: i ? ends[i - 1].dataset.sectionEnd || 'nextPage' : 'nextPage',
        titlePg: this.hfFlags.titlePg ?? onOff(wChild(sectPr, 'titlePg')) === true,
        ...bands,
      };
    });
  }

  /** The final section's properties as they would be saved now (the page setup set here applied). */
  finalSectPr() {
    const sect = (wChild(this.body, 'sectPr') || createW(this.xml, 'sectPr')).cloneNode(true);
    for (const local of PAGE_SETUP_PARTS) {
      const orig = this.originalPage[local];
      const now = wChild(sect, local);
      if (orig) (now || ensureOrdered(sect, local, SECTPR_ORDER)).replaceWith(orig.cloneNode(true));
      else if (now && local !== 'pgSz' && local !== 'pgMar') sect.removeChild(now);
    }
    applySetup(sect, this.pageSetup || {});
    if (this.hfFlags.titlePg != null) setTitlePage(sect, this.hfFlags.titlePg);
    return sect;
  }

  /** Writes the page setup into each section of the body as it will be saved. */
  applyPageSetup() {
    const finalSect = wChild(this.body, 'sectPr');
    // The final section's properties persist between saves: start again from them as opened.
    if (finalSect) {
      // Replaced where they are, so a section's own order is kept.
      for (const local of PAGE_SETUP_PARTS) {
        const orig = this.originalPage[local];
        const now = wChild(finalSect, local);
        if (orig) (now || ensureOrdered(finalSect, local, SECTPR_ORDER)).replaceWith(orig.cloneNode(true));
        else if (now) finalSect.removeChild(now);
      }
    }
    // The others' are in their section breaks, written there by the editor (see setSectionSetup).
    if (finalSect && this.pageSetup) applySetup(finalSect, this.pageSetup);
    // A different first page, as set here for the whole document (Page setup): in every section.
    if (this.hfFlags.titlePg != null) for (const sect of this.body.getElementsByTagNameNS(W, 'sectPr')) setTitlePage(sect, this.hfFlags.titlePg);
  }

  // ---------------------------------------------------------------- inserting page numbers and footnotes

  /** Word's page number field (PAGE), as page nodes rendered as a field in a file is: its result is "1" until Word numbers the pages. */
  newPageField() {
    return this.newField(' PAGE   \\* MERGEFORMAT ', '1');
  }

  /**
   * A field as Word writes one (begin, instruction `instr`, separate, its result `result`, end), as
   * rendered nodes to put at the caret: a page number, or today's date that Word updates (DATE).
   */
  /** A field with no result (an index entry, XE), rendered: its begin, instruction and end markers, hidden. */
  newBareField(instr) {
    const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const p = parseXml(`<w:p xmlns:w="${W}"><w:r><w:fldChar w:fldCharType="begin"/></w:r>`
      + `<w:r><w:instrText xml:space="preserve">${esc(instr)}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`).documentElement;
    const holder = document.createElement('span');
    this.renderInlines(this.xml.importNode(p, true), holder);
    return [...holder.childNodes];
  }

  newField(instr, result) {
    const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const p = parseXml(`<w:p xmlns:w="${W}"><w:r><w:fldChar w:fldCharType="begin"/></w:r>`
      + `<w:r><w:instrText xml:space="preserve">${esc(instr)}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`
      + `<w:r><w:t xml:space="preserve">${esc(result)}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`).documentElement;
    const holder = document.createElement('span');
    this.renderInlines(this.xml.importNode(p, true), holder);
    return [...holder.childNodes];
  }

  /**
   * A new header or footer part (`which`) of `type` ('default', 'first' or 'even') for the first
   * section: one empty paragraph in Word's Header (Footer) style, a default footer centred (for a
   * page number). Saved (with its relationship, content type and the section's reference to it)
   * only while it is in the page and its type is used (see placeNewFooters). Returns the variant.
   */
  newPart(which, type = 'default') {
    const list = which === 'header' ? this.headers : this.footers;
    const dir = this.mainDir ? `${this.mainDir}/` : '';
    let n = 1;
    while (this.zip.file(`${dir}${which}${n}.xml`) || list.some((v) => v.path === `${dir}${which}${n}.xml`)) n++;
    const prefix = which === 'header' ? 'rIdHeader' : 'rIdFooter';
    let rid = `${prefix}1`;
    for (let i = 1; this.rels.has(rid) || list.some((v) => v.rid === rid); i++) rid = `${prefix}${i}`;
    const style = this.styles.ensureBuiltin(which);
    const jc = which === 'footer' && type === 'default' ? '<w:jc w:val="center"/>' : '';
    const root = which === 'header' ? 'hdr' : 'ftr';
    const text = `${XML_HEADER}<w:${root} xmlns:w="${W}" xmlns:r="${NS.r}" ${VML_NS}><w:p><w:pPr><w:pStyle w:val="${style}"/>${jc}</w:pPr></w:p></w:${root}>`;
    const v = { rid, type, section: 1, path: `${dir}${which}${n}.xml`, xml: parseXml(text), rels: new Map(), text, added: true, titlePgInFile: false, shownInFile: true };
    list.push(v);
    this.relabelVariants();
    return v;
  }

  /** A footer for all pages, for a document without one (see newPart). Returns the band. */
  newFooter() {
    this.newPart('footer');
    return this.renderBand('footer');
  }

  /** A header for all pages, for a document without one (see newPart). Returns the band. */
  newHeader() {
    this.newPart('header');
    return this.renderBand('header');
  }

  // ---------------------------------------------------------------- a different first page, odd and even pages

  /** Whether header (footer) variant `v` is used: a first page's only with a different first page, and so on. */
  variantShown(v) {
    if (v.type === 'first') return this.hfFlags.titlePg ?? v.titlePgInFile;
    if (v.type === 'even') return this.hfFlags.evenAndOdd ?? this.evenAndOddInFile;
    return true;
  }

  /** Whether the document has a different first page and different odd and even pages, as it is now. */
  headerFooterFlags() {
    const firstSect = this.body.getElementsByTagNameNS(W, 'sectPr')[0];
    return {
      titlePg: this.hfFlags.titlePg ?? onOff(wChild(firstSect, 'titlePg')) === true,
      evenAndOdd: this.hfFlags.evenAndOdd ?? this.evenAndOddInFile,
    };
  }

  /** Names the variants again (and their hosts under `root`) for the flags as they are now. */
  relabelVariants(root = null) {
    for (const [which, list] of [['header', this.headers], ['footer', this.footers]]) {
      labelVariants(list, { titlePg: (v) => this.hfFlags.titlePg ?? v.titlePgInFile, evenAndOdd: this.hfFlags.evenAndOdd ?? this.evenAndOddInFile });
      list.forEach((v, i) => {
        const host = root?.querySelector(`.docx-hf-part[data-which="${which}"][data-variant="${i}"]`);
        if (host) host.dataset.label = hostLabel(which, v.label);
      });
    }
  }

  /**
   * A different first page and different odd and even pages for the whole document, as Word's
   * Header & Footer options set them: `flags` { titlePg, evenAndOdd } (true, false, or null for as
   * the file has it). Returns the variants (header and footer) of each type now used that the first
   * section doesn't have yet, made new and empty, as Word starts them: [{ which, variant }].
   */
  setHeaderFooterFlags(flags) {
    this.hfFlags = { titlePg: flags.titlePg ?? null, evenAndOdd: flags.evenAndOdd ?? null };
    const made = [];
    const now = this.headerFooterFlags();
    for (const [type, on] of [['first', now.titlePg], ['even', now.evenAndOdd]]) {
      if (!on) continue;
      for (const which of ['header', 'footer']) {
        const list = which === 'header' ? this.headers : this.footers;
        if (list.some((v) => v.type === type && v.section === 1)) continue;
        made.push({ which, type, variant: this.newPart(which, type) });
      }
    }
    this.relabelVariants();
    return made;
  }

  /** The host for header (footer) variant `v`, number `i` in its list: rendered, to edit in the page. */
  renderVariantHost(which, v, i) {
    const host = document.createElement('div');
    host.className = 'docx-hf-part';
    host.contentEditable = 'true';
    host.dataset.which = which;
    host.dataset.variant = String(i);
    host.dataset.part = v.path;
    host.dataset.type = v.type;
    host.dataset.protected = `the ${which}`;
    host.dataset.label = hostLabel(which, v.label);
    this.inPart(v, () => this.renderBlocks(v.xml.documentElement, host));
    if (!host.firstChild) host.append(this.emptyParagraph());
    // Put in the page after the others (switched on): as it saves untouched, so it goes back as it was.
    if (this.partBaseline && !this.partBaseline.has(v.path)) this.partBaseline.set(v.path, this.noteXml(host));
    return host;
  }

  /**
   * Word's text watermark (Design > Watermark): WordArt in the header, behind the text, centred on
   * the page between the margins, as Word writes it (a VML text path, PowerPlusWaterMarkObject), in
   * semi-transparent silver, diagonal (turned 315°) or horizontal. Returns the run, rendered, for the
   * start of a header's first paragraph. `n` tells one header's from another's (their shape ids).
   */
  newWatermark(text, { diagonal = true, font = 'Calibri', n = 1 } = {}) {
    const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    // Word sizes the WordArt to the text: as wide as the margins (a little more, turned), less for a
    // short word, and its height in proportion to the letters (about 3:1 for each).
    const letters = Math.max(1, [...text].length);
    const across = (this.page.width - this.page.left - this.page.right) * 0.75; // pt between the margins
    const w = Math.round(across * (diagonal ? 1.15 : 1) * Math.min(1, 0.55 + letters * 0.04) * 10) / 10;
    const h = Math.round(((w * 3) / letters) * 10) / 10;
    const id = 357831060 + n;
    const xml = `<w:r xmlns:w="${W}" ${VML_NS}><w:rPr><w:noProof/></w:rPr><w:pict>`
      + `<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">`
      + '<v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas>'
      + '<v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>'
      + '<v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>'
      + `<v:shape id="PowerPlusWaterMarkObject${id}" o:spid="_x0000_s${2049 + n}" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:${w}pt;height:${h}pt;${diagonal ? 'rotation:315;' : ''}z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="silver" stroked="f">`
      + `<v:fill opacity=".5"/><v:textpath style="font-family:&quot;${esc(font)}&quot;;font-size:1pt" string="${esc(text)}"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape>`
      + '</w:pict></w:r>';
    const run = this.xml.importNode(parseXml(xml).documentElement, true);
    return this.rawInline(run);
  }

  /**
   * A new footnote or endnote (`kind`): { ref (its reference, for the caret), note (the note, for
   * the notes section) }. The note starts with its number and a space, in Word's Footnote (Endnote)
   * Text style, as Word starts one. A document without such notes gets the part for them when
   * saved. The note is saved only while its reference is in the text (see serializeNotes).
   */
  newNote(kind = 'footnote') {
    if (!this.noteParts[kind]) {
      const dir = this.mainDir ? `${this.mainDir}/` : '';
      const file = `${kind}s`;
      let n = '';
      while (this.zip.file(`${dir}${file}${n}.xml`)) n = Number(n) + 1;
      const sep = (local) => `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:${local}/></w:r></w:p>`;
      const text = `${XML_HEADER}<w:${file} xmlns:w="${W}" xmlns:r="${NS.r}">`
        + `<w:${kind} w:type="separator" w:id="-1">${sep('separator')}</w:${kind}>`
        + `<w:${kind} w:type="continuationSeparator" w:id="0">${sep('continuationSeparator')}</w:${kind}></w:${file}>`;
      const prefix = `rId${NOTE_NAMES[kind]}s`;
      let rid = `${prefix}1`;
      for (let i = 1; this.rels.has(rid); i++) rid = `${prefix}${i}`;
      this.noteParts[kind] = { path: `${dir}${file}${n}.xml`, xml: parseXml(text), rels: new Map(), text, added: true, rid };
    }
    const part = this.noteParts[kind];
    const ids = [...wChildren(part.xml.documentElement, kind).map((el) => Number(wAttr(el, 'id'))),
      ...[...this.newNotes.keys()].filter((k) => k.startsWith(`${kind}:`)).map((k) => Number(k.split(':')[1]))];
    const id = String(Math.max(0, ...ids.filter(Number.isFinite)) + 1);
    const refStyle = this.styles.ensureBuiltin(`${kind}Reference`);
    const textStyle = this.styles.ensureBuiltin(`${kind}Text`);
    const parse = (xml) => this.xml.importNode(parseXml(`<root xmlns:w="${W}">${xml}</root>`).documentElement.firstElementChild, true);
    const mark = `<w:r><w:rPr><w:rStyle w:val="${refStyle}"/></w:rPr>`;
    const ref = this.rawInline(parse(`${mark}<w:${kind}Reference w:id="${id}"/></w:r>`));
    const node = parse(`<w:${kind} w:id="${id}"><w:p><w:pPr><w:pStyle w:val="${textStyle}"/></w:pPr>${mark}<w:${kind}Ref/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r></w:p></w:${kind}>`);
    this.newNotes.set(`${kind}:${id}`, node);
    const note = this.renderNote(kind, node, id, '');
    note.dataset.added = '';
    return { ref, note };
  }

  /**
   * A paragraph's outline level (0 for Heading 1 … 8; 9 or null for body text): its style's, or with
   * `direct` its own first (Word's \u in a table of contents).
   */
  outlineLevel(el, direct = false) {
    const own = direct && el.dataset.ppr != null ? wChild(this.raw[Number(el.dataset.ppr)], 'outlineLvl') : null;
    if (own) return Number(wAttr(own, 'val'));
    const chain = this.styles.chain(el.dataset.style || this.styles.defaultParagraph);
    for (let i = chain.length - 1; i >= 0; i--) {
      const lvl = wChild(chain[i].pPr, 'outlineLvl');
      if (lvl) return Number(wAttr(lvl, 'val'));
      // Word's heading styles are outline levels even where the file doesn't say so.
      const heading = /^heading ([1-9])$/.exec(chain[i].name?.toLowerCase() || '');
      if (heading) return Number(heading[1]) - 1;
    }
    return null;
  }

  /** Paragraphs given as XML (strings in the w namespace), rendered for the page. */
  renderXmlParagraphs(paras) {
    const body = parseXml(bodyXml(paras)).documentElement;
    return elementChildren(body).map((p) => this.renderParagraph(this.xml.importNode(p, true)));
  }

  /** Runs given as XML (an equation's m:oMath too), rendered as nodes to put in a paragraph. */
  renderXmlRuns(runs) {
    const p = parseXml(`<w:p xmlns:w="${W}" xmlns:m="${MATH_NS}">${runs}</w:p>`).documentElement;
    const holder = document.createElement('span');
    this.renderInlines(this.xml.importNode(p, true), holder);
    return [...holder.childNodes];
  }

  /** Adds `source` to the document's sources (with a new tag), returning it as kept. */
  addSource(source) {
    const kept = { ...source, tag: newTag(source, new Set(this.sources.list.map((s) => s.tag))), guid: newGuid() };
    this.sources.list.push(kept);
    this.sources.changed = true;
    return kept;
  }

  /** Sets the citation style (APA or MLA) the document's sources part records. */
  setCitationStyle(style) {
    if (this.sources.style === style) return;
    this.sources.style = style;
    this.sources.changed = true;
  }

  /**
   * Writes the sources part when a source or the style was changed: in place of the one the file
   * had, or as a new one (customXml/itemN.xml, its properties part and relationships, and the
   * main part's relationship to it), as Word adds one.
   */
  async writeSources() {
    const src = this.sources;
    if (!src?.changed) return;
    const xml = sourcesXml(src.style, src.list);
    if (src.path) {
      this.zip.file(src.path, xml);
      return;
    }
    let n = 1;
    while (this.zip.file(`customXml/item${n}.xml`) || this.zip.file(`customXml/itemProps${n}.xml`)) n++;
    const path = `customXml/item${n}.xml`;
    this.zip.file(path, xml);
    this.zip.file(`customXml/itemProps${n}.xml`, sourcesPropsXml(newGuid()));
    this.zip.file(`customXml/_rels/item${n}.xml.rels`, `${XML_HEADER}<Relationships xmlns="${NS.pr}"><Relationship Id="rId1" Type="${CUSTOM_XML_PROPS_REL}" Target="itemProps${n}.xml"/></Relationships>`);
    const rels = parseXml((await this.text(this.relsPath)) || `<Relationships xmlns="${NS.pr}"/>`);
    let id = 1;
    while ([...rels.documentElement.children].some((r) => r.getAttribute('Id') === `rIdSources${id}`)) id++;
    const rel = rels.createElementNS(NS.pr, 'Relationship');
    rel.setAttribute('Id', `rIdSources${id}`);
    rel.setAttribute('Type', CUSTOM_XML_REL);
    rel.setAttribute('Target', `${this.mainDir ? '../' : ''}${path}`);
    rels.documentElement.appendChild(rel);
    this.zip.file(this.relsPath, serializeXml(rels));
    const ct = parseXml(await this.text('[Content_Types].xml'));
    const kids = [...ct.documentElement.children];
    if (!kids.some((d) => d.localName === 'Default' && d.getAttribute('Extension')?.toLowerCase() === 'xml')) {
      const d = ct.createElementNS(NS.ct, 'Default');
      d.setAttribute('Extension', 'xml');
      d.setAttribute('ContentType', 'application/xml');
      ct.documentElement.insertBefore(d, ct.documentElement.firstChild);
    }
    const ov = ct.createElementNS(NS.ct, 'Override');
    ov.setAttribute('PartName', `/customXml/itemProps${n}.xml`);
    ov.setAttribute('ContentType', CUSTOM_XML_PROPS_CT);
    ct.documentElement.appendChild(ov);
    this.zip.file('[Content_Types].xml', serializeXml(ct));
    src.path = path;
  }

  /** A new bookmark called `name`: its start and end markers, hidden in the page as a file's are. */
  newBookmark(name) {
    const id = String(this.nextId++);
    const start = createW(this.xml, 'bookmarkStart');
    start.setAttributeNS(W, 'w:id', id);
    start.setAttributeNS(W, 'w:name', name);
    const end = createW(this.xml, 'bookmarkEnd');
    end.setAttributeNS(W, 'w:id', id);
    return [this.hiddenRaw(start), this.hiddenRaw(end)];
  }

  /** Paragraph `el` of the page as XML now, outside a save (as tableWithEdits does for a table). */
  paragraphXml(el) {
    this.revisions = [];
    this.linksUsed ||= new Map();
    this.serialRels ||= { path: this.relsPath, rels: this.rels };
    const p = this.serializeParagraph(el, [...el.childNodes]);
    this.revisions = [];
    return p;
  }

  /** The section holding the notes of `kind` below the text (empty). */
  notesSection(kind) {
    const section = document.createElement('section');
    section.className = 'docx-notes';
    section.dataset.notes = kind;
    section.dataset.label = `${NOTE_NAMES[kind]}s`; // drawn by CSS, so it is never text in the page
    section.contentEditable = 'false';
    section.dataset.protected = `the ${kind}s`;
    section.title = `${NOTE_NAMES[kind]}s – click a note’s text to edit it`;
    return section;
  }

  /** One note (`node`, its <w:footnote> or <w:endnote>) as an editing host, numbered `n`. */
  renderNote(kind, node, id, n) {
    const part = this.noteParts[kind];
    const rels = this.rels;
    const inNote = this.inNote;
    this.rels = part.rels;
    this.inNote = true;
    try {
      const el = document.createElement('div');
      el.className = 'docx-note';
      el.contentEditable = 'true';
      el.dataset.note = kind;
      el.dataset.noteId = id;
      el.dataset.n = n;
      el.dataset.protected = `${kind} ${n}`;
      this.noteMarkN = n;
      this.renderBlocks(node, el);
      if (!el.firstChild) el.append(this.emptyParagraph());
      return el;
    } finally {
      this.rels = rels;
      this.inNote = inNote;
    }
  }

  /**
   * Adds (on) or takes out (off) a part made this session at `path`: its relationship from the main
   * document (id `rid`) and its content type; the part's file itself is written or removed too.
   */
  async registerPart(path, rid, relType, contentType, on, xml = null) {
    if (on) this.zip.file(path, xml);
    else if (this.zip.file(path)) this.zip.remove(path);
    const rels = parseXml((await this.text(this.relsPath)) || `<Relationships xmlns="${NS.pr}"/>`);
    const ct = parseXml(await this.text('[Content_Types].xml'));
    const rel = [...rels.documentElement.children].find((r) => r.getAttribute('Id') === rid);
    const ov = [...ct.documentElement.children].find((o) => o.localName === 'Override' && o.getAttribute('PartName') === `/${path}`);
    if (on === !!rel && on === !!ov) return;
    if (on && !rel) {
      const el = rels.createElementNS(NS.pr, 'Relationship');
      el.setAttribute('Id', rid);
      el.setAttribute('Type', relType);
      el.setAttribute('Target', this.mainDir ? path.slice(this.mainDir.length + 1) : path);
      rels.documentElement.appendChild(el);
    } else if (!on && rel) rel.remove();
    if (on && !ov) {
      const el = ct.createElementNS(NS.ct, 'Override');
      el.setAttribute('PartName', `/${path}`);
      el.setAttribute('ContentType', contentType);
      ct.documentElement.appendChild(el);
    } else if (!on && ov) ov.remove();
    this.zip.file(this.relsPath, serializeXml(rels));
    this.zip.file('[Content_Types].xml', serializeXml(ct));
  }

  /**
   * A footer (or header) made here is referred to by the first section (and so by the sections after
   * it that have none of their own) while its band is in the page, else by none. Run on the body as
   * saved.
   */
  placeNewFooters(root) {
    const sects = [...this.body.getElementsByTagNameNS(W, 'sectPr')];
    for (const [local, v] of [...this.footers.map((f) => ['footerReference', f]), ...this.headers.map((h) => ['headerReference', h])]) {
      if (!v.added) continue;
      for (const s of sects) for (const ref of wChildren(s, local)) if (ref.getAttributeNS(NS.r, 'id') === v.rid) s.removeChild(ref);
      v.present = !!root.querySelector(`.docx-hf-part[data-part="${CSS.escape(v.path)}"]`) && this.variantShown(v);
      if (!v.present || !sects[0]) continue;
      const ref = createW(this.xml, local);
      ref.setAttributeNS(W, 'w:type', v.type || 'default');
      ref.setAttributeNS(NS.r, 'r:id', v.rid);
      const refs = elementChildren(sects[0]).filter((c) => c.localName === 'headerReference' || c.localName === 'footerReference');
      sects[0].insertBefore(ref, refs.length ? refs[refs.length - 1].nextSibling : sects[0].firstChild);
    }
  }

  /** Writes or takes out the footers and footnotes part made this session (see placeNewFooters, serializeNotes). */
  async writeNewParts() {
    for (const v of this.footers.filter((f) => f.added)) {
      const xml = this.partsOut?.[v.path];
      await this.registerPart(v.path, v.rid, FOOTER_REL, FOOTER_CT, !!(v.present && xml), xml && serializeXml(xml));
    }
    for (const v of this.headers.filter((h) => h.added)) {
      const xml = this.partsOut?.[v.path];
      await this.registerPart(v.path, v.rid, HEADER_REL, HEADER_CT, !!(v.present && xml), xml && serializeXml(xml));
    }
    for (const [kind, rel, ct] of [['footnote', FOOTNOTES_REL, FOOTNOTES_CT], ['endnote', ENDNOTES_REL, ENDNOTES_CT]]) {
      const notes = this.noteParts[kind];
      if (!notes?.added) continue;
      const xml = this.notesOut?.[kind];
      await this.registerPart(notes.path, notes.rid, rel, ct, !!xml, xml && serializeXml(xml));
    }
  }

  css(scope) {
    return this.styles.css(scope);
  }

  /** Style id to use for a builtin block kind ('normal', 'title', 'h1'..). Adds missing styles. */
  styleIdFor(kind) {
    if (kind === 'normal') return null;
    return this.styles.ensureBuiltin(kind);
  }

  /** The document's own paragraph styles for the style menu (see StyleSheet.paragraphChoices). */
  styleChoices() {
    const used = new Set();
    for (const xml of [this.xml, ...Object.values(this.noteParts).map((p) => p.xml)]) {
      for (const el of xml.getElementsByTagNameNS(W, 'pStyle')) used.add(wAttr(el, 'val'));
    }
    return this.styles.paragraphChoices(used);
  }

  /** The style menu's value for a paragraph style: 'normal', 'title', 'h1'–'h3', or 'style:<id>'. */
  blockValue(styleId) {
    const kind = this.blockKind(styleId);
    return ['normal', 'title', 'h1', 'h2', 'h3'].includes(kind) ? kind : `style:${styleId}`;
  }

  styleLabel(styleId) {
    return this.styles.label(styleId);
  }

  /** The element a paragraph in style `styleId` is rendered as: h1–h6 for a title or heading, else p. */
  tagFor(styleId) {
    return HEADING_TAG[this.styles.kindOf(styleId)] || 'p';
  }

  blockKind(styleId) {
    return this.styles.kindOf(styleId) || (styleId && styleId !== this.styles.defaultParagraph ? 'other' : 'normal');
  }

  // ---------------------------------------------------------------- XML -> HTML

  /** Renders the document body into a DocumentFragment of editable HTML. */
  render() {
    this.raw = [];
    this.rPrKeys = new Map();
    this.numbering = new Numbering(this.numberingXml);
    this.commentOrder = new Map(); // comment id -> marker number, in document order
    this.noteCount = { footnote: 0, endnote: 0 };
    this.noteOrder = { footnote: [], endnote: [] }; // note ids in the order the text refers to them
    const frag = document.createDocumentFragment();
    this.renderBody(frag);
    if (!frag.childNodes.length) frag.append(this.emptyParagraph());
    for (const kind of NOTE_KINDS) {
      const notes = this.renderNotes(kind);
      if (notes) frag.append(notes);
    }
    const header = this.renderBand('header');
    if (header) frag.prepend(header);
    const footer = this.renderBand('footer');
    if (footer) frag.append(footer);
    this.applyContextualSpacing(frag);
    this.applyTabStops(frag);
    this.markTracked(frag);
    // Each text box, note, header and footer as it would be saved now: one that still saves the same
    // way is written back as it was. Text boxes first, innermost first: the others hold them.
    this.revisions = [];
    this.textboxBaseline = new Map();
    for (const el of [...frag.querySelectorAll('.textbox')].reverse()) this.textboxBaseline.set(el.dataset.textbox, this.textboxXml(el));
    this.noteBaseline = new Map();
    for (const el of frag.querySelectorAll('.docx-note')) this.noteBaseline.set(noteKey(el), this.noteXml(el));
    this.partBaseline = new Map();
    for (const el of frag.querySelectorAll('.docx-hf-part')) this.partBaseline.set(el.dataset.part, this.noteXml(el));
    return frag;
  }

  /**
   * The body's blocks, section by section: a section set in columns (w:cols) is wrapped in a
   * `div.docx-columns` that shows them, as Word lays it out. The wrapper is only for the page:
   * saving reads the blocks inside it, and the section's settings stay in the file where they are.
   */
  renderBody(frag) {
    let section = document.createDocumentFragment();
    let lastEnd = null; // the paragraph ending the section before
    const close = (sectPr) => {
      // The break before this section is the kind this section starts with (its w:type): a new
      // page unless it says otherwise, as Word reads it.
      if (lastEnd) lastEnd.dataset.sectionEnd = wAttr(wChild(sectPr, 'type'), 'val') || 'nextPage';
      lastEnd = [...section.querySelectorAll('[data-section-end]')].pop() || null;
      const cols = columnsOf(sectPr);
      if (cols && section.firstChild) {
        const div = columnsWrapper(cols);
        div.append(section);
        frag.append(div);
      } else {
        frag.append(section);
      }
      section = document.createDocumentFragment();
    };
    for (const child of elementChildren(this.body)) {
      if (child.namespaceURI === W && child.localName === 'sectPr') continue;
      this.renderBlock(child, section);
      const sectPr = child.namespaceURI === W && child.localName === 'p' && wChild(wChild(child, 'pPr'), 'sectPr');
      if (sectPr) close(sectPr); // the paragraph holding a section break ends its section
    }
    close(wChild(this.body, 'sectPr') || createW(this.xml, 'sectPr'));
  }

  // ---------------------------------------------------------------- headers and footers

  /**
   * The headers (or footers) as a band at the top (bottom) of the page: each one an editing host
   * of its own, so typing, formatting, Undo, Find and Track Changes work there as in the text. The
   * band itself can't be edited or deleted; the view shows the variant chosen in its tabs (first
   * page, other pages, sections). Saved to their own parts by serializeParts.
   */
  renderBand(which, only = null) {
    const all = which === 'header' ? this.headers : this.footers;
    // The ones the document uses: a first page's or even pages' one it doesn't use is put in the
    // page only when switched on (see the editor's applyHeaderFooterFlags).
    const variants = only || all.filter((v) => this.variantShown(v));
    if (!variants.length) return null;
    const band = document.createElement('section');
    band.className = `docx-hf-band ${which}`;
    band.contentEditable = 'false';
    band.dataset.protected = `the ${which}`;
    band.title = `${which === 'header' ? 'Header' : 'Footer'} – click its text to edit it`;
    variants.forEach((v) => band.append(this.renderVariantHost(which, v, all.indexOf(v))));
    return band;
  }

  /** Runs `fn` with the context of a header or footer part: its relationships, its own list numbering. */
  inPart(part, fn) {
    const saved = [this.rels, this.numbering, this.commentOrder, this.noteCount, this.noteOrder, this.inNote];
    this.rels = part.rels;
    this.numbering = new Numbering(this.numberingXml);
    this.commentOrder = new Map();
    this.noteCount = { footnote: 0, endnote: 0 };
    this.noteOrder = null; // a note reference in a header stays read-only
    this.inNote = true; // comments in it stay with the comments panel (see doc.story), as in notes
    try {
      fn();
    } finally {
      [this.rels, this.numbering, this.commentOrder, this.noteCount, this.noteOrder, this.inNote] = saved;
    }
  }

  /**
   * The header and footer parts as they will be saved, from their hosts in the editor content
   * `root`: path -> the part's XML with its new content, for each one that changed.
   */
  serializeParts(root) {
    const out = {};
    for (const v of [...this.headers, ...this.footers]) {
      const host = root.querySelector(`.docx-hf-part[data-part="${CSS.escape(v.path)}"]`);
      if (!host) continue;
      const mainRels = this.serialRels;
      this.serialRels = { path: relsPathOf(v.path), rels: v.rels };
      const nodes = this.serializeBlocks(host);
      this.serialRels = mainRels;
      const s = new XMLSerializer();
      if (nodes.map((n) => s.serializeToString(n)).join('') === this.partBaseline.get(v.path)) continue;
      const xml = parseXml(v.text);
      const top = xml.documentElement;
      for (const c of elementChildren(top)) top.removeChild(c);
      for (const n of nodes) top.appendChild(n); // adopted, so revision ids set later reach them
      if (!nodes.some((n) => n.localName === 'p')) top.appendChild(createW(xml, 'p'));
      out[v.path] = xml;
    }
    return out;
  }

  /** Writes the headers and footers that changed; one changed by an earlier save and no longer gets its original back. */
  writeParts() {
    for (const v of [...this.headers, ...this.footers]) {
      if (v.added) continue; // see writeNewParts
      const xml = this.partsOut?.[v.path];
      if (xml) this.zip.file(v.path, serializeXml(xml));
      else if (v.written) this.zip.file(v.path, v.text);
      v.written = !!xml;
    }
  }

  /** A note, header or footer part as it will be saved, if it changed, else null. */
  partOutAt(path) {
    return this.notesOutAt(path) || this.partsOut?.[path] || null;
  }

  /** The paragraphs of headers and footers that changed, for the review: "First page header", paragraph 1… */
  reviewParts(root, opts) {
    const out = [];
    for (const v of [...this.headers, ...this.footers]) {
      const xml = this.partsOut[v.path];
      if (!xml) continue;
      const host = root.querySelector(`.docx-hf-part[data-part="${CSS.escape(v.path)}"]`);
      const where = host?.dataset.label || 'Header';
      for (const entry of reviewParagraphs(parseXml(v.text).documentElement, xml.documentElement, opts)) out.push({ ...entry, where });
    }
    return out;
  }

  // ---------------------------------------------------------------- text boxes

  /**
   * A text box (or a shape with text) as it is laid out: a frame (not editable itself) with the
   * shape's size, place, fill and outline, holding its text as an editing host of its own. The
   * shape's XML is kept; only its text is written back when it changed (see serializeTextbox).
   */
  renderTextbox(r, info) {
    const span = document.createElement('span');
    span.className = 'textbox';
    span.contentEditable = 'false';
    span.dataset.textbox = this.keep(r);
    span.dataset.wrap = info.wrap;
    span.title = 'Text box – click its text to edit it; its size, place and look stay as they are';
    const w = info.cx / EMU_PER_PX;
    const h = info.cy / EMU_PER_PX;
    this.placePicture(span, info, w, h);
    const st = span.style;
    if (w) st.width = `${w}px`;
    if (h) st.minHeight = `${h}px`; // the text can grow it here, so nothing typed is hidden
    st.padding = info.insets.map((x) => `${Math.max(0, x)}px`).join(' ');
    if (info.fill) st.backgroundColor = `#${info.fill}`;
    if (info.line) st.border = `${info.line.width}px solid #${info.line.color}`;
    if (info.geometry === 'ellipse') st.borderRadius = '50%';
    else if (/^round/.test(info.geometry)) st.borderRadius = '12px';
    if (info.anchor === 'ctr' || info.anchor === 'b') st.justifyContent = info.anchor === 'ctr' ? 'center' : 'flex-end';
    const host = document.createElement('div');
    host.className = 'textbox-content';
    host.contentEditable = 'true';
    const { main } = textboxContents(r);
    this.renderBlocks(main[0], host);
    if (!host.firstChild) host.append(this.emptyParagraph());
    span.append(host);
    return span;
  }

  /** The text of a rendered text box as it would be saved now, as a string. */
  textboxXml(el) {
    const host = el.querySelector(':scope > .textbox-content');
    return host ? this.noteXml(host) : '';
  }

  /**
   * A text box's run: as it was if its text is unchanged; otherwise with the new text, in its shown
   * text and its copies for older readers (given the same text in update, once revision ids are set).
   */
  serializeTextbox(el) {
    const r = this.rawClone(el.dataset.textbox);
    const host = el.querySelector(':scope > .textbox-content');
    if (!host) return r;
    const blocks = this.serializeBlocks(host);
    const s = new XMLSerializer();
    if (blocks.map((n) => s.serializeToString(n)).join('') === this.textboxBaseline?.get(el.dataset.textbox)) return r;
    const { main, copies } = setTextboxText(r, blocks); // adopted, so revision ids set later reach them
    if (copies.length) this.textboxCopies?.push({ main, copies });
    return r;
  }

  // ---------------------------------------------------------------- footnotes and endnotes

  /**
   * The footnotes (or endnotes) the text refers to, in its order, as a section at the end of the
   * page: each note an editing host of its own, numbered as in the text. Null if there are none.
   * The section itself can't be edited or deleted; Word keeps notes in their own part, written by
   * serializeNotes. Comments in notes stay with the comments panel (see doc.story).
   */
  renderNotes(kind) {
    const part = this.noteParts[kind];
    const ids = [...new Set(this.noteOrder[kind])];
    if (!part || !ids.length) return null;
    const byId = new Map(wChildren(part.xml.documentElement, kind).map((n) => [wAttr(n, 'id'), n]));
    const section = this.notesSection(kind);
    for (const id of ids) {
      const node = byId.get(id);
      if (node) section.append(this.renderNote(kind, node, id, this.noteNumber(kind, this.noteOrder[kind].indexOf(id) + 1)));
    }
    return section.firstChild ? section : null;
  }

  noteNumber(kind, n) {
    return kind === 'footnote' ? String(n) : romanLower(n);
  }

  /** The XML a rendered note would be saved as now, as a string. */
  noteXml(el) {
    const s = new XMLSerializer();
    return this.serializeBlocks(el).map((n) => s.serializeToString(n)).join('');
  }

  /**
   * Numbers the note references in the text and their notes in order, as Word does, and hides the
   * notes whose reference was deleted (Word doesn't show them either).
   */
  renumberNotes(root) {
    const count = { footnote: 0, endnote: 0 };
    const numbers = new Map();
    for (const ref of root.querySelectorAll('.note-ref[data-note-ref]')) {
      if (ref.closest('.docx-notes')) continue;
      const kind = ref.dataset.noteRef;
      const n = this.noteNumber(kind, ++count[kind]);
      if (ref.dataset.n !== n) ref.dataset.n = n;
      if (!numbers.has(`${kind}:${ref.dataset.noteId}`)) numbers.set(`${kind}:${ref.dataset.noteId}`, n);
    }
    for (const note of root.querySelectorAll('.docx-note')) {
      const n = numbers.get(noteKey(note));
      if (note.hidden !== !n) note.hidden = !n;
      if (!n) delete note.dataset.n; // numbered again if its reference comes back
      if (!n || note.dataset.n === n) continue;
      note.dataset.n = n;
      note.dataset.protected = `${note.dataset.note} ${n}`;
      for (const mark of note.querySelectorAll('.note-mark')) mark.dataset.n = n;
    }
    for (const section of root.querySelectorAll('.docx-notes')) {
      const empty = !section.querySelector('.docx-note:not([hidden])');
      if (section.hidden !== empty) section.hidden = empty;
    }
  }

  /**
   * The note parts as they will be saved, from the notes in the editor content `root`: for each
   * kind, the part's XML with the notes that changed rewritten, or null if none changed (the part
   * is then written back as it was). Notes not shown (no reference in the text) are kept as they are.
   */
  serializeNotes(root) {
    const out = {};
    for (const kind of NOTE_KINDS) {
      const part = this.noteParts[kind];
      if (!part) continue;
      let xml = null;
      const mainRels = this.serialRels;
      this.serialRels = { path: relsPathOf(part.path), rels: part.rels };
      for (const el of root.querySelectorAll(`.docx-note[data-note="${kind}"]`)) {
        const added = el.dataset.added != null;
        if (added && el.hidden) continue; // a note added here whose reference went: not kept
        const nodes = this.serializeBlocks(el);
        const s = new XMLSerializer();
        if (!added && nodes.map((n) => s.serializeToString(n)).join('') === this.noteBaseline.get(noteKey(el))) continue;
        xml ||= parseXml(part.text);
        let note = wChildren(xml.documentElement, kind).find((n) => wAttr(n, 'id') === el.dataset.noteId);
        if (!note && added) {
          note = xml.documentElement.appendChild(createW(xml, kind));
          note.setAttributeNS(W, 'w:id', el.dataset.noteId);
        }
        if (!note) continue;
        while (note.firstChild) note.removeChild(note.firstChild);
        for (const n of nodes) note.appendChild(n); // adopted, so revision ids set later reach them
        if (!nodes.some((n) => n.localName === 'p')) note.appendChild(createW(xml, 'p'));
      }
      this.serialRels = mainRels;
      out[kind] = xml;
    }
    return out;
  }

  /** Writes the note parts that changed; one changed by an earlier save and no longer gets its original back. */
  writeNotes() {
    for (const kind of NOTE_KINDS) {
      const part = this.noteParts[kind];
      if (!part) continue;
      const xml = this.notesOut?.[kind];
      if (part.added) continue; // see writeNewParts
      if (xml) this.zip.file(part.path, serializeXml(xml));
      else if (part.written) this.zip.file(part.path, part.text);
      part.written = !!xml;
    }
  }

  /** The note part at `path` as it will be saved, if its notes changed, else null. */
  notesOutAt(path) {
    const kind = NOTE_KINDS.find((k) => this.noteParts[k]?.path === path);
    return (kind && this.notesOut?.[kind]) || null;
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
    if (kind === 'locked') {
      const info = this.lockedChanges(el);
      return info && { kind: 'locked', mark: false, own: false, author: info.authors.join(', '), ...info };
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
    if (item.kind === 'locked') return this.resolveLocked(el, accept);
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

  // ---------------------------------------------------------------- tracked changes in locked content

  /**
   * Content kept as it is (a locked or data-bound field, a locked cell, a field, a preserved block)
   * holding tracked changes: { count, authors, what } for "the 2 tracked changes in the locked field
   * “Company”", or null.
   */
  lockedChanges(el) {
    const node = el.tagName === 'TD' ? this.cellXml(el) : this.raw[Number(el.dataset.raw)];
    if (!node) return null;
    const count = countRevisions(node);
    if (!count) return null;
    let what;
    if (el.tagName === 'TD') what = 'a locked cell';
    else if (node.localName === 'sdt' && node.namespaceURI === W) {
      const info = controlInfo(node);
      const kind = info.readOnly === 'locked in Word' ? 'the locked field' : info.readOnly?.startsWith('linked') ? 'the field linked to data' : 'the field';
      what = `${kind}${info.label ? ` “${info.label}”` : ''}`;
    } else {
      const name = describe(node);
      what = name.startsWith('Content Intact') ? 'content kept as it is' : `the ${name.replace(/^[A-Z](?=[a-z])/, (c) => c.toLowerCase())}`;
    }
    return { count, authors: revisionAuthors(node), what };
  }

  /**
   * Accepts or rejects every tracked change in locked content at once (as Word's Accept All in it
   * would), in its XML; the content is rendered again, still locked.
   */
  resolveLocked(el, accept) {
    if (el.tagName === 'TD') {
      const table = el.closest('table');
      const r = this.tableRows(table).indexOf(el.parentElement);
      const c = [...el.parentElement.cells].indexOf(el);
      this.editTable(table, (tbl) => {
        const tc = cellsOf(rowsOf(tbl)[r])[c];
        const copy = tc.cloneNode(true); // rendered again
        tc.replaceWith(copy);
        resolveRevisionsIn(copy, accept);
      });
      return;
    }
    const node = resolveRevisionsIn(this.raw[Number(el.dataset.raw)].cloneNode(true), accept);
    if (!node) {
      el.remove();
      return;
    }
    let fresh;
    if (node.localName === 'sdt' && node.namespaceURI === W) fresh = this.renderControl(node, el.tagName === 'DIV');
    else if (el.tagName === 'DIV') fresh = this.rawBlock(node);
    else fresh = this.rawInline(node);
    const holder = document.createElement('div');
    holder.append(fresh);
    this.markTracked(holder, false);
    el.replaceWith(fresh);
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
    for (const child of elementChildren(container)) this.renderBlock(child, parent);
  }

  renderBlock(child, parent) {
    const local = child.namespaceURI === W ? child.localName : null;
    if (local === 'p') parent.append(this.renderParagraph(child));
    else if (local === 'tbl') parent.append(this.renderTable(child));
    else if (local === 'sectPr' || local === 'tblPr' || local === 'tcPr' || local === 'trPr') return;
    else if (local === 'sdt') parent.append(this.renderControl(child, true));
    else parent.append(this.rawBlock(child));
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
    // Right to left, from the paragraph or its style: the browser lays out and edits it so.
    const own = onOff(wChild(pPr, 'bidi'));
    const rtl = own ?? this.styles.chain(styleId || this.styles.defaultParagraph).some((s) => onOff(wChild(s.pPr, 'bidi')) === true);
    if (rtl) el.dir = 'rtl';
    // For line numbers (editor/lineNumbers.js): where a section ends, where Word starts a page, and
    // paragraphs that don't count.
    const chainHas = (local) => {
      const v = onOff(wChild(pPr, local));
      return v ?? this.styles.chain(styleId || this.styles.defaultParagraph).some((s) => onOff(wChild(s.pPr, local)) === true);
    };
    // Its value is the kind of break (nextPage unless the section says otherwise), for printing.
    if (wChild(pPr, 'sectPr')) el.dataset.sectionEnd = wAttr(wChild(wChild(pPr, 'sectPr'), 'type'), 'val') || 'nextPage';
    if (chainHas('pageBreakBefore')) el.dataset.pageBreakBefore = '';
    // For the page estimate (editor/pages.js): a paragraph kept with the next, or kept whole.
    if (chainHas('keepNext')) el.dataset.keepNext = '';
    if (chainHas('keepLines')) el.dataset.keepLines = '';
    if (chainHas('suppressLineNumbers')) el.dataset.noLineNumbers = '';
    if (pPr) {
      el.dataset.ppr = this.keep(pPr);
      Object.assign(el.style, cssToStyle(this.styles.pPrCss(pPr, rtl)));
      // A drop cap is a paragraph of its own, in a frame the next paragraphs flow around.
      const frame = wChild(pPr, 'framePr');
      const dropCap = wAttr(frame, 'dropCap');
      if (dropCap === 'drop' || dropCap === 'margin') {
        el.classList.add('drop-cap');
        el.dataset.dropCap = dropCap;
        el.style.paddingRight = `${(Number(wAttr(frame, 'hSpace')) || 60) / 20}pt`;
      }
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
      // A wrapper's own properties (smartTagPr, customXmlPr) stay with its shell (see data-wrap).
      if (local === 'pPr' || local === 'proofErr' || local === 'smartTagPr' || local === 'customXmlPr') continue;
      const mark = local === 'r' && fieldMarker(child);
      if (mark) {
        const marker = this.hiddenRaw(child);
        const form = mark.type === 'begin' ? formFieldInfo(child) : null;
        if (form) this.showFormField(marker, form);
        parent.append(marker);
        if (mark.type === 'begin') fields.push({ instr: '', outer: parent, form });
        else if (mark.type === 'instr' && fields.length) fields[fields.length - 1].instr += mark.text;
        else if (mark.type === 'separate' && fields.length) {
          fields[fields.length - 1].separated = true;
          const wrap = document.createElement('span');
          wrap.className = 'field';
          const { instr, form: ff } = fields[fields.length - 1];
          wrap.title = fieldTitle(instr);
          if (ff?.kind === 'text') {
            // A text form field: its result is its value, typed in (alone, while only filling in
            // forms is allowed: see the editor's applyLock).
            wrap.classList.add('ff-text');
            wrap.title = formFieldTitle(ff);
            if (!ff.enabled) wrap.dataset.ffOff = '';
          } else if (ff) {
            // A check box's or list's result, if a file has one, isn't what Word shows: the marker is.
            wrap.classList.add('ff-shadow');
            wrap.contentEditable = 'false';
          }
          wrap.dataset.instr = instr.trim(); // for updating it here (see fieldUpdates.js)
          // A page number: each page's own, which Word works out (and printing here does). Shown
          // as # rather than the 1 it holds, and the page count as about how many there are
          // (the editor sets data-show); the result in the file isn't changed, nor typed into.
          const pageKind = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(instr)?.[1].toUpperCase();
          if (pageKind && !ff) {
            wrap.classList.add('page-field');
            wrap.dataset.show = pageKind === 'PAGE' ? '#' : '';
            wrap.contentEditable = 'false';
          }
          const ref = referenceManager(instr);
          if (ref && ref.source !== 'Word') wrap.dataset.refManager = `${ref.source} ${ref.kind}`;
          parent.append(wrap);
          parent = wrap;
        } else if (mark.type === 'end' && fields.length) {
          const field = fields.pop();
          // A page number without a result (written by software other than Word, which works
          // it out on opening): its end marks it, to show # and print numbered (see markBarePageField).
          if (!field.separated && !field.form) markBarePageField(marker, field.instr);
          parent = field.outer;
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
        // Smart tags, custom XML, and right-to-left embeddings (w:dir) and overrides (w:bdo) wrap
        // runs: their runs are edited as text, inside a span that keeps the wrapper (and its
        // properties) to write around them again (see serializeInline).
        const span = document.createElement('span');
        span.className = 'docx-wrap';
        const shell = child.cloneNode(false);
        for (const pr of elementChildren(child)) if (/Pr$/.test(pr.localName)) shell.append(pr.cloneNode(true));
        span.dataset.wrap = this.keep(shell);
        const dir = wAttr(child, 'val');
        if ((local === 'dir' || local === 'bdo') && (dir === 'rtl' || dir === 'ltr')) {
          span.dir = dir;
          if (local === 'bdo') span.style.unicodeBidi = 'bidi-override';
        }
        this.renderInlines(child, span);
        parent.append(span);
      } else {
        parent.append(this.rawInline(child));
      }
    }
  }

  /**
   * Shows a legacy check box or drop-down list on its begin marker (see formFields.js): the marker
   * stays a hidden marker to the editor (and field updates), and the page draws the box or the
   * chosen entry from its data-ff-* attributes (CSS ::before), so no text is added to the paragraph.
   */
  showFormField(marker, info) {
    if (info.kind === 'text') return;
    marker.classList.add(info.kind === 'checkbox' ? 'ff-check' : 'ff-drop');
    marker.title = formFieldTitle(info);
    marker.setAttribute('aria-label', marker.title);
    if (!info.enabled) marker.dataset.ffOff = '';
    if (info.kind === 'checkbox') {
      marker.dataset.checked = info.checked ? '1' : '0';
      marker.setAttribute('role', 'checkbox');
      marker.setAttribute('aria-checked', String(info.checked));
    } else {
      // Word shows five spaces' width for a list with no entries.
      marker.dataset.ffText = info.entries[info.selected] ?? '\u2002'.repeat(5);
      marker.setAttribute('role', 'button');
      marker.setAttribute('aria-haspopup', 'listbox');
    }
    marker.tabIndex = -1; // focused by Tab while only filling in forms is allowed
  }

  /** The legacy form field a rendered begin marker (.ff-check, .ff-drop) starts. */
  formField(marker) {
    return formFieldInfo(this.raw[Number(marker.dataset.raw)]);
  }

  /**
   * Changes a legacy check box or drop-down list: `change(run)` changes a copy of its begin marker,
   * which replaces it (so Undo stays exact), and it is shown again.
   */
  changeFormField(marker, change) {
    const run = this.raw[Number(marker.dataset.raw)].cloneNode(true);
    change(run);
    marker.dataset.raw = this.keep(run);
    this.showFormField(marker, formFieldInfo(run));
  }

  toggleFormCheck(marker) {
    const on = !this.formField(marker).checked;
    this.changeFormField(marker, (run) => setFormCheck(run, on));
  }

  chooseFormEntry(marker, index) {
    this.changeFormField(marker, (run) => setFormDropdown(run, index));
  }

  hiddenRaw(node) {
    const span = document.createElement('span');
    span.contentEditable = 'false';
    span.dataset.raw = this.keep(node);
    span.className = 'raw-inline raw-hidden';
    // Comments in notes are the comments panel's (see doc.story): their markers are only kept.
    if (!this.inNote) markCommentRange(span, node);
    return span;
  }

  rawInline(node) {
    const local = node.localName;
    if (HIDDEN_INLINE.has(local)) return this.hiddenRaw(node);
    const commentRef = local === 'r' && onlyChild(node, 'commentReference');
    if (commentRef && this.inNote) return this.hiddenRaw(node);
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
    // A whole page-number field in one run (begin, instruction, end), as some software writes it:
    // kept as it is, shown and printed as a page number Word hasn't numbered yet.
    const bare = local === 'r' && bareFieldInstr(node);
    if (bare != null) {
      const span = this.hiddenRaw(node);
      if (markBarePageField(span, bare)) return span;
    }
    const mark = local === 'r' && watermarkInfo(node);
    if (mark) {
      // A watermark: the page draws it over each page (see the editor's watermark layer).
      const span = this.hiddenRaw(node);
      span.classList.add('watermark');
      span.dataset.watermark = mark.text;
      span.dataset.diagonal = mark.diagonal ? '1' : '';
      span.dataset.color = mark.color;
      span.dataset.font = mark.font;
      span.title = `Watermark “${mark.text}” – change it in Page setup`;
      return span;
    }
    const box = local === 'r' && textboxInfo(node, this.styles.themeColors);
    if (box) return this.renderTextbox(node, box);
    const span = document.createElement('span');
    span.contentEditable = 'false';
    span.dataset.raw = this.keep(node);
    span.className = 'raw-inline';
    // Images (DrawingML or legacy VML)
    const blips = [...node.getElementsByTagNameNS(NS.a, 'blip'), ...node.getElementsByTagNameNS(NS.v, 'imagedata')];
    // A line, an arrow, a rectangle or an ellipse without text: drawn, and laid out as a picture is.
    const shape = local === 'r' && !blips.length && pictureInfo(node) && shapeOf(node, this.styles.themeColors);
    if (shape) {
      span.classList.add('picture', 'shape');
      span.dataset.shape = shape.prst;
      span.title = `${/line|Connector/i.test(shape.prst) ? (shape.head || shape.tail ? 'Arrow' : 'Line') : 'Shape'} – click to select it: drag it to move it, a handle to resize it, or use Layout Options`;
      this.stylePicture(span);
      return span;
    }
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
      const id = wAttr(noteRef, 'id');
      this.noteOrder?.[kind].push(id);
      span.classList.add('note-ref');
      span.dataset.n = this.noteNumber(kind, n);
      const text = this.notes[kind].get(id);
      if (this.noteParts[kind] && this.noteOrder) {
        span.dataset.noteRef = kind;
        span.dataset.noteId = id;
        span.title = `${NOTE_NAMES[kind]}${text ? `: ${text}` : ''} – click to edit it (${kind}s are below the text)`;
      } else {
        span.title = `${NOTE_NAMES[kind]} ${span.dataset.n}${text ? `: ${text}` : ''} – edit it in Word`;
      }
      return span;
    }
    const noteMark = this.inNote && local === 'r' && (wChild(node, 'footnoteRef') || wChild(node, 'endnoteRef'));
    if (noteMark) {
      // The note's own number, at its start: Word numbers it, so it stays (and says where it points).
      span.classList.add('note-ref', 'note-mark');
      span.dataset.n = this.noteMarkN;
      span.dataset.protected = 'the note’s number';
      span.title = 'The note’s number – Word numbers notes; click it to go to the text';
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
    if (isEquation(node) && this.renderEquation(span, node)) return span;
    span.textContent = plainText(node) || '⧉';
    return span;
  }

  /**
   * Shows an equation as MathML in `span` (on its own line for an equation paragraph, m:oMathPara).
   * False if it can't be drawn, so it shows as text instead. Its XML is kept as it was either way.
   */
  renderEquation(span, node) {
    try {
      const maths = ommlToMathml(node);
      if (!maths.length) return false;
      span.classList.add('equation');
      if (maths.some((m) => m.getAttribute('display') === 'block')) span.classList.add('display');
      // One the Equation dialog can write back (as LaTeX) is edited there; the rest stay Word's.
      if (ommlToLatex(node)) {
        span.dataset.editable = '1';
        span.title = 'Equation – double-click to edit it';
      }
      span.append(...maths);
      return true;
    } catch {
      return false;
    }
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

  renderRun(r, parent) {
    const kids = elementChildren(r);
    const isSimple = (k) => k.namespaceURI === W && SIMPLE_RUN_CHILDREN.has(k.localName)
      && !(k.localName === 'br' && wAttr(k, 'type') && wAttr(k, 'type') !== 'textWrapping');
    const simple = kids.every(isSimple);
    const pieces = simple ? [] : runPieces(kids, isSimple);
    if (pieces.length > 1) {
      // Pictures, text boxes or objects with other content in one run (as LibreOffice writes them;
      // Word gives each its own): shown as a run each, with the run's properties, which means the
      // same. So each is drawn, and the text beside them stays text.
      const rPr = wChild(r, 'rPr');
      for (const piece of pieces) {
        const run = r.cloneNode(false);
        if (rPr) run.append(rPr.cloneNode(true));
        for (const k of piece) run.append(k.cloneNode(true));
        this.renderRun(run, parent);
      }
      return;
    }
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
        case 'lastRenderedPageBreak': {
          // Where Word began a page when it last saved: a character offset in the paragraph.
          const block = target.closest?.('p, h1, h2, h3, h4, h5, h6');
          if (block) block.dataset.pageBreaks = [block.dataset.pageBreaks, block.textContent.length + text.length].filter((v) => v != null && v !== '').join(',');
          break;
        }
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
    // A right-to-left run (w:rtl): its characters, digits and punctuation run right to left.
    if (onOff(wChild(rPr, 'rtl'))) span.dir = 'rtl';
    // Hidden text (w:vanish): shown faintly here, listed in the review before release.
    if (onOff(wChild(rPr, 'vanish'))) span.dataset.hidden = '';
    // The run's language, for the browser's spelling check (and hyphenation, screen readers).
    const lang = wAttr(wChild(rPr, 'lang'), 'val');
    if (lang && /^[a-z]{2,3}(-[A-Za-z0-9]{2,8})*$/i.test(lang)) span.lang = lang;
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
    // A right-to-left table: its first column at the right.
    if (onOff(wChild(wChild(tbl, 'tblPr'), 'bidiVisual'))) table.dir = 'rtl';
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
    // The row's height (w:trHeight, twips): at least that high, as Word lays it out (an exact
    // height is shown as a least one: the text is never cut off here). Only shown; the row's
    // properties are kept as they are.
    const height = Number(wAttr(wChild(wChild(tr, 'trPr'), 'trHeight'), 'val'));
    if (height > 0 && wAttr(wChild(wChild(tr, 'trPr'), 'trHeight'), 'hRule') !== 'auto') row.style.height = `${height / 15}px`;
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
    // Cells merged across the row with w:hMerge: the first spans them, the others are hidden (but
    // kept, as in the file, so rendered cells still match the row's cells).
    let restart = null;
    cellsOf(tr).forEach((tc, j) => {
      const td = row.cells[j];
      if (!td) return;
      const h = hMerge(tc);
      if (h === 'continue' && restart) {
        restart.colSpan += td.colSpan || 1;
        td.style.display = 'none';
        td.classList.add('hmerge');
      } else {
        td.style.display = '';
        td.classList.remove('hmerge');
        restart = h === 'restart' ? td : null;
      }
    });
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
    // The cell's own borders, over its table's: a side set to none ("nil") hides the table's.
    const borders = cellBorders(tc);
    for (const side of ['top', 'left', 'bottom', 'right']) {
      const prop = `border${side[0].toUpperCase()}${side.slice(1)}`;
      td.style[prop] = borders[side] ? borderCss(borders[side]) || 'none' : '';
    }
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
    // A tracked merge across a row is a tracked change of the cells' properties that sets w:hMerge.
    const merge = wChild(tcPr, 'cellMerge') || (isTrackedHMerge(tc) ? wChild(tcPr, 'tcPrChange') : null);
    if (merge) {
      // A tracked merge or split down a column: its cells share a key (author and time).
      td.dataset.revMerge = merge.localName === 'tcPrChange' || wAttr(merge, 'vMerge') === 'cont' ? 'merge' : 'split';
      td.dataset.revMergeKey = `${wAttr(merge, 'author') || ''}|${wAttr(merge, 'date') || ''}`;
      if (!rev) td.title = `Cells ${td.dataset.revMerge === 'merge' ? 'merged' : 'split'}${byWhom({ author: wAttr(merge, 'author'), date: wAttr(merge, 'date') })} – tracked change: Accept or Reject on the toolbar`;
    } else if (td.dataset.revMerge) {
      delete td.dataset.revMerge;
      delete td.dataset.revMergeKey;
      if (!rev) td.removeAttribute('title');
    }
    const change = !isTrackedHMerge(tc) && propertyChange(tc);
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
   * A picture from a file (PNG, JPEG, GIF, BMP, or SVG with a PNG beside it) as Word inserts one: in line with the text, at its
   * own size or narrowed to fit between the margins. Its file goes into word/media when saved (if
   * it is still in the document). Returns { span } (the rendered picture) or { reason }.
   */
  /**
   * A new check box at the caret, as Word's Developer tab inserts one: a content control with
   * w14:checkbox (☐ / ☒ in MS Gothic, as Word writes it), rendered as an editable control.
   */
  newCheckbox() {
    const used = new Set([...this.xml.getElementsByTagNameNS(W, 'id')].map((e) => wAttr(e, 'val')));
    let id;
    do id = String(Math.floor(Math.random() * 2 ** 31) - 2 ** 30); while (used.has(id));
    const font = '<w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/>';
    const xml = `<w:sdt xmlns:w="${W}" xmlns:w14="${W14}"><w:sdtPr><w:id w:val="${id}"/>`
      + '<w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>'
      + `</w:sdtPr><w:sdtContent><w:r><w:rPr>${font}</w:rPr><w:t>\u2610</w:t></w:r></w:sdtContent></w:sdt>`;
    const sdt = this.xml.importNode(parseXml(xml).documentElement, true);
    for (const a of [...sdt.attributes]) if (a.name.startsWith('xmlns')) sdt.removeAttributeNode(a); // declared where needed when saved
    return this.renderControl(sdt, false);
  }

  /**
   * A new picture from `file`, in line with the text: at its size (or `width`, in pixels at 96 per
   * inch, keeping its proportions), narrowed to fit between the margins, with alt text `descr`.
   */
  async addPicture(file, { width, descr, part = null } = {}) {
    const type = file.type || (/\.svg$/i.test(file.name || '') ? 'image/svg+xml' : '');
    const ext = PICTURE_TYPES[type];
    if (!ext) return { reason: `${file.name ? `“${file.name}”` : 'This file'} isn’t a picture Word shows everywhere: insert a PNG, JPEG, GIF, BMP or SVG picture` };
    let data = new Uint8Array(await file.arrayBuffer());
    const svg = type === 'image/svg+xml';
    let url = URL.createObjectURL(new Blob([data], { type }));
    const img = new Image();
    let size = await new Promise((resolve) => {
      img.onload = () => resolve(img.naturalWidth ? { w: img.naturalWidth, h: img.naturalHeight } : null);
      img.onerror = () => resolve(null);
      img.src = url;
    });
    // An SVG: shown (and kept for readers before Office 2016) as a PNG drawn from it here, at
    // twice its size for a sharp print, the SVG itself beside it, as Word inserts one.
    let svgData = null;
    if (svg && size) {
      const png = await svgToPng(img, size);
      if (!png) size = null;
      else {
        URL.revokeObjectURL(url);
        svgData = data;
        data = png;
        url = URL.createObjectURL(new Blob([png], { type: 'image/png' }));
      }
    }
    if (!size) {
      URL.revokeObjectURL(url);
      return { reason: `${file.name ? `“${file.name}”` : 'This picture'} couldn’t be read${svg ? ' (an SVG needs a width and height, or a viewBox)' : ''}` };
    }
    const maxWidth = this.textWidth / 15;
    let { w, h } = size;
    if (width) {
      h = Math.max(1, Math.round((h * width) / w));
      w = width;
    }
    if (w > maxWidth) {
      h = Math.max(1, Math.round((h * maxWidth) / w));
      w = maxWidth;
    }
    const { rid, n } = this.addMedia(data, svgData ? 'image/png' : type, url, part);
    const svgRid = svgData ? this.addMedia(svgData, 'image/svg+xml', URL.createObjectURL(new Blob([svgData], { type })), part).rid : null;
    const blipExt = svgRid ? `<a:extLst><a:ext uri="${SVG_EXT}"><asvg:svgBlip xmlns:asvg="${ASVG}" r:embed="${svgRid}"/></a:ext></a:extLst>` : '';
    const id = await this.nextDocPrId();
    const cx = Math.round(w * 9525);
    const cy = Math.round(h * 9525);
    const name = (file.name || `image${n}.${svgData ? 'svg' : ext}`).replace(/[<>&"]/g, '');
    const xml = `<w:r xmlns:w="${W}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${PIC}" xmlns:r="${NS.r}"><w:drawing>`
      + `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
      + `<wp:docPr id="${id}" name="Picture ${id}"${descr ? ` descr="${descr.replace(/[<>&"]/g, '')}"` : ''}/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>`
      + `<a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>`
      + `<pic:blipFill><a:blip r:embed="${rid}">${blipExt}</a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
      + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
      + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const r = this.xml.importNode(parseXml(xml).documentElement, true);
    for (const a of [...r.attributes]) if (a.name.startsWith('xmlns')) r.removeAttributeNode(a); // declared where needed when saved
    // In a header or footer, drawn with that part's relationships, where its picture file is.
    let span = null;
    if (part) this.inPart(part, () => { span = this.rawInline(r); });
    else span = this.rawInline(r);
    return { span };
  }

  /** The header or footer part (as in this.headers, this.footers) a rendered element is in, or null. */
  partOf(el) {
    const path = el?.closest?.('.docx-hf-part')?.dataset.part;
    return (path && [...this.headers, ...this.footers].find((v) => v.path === path)) || null;
  }

  /**
   * A picture file added this session (`data`, of MIME `type`; `url` shows it): a new file in
   * word/media and a relationship to it, written on save if a picture uses it. Returns { rid, path, n }.
   */
  addMedia(data, type, url = URL.createObjectURL(new Blob([data], { type })), part = null) {
    const ext = PICTURE_TYPES[type] || 'png';
    let n = 1;
    const taken = (path) => this.zip.file(path) || [...this.newMedia.values()].some((m) => m.path === path);
    while (taken(`${this.mainDir ? this.mainDir + '/' : ''}media/image${n}.${ext}`)) n++;
    const path = `${this.mainDir ? this.mainDir + '/' : ''}media/image${n}.${ext}`;
    // Relative to the part that uses it: a header or footer next to the main part (as Word puts
    // them) shares its media folder; one elsewhere points at the file from the package's root.
    const owner = part ? part.path : this.mainPath;
    const ownerDir = owner.includes('/') ? owner.slice(0, owner.lastIndexOf('/')) : '';
    const target = ownerDir === (this.mainDir || '') ? `media/image${n}.${ext}` : `/${path}`;
    const rels = part ? part.rels : this.rels;
    const relsPath = part ? relsPathOf(part.path) : this.relsPath;
    let k = 1;
    while (rels.has(`rIdImage${k}`)) k++;
    const rid = `rIdImage${k}`;
    rels.set(rid, { type: IMAGE_REL, target, external: false, path });
    this.mediaUrls.set(path, url);
    this.newMedia.set(`${relsPath}|${rid}`, { rid, relsPath, part: part?.path || null, path, target, data, ext, type });
    return { rid, path, n };
  }

  /**
   * The address the editor shows a picture file (by relationship id) at, or null: in the part
   * `span` is in (a header or footer has relationships of its own), else in the part being drawn.
   */
  mediaUrl(rid, span = null) {
    const rels = this.partOf(span)?.rels || this.rels;
    const rel = rid && rels.get(rid);
    return (rel && this.mediaUrls.get(rel.path)) || null;
  }

  /**
   * A new line, arrow or text box (see pictures.js newShapeRun), rendered: a line or arrow in front of
   * the text, laid out as a picture is; a text box in line with the text, its text edited here.
   */
  async addShape(kind, opts) {
    const id = await this.nextDocPrId();
    const r = this.xml.importNode(parseXml(newShapeRun(kind, { id, ...opts })).documentElement, true);
    for (const a of [...r.attributes]) if (a.name.startsWith('xmlns')) r.removeAttributeNode(a); // declared where needed when saved
    return this.rawInline(r);
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
    // Used: shown by a picture, or kept as the original of a picture with an artistic effect, in
    // the part the picture was put in (a header or footer as it will be saved, if it changed).
    const embeds = (root) => new Set(root ? [...root.getElementsByTagName('*')].map((el) => el.getAttributeNS(NS.r, 'embed')).filter(Boolean) : []);
    const usedIn = new Map();
    const used = (m) => {
      // A footer added here and taken out again before saving isn't written, nor its pictures.
      if (m.part && [...this.headers, ...this.footers].some((v) => v.path === m.part && v.added && !v.present)) return false;
      const key = m.part || '';
      if (!usedIn.has(key)) usedIn.set(key, embeds(m.part ? this.partsOut?.[m.part]?.documentElement : this.body));
      return usedIn.get(key).has(m.rid);
    };
    const ct = parseXml(await this.text('[Content_Types].xml'));
    const relsDocs = new Map();
    let ctChanged = false;
    for (const m of this.newMedia.values()) {
      if (!relsDocs.has(m.relsPath)) relsDocs.set(m.relsPath, { xml: parseXml((await this.text(m.relsPath)) || `<Relationships xmlns="${NS.pr}"/>`), changed: false });
      const entry = relsDocs.get(m.relsPath);
      const rels = entry.xml;
      const { rid } = m;
      const rel = [...rels.documentElement.children].find((r) => r.getAttribute('Id') === rid);
      if (used(m)) {
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
          ctChanged = true;
        }
        entry.changed = true;
      } else if (rel) {
        this.zip.remove(m.path);
        rel.remove();
        entry.changed = true;
      }
    }
    for (const [path, { xml, changed }] of relsDocs) {
      if (!changed) continue;
      // A part that had no relationships before gets none back once its pictures are gone.
      if (!xml.documentElement.firstElementChild && !this.openedEntries.has(path)) this.zip.remove(path);
      else this.zip.file(path, serializeXml(xml));
    }
    if (ctChanged) this.zip.file('[Content_Types].xml', serializeXml(ct));
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
    const w = info.cx / EMU_PER_PX;
    const h = info.cy / EMU_PER_PX;
    span.dataset.wrap = info.wrap;
    span.removeAttribute('style');
    this.placePicture(span, info, w, h);
    if (span.dataset.shape != null) {
      drawShape(span, { w, h, look, shape: shapeOf(this.raw[Number(span.dataset.raw)], this.styles.themeColors), inline: info.wrap === 'inline' });
      return;
    }
    drawPicture(span, { w, h, look, alt: info.descr, url: this.mediaUrl(blipOf(this.raw[Number(span.dataset.raw)]), span), inline: info.wrap === 'inline' });
  }

  /** Where a picture goes: in line, floating left or right, on a line of its own, or behind or in front of the text. */
  placePicture(span, info, w, h) {
    if (info.wrap === 'inline') return;
    const st = span.style;
    const page = this.page;
    const content = page.width - page.left - page.right;
    const px = (emu) => (emu || 0) / EMU_PER_PX;
    // Horizontal place, measured from the left margin, where the picture's paragraph (its containing
    // block: see .docx-sheet p in the editor's styles) begins. Word measures it from what relativeFrom
    // names: the margins or column (the text's width), the page, or the left or right margin itself.
    const { h: hp, v: vp } = info;
    const [start, width] = {
      page: [-page.left, page.width],
      leftMargin: [-page.left, page.left],
      insideMargin: [-page.left, page.left],
      rightMargin: [content, page.right],
      outsideMargin: [content, page.right],
    }[hp.relativeFrom] || [0, content]; // margin, column, character
    let left = start;
    if (hp.align === 'center') left = start + (width - w) / 2;
    else if (hp.align === 'right' || hp.align === 'outside') left = start + width - w;
    else if (hp.align == null) left = start + px(hp.offset);
    const fromParagraph = !vp.relativeFrom || ['paragraph', 'line'].includes(vp.relativeFrom);
    const top = vp.align == null ? px(vp.offset) : 0;
    if (info.wrap === 'behind' || info.wrap === 'front') {
      st.position = 'absolute';
      // From the paragraph, which starts at the margin; never starting past the page's right edge
      // (--hf-right: the right margin as shown), which only matters where the page is narrower than
      // the paper, on a phone: there the picture comes in to the edge rather than off the page.
      st.left = `min(${left}px, calc(100% + var(--hf-right, ${page.right}px) - ${w}px))`;
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

  setPictureGeometry(span, prst) {
    this.changePicture(span, (run) => setGeometry(run, prst));
  }

  setPicture3D(span, format) {
    this.changePicture(span, (run) => set3D(run, format));
  }

  /** Applies a picture style (see pictures.js PICTURE_STYLES), or takes it off with null. */
  setPictureStyle(span, style) {
    this.changePicture(span, (run) => applyPictureStyle(run, style));
  }

  /** The picture file a picture shows without its artistic effect: its relationship id and address. */
  pictureOriginal(span) {
    const run = this.raw[Number(span.dataset.raw)];
    const rid = pictureLook(run).artistic?.original || blipOf(run);
    return { rid, url: this.mediaUrl(rid, span) };
  }

  /**
   * An artistic effect on a picture: `png` (bytes) is the picture drawn with it, added as a new
   * picture file; the original is kept (see pictures.js setArtistic). `effect` null takes it off.
   */
  setPictureArtistic(span, effect, png = null, amount = null) {
    const { rid: original } = this.pictureOriginal(span);
    if (!effect) {
      this.changePicture(span, (run) => clearArtistic(run));
      return;
    }
    const { rid } = this.addMedia(png, 'image/png', undefined, this.partOf(span));
    this.changePicture(span, (run) => setArtistic(run, { effect, rendered: rid, original, amount }));
  }

  setPictureDescription(span, text) {
    this.changePicture(span, (run) => setDescription(run, text));
  }

  // ---------------------------------------------------------------- table structure

  /** The rendered rows of a table, in order (they stand for its <w:tr> elements, in order). */
  tableRows(table) {
    return [...table.querySelectorAll(':scope > tbody > tr, :scope > tr')];
  }

  /**
   * Why a table's structure can't be changed here, or null. `scope` limits the fields that count to
   * the rows and cells a change touches (see tables.js lockedParts); none, the whole table.
   */
  tableLocked(table, scope = undefined) {
    if (table?.dataset.raw == null) return 'This table can’t be changed here: change it in Word';
    return lockedParts(this.raw[Number(table.dataset.raw)], scope);
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
    const reason = this.tableLocked(table, { rows: [index, index] }); // the row copied, and where the copies go
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
    if (res.reason) return { row: null, reason: res.reason };
    return { row: res.rows[0], reason: null };
  }

  /**
   * Deletes rendered rows `from` to `to`. While tracking they are marked deleted (rows inserted
   * this session go); otherwise they go, and so does a table left with no rows. Returns
   * { removed } (the table went) or { reason }.
   */
  deleteTableRows(table, from, to, tracked = null) {
    const reason = this.tableLocked(table, { rows: [from, to] });
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
    const at = right ? start + span : start;
    // Every row gets a cell (so no row may be a locked field), and cells across the place widen.
    const reason = this.tableLocked(table, { rows: null, cells: (c) => c.start < at && at < c.start + c.span });
    if (reason) return { reason };
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
    const reason = this.tableLocked(table, { rows: null, cells: (c) => c.start <= to && c.start + c.span - 1 >= from });
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
  /**
   * Shading or borders for the cells in rows `r0`–`r1` and grid columns `c0`–`c1` (grown to whole
   * cells), as Word's Shading and Borders buttons set them, tracked while tracking: `kind` is
   * 'shading' (`value` a colour RRGGBB, or null for none) or 'borders' (`value` 'all', 'outside',
   * 'bottom' or 'none'). Returns { reason } if they can't be changed.
   */
  formatTableCells(table, [r0, r1], [c0, c1], kind, value, tracked = null) {
    if (table?.dataset.raw == null) return { reason: 'This table can’t be changed here: change it in Word' };
    const tbl = this.raw[Number(table.dataset.raw)];
    const area = mergeArea(tbl, r0, r1, c0, c1) || { r0, r1, c0, c1 };
    const reason = this.tableLocked(table, {
      rows: [area.r0, area.r1],
      cells: (c, r) => r >= area.r0 && r <= area.r1 && c.start <= area.c1 && c.start + c.span - 1 >= area.c0,
    });
    if (reason) return { reason };
    const change = kind === 'shading'
      ? (tc) => setShading(tc, value)
      : (tc, edges) => setBorders(tc, {
        all: { top: true, left: true, bottom: true, right: true },
        none: { top: false, left: false, bottom: false, right: false },
        outside: Object.fromEntries(Object.entries(edges).filter(([, on]) => on)),
        bottom: edges.bottom ? { bottom: true } : {},
      }[value]);
    const rev = tracked && this.rowRevisionMaker(tracked);
    this.editTable(table, (t) => formatCells(t, area, change, rev));
    return {};
  }

  mergeTableCells(table, [r0, r1], [c0, c1], tracked = null) {
    if (table?.dataset.raw == null) return { reason: 'This table can’t be changed here: change it in Word' };
    const tbl = this.raw[Number(table.dataset.raw)];
    const area = mergeArea(tbl, r0, r1, c0, c1);
    if (!area) return { reason: 'These cells can’t be merged: select a rectangle of cells' };
    const reason = this.tableLocked(table, {
      rows: [area.r0, area.r1],
      cells: (c, r) => r >= area.r0 && r <= area.r1 && c.start <= area.c1 && c.start + c.span - 1 >= area.c0,
    });
    if (reason) return { reason };
    const cells = areaCells(tbl, area);
    if (cells.flat().length < 2) return { reason: 'Select the cells to merge: more than one' };
    const domRows = this.tableRows(table);
    const tds = cells.map((row, i) => row.map((c) => domRows[area.r0 + i].cells[c.index]));
    if (tds.flat().some((td) => !td || td.dataset.revCell || td.dataset.revMerge || td.classList.contains('hmerge')) || domRows.slice(area.r0, area.r1 + 1).some((r) => r.dataset.revRow)) {
      return { reason: 'These cells are a tracked change: accept or reject it first' };
    }
    // Cells that are fields merge too (the merged cell is the first one's field), unless a field
    // can't be edited here: a locked one, one linked to data, a check box or a list.
    if (tds.flat().some((td) => td.contentEditable === 'false')) {
      return { reason: 'A cell to merge is a field that can’t be changed here (locked, linked to data, or not a text field): merge it in Word' };
    }
    if (tracked && tds.flat().slice(1).some((td) => td.querySelector('[data-tracked], [data-tc], [data-tc-mark], [data-rev-mark]'))) {
      return { reason: 'These cells hold tracked changes: accept or reject them first' };
    }
    const target = tds[0][0];
    // A field still showing its placeholder counts as empty: its placeholder text isn't moved.
    const empty = (td) => td.classList.contains('showing-placeholder') || (!td.textContent.trim() && !td.querySelector('img, .raw-inline, .sdt, table'));
    const clear = (td) => {
      td.replaceChildren();
      td.classList.remove('showing-placeholder');
    };
    // Their content goes into the merged cell, one cell after another; tracked, it is copied in as
    // an insertion and left as a deletion where it was (so rejecting the merge puts it back).
    for (const td of tds.flat().slice(1)) {
      if (empty(td)) continue;
      if (!tracked) {
        if (empty(target)) clear(target);
        target.append(...td.childNodes);
        td.replaceChildren(this.emptyParagraph());
        continue;
      }
      if (empty(target)) clear(target);
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
    if (!target.firstElementChild) target.append(this.emptyParagraph());
    const rev = tracked && this.rowRevisionMaker(tracked);
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
    if (td.dataset.revCell || td.dataset.revMerge || td.parentElement.dataset.revRow) return { reason: 'This cell is a tracked change: accept or reject it first' };
    const rowIndex = this.tableRows(table).indexOf(td.parentElement);
    const index = [...td.parentElement.cells].indexOf(td);
    // The cell and its row; split into columns, the other rows may span new grid columns too.
    const reason = this.tableLocked(table, { rows: cols > 1 ? null : [rowIndex, rowIndex], cells: (c, r) => r === rowIndex && layout(rowsOf(this.raw[Number(table.dataset.raw)])[r]).indexOf(c) === index });
    if (reason) return { reason };
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
    this.editTable(table, (tbl) => {
      const tcs = at.map(([r, i]) => cellsOf(rowsOf(tbl)[r])[i]).filter(Boolean);
      for (const tc of tcs) {
        resolveMerge(tc, accept);
        if (isTrackedHMerge(tc)) resolveCellProps(tc, accept);
      }
      if (accept) normalizeHMerge(tbl); // accepted across a row: one cell spanning the columns, as Word writes it
      fitWidths(tbl, { before: Infinity, textWidth: 0 });
    });
  }

  /** Deletes a whole table: its rows marked deleted while tracking, else the table goes. */
  deleteTable(table, tracked = null) {
    const rows = this.tableRows(table);
    if (!tracked) {
      if (table?.dataset.raw == null) return { reason: 'This table can’t be changed here: change it in Word' };
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
  newTable(rows, cols, tracked = null, { widths = null, merges = [] } = {}) {
    const styleId = this.styles.ensureTableGrid();
    const tbl = newTableXml(this.xml, rows, cols, this.textWidth, styleId, tracked && this.rowRevisionMaker(tracked), widths);
    // Merged cells (a pasted table's), as Word writes a merge: { r0, r1, c0, c1 } each. They are
    // part of the new table, so not tracked apart from its rows.
    for (const area of merges) mergeCells(tbl, area);
    if (merges.length) fitWidths(tbl, { before: Infinity, textWidth: 0 });
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
    this.applyTabStops(root);
    this.renumberNotes(root);
    if (!this.numberingXml && !root.querySelector('.list')) return;
    // Each header and footer numbers its own lists; the body (with its notes) numbers the rest.
    for (const host of root.querySelectorAll('.docx-hf-part')) {
      this.numberBlocks(host.querySelectorAll('p,h1,h2,h3,h4,h5,h6'), new Numbering(this.numberingXml));
    }
    const numbering = new Numbering(this.numberingXml);
    this.numbering = numbering;
    this.numberBlocks([...root.querySelectorAll('p,h1,h2,h3,h4,h5,h6')].filter((el) => !el.closest('.docx-hf-band')), numbering);
  }

  numberBlocks(blocks, numbering) {
    for (const el of blocks) {
      const list = this.listOf(el);
      const num = list && numbering.next(list.numId, list.ilvl);
      // Indentation follows the list level, unless the paragraph keeps its own (as when rendered).
      const edited = this.listEdited(el);
      if (num) {
        el.classList.add('list');
        if (el.dataset.marker !== num.marker) el.dataset.marker = num.marker;
        const pPr = el.dataset.ppr != null ? this.raw[Number(el.dataset.ppr)] : null;
        if ((edited && el.dataset.indSet == null) || !wChild(pPr, 'ind')) this.indentForList(el, num, list.ilvl);
        else el.style.paddingLeft = this.styles.pPrCss(pPr)['padding-left'] || '36pt';
      } else if (el.classList.contains('list')) {
        // Out of the list: back to the paragraph's own indentation (none once its list was edited).
        el.classList.remove('list');
        delete el.dataset.marker;
        if (el.dataset.markerInline != null) {
          delete el.dataset.markerInline;
          el.style.textAlign = '';
        }
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
    // A level centred on its own line ("Pasal 1", "BAB I"): its number in the line, not hanging.
    const jc = wAttr(wChild(num.pPr, 'jc'), 'val');
    if (jc === 'center' && !wChild(this.raw[Number(el.dataset.ppr)], 'jc')) {
      el.dataset.markerInline = '';
      el.style.textAlign = 'center';
      el.style.paddingLeft = '0px';
    } else if (el.dataset.markerInline != null) {
      delete el.dataset.markerInline;
      el.style.textAlign = '';
    }
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
  /** A list level's { fmt ('decimal', 'bullet'…), text ('%1.', '•'…) }, or null. */
  listLevel(numId, ilvl) {
    this.kinds ||= new Numbering(this.numberingXml);
    return this.kinds.levels(numId)?.get(ilvl) || null;
  }

  hasListLevel(numId, ilvl) {
    this.kinds ||= new Numbering(this.numberingXml);
    return ilvl >= 0 && ilvl <= MAX_LEVEL && this.kinds.hasLevel(numId, ilvl);
  }

  /** Puts a block in list `numId` at level `ilvl`. */
  setList(el, numId, ilvl = 0) {
    delete el.dataset.indSet; // the list level's indentation replaces one set in the Paragraph dialog
    el.dataset.numId = numId;
    el.dataset.ilvl = String(Math.max(0, Math.min(MAX_LEVEL, ilvl)));
  }

  /**
   * Takes a block out of its list, as Word does when you press Enter on an empty item: a paragraph
   * whose list came from a list style ("List Bullet", "List Paragraph") goes back to Normal.
   */
  removeList(el) {
    delete el.dataset.indSet;
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

  /**
   * A numId for a new list of `kind` ('bullet' or 'number'); bullets share one list. `first`, for a
   * list started by typing ("a) ", "- "): its first level's look and start (see numbering.addList).
   */
  /** A numId for a new list of a ready-made multilevel kind (numbering.LIST_PRESETS: 'pasal', 'bab', 'outline'). */
  newPresetList(preset) {
    if (!this.numberingXml) this.numberingXml = parseXml(`<w:numbering xmlns:w="${W}"/>`);
    const numId = addPresetList(this.numberingXml, preset, this.listDefs);
    this.numberingChanged = true;
    this.kinds = null;
    return numId;
  }

  /**
   * Numbers the document's headings by their styles, as Word's "1 Heading 1, 1.1 Heading 2…"
   * multilevel list does: a new list whose levels are linked to Heading 1–9, and each heading style
   * numbered by it (Heading 1–3 added if the document lacks them). The paragraphs aren't changed:
   * their styles number them, in Word as here. Returns the heading styles' ids by level.
   */
  numberHeadings() {
    if (!this.numberingXml) this.numberingXml = parseXml(`<w:numbering xmlns:w="${W}"/>`);
    for (const key of ['h1', 'h2', 'h3']) this.styles.ensureBuiltin(key);
    const headings = Array.from({ length: MAX_LEVEL + 1 }, (_, i) => {
      for (const s of this.styles.byId.values()) if (s.type === 'paragraph' && s.name.toLowerCase() === `heading ${i + 1}`) return s.id;
      return null;
    });
    const numId = addHeadingList(this.numberingXml, headings);
    headings.forEach((id, i) => id && this.styles.setStyleNumbering(id, numId, i));
    this.numberingChanged = true;
    this.kinds = null;
    return headings;
  }

  /**
   * Adds a style made from the formatting at `range` in paragraph `block`, as Word's Create New
   * Style from Formatting does: a paragraph style (`type` 'paragraph') takes the paragraph's own
   * formatting and its first run's, based on the paragraph's style; a character style the first
   * selected run's. Lists, a section, a tracked change's record and a style already applied aren't
   * taken. Returns { id, pPr, rPr } (what the style was given, for taking it off the text).
   */
  newStyleFrom(range, block, { name, type }) {
    const holder = block.cloneNode(false);
    const r = range.cloneRange();
    if (type === 'paragraph' || r.collapsed) r.selectNodeContents(block);
    holder.append(r.cloneContents());
    const p = this.paragraphXml(holder);
    // A character style: the run properties of the first selected text, as it is in the page.
    const first = (() => {
      if (type !== 'character') return null;
      const walker = document.createTreeWalker(range.commonAncestorContainer, NodeFilter.SHOW_TEXT);
      for (let n = walker.currentNode.nodeType === 3 ? walker.currentNode : walker.nextNode(); n; n = walker.nextNode()) {
        if (n.data && range.intersectsNode(n)) return n;
      }
      return null;
    })();
    const run = wChildren(p, 'r').find((x) => wChild(x, 't')) || wChildren(p, 'r')[0];
    const found = first ? this.rPrOf(first) : run && wChild(run, 'rPr');
    const rPr = found ? found.cloneNode(true) : null;
    if (rPr) removeW(rPr, 'rStyle', 'rPrChange', 'ins', 'del', 'moveFrom', 'moveTo');
    let pPr = null;
    if (type === 'paragraph' && wChild(p, 'pPr')) {
      pPr = wChild(p, 'pPr').cloneNode(true);
      removeW(pPr, 'pStyle', 'numPr', 'rPr', 'sectPr', 'pPrChange', 'framePr', 'cnfStyle');
    }
    const basedOn = type === 'paragraph' ? (block.dataset.style || this.styles.defaultParagraph) : this.defaultCharacterStyle();
    const id = this.styles.addStyle({ name, type, basedOn, pPr, rPr });
    return { id, pPr, rPr };
  }

  /** True if two properties (<w:b/>, <w:rFonts …/>) say the same, whatever the order of their attributes. */
  sameProperty(a, b) {
    return a.localName === b.localName && this.sameProps(wrapIn(this.xml, a), wrapIn(this.xml, b));
  }

  /**
   * The document's own styles for code, by name, as technical documents have them: { paragraph
   * ("Code", "Code Block", "Source Code", "HTML Preformatted"…), character ("HTML Code", "Inline
   * Code", "Code Char"…) }, each a style id or null.
   */
  codeStyles() {
    const find = (type, re) => [...this.styles.byId.values()].find((st) => st.type === type && re.test(st.name.trim()))?.id || null;
    return {
      paragraph: find('paragraph', /^(code|code ?block|source ?code|source|listing|program listing|html preformatted|preformatted( text)?|code example)$/i),
      character: find('character', /^(html code|inline ?code|code( char)?|source ?code( char)?|code text|html preformatted char)$/i),
    };
  }

  /** Word's Default Paragraph Font (the default character style), or null. */
  defaultCharacterStyle() {
    for (const st of this.styles.byId.values()) if (st.type === 'character' && st.isDefault) return st.id;
    return null;
  }

  /** True if a style is named `name` (in any case), as Word allows a name once. */
  styleNamed(name) {
    const low = name.trim().toLowerCase();
    return [...this.styles.byId.values()].some((st) => st.name.toLowerCase() === low || this.styles.label(st.id).toLowerCase() === low);
  }

  /**
   * Takes paragraph `el`'s own properties that say what `pPr` (a style's) says off it, so the style
   * gives them, as Word does when a paragraph is made into a new style.
   */
  absorbParagraph(el, pPr) {
    if (!pPr || el.dataset.ppr == null) return;
    const own = this.rawClone(el.dataset.ppr);
    for (const c of elementChildren(pPr)) {
      const mine = wChild(own, c.localName);
      if (mine && this.sameProperty(mine, c)) own.removeChild(mine);
    }
    el.dataset.ppr = this.keep(own);
    const css = this.styles.pPrCss(own, el.dir === 'rtl');
    for (const prop of ['margin-top', 'margin-bottom', 'line-height', 'padding-left', 'padding-right', 'text-indent', 'text-align', 'border-top', 'border-bottom', 'border-left', 'border-right', 'background-color']) {
      el.style.setProperty(prop, css[prop] || '');
    }
  }

  /** True if the document's Heading 1 is numbered by its style (see numberHeadings). */
  get headingsNumbered() {
    const h1 = this.styles.findBuiltin('h1');
    return !!(h1 && this.styleList(h1));
  }

  /** Takes the numbering off the heading styles (numbered by numberHeadings, or in Word). */
  unnumberHeadings() {
    for (const s of this.styles.byId.values()) {
      if (s.type === 'paragraph' && /^heading [1-9]$/.test(s.name.toLowerCase()) && wChild(s.pPr, 'numPr')) this.styles.setStyleNumbering(s.id, null);
    }
  }

  newList(kind, first = null) {
    if (kind === 'bullet' && !first && this.listDefs.bulletNumId) return this.listDefs.bulletNumId;
    if (!this.numberingXml) this.numberingXml = parseXml(`<w:numbering xmlns:w="${W}"/>`);
    const numId = addList(this.numberingXml, kind, this.listDefs, first);
    if (kind === 'bullet' && !first) this.listDefs.bulletNumId = numId;
    this.numberingChanged = true;
    this.kinds = null;
    return numId;
  }

  // ---------------------------------------------------------------- HTML -> XML

  /** Writes the editor content back into document.xml (in memory). */
  update(root) {
    const sectPr = wChild(this.body, 'sectPr');
    this.revisions = []; // <w:ins>/<w:del> written for this session's tracked changes
    this.linksUsed = new Map(); // rels path -> Set of the ids of this session's links still in the text
    this.serialRels = { path: this.relsPath, rels: this.rels };
    this.textboxCopies = []; // text boxes whose new text goes to their copies for older readers too
    const nodes = this.serializeBlocks(root);
    this.notesOut = this.serializeNotes(root);
    this.partsOut = this.serializeParts(root);
    this.revisions.forEach((rev, i) => rev.setAttributeNS(W, 'w:id', String(this.nextId + i)));
    // Copies of edited text boxes, with their own ids for this session's tracked changes in them.
    let nextCopyId = this.nextId + this.revisions.length;
    const ours = new Set(this.revisions);
    for (const { main, copies } of this.textboxCopies) {
      const from = [...main.getElementsByTagName('*')];
      for (const copy of syncCopies(main, copies)) {
        const to = [...copy.getElementsByTagName('*')];
        from.forEach((n, i) => { if (ours.has(n) && to[i]) to[i].setAttributeNS(W, 'w:id', String(nextCopyId++)); });
      }
    }
    while (this.body.firstChild) this.body.removeChild(this.body.firstChild);
    for (const n of nodes) this.body.appendChild(n);
    if (!nodes.some((n) => n.localName === 'p')) this.body.appendChild(createW(this.xml, 'p'));
    if (sectPr) this.body.appendChild(sectPr);
    this.applyPageSetup();
    this.placeNewFooters(root);
    this.commentsRemoved = removedComments(this.comments, {
      wasInBody: this.commentsInBody, present: commentIdsIn(this.body), elsewhere: this.commentsElsewhere,
    });
    pruneCommentMarkers(this.body, this.commentsRemoved, this.commentRangesInBody);
    for (const id of this.storyRemoved()) this.commentsRemoved.add(id);
    // A comment in a note, header or footer goes with its text, as in the body: its mark was deleted with it.
    const outs = [
      ...NOTE_KINDS.filter((k) => this.notesOut[k]).map((k) => [this.noteParts[k].path, this.notesOut[k]]),
      ...Object.entries(this.partsOut),
    ];
    for (const [partPath, xml] of outs) {
      const present = commentIdsIn(xml);
      for (const [id, path] of this.commentStory) if (path === partPath && !present.has(id)) this.commentsRemoved.add(id);
    }
    for (let grew = true; grew;) {
      grew = false;
      for (const c of this.comments) {
        if (c.parentId && this.commentsRemoved.has(c.parentId) && !this.commentsRemoved.has(c.id)) {
          this.commentsRemoved.add(c.id);
          grew = true;
        }
      }
    }
    this.commentEdits = this.commentOverrides(root);
    // Edits to comments in headers, footers and notes, where they differ from the comment as it was.
    for (const [id, text] of Object.entries(this.story.edits)) {
      const c = this.comments.find((x) => x.id === id);
      if (c && text !== this.commentText(c)) this.commentEdits.set(id, { ...this.commentEdits.get(id), text });
    }
    for (const [id, done] of Object.entries(this.story.done)) {
      const c = this.comments.find((x) => x.id === id);
      if (c && done !== !!c.done) this.commentEdits.set(id, { ...this.commentEdits.get(id), done });
    }
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
    // Formatting changed with tracking this session: the paragraphs are tracked too.
    let formatChanges = 0;
    const outs = [this.xml, ...NOTE_KINDS.map((k) => this.notesOut[k]).filter(Boolean), ...Object.values(this.partsOut)];
    for (const xml of outs) {
      for (const local of ['rPrChange', 'pPrChange']) {
        for (const change of xml.getElementsByTagNameNS(W, local)) {
          if (Number(wAttr(change, 'id')) < this.firstNewId) continue;
          formatChanges++;
          for (let n = change.parentNode; n?.nodeType === 1; n = n.parentNode) {
            if (n.namespaceURI === W && n.localName === 'p') { trackedParagraphs.add(n); break; }
          }
        }
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
    const tracked = (p) => trackedParagraphs.has(p);
    return {
      paragraphs: [
        ...reviewParagraphs(this.originalBody, this.body, { tracked, styleName }),
        ...this.reviewNotes(root, { tracked, styleName }),
        ...this.reviewParts(root, { tracked, styleName }),
      ],
      commentsRemoved: this.comments.filter((c) => this.commentsRemoved.has(c.id) && !c.isNew),
      commentsAdded: this.newComments.filter((c) => this.presentComments().has(c.id) && !this.commentsRemoved.has(c.id))
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
      formatChanges,
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
      fields: { ...fieldReview(this.originalBody, this.body), updateOnOpen: this.updateFields, alreadyUpdating: this.updateFieldsSet },
    };
  }

  /** The paragraphs of footnotes and endnotes that changed, for the review: "Footnote 2", paragraph 1… */
  reviewNotes(root, opts) {
    const out = [];
    for (const kind of NOTE_KINDS) {
      const xml = this.notesOut[kind];
      if (!xml) continue;
      const before = parseXml(this.noteParts[kind].text);
      const find = (doc, id) => wChildren(doc.documentElement, kind).find((n) => wAttr(n, 'id') === id);
      for (const el of root.querySelectorAll(`.docx-note[data-note="${kind}"]`)) {
        const a = find(before, el.dataset.noteId);
        const b = find(xml, el.dataset.noteId);
        if (!a || !b) continue;
        for (const entry of reviewParagraphs(a, b, opts)) out.push({ ...entry, where: `${NOTE_NAMES[kind]} ${el.dataset.n}` });
      }
    }
    return out;
  }

  /**
   * Line numbering of each section, in order (w:lnNumType in each w:sectPr): { countBy, start,
   * distance (px), restart } or null. The first number shown is start + 1, as Word's "Start at".
   */
  lineNumbering() {
    const sectPrs = [...this.body.getElementsByTagNameNS(W, 'sectPr')]
      .filter((s) => s.parentNode === this.body || (s.parentNode.localName === 'pPr' && s.parentNode.parentNode?.localName === 'p'));
    return sectPrs.map((s) => {
      const ln = wChild(s, 'lnNumType');
      if (!ln) return null;
      const countBy = Number(wAttr(ln, 'countBy')) || 1;
      const distance = wAttr(ln, 'distance');
      return {
        countBy,
        start: Number(wAttr(ln, 'start')) || 0,
        distance: distance != null && distance !== '' ? Number(distance) / 15 : 24, // twips to px; Word's auto is a quarter inch
        restart: wAttr(ln, 'restart') || 'newPage',
      };
    });
  }

  /** The kind of file: .docx, a template (.dotx) or with macros (.docm, .dotm), from its main part. */
  async readFormat() {
    const ct = await this.text('[Content_Types].xml');
    const main = ct && [...parseXml(ct).getElementsByTagName('Override')].find((o) => o.getAttribute('PartName') === `/${this.mainPath}`);
    const type = main?.getAttribute('ContentType') || '';
    const byType = /macroEnabledTemplate/i.test(type) ? 'dotm' : /macroEnabled/i.test(type) ? 'docm' : /template\.main/i.test(type) ? 'dotx' : /document\.main/i.test(type) ? 'docx' : null;
    const ext = byType || (/\.(dotx|docm|dotm)$/i.exec(this.name)?.[1].toLowerCase()) || 'docx';
    this.format = {
      ext,
      mime: FORMAT_MIME[ext],
      template: ext === 'dotx' || ext === 'dotm',
      macros: ext === 'docm' || ext === 'dotm' || Object.keys(this.zip.files).some((p) => /vbaProject\.bin$/i.test(p)),
    };
    this.openedFormat = this.format;
    this.openedMainType = type;
  }

  /** Whether this template is being made into a new document (see newFromTemplate). */
  get fromTemplate() {
    return this.format !== this.openedFormat;
  }

  /**
   * A template opened here saves as the template itself, unless it is made into a new document,
   * as Word makes one when a template is double-clicked: a .dotx gives a .docx and a .dotm a .docm,
   * with the main part's content type changed to match and everything else as it is. The template
   * the file was opened from is never changed. `on` false goes back to saving the template.
   */
  newFromTemplate(on) {
    const from = this.openedFormat;
    if (!from.template) return;
    if (!on) {
      this.format = from;
      return;
    }
    const ext = DOCUMENT_OF[from.ext];
    this.format = { ...from, ext, mime: FORMAT_MIME[ext], template: false };
  }

  /** The main part's content type, to match the kind of file being saved (see newFromTemplate). */
  async writeMainType() {
    const text = await this.text('[Content_Types].xml');
    if (!text) return;
    const ct = parseXml(text);
    const main = [...ct.getElementsByTagName('Override')].find((o) => o.getAttribute('PartName') === `/${this.mainPath}`);
    // As opened, unless the template is being made into a document.
    const type = this.fromTemplate ? MAIN_PART_TYPE[this.format.ext] : this.openedMainType;
    if (!main || !type || main.getAttribute('ContentType') === type) return;
    main.setAttribute('ContentType', type);
    this.zip.file('[Content_Types].xml', serializeXml(ct));
  }

  /** Produces a .docx Blob from the current editor content. */
  async save(root, { plain = false } = {}) {
    await this.write(root);
    let blob;
    if (this.signatures) {
      // Unchanged, the file goes back exactly as it was opened, so its signatures stay valid.
      this.signaturesRemoved = this.changed;
      if (!this.signaturesRemoved) blob = new Blob([this.sourceBytes], { type: this.format.mime });
      // Changed, they no longer match the document: Word removes them when a signed document is
      // edited, and so does Intact, rather than send a file that Word shows as invalidly signed.
      else await removeSignatures(this.zip, this.signatures);
    }
    blob ||= await this.zip.generateAsync({ type: 'blob', mimeType: this.format.mime, compression: 'DEFLATE' });
    // With a password, encrypted with it, as Word saves such a document; `plain` for what is made
    // from the document on this device (mail merge, a redline), which encrypts its own results.
    return plain ? blob : this.sealed(blob, this.changed);
  }

  /**
   * A copy of the document as edited, to keep as a draft (see drafts.js): the edits only, without
   * what saving adds for the recipient (the properties' record of the edit, personal information
   * removed), so the download made later from the draft records them once, as it would have. A
   * signed document's signatures go from the copy, as from any changed download; the document
   * itself keeps them until it is saved. Null when there is nothing to keep: no change at all.
   */
  async draft(root) {
    const keep = { updateProperties: this.updateProperties, removePersonal: this.removePersonal };
    Object.assign(this, { updateProperties: false, removePersonal: false });
    try {
      await this.write(root);
    } finally {
      Object.assign(this, keep);
    }
    if (!this.changed && !this.encryptionChanged) return null;
    let zip = this.zip;
    if (this.signatures) {
      zip = await JSZip.loadAsync(await this.zip.generateAsync({ type: 'uint8array' }));
      await removeSignatures(zip, this.signatures);
    }
    // A password-protected document's draft is encrypted with its password too: nothing of it is
    // kept on the device in the clear (restoring it asks for the password).
    return this.sealed(await zip.generateAsync({ type: 'blob', mimeType: this.format.mime, compression: 'DEFLATE' }), true);
  }

  /**
   * For the review of a signed document: how many signatures it has, and whether saving now would
   * change it (and so remove them).
   */
  async signatureStatus(root) {
    if (!this.signatures) return null;
    await this.write(root);
    return { count: this.signatures.count, removed: this.changed };
  }

  /**
   * Whether the package written so far differs in meaning from the one opened. XML parts are
   * compared as serialized (rewriting an untouched part only reorders its markup); other parts byte
   * by byte. The package relationships, content types and signatures don't count: they change only
   * when signatures are removed, which follows from a change elsewhere.
   */
  async packageChanged() {
    // The core and app properties change because something else did (see writeCore, writePersonal).
    const skip = new Set(['_rels/.rels', '[Content_Types].xml', ...(this.signatures?.parts || []),
      this.propertyPaths?.core, this.propertyPaths?.app].filter(Boolean));
    const now = Object.entries(this.zip.files).filter(([path, f]) => !f.dir && !skip.has(path));
    const before = [...this.openedEntries].filter(([path, f]) => !f.dir && !skip.has(path));
    if (now.length !== before.length) return true;
    for (const [path, f] of now) {
      const old = this.openedEntries.get(path);
      if (!old) return true;
      if (old === f) continue; // never rewritten
      if (path === this.mainPath && this.mainUnchanged) continue;
      if (path === this.mainPath && this.pristineMain != null) {
        if ((await f.async('string')) !== this.pristineMain) return true;
      } else if (/\.(xml|rels)$/i.test(path)) {
        const [a, b] = await Promise.all([old.async('string'), f.async('string')]);
        if (a !== b && serializeXml(parseXml(a)) !== serializeXml(parseXml(b))) return true;
      } else {
        const [a, b] = await Promise.all([old.async('uint8array'), f.async('uint8array')]);
        if (a.length !== b.length || a.some((x, i) => x !== b[i])) return true;
      }
    }
    return false;
  }

  /**
   * Notes what the main part saves as before anything is edited (`root` as rendered), so that
   * packageChanged tells an edit from the tidying any save does (xml:space, runs merged), as notes,
   * headers and footers are compared with how they saved when rendered.
   */
  markPristine(root) {
    this.update(root);
    this.pristineMain = serializeXml(this.xml);
  }

  /** Writes the current editor content into the package: every part a save changes. */
  async write(root) {
    this.update(root);
    // The text untouched (it saves as it did when rendered), the part goes back byte for byte, as
    // unedited headers, footers and notes do: an edit elsewhere doesn't rewrite it.
    const main = serializeXml(this.xml);
    const opened = this.openedEntries.get(this.mainPath);
    this.mainUnchanged = main === this.pristineMain && !this.restored && !!opened;
    this.zip.file(this.mainPath, this.mainUnchanged ? await opened.async('uint8array') : main);
    await this.writeMedia();
    await this.writeLinks();
    await this.writeStyles();
    if (this.numberingChanged) await this.writeNumbering();
    await this.writeComments();
    this.writeNotes();
    this.writeParts();
    await this.writeNewParts();
    await this.writeSources();
    this.writeStories();
    await this.writeSettings();
    await this.writeForgottenLinks(); // once every part is written: whether a part still uses each
    await this.restorePersonal();
    await this.writeMainType();
    // Whether the document now differs from the one opened: only then are its properties updated.
    // A restored draft already holds edits to the file first opened, so it counts as changed too,
    // and so does a new document made from a template, which is no longer the file opened.
    await this.writeCustom();
    this.changed = this.restored || this.removePersonal || this.newTitle != null || !!this.propEdits.core || this.fromTemplate || await this.packageChanged();
    await this.writeCore();
    await this.writePersonal();
  }

  /**
   * Records the edit in docProps/core.xml, as Word does on saving: modified now, last modified by
   * whoever edited (if they gave their name), the revision number one up; and the title given here.
   * Always worked out from the properties as opened, so saving twice counts one revision; with
   * nothing changed (or updating switched off), the part as opened goes back.
   */
  async writeCore() {
    const edits = this.propEdits?.core || null;
    // Without a core part, one is added for properties given here (as Word adds one on saving).
    if (!this.propertyPaths?.core && ((edits && Object.values(edits).some(Boolean)) || this.newTitle)) {
      await this.addPackagePart('docProps/core.xml', CORE_PROPS_REL, CORE_PROPS_CT, EMPTY_CORE);
      this.propertyPaths.core = 'docProps/core.xml';
      this.coreOriginal = EMPTY_CORE;
      this.coreAdded = true;
    }
    const path = this.propertyPaths?.core;
    if (!path || this.coreOriginal == null) return;
    const update = this.changed && this.updateProperties;
    if (!update && this.newTitle == null && !edits) {
      if (this.coreWritten && this.zip.files[path] === this.coreWritten) this.zip.file(path, this.coreOriginal);
      this.coreWritten = null;
      return;
    }
    let xml = updateCore(this.coreOriginal, update ? { modified: new Date(), by: this.modifiedBy, revision: true } : {}, edits ? null : this.newTitle);
    if (edits) xml = setCore(xml, edits);
    this.zip.file(path, xml);
    this.coreWritten = this.zip.files[path];
  }

  /**
   * Writes the custom properties given here (Document properties…) into docProps/custom.xml: in
   * the part the file had, or a new one (with its relationship and content type); with none given
   * (or Undo back to none), the part goes back as it was.
   */
  async writeCustom() {
    const list = this.propEdits?.custom;
    const path = this.propertyPaths?.custom;
    if (!list) {
      if (path && this.customWritten && this.zip.files[path] === this.customWritten) {
        if (this.customOriginal != null) this.zip.file(path, this.customOriginal);
      }
      this.customWritten = null;
      return;
    }
    if (!path) {
      await this.addPackagePart('docProps/custom.xml', CUSTOM_PROPS_REL, CUSTOM_PROPS_CT, setCustom(null, list));
      this.propertyPaths.custom = 'docProps/custom.xml';
    } else this.zip.file(path, setCustom(this.customOriginal, list));
    this.customWritten = this.zip.files[this.propertyPaths.custom];
  }

  /** Adds a part of the package (not the main part's): its file, its relationship in _rels/.rels and its content type. */
  async addPackagePart(path, relType, contentType, xml) {
    this.zip.file(path, xml);
    const rels = parseXml((await this.text('_rels/.rels')) || `<Relationships xmlns="${NS.pr}"/>`);
    if (![...rels.documentElement.children].some((r) => r.getAttribute('Type') === relType)) {
      let id = 1;
      while ([...rels.documentElement.children].some((r) => r.getAttribute('Id') === `rIdProps${id}`)) id++;
      const rel = rels.createElementNS(NS.pr, 'Relationship');
      rel.setAttribute('Id', `rIdProps${id}`);
      rel.setAttribute('Type', relType);
      rel.setAttribute('Target', path);
      rels.documentElement.appendChild(rel);
      this.zip.file('_rels/.rels', serializeXml(rels));
    }
    const ct = parseXml(await this.text('[Content_Types].xml'));
    if (![...ct.documentElement.children].some((d) => d.getAttribute('PartName') === `/${path}`)) {
      const ov = ct.createElementNS(NS.ct, 'Override');
      ov.setAttribute('PartName', `/${path}`);
      ov.setAttribute('ContentType', contentType);
      ct.documentElement.appendChild(ov);
      this.zip.file('[Content_Types].xml', serializeXml(ct));
    }
  }

  /**
   * The properties Document properties… edits, as they are now: { core: { title, subject, keywords,
   * category, contentStatus, description }, custom: [{ name, value }] } (custom ones Office manages,
   * a sensitivity label's, left out).
   */
  editableProperties() {
    const core = this.propEdits.core || readCore(this.coreOriginal);
    if (!this.propEdits.core && this.newTitle != null) core.title = this.newTitle;
    const custom = this.propEdits.custom || (this.properties?.custom || []).filter((p) => !managedProperty(p.name)).map(({ name, value }) => ({ name, value }));
    return { core: { ...core }, custom: custom.map((p) => ({ ...p })) };
  }

  /** Sets the properties to save (see editableProperties); the review's list follows. */
  setProperties({ core, custom }) {
    this.propEdits = { core: core ? { ...core } : null, custom: custom ? custom.map((p) => ({ ...p })) : null };
    if (core) this.newTitle = null;
    this.refreshPropertyList();
  }

  /** The properties as they are now in the list the review shows (doc.properties). */
  refreshPropertyList() {
    if (!this.propertiesAsOpened) this.propertiesAsOpened = { core: this.properties.core, custom: this.properties.custom };
    const { core, custom } = this.propEdits;
    const labels = Object.fromEntries(EDITABLE_CORE.map(([key, , , label]) => [label, key]));
    this.properties.core = core
      ? [...this.propertiesAsOpened.core.filter((p) => !(p.label in labels)), ...EDITABLE_CORE.filter(([key]) => core[key]).map(([key, , , label]) => ({ label, value: core[key] }))]
      : this.propertiesAsOpened.core;
    this.properties.custom = custom
      ? [...this.propertiesAsOpened.custom.filter((p) => managedProperty(p.name)), ...custom]
      : this.propertiesAsOpened.custom;
  }

  /** For Undo: the properties edited here, as a value (see setPropertyState). */
  propertyState() {
    return this.propEdits.core || this.propEdits.custom ? this.propEdits : null;
  }

  setPropertyState(state) {
    this.propEdits = { core: state?.core ? { ...state.core } : null, custom: state?.custom ? state.custom.map((p) => ({ ...p })) : null };
    this.refreshPropertyList();
  }

  /**
   * The value a DOCPROPERTY field shows for property `name` (Title, Subject, Keywords, Category,
   * Comments, or a custom property's name, in any case), or null if the document has none.
   */
  propertyValue(name) {
    const { core, custom } = this.editableProperties();
    const key = EDITABLE_CORE.find(([, , , , field]) => field && field.toLowerCase() === String(name).toLowerCase())?.[0];
    if (key) return core[key];
    const p = custom.find((c) => c.name.toLowerCase() === String(name).toLowerCase());
    return p ? p.value : null;
  }

  /** What saving records in the properties, for the review: { available, on, by, revision, lastBy }. */
  propertiesUpdate() {
    const core = this.coreOriginal ? parseXml(this.coreOriginal).documentElement : null;
    const value = (local) => core?.getElementsByTagNameNS('*', local)[0]?.textContent.trim() || '';
    const revision = Number(value('revision'));
    return {
      available: !!core,
      changed: this.changed,
      on: this.updateProperties,
      by: this.modifiedBy,
      lastBy: value('lastModifiedBy'),
      revision: Number.isInteger(revision) && revision > 0 ? revision + 1 : null,
    };
  }

  /**
   * Removes personal information from the parts that hold it, when asked for (as Word's Document
   * Inspector does): names on comments and tracked changes, and the author, last saved by, company
   * and manager properties. A part changed by an earlier save and not rewritten since goes back to
   * how it was first, so switching the option off restores everything.
   */
  async restorePersonal() {
    for (const [path, { written, original }] of this.anonymized) {
      if (this.zip.files[path] === written) this.zip.file(path, await original.async('uint8array'));
    }
    this.anonymized.clear();
  }

  async writePersonal() {
    await this.restorePersonal();
    if (!this.removePersonal) return;
    for (const path of personalParts(this.zip, this.mainDir, this.propertyPaths)) {
      const original = this.zip.files[path];
      if (!original) continue;
      const out = anonymize(await original.async('string'));
      if (out == null) continue;
      this.zip.file(path, out);
      this.anonymized.set(path, { written: this.zip.files[path], original });
    }
  }

  /**
   * What a document can carry that isn't on the page, for the review's "Before you release":
   * tracked changes (and who made them), the document's properties, custom XML data, embedded files
   * and macros. Comments and hidden text are counted by the editor, which shows them.
   */
  inspect() {
    const outs = [this.xml, ...NOTE_KINDS.map((k) => this.notesOut?.[k]).filter(Boolean), ...Object.values(this.partsOut || {})];
    const authors = new Set();
    let revisions = 0;
    for (const xml of outs) {
      const root = xml.documentElement || xml;
      revisions += countRevisions(root);
      for (const a of revisionAuthors(root)) authors.add(a);
    }
    const files = Object.keys(this.zip.files);
    return {
      revisions,
      revisionAuthors: [...authors],
      properties: this.properties,
      removePersonal: this.removePersonal,
      customXml: files.filter((p) => /^customXml\/item\d+\.xml$/i.test(p)).length,
      embedded: files.filter((p) => /\/embeddings\/[^/]+$/i.test(p)).length,
      macros: files.some((p) => /vbaProject\.bin$/i.test(p)),
    };
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
   * Saves whether Track Changes is on (w:trackRevisions), so Word keeps tracking for the recipient,
   * and, if asked for in the review, that Word should update fields on opening (w:updateFields).
   * settings.xml is only rewritten for one of those; when neither applies any more, the original
   * goes back.
   */
  async writeSettings() {
    const tracking = this.tracking !== this.trackRevisions;
    const fields = this.updateFields && !this.updateFieldsSet;
    // Word keeps removing personal information on its own saves (w:removePersonalInformation).
    const personal = this.removePersonal && !this.removePersonalSet;
    const evenOdd = this.hfFlags.evenAndOdd != null && this.hfFlags.evenAndOdd !== this.evenAndOddInFile;
    if (!tracking && !fields && !personal && !evenOdd) {
      if (this.settingsChanged) this.zip.file(this.settingsPath, this.settingsText);
      this.settingsChanged = false;
      return;
    }
    this.settingsChanged = true;
    if (!this.settingsPath) {
      this.settingsText = `${XML_HEADER}<w:settings xmlns:w="${W}"/>`;
      this.settingsPath = await this.addPart('settings.xml', this.settingsText, SETTINGS_REL, SETTINGS_CT);
    }
    const doc = parseXml(this.settingsText);
    const root = doc.documentElement;
    if (tracking && this.tracking) {
      const before = wChildren(root).filter((c) => BEFORE_TRACK_REVISIONS.has(c.localName)).pop();
      root.insertBefore(createW(doc, 'trackRevisions'), before ? before.nextSibling : root.firstChild);
    } else if (tracking) {
      removeW(root, 'trackRevisions');
    }
    if (personal) {
      const before = wChildren(root).filter((c) => ['writeProtection', 'view', 'zoom'].includes(c.localName)).pop();
      root.insertBefore(createW(doc, 'removePersonalInformation'), before ? before.nextSibling : root.firstChild);
    }
    if (evenOdd) {
      removeW(root, 'evenAndOddHeaders');
      if (this.hfFlags.evenAndOdd) {
        const after = [...root.children].find((c) => c.namespaceURI !== W || AFTER_EVEN_AND_ODD.has(c.localName));
        root.insertBefore(createW(doc, 'evenAndOddHeaders'), after || null);
      }
    }
    if (fields) {
      removeW(root, 'updateFields');
      const update = createW(doc, 'updateFields');
      update.setAttributeNS(W, 'w:val', 'true');
      // In schema order: before the settings that follow it (and any extension elements).
      const after = [...root.children].find((c) => c.namespaceURI !== W || AFTER_UPDATE_FIELDS.has(c.localName));
      root.insertBefore(update, after || null);
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
    const present = this.presentComments();
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
      if (node.nodeType === 1 && node.classList.contains('docx-notes')) continue; // saved to their own part
      if (node.nodeType === 1 && node.classList.contains('docx-hf-band')) continue; // headers and footers too
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
    // In a right-to-left paragraph alignment is logical: start (Word's "left") is at the right.
    const rtl = el?.dir === 'rtl';
    if (jc && jcToCss(wAttr(jc, 'val'), rtl) === align) {
      // unchanged
    } else if (!align || align === 'start') {
      removeW(pPr, 'jc');
    } else {
      const val = (rtl ? { left: 'right', right: 'left', end: 'right', center: 'center', justify: 'both' }
        : { left: 'left', center: 'center', right: 'right', justify: 'both', end: 'right' })[align] || 'left';
      ensureOrdered(pPr, 'jc', PPR_ORDER).setAttributeNS(W, 'w:val', val);
    }
    if (el && this.listEdited(el)) {
      // Write the new list membership; the list level's indentation replaces direct indentation,
      // unless indentation was set since (in the Paragraph dialog).
      if (el.dataset.indSet == null) removeW(pPr, 'ind');
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
    if (el.dataset.textbox != null) {
      out.appendChild(this.serializeTextbox(el));
      state.last = null;
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
    if (el.dataset.wrap != null) {
      // A smart tag, custom XML or direction wrapper: its runs, inside it again.
      const wrap = this.rawClone(el.dataset.wrap);
      const before = wrap.childNodes.length;
      const inner = { last: null };
      for (const c of el.childNodes) this.serializeInline(c, wrap, ctx, block, inner);
      if (wrap.childNodes.length > before) out.appendChild(wrap);
      state.last = null;
      return;
    }
    if (el.tagName === 'A' && (el.dataset.hl != null || el.dataset.link != null)) {
      out.appendChild(this.serializeLink(el, ctx, block));
      if (!out.lastChild.childNodes.length) out.lastChild.remove();
      state.last = null;
      return;
    }
    if (el.tagName === 'IMG' || el.tagName === 'SCRIPT' || el.tagName === 'STYLE') return;
    const next = this.inlineCtx(el, ctx);
    for (const c of el.childNodes) this.serializeInline(c, out, next, block, state);
  }

  /** The formatting context inside element `el` (a run's span or a formatting tag), from `ctx` outside it. */
  inlineCtx(el, ctx) {
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
    return next;
  }

  // ---------------------------------------------------------------- run formatting

  /** The formatting elements around text node `node` in its run, outermost first (up to a paragraph, link, field…). */
  formattingAround(node) {
    const chain = [];
    for (let el = node.parentElement; el && isFormatting(el); el = el.parentElement) chain.unshift(el);
    return chain;
  }

  /** The <w:rPr> text node `node` would be saved with now (a new element), or null for none. */
  rPrOf(node) {
    let ctx = {};
    for (const el of this.formattingAround(node)) ctx = this.inlineCtx(el, ctx);
    return this.buildRPr(ctx);
  }

  /** A copy of run properties `rPr` (or empty ones) changed by `mutate`, or null if left empty. */
  mutatedRPr(rPr, mutate) {
    const next = rPr ? rPr.cloneNode(true) : createW(this.xml, 'rPr');
    mutate(next);
    return isEmptyElement(next) ? null : next;
  }

  /**
   * Makes text node `node` a run of its own with properties `rPr` (null: none): it is taken out of
   * the formatting around it and put in a run span rendered from `rPr`, so what shows is what is
   * saved. A tracked formatting change in it is labelled as rendering labels one.
   */
  flattenRun(node, rPr) {
    const top = this.formattingAround(node)[0] || null;
    if (top) isolateIn(node, top); // the run's formatting elements now hold only this text
    let made = node;
    let target = null;
    if (rPr) {
      ({ span: made, target } = this.runSpan(rPr));
      const change = formattingChange(rPr);
      if (change) {
        made.dataset.fmt = 'run';
        made.title = `Formatting changed${byWhom(change)} – tracked change: Accept or Reject on the toolbar`;
      }
    }
    (top || node).before(made);
    if (target) target.append(node);
    top?.remove();
  }

  /** Run (or paragraph) properties without a tracked change's record, for comparing: a copy, or null. */
  withoutChange(props) {
    if (!props) return null;
    const c = props.cloneNode(true);
    removeW(c, 'rPrChange', 'pPrChange');
    return isEmptyElement(c) ? null : c;
  }

  /** True if two <w:rPr> / <w:pPr> say the same (null for none), whatever the order of their attributes. */
  sameProps(a, b) {
    const canon = (el) => {
      const attrs = [...el.attributes].filter((x) => !x.name.startsWith('xmlns')).map((x) => `${x.localName}=${x.value}`).sort().join(' ');
      return `<${el.localName} ${attrs}>${elementChildren(el).map(canon).join('')}</>`;
    };
    // No properties and empty properties say the same; within them, <w:i/> says something.
    const top = (el) => (!el || isEmptyElement(el) ? '' : elementChildren(el).map(canon).join(''));
    return top(a) === top(b);
  }

  /**
   * Run properties `now` recording `old` as a tracked formatting change by { author, date }
   * (<w:rPrChange>, last in <w:rPr>). A change already recorded is kept, as Word keeps the first
   * formatting; formatting changed back to it records nothing.
   */
  trackedRPr(old, now, { author, date }) {
    const prior = old && wChild(old, 'rPrChange');
    const base = now ? now.cloneNode(true) : createW(this.xml, 'rPr');
    removeW(base, 'rPrChange');
    const original = prior ? wChild(prior, 'rPr') : this.withoutChange(old);
    if (this.sameProps(original, base)) return isEmptyElement(base) ? null : base;
    base.appendChild(prior ? prior.cloneNode(true) : this.formatChange('rPrChange', original, author, date));
    return base;
  }

  /** A new <w:rPrChange> / <w:pPrChange> holding the old properties (`props`, or none). */
  formatChange(kind, props, author, date) {
    const change = createW(this.xml, kind);
    change.setAttributeNS(W, 'w:id', String(this.nextId++));
    change.setAttributeNS(W, 'w:author', author || 'Author');
    if (date) change.setAttributeNS(W, 'w:date', date);
    const inner = props ? props.cloneNode(true) : createW(this.xml, kind === 'rPrChange' ? 'rPr' : 'pPr');
    if (kind === 'pPrChange') removeW(inner, 'rPr', 'sectPr', 'pPrChange'); // the old paragraph properties only
    change.appendChild(inner);
    return change;
  }

  /** The <w:pPr> paragraph `el` would be saved with now (a new element), or null; its break's tracking left out. */
  blockPPr(el) {
    const c = el.cloneNode(false);
    delete c.dataset.tcMark;
    return wChild(this.serializeParagraph(c, []), 'pPr');
  }

  /**
   * After a change to paragraph `el`'s style, alignment or list while tracking: its properties record
   * `old` (from blockPPr before the change) as a tracked change by { author, date } (<w:pPrChange>).
   */
  trackParagraph(el, old, { author, date }) {
    const now = this.blockPPr(el) || createW(this.xml, 'pPr');
    const prior = old && wChild(old, 'pPrChange');
    removeW(now, 'pPrChange');
    const strip = (p) => {
      if (!p) return null;
      const c = p.cloneNode(true);
      removeW(c, 'rPr', 'sectPr', 'pPrChange');
      return c;
    };
    const original = prior ? wChild(prior, 'pPr') : strip(old);
    if (!this.sameProps(strip(original), strip(now))) now.appendChild(prior ? prior.cloneNode(true) : this.formatChange('pPrChange', original, author, date));
    el.dataset.ppr = this.keep(now);
    if (formattingChange(now)) el.dataset.fmtPara = '1';
    else delete el.dataset.fmtPara;
  }

  // ---------------------------------------------------------------- spacing and indentation

  /**
   * The spacing and indentation paragraph `el` has, as Word works it out: its own properties, then
   * its list level's, its style's (and those it is based on), then the document's defaults. In
   * twips: { before, after, line, lineRule ('auto': line in 240ths of a line; 'exact', 'atLeast'),
   * left, right, firstLine, hanging }.
   */
  paragraphFormat(el) {
    const own = el.dataset.ppr != null ? this.raw[Number(el.dataset.ppr)] : null;
    const list = this.listOf(el);
    this.kinds ||= new Numbering(this.numberingXml);
    const level = list ? this.kinds.levels(list.numId)?.get(list.ilvl)?.pPr : null;
    // A paragraph whose list changed here is saved with its list level's indentation (see serializeParagraph).
    const ownInd = this.listEdited(el) && el.dataset.indSet == null ? null : own;
    const styles = this.styles.chain(el.dataset.style || this.styles.defaultParagraph).map((st) => st.pPr).reverse();
    const sources = [own, level, ...styles, this.styles.defaultPPr];
    const find = (local, attrs, from = sources) => {
      for (const pPr of from) {
        const e = wChild(pPr, local);
        if (e && attrs.some((a) => wAttr(e, a) != null)) return attrs.map((a) => wAttr(e, a));
      }
      return attrs.map(() => null);
    };
    const num = (v, d = 0) => (v == null || Number.isNaN(Number(v)) ? d : Number(v));
    const indFrom = [ownInd, level, ...styles, this.styles.defaultPPr];
    const [before] = find('spacing', ['before']);
    const [after] = find('spacing', ['after']);
    const [line, lineRule] = find('spacing', ['line', 'lineRule']);
    const [left, start] = find('ind', ['left', 'start'], indFrom);
    const [right, end] = find('ind', ['right', 'end'], indFrom);
    const [firstLine, hanging] = find('ind', ['firstLine', 'hanging'], indFrom);
    return {
      before: num(before), after: num(after), line: num(line, 240), lineRule: line == null ? 'auto' : lineRule || 'auto',
      left: num(left ?? start), right: num(right ?? end), firstLine: hanging != null ? 0 : num(firstLine), hanging: num(hanging),
      ...this.paragraphBreaks(sources),
    };
  }

  /**
   * Word's Line and Page Breaks: keepNext, keepLines and pageBreakBefore, true or false, from the
   * first of `sources` (<w:pPr>s, the paragraph's own first) that says (an on/off property: absent
   * is off).
   */
  paragraphBreaks(sources) {
    const out = {};
    for (const local of BREAK_PROPS) {
      out[local] = false;
      for (const pPr of sources) {
        const v = onOff(wChild(pPr, local));
        if (v !== undefined) {
          out[local] = v;
          break;
        }
      }
    }
    return out;
  }

  /**
   * Sets paragraph `el`'s own spacing and indentation, as Word's Paragraph dialog does: `change` has
   * any of before, after, line, lineRule, left, right, firstLine, hanging (twips; null removes the
   * paragraph's own value, so its style's applies), and `tabs`, the tab stops it is to have (see
   * tabStops). Only what `change` names is written; everything
   * else in the paragraph's properties is kept as it is.
   */
  setParagraphFormat(el, change) {
    const pPr = el.dataset.ppr != null ? this.rawClone(el.dataset.ppr) : createW(this.xml, 'pPr');
    const has = (k) => change[k] !== undefined;
    const write = (local, values, drop) => {
      if (!Object.keys(values).some(has)) return;
      let e = wChild(pPr, local) || ensureOrdered(pPr, local, PPR_ORDER);
      // Values that would win over the ones written (Word reads "lines" and auto spacing first).
      for (const a of drop) e.removeAttributeNS(W, a);
      for (const [k, attr] of Object.entries(values)) {
        if (!has(k)) continue;
        const v = change[k];
        if (v == null) e.removeAttributeNS(W, attr);
        else e.setAttributeNS(W, `w:${attr}`, k === 'lineRule' ? v : String(Math.round(v)));
      }
      if (!e.attributes.length) {
        pPr.removeChild(e);
        e = null;
      }
    };
    const dropSpacing = [];
    if (has('before')) dropSpacing.push('beforeLines', 'beforeAutospacing');
    if (has('after')) dropSpacing.push('afterLines', 'afterAutospacing');
    write('spacing', { before: 'before', after: 'after', line: 'line', lineRule: 'lineRule' }, dropSpacing);
    const dropInd = [];
    if (has('left')) dropInd.push('start', 'leftChars', 'startChars');
    if (has('right')) dropInd.push('end', 'rightChars', 'endChars');
    if (has('firstLine') || has('hanging')) dropInd.push('firstLineChars', 'hangingChars', has('firstLine') ? 'hanging' : 'firstLine');
    write('ind', { left: 'left', right: 'right', firstLine: 'firstLine', hanging: 'hanging' }, dropInd);
    if (has('tabs')) this.writeTabs(pPr, change.tabs, this.tabStops(el).inherited);
    // Line and page breaks: the paragraph's own only where its style (or list level) doesn't say the
    // same, off as <w:keepNext w:val="0"/> where the style has it on, as Word writes them.
    if (BREAK_PROPS.some(has)) {
      const list = this.listOf(el);
      this.kinds ||= new Numbering(this.numberingXml);
      const level = list ? this.kinds.levels(list.numId)?.get(list.ilvl)?.pPr : null;
      const styles = this.styles.chain(el.dataset.style || this.styles.defaultParagraph).map((st) => st.pPr).reverse();
      const inherited = this.paragraphBreaks([level, ...styles, this.styles.defaultPPr]);
      for (const local of BREAK_PROPS) {
        if (!has(local)) continue;
        removeW(pPr, local);
        const on = !!change[local];
        if (on !== inherited[local]) {
          const e = ensureOrdered(pPr, local, PPR_ORDER);
          if (!on) e.setAttributeNS(W, 'w:val', '0');
        }
        if (on) el.dataset[local] = '';
        else delete el.dataset[local];
      }
    }
    el.dataset.ppr = this.keep(pPr);
    if (['left', 'right', 'firstLine', 'hanging'].some(has) && this.listOf(el)) el.dataset.indSet = '';
    // Show it: the properties these set, from the paragraph's own (its style's come from the stylesheet).
    const css = this.styles.pPrCss(pPr, el.dir === 'rtl');
    for (const prop of ['margin-top', 'margin-bottom', 'line-height', 'padding-left', 'padding-right', 'text-indent']) {
      el.style.setProperty(prop, css[prop] || '');
    }
    this.applyTabStops({ querySelectorAll: () => [el] });
  }

  // ---------------------------------------------------------------- tab stops

  /** Word's default tab stops, in twips (w:defaultTabStop in settings.xml; half an inch if unset). */
  get defaultTabStop() {
    if (this._defaultTab == null) {
      const m = /<w:defaultTabStop\b[^>]*\bw:val="(\d+)"/.exec(this.settingsText || '');
      this._defaultTab = m && Number(m[1]) > 0 ? Number(m[1]) : 720;
    }
    return this._defaultTab;
  }

  /**
   * The tab stops paragraph `el` has, as Word works them out: its styles' (from the base style down),
   * its list level's, then its own, each adding stops and clearing ones before (w:tab w:val="clear").
   * Returns { stops, inherited }: [{ pos (twips from the margin), align ('left', 'center', 'right',
   * 'decimal' or 'bar'), leader ('none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot') }],
   * sorted; `inherited` are those without the paragraph's own.
   */
  tabStops(el) {
    const own = el.dataset.ppr != null ? this.raw[Number(el.dataset.ppr)] : null;
    const list = this.listOf(el);
    this.kinds ||= new Numbering(this.numberingXml);
    const level = list ? this.kinds.levels(list.numId)?.get(list.ilvl)?.pPr : null;
    const styles = this.styles.chain(el.dataset.style || this.styles.defaultParagraph).map((st) => st.pPr);
    const apply = (stops, pPr) => {
      for (const t of wChildren(wChild(pPr, 'tabs'), 'tab')) {
        const pos = Number(wAttr(t, 'pos'));
        if (!Number.isFinite(pos)) continue;
        stops.delete(pos);
        const val = wAttr(t, 'val');
        if (val === 'clear' || val === 'num') continue;
        const align = { start: 'left', end: 'right' }[val] || val || 'left';
        stops.set(pos, { pos, align, leader: wAttr(t, 'leader') || 'none' });
      }
      return stops;
    };
    const sorted = (m) => [...m.values()].sort((a, b) => a.pos - b.pos);
    const base = [this.styles.defaultPPr, ...styles, level].reduce(apply, new Map());
    const inherited = sorted(base);
    return { inherited, stops: sorted(apply(new Map(base), own)) };
  }

  /**
   * The paragraph's own <w:tabs> that give it exactly `stops` (see tabStops) over those it inherits:
   * a stop for each one it doesn't inherit as it is, and a clear for each inherited one it hasn't.
   * Written into (a copy of) `pPr`; no <w:tabs> if it needs none.
   */
  writeTabs(pPr, stops, inherited) {
    removeW(pPr, 'tabs');
    const tabs = [];
    const same = (a, b) => a.pos === b.pos && a.align === b.align && a.leader === b.leader;
    for (const t of inherited) if (!stops.some((s) => s.pos === t.pos)) tabs.push({ pos: t.pos, clear: true });
    for (const s of stops) if (!inherited.some((t) => same(t, s))) tabs.push(s);
    if (!tabs.length) return;
    const el = ensureOrdered(pPr, 'tabs', PPR_ORDER);
    for (const t of tabs.sort((a, b) => a.pos - b.pos)) {
      const attrs = t.clear ? { val: 'clear', pos: t.pos } : { val: t.align, ...(t.leader && t.leader !== 'none' ? { leader: t.leader } : {}), pos: t.pos };
      el.appendChild(createW(this.xml, 'tab', attrs));
    }
  }

  /**
   * How the page shows a paragraph's tabs: the browser puts tab stops at multiples of one width
   * (CSS tab-size), from where the paragraph's text starts. That is Word's default stops, or, where
   * the paragraph's first stop of its own is a left one, that stop: exactly where Word puts its
   * first tab (and the next where its stops are evenly spaced, as in "Name:\tAddress:\t" forms).
   */
  tabSize(el) {
    const { stops } = this.tabStops(el);
    const left = this.paragraphFormat(el).left;
    const first = stops.find((s) => s.pos > left && s.align !== 'bar');
    const width = first && first.align === 'left' ? first.pos - left : this.defaultTabStop;
    return `${width / 20}pt`;
  }

  /**
   * Lays out the tabs of the paragraphs under `root` that have any: the paragraph's tab-size (see
   * tabSize), then, in the page, each tab where Word puts it, with its leader (editor/tabLayout.js).
   */
  applyTabStops(root) {
    for (const el of root.querySelectorAll('p, h1, h2, h3, h4, h5, h6')) {
      if (!el.textContent.includes('\t')) {
        if (el.style.tabSize) el.style.removeProperty('tab-size');
        if (el.dataset.tabLaid != null || el.querySelector('[data-tab-laid]')) layoutTabs(el, { stops: [], left: 0, defaultTab: this.defaultTabStop });
        continue;
      }
      const size = this.tabSize(el);
      if (el.style.tabSize !== size) el.style.tabSize = size;
      if (el.isConnected) layoutTabs(el, { stops: this.tabStops(el).stops, left: this.paragraphFormat(el).left, defaultTab: this.defaultTabStop });
    }
  }

  /** The XML parts with text: the body, the notes, headers and footers (as loaded), and the styles. */
  textParts() {
    return [this.xml, this.styles.doc, ...Object.values(this.noteParts).map((p) => p.xml), ...this.headers.map((h) => h.xml), ...this.footers.map((f) => f.xml)];
  }

  /**
   * The fonts to offer: the theme's heading and body fonts (as Word lists them), then the fonts the
   * document's text and styles use, then Word's usual fonts:
   * [{ value ('theme:major', 'theme:minor' or a name), label, name, group ('theme', 'document', 'standard') }].
   */
  fontChoices() {
    const { major, minor } = this.styles.fonts;
    const out = [
      { value: 'theme:major', label: `${major} (Headings)`, name: major },
      { value: 'theme:minor', label: `${minor} (Body)`, name: minor },
    ];
    const names = new Set();
    for (const xml of this.textParts()) {
      for (const f of xml.getElementsByTagNameNS(W, 'rFonts')) {
        const name = wAttr(f, 'ascii') || wAttr(f, 'hAnsi');
        if (name && !SYMBOL_FONTS.test(name)) names.add(name);
      }
    }
    for (const name of [...names].filter((n) => n !== major && n !== minor).sort((a, b) => a.localeCompare(b))) out.push({ value: name, label: name, name, group: 'document' });
    // Then Word's usual fonts, those Windows and macOS have, for a document with few of its own (a
    // blank one, say): after the document's, so a brand template's fonts stay first.
    const listed = new Set(out.map((f) => f.name));
    for (const name of STANDARD_FONTS) if (!listed.has(name)) out.push({ value: name, label: name, name, group: 'standard' });
    out[0].group = 'theme';
    out[1].group = 'theme';
    return out;
  }

  /** Font sizes in points: { used } in the document's text and styles, and Word's usual { others }. */
  sizeChoices() {
    const used = new Set();
    for (const xml of this.textParts()) {
      for (const sz of xml.getElementsByTagNameNS(W, 'sz')) {
        const v = Number(wAttr(sz, 'val'));
        if (v > 0 && sz.parentNode?.localName === 'rPr') used.add(v / 2);
      }
    }
    const sorted = [...used].sort((a, b) => a - b);
    return { used: sorted, others: WORD_SIZES.filter((x) => !used.has(x)) };
  }

  /**
   * The colours to offer for text: the theme's ten (as Word's palette shows them, with the name
   * w:themeColor takes) and the other colours the document's text uses:
   * { theme: [{ hex, theme, label }], used: [{ hex, label }] }.
   */
  colorChoices() {
    const t = { ...DEFAULT_THEME_COLORS, ...this.styles.themeColors };
    const theme = THEME_SLOTS.map(([slot, name, label]) => ({ hex: t[slot], theme: name, label }));
    const themeHex = new Set(theme.map((c) => c.hex));
    const used = new Set();
    for (const xml of this.textParts()) {
      for (const c of xml.getElementsByTagNameNS(W, 'color')) {
        const v = wAttr(c, 'val');
        if (c.parentNode?.localName === 'rPr' && !wAttr(c, 'themeColor') && v && /^[0-9a-f]{6}$/i.test(v) && !themeHex.has(v.toUpperCase())) used.add(v.toUpperCase());
      }
    }
    return { theme, used: [...used].sort().map((hex) => ({ hex, label: `#${hex}` })) };
  }

  /** The document's character styles to offer: those in Word's gallery or used, not hidden: [{ id, label }]. */
  characterStyles() {
    const used = new Set();
    for (const xml of this.textParts()) for (const el of xml.getElementsByTagNameNS(W, 'rStyle')) used.add(wAttr(el, 'val'));
    const out = [];
    for (const st of this.styles.byId.values()) {
      if (st.type !== 'character' || st.isDefault || !st.id) continue;
      if (/^(placeholder text|footnote reference|endnote reference|annotation reference|comment reference|line number|page number)$/i.test(st.name)) continue;
      if (used.has(st.id) || (st.quick && (!st.hidden || st.unhideWhenUsed))) out.push({ id: st.id, label: this.styles.label(st.id), priority: st.priority });
    }
    out.sort((a, b) => a.priority - b.priority || a.label.localeCompare(b.label));
    return out.map(({ id, label }) => ({ id, label }));
  }

  /**
   * A link as <w:hyperlink>: one from the file keeps its element (and relationship) as it was; a
   * new one (data-link="new", pasted or inserted) or one given a new address (data-link="changed")
   * points at a relationship to its address in the part being written, and a new one's text takes
   * Word's Hyperlink character style.
   */
  serializeLink(el, ctx, block, wrapRuns = null) {
    const hl = el.dataset.hl != null ? this.rawClone(el.dataset.hl) : createW(this.xml, 'hyperlink');
    if (el.dataset.link != null) {
      hl.removeAttributeNS(NS.r, 'id');
      hl.removeAttributeNS(W, 'anchor');
      hl.setAttributeNS(NS.r, 'r:id', this.linkRel(el.getAttribute('href')));
      if (el.dataset.hl == null) hl.setAttributeNS(W, 'w:history', '1');
    }
    const inner = { last: null };
    const linkCtx = el.dataset.link === 'new' ? { ...ctx, link: true } : ctx;
    const into = wrapRuns ? hl.appendChild(wrapRuns()) : hl;
    for (const c of el.childNodes) this.serializeInline(c, into, linkCtx, block, inner);
    if (wrapRuns && !into.childNodes.length) hl.removeChild(into);
    return hl;
  }

  /**
   * A link from the file (`a`, rendered from its <w:hyperlink>) was taken out of the text by
   * redaction: on saving, its relationship (the address) goes from its part's relationships too,
   * unless the part still points at it (Undo brought it back), so the address isn't left in the
   * package for anyone to read. A link of this session needs nothing: writeLinks drops it.
   */
  forgetLink(a) {
    if (a.dataset.hl == null) return;
    const rid = this.raw[Number(a.dataset.hl)]?.getAttributeNS(NS.r, 'id');
    const kind = a.closest('.docx-notes')?.dataset.notes;
    const path = this.partOf(a)?.path || (kind ? this.noteParts[kind]?.path : this.mainPath);
    if (rid && path && !this.forgottenLinks.has(`${path}|${rid}`)) this.forgottenLinks.set(`${path}|${rid}`, { path, rid, rel: null });
  }

  /** Takes forgotten links' relationships out of the package, and puts back any that are used again. */
  async writeForgottenLinks() {
    for (const link of this.forgottenLinks.values()) {
      const part = await this.text(link.path);
      const relsPath = relsPathOf(link.path);
      const text = await this.text(relsPath);
      if (!part || !text) continue;
      const used = new RegExp(`:id="${link.rid}"`).test(part);
      const rels = parseXml(text);
      const root = rels.documentElement;
      const rel = [...root.children].find((r) => r.getAttribute('Id') === link.rid);
      if (rel && !used && rel.getAttribute('Type') === HYPERLINK_REL) {
        link.rel = rel.cloneNode(true);
        rel.remove();
      } else if (!rel && used && link.rel) {
        root.appendChild(rels.importNode(link.rel, true));
      } else continue;
      this.zip.file(relsPath, serializeXml(rels));
    }
  }

  /** The relationship id for a link of this session to `href`, in the part being written. */
  linkRel(href) {
    const { path, rels } = this.serialRels;
    if (!this.linkIds.has(path)) this.linkIds.set(path, new Map());
    const ids = this.linkIds.get(path);
    let rid = ids.get(href);
    if (!rid) {
      const taken = new Set([...(rels?.keys() || []), ...ids.values(), ...[...this.newMedia.values()].map((m) => m.rid)]);
      let n = 1;
      while (taken.has(`rIdLink${n}`)) n++;
      rid = `rIdLink${n}`;
      ids.set(href, rid);
    }
    if (!this.linksUsed.has(path)) this.linksUsed.set(path, new Set());
    this.linksUsed.get(path).add(rid);
    return rid;
  }

  /**
   * Writes the relationships of this session's links into each part's relationships, as external
   * hyperlinks: those still in the text are there, those deleted are taken out again. A part none of
   * whose links are left is written back as it was (when it had been changed).
   */
  async writeLinks() {
    for (const [path, ids] of this.linkIds) {
      const used = this.linksUsed.get(path) || new Set();
      const ours = new Set(ids.values());
      const text = await this.text(path);
      if (!text && !used.size) continue;
      const rels = parseXml(text || `<Relationships xmlns="${NS.pr}"/>`);
      const root = rels.documentElement;
      let changed = false;
      for (const r of [...root.children]) {
        if (ours.has(r.getAttribute('Id')) && !used.has(r.getAttribute('Id'))) { r.remove(); changed = true; }
      }
      for (const [href, rid] of ids) {
        if (!used.has(rid) || [...root.children].some((r) => r.getAttribute('Id') === rid)) continue;
        const rel = rels.createElementNS(NS.pr, 'Relationship');
        rel.setAttribute('Id', rid);
        rel.setAttribute('Type', HYPERLINK_REL);
        rel.setAttribute('Target', href);
        rel.setAttribute('TargetMode', 'External');
        root.appendChild(rel);
        changed = true;
      }
      if (changed) this.zip.file(path, serializeXml(rels));
    }
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
    // A change of the same kind inside one (text typed into the session's own insertion, as when
    // typing over part of it): its text is part of the outer change; Word doesn't nest them.
    if (out.namespaceURI === W && out.localName === kind) {
      for (const c of el.childNodes) this.serializeInline(c, out, ctx, block, state);
      return;
    }
    const prev = state.revision;
    const wrap = prev && prev.kind === kind && prev.author === author && out.lastChild === prev.el
      ? prev.el : this.newRevision(kind, author, el.dataset.tcDate);
    const inner = { last: null };
    let current = wrap;
    const written = [wrap]; // what holds this change's text: deleted text is <w:delText> in all of it
    for (const c of el.childNodes) {
      if (c.nodeType === 1 && c.tagName === 'A' && c.dataset.link === 'new') {
        // A link added while tracking: the link holds the tracked change, as Word writes it.
        if (current.childNodes.length && current.parentNode !== out) out.appendChild(current);
        written.push(out.appendChild(this.serializeLink(c, ctx, block, () => this.newRevision(kind, author, el.dataset.tcDate))));
        current = this.newRevision(kind, author, el.dataset.tcDate);
        written.push(current);
        continue;
      }
      this.serializeInline(c, current, ctx, block, inner);
    }
    if (current !== wrap) {
      if (current.childNodes.length) out.appendChild(current);
      else this.revisions.splice(this.revisions.indexOf(current), 1);
    }
    if (kind === 'del') {
      for (const [from, to] of [['t', 'delText'], ['instrText', 'delInstrText']]) {
        for (const t of written.flatMap((w) => [...w.getElementsByTagNameNS(W, from)])) {
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
    // Runs whose properties say the same are joined, wherever their properties are kept.
    const key = `${this.rPrKey(ctx.rpr)}|${ctx.b}|${ctx.i}|${ctx.u}|${ctx.s}|${ctx.va}|${ctx.hl}|${!!ctx.link}`;
    if (state.last && state.last.key === key && state.last.run.parentNode === out && out.lastChild === state.last.run) {
      const prev = state.last.run.lastChild;
      // Text after text in the same run is one <w:t>, as Word writes it.
      if (child.localName === 't' && prev?.localName === 't' && prev.namespaceURI === W) prev.textContent += child.textContent;
      else state.last.run.appendChild(child);
      return;
    }
    const run = createW(this.xml, 'r');
    const rPr = this.buildRPr(ctx);
    if (rPr) run.appendChild(rPr);
    run.appendChild(child);
    out.appendChild(run);
    state.last = { key, run };
  }

  /** What the run properties kept at raw index `index` say, as a string (cached): equal for equal properties. */
  rPrKey(index) {
    if (index == null) return '';
    this.rPrKeys ||= new Map();
    const rPr = this.raw[Number(index)];
    if (!this.rPrKeys.has(index)) this.rPrKeys.set(index, isEmptyElement(rPr) ? '' : new XMLSerializer().serializeToString(rPr).replace(/ xmlns(:\w+)?="[^"]*"/g, ''));
    return this.rPrKeys.get(index);
  }

  buildRPr(ctx) {
    const rPr = ctx.rpr != null ? this.rawClone(ctx.rpr) : createW(this.xml, 'rPr');
    if (ctx.link) ensureOrdered(rPr, 'rStyle', RPR_ORDER).setAttributeNS(W, 'w:val', this.styles.ensureHyperlink());
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
/** A section's columns (w:cols), for showing them: { count, gap (px), rule } or null for one column. */
function columnsOf(sectPr) {
  const cols = wChild(sectPr, 'cols');
  if (!cols) return null;
  const count = Number(wAttr(cols, 'num')) || wChildren(cols, 'col').length || 1;
  if (count < 2) return null;
  const space = wAttr(cols, 'space');
  return { count: Math.min(count, 45), gap: (space != null ? Number(space) || 0 : 720) / 15, rule: ['1', 'true', 'on'].includes(wAttr(cols, 'sep') || '') };
}

/** The page's wrapper for a section in columns `cols` (see columnsOf): only for the page, never saved. */
export function columnsWrapper(cols) {
  const div = document.createElement('div');
  div.className = 'docx-columns';
  div.dataset.cols = String(cols.count);
  div.style.columnCount = String(cols.count);
  div.style.columnGap = `${cols.gap}px`;
  if (cols.rule) div.style.columnRule = '1px solid #000';
  return div;
}

const FORMAT_TAGS = new Set(['SPAN', 'B', 'STRONG', 'I', 'EM', 'U', 'INS', 'S', 'STRIKE', 'DEL', 'SUP', 'SUB', 'MARK', 'FONT']);

/** True for an element that only formats text: a run's span or a formatting tag (not a field, tracked change, link…). */
function isFormatting(el) {
  if (!FORMAT_TAGS.has(el.tagName) || el.hasAttribute('contenteditable')) return false;
  const d = el.dataset;
  if (d.tc != null || d.raw != null || d.sdt != null || d.textbox != null || d.hl != null || d.link != null) return false;
  return !el.classList.contains('field') && !el.classList.contains('raw-inline');
}

/**
 * Splits `ancestor` around `node` so that it holds only `node` (and what wraps it). A piece split off
 * with nothing in it (only the empty formatting elements a split leaves) is dropped.
 */
function isolateIn(node, ancestor) {
  const split = (setRange, place) => {
    const r = document.createRange();
    setRange(r);
    if (r.collapsed) return;
    const copy = ancestor.cloneNode(false);
    copy.append(r.extractContents());
    if (copy.textContent || copy.querySelector('br, img, [data-raw], [data-sdt], [data-textbox]')) place(copy);
  };
  split((r) => { r.setStart(ancestor, 0); r.setEndBefore(node); }, (c) => ancestor.before(c));
  split((r) => { r.setStartAfter(node); r.setEnd(ancestor, ancestor.childNodes.length); }, (c) => ancestor.after(c));
}

const SYMBOL_FONTS = /^(Symbol|Wingdings( \d)?|Webdings|MS Gothic|MS Mincho|Marlett)$/i;
const WORD_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
/** Word's theme colour slots in the order of its palette: [slot, w:themeColor, label]. */
const THEME_SLOTS = [
  ['lt1', 'background1', 'White, Background 1'], ['dk1', 'text1', 'Black, Text 1'], ['lt2', 'background2', 'Background 2'],
  ['dk2', 'text2', 'Text 2'], ['accent1', 'accent1', 'Accent 1'], ['accent2', 'accent2', 'Accent 2'], ['accent3', 'accent3', 'Accent 3'],
  ['accent4', 'accent4', 'Accent 4'], ['accent5', 'accent5', 'Accent 5'], ['accent6', 'accent6', 'Accent 6'],
];
const DEFAULT_THEME_COLORS = {
  dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6', accent1: '4472C4', accent2: 'ED7D31',
  accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47',
};

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

/**
 * The instruction of the whole field run `r` holds, begin to end with no result (no separate), or
 * null: some software writes a page number so, in one run.
 */
function bareFieldInstr(r) {
  const kids = elementChildren(r).filter((k) => !(k.namespaceURI === W && (k.localName === 'rPr' || k.localName === 'lastRenderedPageBreak')));
  if (kids.length < 3 || kids.some((k) => k.namespaceURI !== W || (k.localName !== 'fldChar' && k.localName !== 'instrText'))) return null;
  const types = kids.filter((k) => k.localName === 'fldChar').map((k) => wAttr(k, 'fldCharType'));
  if (types.length !== 2 || types[0] !== 'begin' || types[1] !== 'end' || kids[0].localName !== 'fldChar' || kids[kids.length - 1].localName !== 'fldChar') return null;
  return kids.filter((k) => k.localName === 'instrText').map((k) => k.textContent).join('');
}

/**
 * Marks `span` (a hidden marker) as a page number (PAGE, NUMPAGES, SECTIONPAGES) without a result,
 * if `instr` is one: it shows #, as the editor shows page numbers, and printing numbers it (see
 * editor/printBands.js). Returns whether it is one.
 */
function markBarePageField(span, instr) {
  const kind = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(instr || '')?.[1].toUpperCase();
  if (!kind) return false;
  span.classList.add('page-field-bare');
  span.dataset.instr = instr.trim();
  span.title = `${kind === 'PAGE' ? 'Page number' : kind === 'NUMPAGES' ? 'Number of pages' : 'Pages in section'} – Word works it out when it opens the file`;
  return true;
}

/** The run's only content element if it is `local` (ignoring run properties). */
function onlyChild(r, local) {
  const kids = elementChildren(r).filter((k) => !(k.namespaceURI === W && k.localName === 'rPr'));
  return kids.length === 1 && kids[0].namespaceURI === W && kids[0].localName === local ? kids[0] : null;
}

/** "footnote:2": a rendered note's kind and id. */
function noteKey(el) {
  return `${el.dataset.note}:${el.dataset.noteId}`;
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

/** A reply's marks in a part's XML, next to its comment's: inside its range, its mark just after. */
function addReplyMarkers(xml, parentId, id) {
  const find = (local) => [...xml.getElementsByTagNameNS(W, local)].find((el) => wAttr(el, 'id') === parentId);
  const marker = (local) => {
    const el = xml.createElementNS(W, `w:${local}`);
    el.setAttributeNS(W, 'w:id', id);
    return el;
  };
  const ref = find('commentReference');
  const refRun = ref?.parentNode?.localName === 'r' ? ref.parentNode : ref;
  if (!refRun) return;
  const start = find('commentRangeStart');
  const end = find('commentRangeEnd');
  if (start) start.after(marker('commentRangeStart'));
  else refRun.before(marker('commentRangeStart'));
  if (end) end.after(marker('commentRangeEnd'));
  else refRun.before(marker('commentRangeEnd'));
  const run = xml.createElementNS(W, 'w:r');
  run.appendChild(marker('commentReference'));
  refRun.after(run);
}

/** " by Ann on 2026-01-05", from a revision's author and date. */
function byWhom({ author, date }) {
  return `${author ? ` by ${author}` : ''}${date ? ` on ${date.slice(0, 10)}` : ''}`;
}
