// The parts of a minimal, valid blank Word document (A4, Calibri 11pt, Normal/Title/Heading styles),
// or of a starter document (see starters.js): its body, the styles it uses, a bulleted list, a
// footer with the page number, and its paper (A4 or Letter).
import { XML_HEADER } from './xml.js';
import { BUILTIN_STYLES } from './styles.js';

// Word 2010's namespace, for a starter's check boxes (w14:checkbox), ignorable by older readers.
const W14_NS = ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"';
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

// Word's Table Grid (as styles.ensureTableGrid adds it) and a bulleted list's numbering (Word's
// default bullet: Symbol's  at 0.63 cm, hanging), for the starters that use them.
const TABLE_GRID = '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/>'
  + '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>'
  + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('')
  + '</w:tblBorders></w:tblPr></w:style>';
const LIST_PARAGRAPH = '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>';
const BULLET_LIST = '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>'
  + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\uf0b7"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>'
  + '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/></w:rPr></w:lvl>'
  + '</w:abstractNum>';
// Questions and their answer choices, as a test paper numbers them: 1., 2., 3. and under each A., B.,
// C., D. (Word's numbered list with its second level in capital letters), hanging, so the text lines up.
const QUESTION_LIST = '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>'
  + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="397" w:hanging="397"/></w:pPr></w:lvl>'
  + '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="794" w:hanging="397"/></w:pPr></w:lvl>'
  + '</w:abstractNum>';
/**
 * A numbered list as official documents number their points (see starters.js), its first level in
 * Word's format `fmt` ('lowerLetter' for a., b.; 'decimal' for 1., 2.; 'upperLetter' for A., B.) and
 * its second in 1., 2. under it, each hanging 0.75 cm, so the text lines up after the number.
 */
const pointList = (id, fmt) => `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>`
  + `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${LIST_INDENT}" w:hanging="${LIST_INDENT}"/></w:pPr></w:lvl>`
  + `<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${2 * LIST_INDENT}" w:hanging="${LIST_INDENT}"/></w:pPr></w:lvl>`
  + '</w:abstractNum>';
/** Where a point list's text starts (its hanging indent), in twips: 0.75 cm. */
export const LIST_INDENT = 425;
/** The numId of the `i`th of a starter's point lists (`lists`, see blankDocxFiles). */
export const listId = (i) => 20 + i;

/**
 * The numbering part: the bulleted list (numId 1); for a test paper, `questions` lists of
 * questions (numIds 2, 3…), each starting again at 1, as each part of a test numbers its own; and
 * an official document's `lists`, a format each (see pointList), numbered from listId(0) on, each
 * starting at 1 (the points of Menimbang, then of Mengingat…).
 */
const numberingXml = (questions = 0, lists = []) => `<w:numbering ${W_NS}>${BULLET_LIST}${questions ? QUESTION_LIST : ''}`
  + lists.map((fmt, i) => pointList(listId(i), fmt)).join('')
  + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
  + Array.from({ length: questions }, (_, i) => `<w:num w:numId="${i + 2}"><w:abstractNumId w:val="1"/>${i ? '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>' : ''}</w:num>`).join('')
  + lists.map((_, i) => `<w:num w:numId="${listId(i)}"><w:abstractNumId w:val="${listId(i)}"/></w:num>`).join('')
  + '</w:numbering>';
/** A section's reference to the header blankDocxFiles makes (`header`), for a section in the body to have it too. */
export const HEADER_REF = '<w:headerReference w:type="default" r:id="rId5"/>';
const FOOTER = `<w:ftr ${W_NS}><w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>`
  + '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE   \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
  + '<w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>';
const PAPERS = { A4: { w: 11906, h: 16838 }, Letter: { w: 12240, h: 15840 } };

/** The width between the margins, in twips, of a starter on `paper` (2.54 cm margins), turned for `landscape`. */
export const textWidth = (paper = 'A4', landscape = false) => (landscape ? PAPERS[paper].h : PAPERS[paper].w) - 2 * 1440;

/**
 * A section's page and columns, as Word writes them in a w:sectPr: on `paper`, turned for
 * `landscape`, with `margins` (twips), starting as `type` says ('continuous' for a section on the
 * same page; none, a new page), in `cols` equal columns with a line between them (`sep`); its page
 * numbers in Word's format `numFmt` ('numberInDash' for - 2 -), from `numStart` (else on from the
 * section before); its header and footer `refs` (references, as HEADER_REF), with a different first
 * page (`titlePg`) where asked.
 */
export function sectionXml({ paper = 'A4', landscape = false, margins = null, type = null, cols = 1, sep = false, numFmt = null, numStart = null, refs = '', titlePg = false } = {}) {
  const { w, h } = PAPERS[paper] || PAPERS.A4;
  const m = { top: 1440, right: 1440, bottom: 1440, left: 1440, ...margins };
  const size = landscape ? `<w:pgSz w:w="${h}" w:h="${w}" w:orient="landscape"/>` : `<w:pgSz w:w="${w}" w:h="${h}"/>`;
  const num = numFmt || numStart != null ? `<w:pgNumType${numFmt ? ` w:fmt="${numFmt}"` : ''}${numStart != null ? ` w:start="${numStart}"` : ''}/>` : '';
  return `${refs}${type ? `<w:type w:val="${type}"/>` : ''}${size}<w:pgMar w:top="${m.top}" w:right="${m.right}" w:bottom="${m.bottom}" w:left="${m.left}" w:header="708" w:footer="708" w:gutter="0"/>`
    + `${num}<w:cols${cols > 1 ? ` w:num="${cols}"` : ''} w:space="708"${cols > 1 && sep ? ' w:sep="1"' : ''}/>${titlePg ? '<w:titlePg/>' : ''}`;
}

/**
 * The files of a new document: blank, or with `body` (the body's XML, without its sectPr), the
 * `styles` it needs beyond Normal, Title and Heading 1–3 ('tableGrid', 'listParagraph', 'footer'),
 * a bulleted list (`bullets`), a centred page number in the footer (`footer`), on `paper`. A
 * starter with a look of its own (a manuscript) also gives `font` ({ name, sz } in half-points: the
 * document's default), `normal` (Normal's w:pPr, its inner XML), `heading1` (Heading 1's own w:pPr
 * and w:rPr, in place of Word's), `header` (the paragraphs of a header on every page) and `titlePg`
 * (a different first page, its header left empty, as a title page has it). `margins` ({ top, right,
 * bottom, left } in twips) are 2.54 cm unless given; `language` is the text's (w:lang), en-US unless given.
 * A test paper also gives `questions` (how many numbered lists of questions it has, see
 * numberingXml), an official document `lists` (its numbered points' formats, see numberingXml) and `section` (sectionXml's options for the final section: a page turned
 * `landscape`, or a section that starts on the same page, after others in the body in columns).
 */
export function blankDocxFiles({ body = '<w:p/>', styles = [], bullets = false, questions = 0, lists = [], footer = false, paper = 'A4', title = '', font = null, normal = '', heading1 = null, header = '', titlePg = false, margins = null, language = 'en-US', section = {} } = {}) {
  const numbered = bullets || questions > 0 || lists.length > 0;
  const builtin = (k) => BUILTIN_STYLES[k].xml('<w:basedOn w:val="Normal"/>').replace('__NORMAL__', 'Normal');
  const h1 = heading1
    ? builtin('h1').replace(/<w:pPr>.*?<\/w:pPr>\s*<w:rPr>.*?<\/w:rPr>/s, `<w:pPr>${heading1.pPr}</w:pPr><w:rPr>${heading1.rPr}</w:rPr>`)
    : builtin('h1');
  const builtins = [builtin('title'), h1, builtin('h2'), builtin('h3')].join('')
    + (styles.includes('tableGrid') ? TABLE_GRID : '')
    + (bullets || styles.includes('listParagraph') ? LIST_PARAGRAPH : '')
    + (footer ? builtin('footer') : '')
    + (header ? builtin('header') : '');
  const fontName = font?.name || 'Calibri';
  const fontSize = font?.sz || 22;
  const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const files = {
    '[Content_Types].xml': XML_HEADER + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Default Extension="jpeg" ContentType="image/jpeg"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
${numbered ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>\n' : ''}${footer ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>\n' : ''}${header ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>\n' : ''}</Types>`,
    '_rels/.rels': XML_HEADER + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`,
    'docProps/core.xml': XML_HEADER + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>${esc(title)}</dc:title><dc:creator></dc:creator>
<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`,
    'docProps/app.xml': XML_HEADER + `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Intact</Application></Properties>`,
    'word/_rels/document.xml.rels': XML_HEADER + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
${numbered ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>\n' : ''}${footer ? '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>\n' : ''}${header ? '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>\n' : ''}</Relationships>`,
    'word/settings.xml': XML_HEADER + `<w:settings ${W_NS}><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
    'word/document.xml': XML_HEADER + `<w:document ${W_NS}${body.includes('<w14:') ? W14_NS : ''}><w:body>${body}
<w:sectPr>${sectionXml({ paper, margins, ...section, refs: `${header ? HEADER_REF : ''}${footer ? '<w:footerReference w:type="default" r:id="rId4"/>' : ''}`, titlePg })}</w:sectPr>
</w:body></w:document>`,
    'word/styles.xml': XML_HEADER + `<w:styles ${W_NS}>
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="${fontName}" w:hAnsi="${fontName}" w:eastAsia="${fontName}" w:cs="${fontName}"/><w:sz w:val="${fontSize}"/><w:szCs w:val="${fontSize}"/><w:lang w:val="${/^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*$/.test(language) ? language : 'en-US'}"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>${normal ? `<w:pPr>${normal}</w:pPr>` : ''}</w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
${builtins.replace(/w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/g, font ? `w:ascii="${fontName}" w:hAnsi="${fontName}"` : 'w:ascii="Calibri Light" w:hAnsi="Calibri Light"')}
</w:styles>`,
  };
  if (numbered) files['word/numbering.xml'] = XML_HEADER + numberingXml(questions, lists);
  if (footer) files['word/footer1.xml'] = XML_HEADER + FOOTER;
  if (header) files['word/header1.xml'] = XML_HEADER + `<w:hdr ${W_NS}>${header}</w:hdr>`;
  return files;
}
