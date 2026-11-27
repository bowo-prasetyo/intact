// Shared helpers: build .docx fixtures from XML, open them in the app, download the result and
// compare documents by meaning rather than by bytes.
import fs from 'node:fs';
import JSZip from 'jszip';
import { expect } from '@playwright/test';

export const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
  + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
  + 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

const PLACEHOLDER = '<w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>';

/** One inline content control of each kind, as Word writes them in a form. */
export const FIELDS = `<w:sdt><w:sdtPr><w:rPr><w:b/></w:rPr><w:alias w:val="Client name"/><w:tag w:val="client"/><w:id w:val="101"/><w:showingPlcHdr/><w:text/></w:sdtPr>`
  + `<w:sdtContent><w:r>${PLACEHOLDER}<w:t>Click or tap here to enter text.</w:t></w:r></w:sdtContent></w:sdt>`
  + `<w:r><w:t xml:space="preserve"> </w:t></w:r>`
  + `<w:sdt><w:sdtPr><w:id w:val="102"/><w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr>`
  + `<w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>`
  + `<w:sdt><w:sdtPr><w:alias w:val="Plan"/><w:id w:val="103"/><w:showingPlcHdr/><w:dropDownList w:lastValue=""><w:listItem w:displayText="Basic plan" w:value="basic"/><w:listItem w:displayText="Premium plan" w:value="premium"/></w:dropDownList></w:sdtPr>`
  + `<w:sdtContent><w:r>${PLACEHOLDER}<w:t>Choose an item.</w:t></w:r></w:sdtContent></w:sdt>`
  + `<w:sdt><w:sdtPr><w:alias w:val="Start date"/><w:id w:val="104"/><w:showingPlcHdr/><w:date><w:dateFormat w:val="d MMMM yyyy"/><w:lid w:val="en-GB"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date></w:sdtPr>`
  + `<w:sdtContent><w:r>${PLACEHOLDER}<w:t>Click or tap to enter a date.</w:t></w:r></w:sdtContent></w:sdt>`
  + `<w:sdt><w:sdtPr><w:alias w:val="Company"/><w:id w:val="105"/><w:dataBinding w:xpath="/root/company" w:storeItemID="{00000000-0000-0000-0000-000000000001}"/><w:text/></w:sdtPr>`
  + `<w:sdtContent><w:r><w:t>ACME Corp</w:t></w:r></w:sdtContent></w:sdt>`;

const KS_NUM = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';

// One paragraph per feature, so a failure points at the feature.
/** A document with one paragraph per feature, for round-trip and review tests. */
export const KITCHEN_SINK = `
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:after="120"/></w:pPr>
  <w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b w:val="1"/><w:color w:val="C00000"/><w:sz w:val="28"/></w:rPr><w:t>Bold red Georgia</w:t></w:r>
  <w:r><w:t xml:space="preserve"> plain </w:t></w:r>
  <w:r><w:rPr><w:i/><w:u w:val="double"/></w:rPr><w:t>italic double-underline</w:t></w:r>
  <w:r><w:rPr><w:rStyle w:val="Strong"/><w:b w:val="0"/></w:rPr><w:t xml:space="preserve"> style with bold off</w:t></w:r>
  <w:r><w:rPr><w:dstrike/></w:rPr><w:t xml:space="preserve"> double strike</w:t></w:r>
  <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Tab:</w:t></w:r><w:r><w:tab/><w:t>after tab, line</w:t><w:br/><w:t>break, §\u00a05 (non-breaking space) and a ZWSP\u200bhere</w:t></w:r></w:p>
<w:p>${KS_NUM}<w:r><w:t>First item</w:t></w:r></w:p>
<w:p>${KS_NUM}<w:r><w:t>Second item</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>Bullet from a list style</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:ilvl w:val="0"/></w:numPr><w:ind w:left="500"/></w:pPr><w:r><w:t>Style bullet with its own level and indent</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>List style with numbering switched off</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Link: </w:t></w:r><w:hyperlink r:id="rId3" w:history="1"><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>example</w:t></w:r></w:hyperlink><w:bookmarkStart w:id="0" w:name="mark"/><w:r><w:t xml:space="preserve"> bookmarked</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>
<w:p><w:r><w:t xml:space="preserve">Kept </w:t></w:r><w:ins w:id="1" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:t>inserted by Ann</w:t></w:r></w:ins><w:del w:id="2" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve"> deleted by Ann</w:delText></w:r></w:del></w:p>
<w:p><w:r><w:t xml:space="preserve">Page: </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> of </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Image: </w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="95250" cy="95250"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId4"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Text box: </w:t></w:r><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml"><mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="1828800" cy="457200"/><wp:docPr id="2" name="Text Box 2"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Boxed</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape style="width:144pt;height:36pt" fillcolor="#fff2cc"><v:textbox><w:txbxContent><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Boxed</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>
<w:p><w:r><w:t>Before break</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>After break</w:t></w:r></w:p>
<w:sdt><w:sdtPr><w:alias w:val="Clause"/><w:tag w:val="clause"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Clause in a content control</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Locked clause</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:p><w:r><w:t xml:space="preserve">Fields: </w:t></w:r>${FIELDS}</w:p>
<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
  <w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Merged header</w:t></w:r></w:p></w:tc></w:tr>
  <w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Tall</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>
  <w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>B3</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
  <w:sdt><w:sdtPr><w:alias w:val="Rows"/><w:id w:val="301"/></w:sdtPr><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>Row field</w:t></w:r></w:p></w:tc>
  <w:sdt><w:sdtPr><w:alias w:val="Cell"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:tc><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Cell field</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt></w:tr></w:sdtContent></w:sdt>
</w:tbl>
<w:p><w:pPr><w:keepNext/><w:keepLines/></w:pPr><w:bookmarkStart w:id="9" w:name="Seals"/><w:r><w:t xml:space="preserve">Kept with the next, indexed</w:t></w:r><w:bookmarkEnd w:id="9"/><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> XE "Pump:seals" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve">, version </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DOCPROPERTY  Version  \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1.0</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve">, see </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF Seals \\r \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
<w:p><w:r><w:t>Last paragraph</w:t></w:r></w:p>`;

// A 1×1 transparent PNG.
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAwS2OUAAAAABJRU5ErkJggg==', 'base64');

/**
 * Builds a .docx (as a Buffer) whose body is `bodyXml`, with numbering, a hyperlink target and an
 * image. `parts` adds parts ({ 'word/header1.xml': xml }), `rels` relationships of the main document
 * and `sectPr` the start of its section properties (such as header references).
 */
export async function buildDocx(bodyXml, { parts = {}, rels = '', sectPr = '' } = {}) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
${rels}</Relationships>`);
  for (const [path, content] of Object.entries(parts)) zip.file(path, content);
  if (!parts['word/styles.xml']) zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W_NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr><w:ind w:left="360" w:hanging="360"/></w:pPr></w:style>
<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:rPr><w:color w:val="808080"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>
</w:styles>`);
  zip.file('word/numbering.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${W_NS}>
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`);
  zip.file('word/media/image1.png', PNG);
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W_NS}><w:body>${bodyXml}
<w:sectPr>${sectPr}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
</w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** Opens a .docx buffer in the editor via the file picker. */
/**
 * The names a download given `expected` may arrive under in `browserName`. WebKit's build for Linux
 * on CI writes the spaces of a suggested file name as underscores ("Lease_v2.docx"); elsewhere, and
 * in Safari on a Mac, the name is as the page gave it. Anything else is still a wrong name.
 */
export function downloadNames(expected, browserName) {
  return browserName === 'webkit' ? [expected, expected.replace(/ /g, '_')] : [expected];
}

export async function openInEditor(page, buffer, name = 'fixture.docx') {
  page.on('dialog', (d) => d.accept());
  await page.goto('./');
  await page.setInputFiles('input[type=file]', { name, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer });
  await expect(page.locator('.docx-page')).toBeVisible();
}

export async function documentXmlOf(buffer) {
  return (await JSZip.loadAsync(buffer)).file('word/document.xml').async('string');
}

/** Clicks Download, then Download in "Review your changes", and returns the downloaded file's path. */
export async function downloadFile(page) {
  await page.locator('.app-header').getByRole('button', { name: 'Download' }).click();
  const review = page.getByRole('dialog', { name: 'Review your changes' });
  const [dl] = await Promise.all([page.waitForEvent('download'), review.getByRole('button', { name: 'Download' }).click()]);
  return dl.path();
}

/** Clicks Download (through the review) and returns the saved package and its word/document.xml. */
export async function download(page) {
  const path = await downloadFile(page);
  const zip = await JSZip.loadAsync(fs.readFileSync(path));
  return { zip, xml: await zip.file('word/document.xml').async('string') };
}

/**
 * Reduces document.xml to what it means, so a faithful save compares equal even though the bytes
 * differ. Normalisations, each of which Word treats as equivalent:
 *  - adjacent runs with identical properties are merged, and empty runs are dropped;
 *  - <w:t> text is compared, not its xml:space attribute or how it is split;
 *  - on/off properties written as <w:b w:val="1"/>, "true" or "on" equal <w:b/>;
 *  - proofErr and lastRenderedPageBreak (spelling marks and layout caches Word regenerates) and
 *    revision-session ids (rsid*) and paragraph ids on <w:p>/<w:r> are ignored.
 * Runs in the page because Node has no DOMParser.
 */
export function canonicalize(page, xml) {
  return page.evaluate((text) => {
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const IGNORED = new Set(['proofErr', 'lastRenderedPageBreak']);
    const isW = (n, local) => n.nodeType === 1 && n.namespaceURI === W && (!local || n.localName === local);
    const kids = (n) => [...n.childNodes].filter((c) => c.nodeType === 1 && !(isW(c) && IGNORED.has(c.localName)));
    const attrs = (el) => [...el.attributes]
      .filter((a) => !/^w:rsid|^w14:paraId$|^w14:textId$/.test(a.name) && !a.name.startsWith('xmlns'))
      .map((a) => [a.name, a.value])
      .filter(([n, v]) => !(n === 'w:val' && ['1', 'true', 'on'].includes(v)))
      .sort();

    const generic = (el) => {
      if (isW(el, 'p')) return { p: paragraph(el) };
      const children = kids(el);
      return children.length
        ? { tag: el.nodeName, attrs: attrs(el), kids: children.map(generic) }
        : { tag: el.nodeName, attrs: attrs(el), text: el.textContent };
    };

    const inlines = (container) => {
      const out = [];
      const pushText = (rPr, t) => {
        if (!t) return; // empty runs are dropped before merging, so they don't keep neighbours apart
        const last = out[out.length - 1];
        if (last && last.text !== undefined && last.rPr === rPr) last.text += t;
        else out.push({ rPr, text: t });
      };
      for (const c of kids(container)) {
        if (isW(c, 'pPr')) continue;
        if (isW(c, 'r')) {
          const rPrEl = kids(c).find((k) => isW(k, 'rPr'));
          const rPr = rPrEl && kids(rPrEl).length ? JSON.stringify(generic(rPrEl)) : '';
          for (const k of kids(c)) {
            if (isW(k, 'rPr')) continue;
            if (isW(k, 't')) pushText(rPr, k.textContent);
            else if (isW(k, 'tab')) pushText(rPr, '\t');
            else if ((isW(k, 'br') && !k.getAttributeNS(W, 'type')) || isW(k, 'cr')) pushText(rPr, '\n');
            else out.push({ rPr, el: generic(k) });
          }
        } else if (isW(c, 'hyperlink')) {
          out.push({ hyperlink: attrs(c), kids: inlines(c) });
        } else if (isW(c, 'sdt')) {
          const pr = kids(c).find((k) => isW(k, 'sdtPr'));
          const content = kids(c).find((k) => isW(k, 'sdtContent'));
          out.push({ sdt: pr ? generic(pr) : null, kids: content ? inlines(content) : [] });
        } else {
          out.push(generic(c));
        }
      }
      return out.filter((x) => x.text !== '');
    };

    const paragraph = (p) => {
      const pPr = kids(p).find((k) => isW(k, 'pPr'));
      return { pPr: pPr && kids(pPr).length ? generic(pPr) : null, inlines: inlines(p) };
    };

    const doc = new DOMParser().parseFromString(text, 'application/xml');
    // The body's blocks; in a header, footer or notes part, its root's (paragraphs, tables, notes).
    return kids(doc.getElementsByTagNameNS(W, 'body')[0] || doc.documentElement).map(generic);
  }, xml);
}
