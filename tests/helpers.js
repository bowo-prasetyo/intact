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

// A 1×1 transparent PNG.
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAwS2OUAAAAABJRU5ErkJggg==', 'base64');

/** Builds a .docx (as a Buffer) whose body is `bodyXml`, with numbering, a hyperlink target and an image. */
export async function buildDocx(bodyXml) {
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
</Relationships>`);
  zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
</w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** Opens a .docx buffer in the editor via the file picker. */
export async function openInEditor(page, buffer, name = 'fixture.docx') {
  page.on('dialog', (d) => d.accept());
  await page.goto('./');
  await page.setInputFiles('input[type=file]', { name, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer });
  await expect(page.locator('.docx-page')).toBeVisible();
}

export async function documentXmlOf(buffer) {
  return (await JSZip.loadAsync(buffer)).file('word/document.xml').async('string');
}

/** Clicks Download and returns the saved package's word/document.xml. */
export async function download(page) {
  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: 'Download' }).click()]);
  const zip = await JSZip.loadAsync(fs.readFileSync(await dl.path()));
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
    return kids(doc.getElementsByTagNameNS(W, 'body')[0]).map(generic);
  }, xml);
}
