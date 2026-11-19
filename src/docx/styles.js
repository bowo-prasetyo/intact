// Reads styles.xml / theme1.xml to (a) find heading styles and (b) produce CSS that makes
// the editor look roughly like Word. Also injects built-in heading styles that are missing.
import { NS, wChild, wChildren, wAttr, onOff, parseXml } from './xml.js';

const twipsToPt = (v) => Number(v) / 20;

export const BUILTIN_STYLES = {
  title: { name: 'title', id: 'Title', xml: (basedOn) => `
    <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>${basedOn}<w:next w:val="__NORMAL__"/><w:uiPriority w:val="10"/><w:qFormat/>
    <w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr>
    <w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>` },
  ...Object.fromEntries([1, 2, 3, 4, 5, 6].map((n) => {
    const sz = [40, 32, 28, 24, 22, 22][n - 1];
    const before = [360, 160, 160, 80, 80, 40][n - 1];
    return [`h${n}`, { name: `heading ${n}`, id: `Heading${n}`, xml: (basedOn) => `
    <w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/>${basedOn}<w:next w:val="__NORMAL__"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="80"/><w:outlineLvl w:val="${n - 1}"/></w:pPr>
    <w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/>${n >= 4 ? '<w:i/>' : ''}<w:color w:val="2F5496"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr></w:style>` }];
  })),
  // Added with what uses them: a header (for a watermark) or footer made here, and new footnotes, as Word defines them.
  header: { name: 'header', id: 'Header', xml: (basedOn) => `
    <w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/>${basedOn}<w:uiPriority w:val="99"/><w:unhideWhenUsed/>
    <w:pPr><w:tabs><w:tab w:val="center" w:pos="4513"/><w:tab w:val="right" w:pos="9026"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>` },
  footer: { name: 'footer', id: 'Footer', xml: (basedOn) => `
    <w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/>${basedOn}<w:uiPriority w:val="99"/><w:unhideWhenUsed/>
    <w:pPr><w:tabs><w:tab w:val="center" w:pos="4513"/><w:tab w:val="right" w:pos="9026"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>` },
  footnoteText: { name: 'footnote text', id: 'FootnoteText', xml: (basedOn) => `
    <w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/>${basedOn}<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>
    <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>` },
  footnoteReference: { name: 'footnote reference', id: 'FootnoteReference', type: 'character', xml: () => `
    <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>` },
  endnoteText: { name: 'endnote text', id: 'EndnoteText', xml: (basedOn) => `
    <w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/>${basedOn}<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>
    <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>` },
  endnoteReference: { name: 'endnote reference', id: 'EndnoteReference', type: 'character', xml: () => `
    <w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>` },
  // A table of contents' entries, by level, as Word defines them (added for a level its TOC lacks).
  ...Object.fromEntries([1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => [`toc${n}`, { name: `toc ${n}`, id: `TOC${n}`, xml: (basedOn) => `
    <w:style w:type="paragraph" w:styleId="TOC${n}"><w:name w:val="toc ${n}"/>${basedOn}<w:next w:val="__NORMAL__"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/>
    <w:pPr><w:spacing w:after="100"/>${n > 1 ? `<w:ind w:left="${220 * (n - 1)}"/>` : ''}</w:pPr></w:style>` }])),
};

export class StyleSheet {
  constructor(stylesDoc, themeDoc) {
    this.doc = stylesDoc;
    this.changed = false; // set when a style is added; styles.xml is only rewritten then
    this.fonts = { major: 'Calibri Light', minor: 'Calibri' };
    if (themeDoc) {
      const major = themeDoc.getElementsByTagNameNS(NS.a, 'majorFont')[0];
      const minor = themeDoc.getElementsByTagNameNS(NS.a, 'minorFont')[0];
      const latin = (f) => f && f.getElementsByTagNameNS(NS.a, 'latin')[0]?.getAttribute('typeface');
      this.fonts.major = latin(major) || this.fonts.major;
      this.fonts.minor = latin(minor) || this.fonts.minor;
    }
    this.themeColors = themeColors(themeDoc);
    this.reindex();
  }

  reindex() {
    this.byId = new Map();
    this.defaultParagraph = null;
    const root = this.doc.documentElement;
    for (const s of wChildren(root, 'style')) {
      const info = {
        id: wAttr(s, 'styleId'),
        type: wAttr(s, 'type'),
        name: wAttr(wChild(s, 'name'), 'val') || '',
        basedOn: wAttr(wChild(s, 'basedOn'), 'val'),
        pPr: wChild(s, 'pPr'),
        rPr: wChild(s, 'rPr'),
        isDefault: ['1', 'true', 'on'].includes(wAttr(s, 'default') || ''),
        // For the style menu: hidden from Word's lists, in its style gallery, and its place there.
        hidden: onOff(wChild(s, 'semiHidden')) === true || onOff(wChild(s, 'hidden')) === true,
        unhideWhenUsed: onOff(wChild(s, 'unhideWhenUsed')) === true,
        quick: onOff(wChild(s, 'qFormat')) === true,
        priority: Number(wAttr(wChild(s, 'uiPriority'), 'val') ?? 99),
      };
      this.byId.set(info.id, info);
      if (info.type === 'paragraph' && info.isDefault) this.defaultParagraph = info.id;
    }
    const dd = wChild(root, 'docDefaults');
    this.defaultRPr = wChild(wChild(dd, 'rPrDefault'), 'rPr');
    this.defaultPPr = wChild(wChild(dd, 'pPrDefault'), 'pPr');
  }

  /** Style id for a builtin key ('title', 'h1'..'h6'), found by name (works for localized ids). */
  findBuiltin(key) {
    const def = BUILTIN_STYLES[key];
    for (const s of this.byId.values()) {
      if (s.type === (def.type || 'paragraph') && s.name.toLowerCase() === def.name) return s.id;
    }
    return null;
  }

  /** Makes sure the builtin style exists in styles.xml, returning its id. */
  ensureBuiltin(key) {
    const found = this.findBuiltin(key);
    if (found) return found;
    const def = BUILTIN_STYLES[key];
    let id = def.id;
    while (this.byId.has(id)) id += '0';
    const normal = this.defaultParagraph;
    const basedOn = normal ? `<w:basedOn w:val="${normal}"/>` : '';
    let xml = def.xml(basedOn).replace(`w:styleId="${def.id}"`, `w:styleId="${id}"`);
    xml = normal ? xml.replace('__NORMAL__', normal) : xml.replace(/<w:next w:val="__NORMAL__"\/>/, '');
    const frag = parseXml(`<root xmlns:w="${NS.w}">${xml}</root>`).documentElement;
    const el = this.doc.importNode(wChild(frag, 'style'), true);
    this.doc.documentElement.appendChild(el);
    this.changed = true;
    this.reindex();
    return id;
  }

  /**
   * The id of Word's "Table Grid" table style (thin borders on every cell), added to styles.xml if
   * it isn't there, as Word adds it when it inserts a table.
   */
  ensureTableGrid() {
    for (const s of this.byId.values()) if (s.type === 'table' && s.name.toLowerCase() === 'table grid') return s.id;
    let id = 'TableGrid';
    while (this.byId.has(id)) id += '0';
    const normal = [...this.byId.values()].find((s) => s.type === 'table' && (s.isDefault || s.name.toLowerCase() === 'normal table'));
    const border = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;
    const xml = `<w:style w:type="table" w:styleId="${id}"><w:name w:val="Table Grid"/>${normal ? `<w:basedOn w:val="${normal.id}"/>` : ''}<w:uiPriority w:val="39"/>`
      + '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
      + `<w:tblPr><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders></w:tblPr></w:style>`;
    const frag = parseXml(`<root xmlns:w="${NS.w}">${xml}</root>`).documentElement;
    this.doc.documentElement.appendChild(this.doc.importNode(wChild(frag, 'style'), true));
    this.changed = true;
    this.reindex();
    return id;
  }

  /**
   * The document's own paragraph styles for the style menu, after Normal, Title and Heading 1–3:
   * [{ id, label }]. As Word's style gallery offers them: headings (4–9) the file defines, styles
   * in the gallery (w:qFormat) that aren't hidden, and any style the text already uses (`used`, a
   * Set of ids). Headings first, then in Word's order (w:uiPriority), then by name.
   */
  paragraphChoices(used = new Set()) {
    const out = [];
    for (const s of this.byId.values()) {
      if (s.type !== 'paragraph' || !s.id || s.id === this.defaultParagraph) continue;
      if (['title', 'h1', 'h2', 'h3'].includes(this.kindOf(s.id))) continue;
      const heading = /^heading ([1-9])$/i.exec(s.name);
      const offered = used.has(s.id) || heading || (s.quick && (!s.hidden || s.unhideWhenUsed));
      if (offered) out.push({ id: s.id, label: this.label(s.id), heading: heading ? Number(heading[1]) : 10, priority: s.priority });
    }
    out.sort((a, b) => a.heading - b.heading || a.priority - b.priority || a.label.localeCompare(b.label));
    return out.map(({ id, label }) => ({ id, label }));
  }

  /** A style's name as Word shows it: built-in names are stored in lower case ("heading 4", "toc 1"). */
  label(styleId) {
    const name = this.byId.get(styleId)?.name || styleId || '';
    if (/[A-Z]/.test(name)) return name;
    return name.replace(/\b\w/g, (c) => c.toUpperCase()).replace(/^Toc\b/, 'TOC');
  }

  /**
   * The id of Word's "Hyperlink" character style (blue, underlined), added to styles.xml if it isn't
   * there, as Word adds it with the first link.
   */
  ensureHyperlink() {
    for (const s of this.byId.values()) if (s.type === 'character' && s.name.toLowerCase() === 'hyperlink') return s.id;
    let id = 'Hyperlink';
    while (this.byId.has(id)) id += '0';
    const font = [...this.byId.values()].find((s) => s.type === 'character' && s.isDefault);
    const xml = `<w:style w:type="character" w:styleId="${id}"><w:name w:val="Hyperlink"/>${font ? `<w:basedOn w:val="${font.id}"/>` : ''}<w:uiPriority w:val="99"/><w:unhideWhenUsed/>`
      + '<w:rPr><w:color w:val="0563C1" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr></w:style>';
    const frag = parseXml(`<root xmlns:w="${NS.w}">${xml}</root>`).documentElement;
    this.doc.documentElement.appendChild(this.doc.importNode(wChild(frag, 'style'), true));
    this.changed = true;
    this.reindex();
    return id;
  }

  /**
   * The document's language, as a BCP 47 tag ("en-GB"): its default paragraph style's, or else the
   * document defaults'. Null if it names none.
   */
  language() {
    for (const s of this.chain(this.defaultParagraph).reverse()) {
      const lang = wAttr(wChild(s.rPr, 'lang'), 'val');
      if (lang) return lang;
    }
    return wAttr(wChild(this.defaultRPr, 'lang'), 'val') || null;
  }

  /** Classifies a paragraph style id as 'title', 'h1'..'h6' or null. */
  kindOf(styleId) {
    const s = styleId && this.byId.get(styleId);
    if (!s) return null;
    const name = s.name.toLowerCase();
    if (name === 'title') return 'title';
    const m = /^heading ([1-6])$/.exec(name);
    return m ? `h${m[1]}` : null;
  }

  /** Whether the style (or its basedOn chain) ignores spacing next to paragraphs of the same style. */
  contextualSpacing(styleId) {
    const chain = this.chain(styleId || this.defaultParagraph);
    for (let i = chain.length - 1; i >= 0; i--) {
      const v = onOff(wChild(chain[i].pPr, 'contextualSpacing'));
      if (v !== undefined) return v;
    }
    return !!onOff(wChild(this.defaultPPr, 'contextualSpacing'));
  }

  isListParagraph(styleId) {
    return this.byId.get(styleId)?.name.toLowerCase() === 'list paragraph';
  }

  /** Walks the basedOn chain (most-derived last). */
  chain(styleId) {
    const out = [];
    const seen = new Set();
    let s = styleId && this.byId.get(styleId);
    while (s && !seen.has(s.id)) {
      seen.add(s.id);
      out.unshift(s);
      s = s.basedOn && this.byId.get(s.basedOn);
    }
    return out;
  }

  /** numPr inherited from the paragraph style chain, if any. */
  styleNumPr(styleId) {
    const chain = this.chain(styleId || this.defaultParagraph);
    for (let i = chain.length - 1; i >= 0; i--) {
      const np = wChild(chain[i].pPr, 'numPr');
      if (np) return np;
    }
    return null;
  }

  fontFor(rFonts) {
    if (!rFonts) return null;
    const ascii = wAttr(rFonts, 'ascii') || wAttr(rFonts, 'hAnsi');
    if (ascii) return ascii;
    const theme = wAttr(rFonts, 'asciiTheme') || wAttr(rFonts, 'hAnsiTheme');
    if (theme) return theme.startsWith('major') ? this.fonts.major : this.fonts.minor;
    return null;
  }

  /**
   * CSS declarations for run properties. `withToggles` includes what the editor shows as tags
   * instead: bold, italic, underline, strikethrough, superscript, subscript and highlight.
   */
  rPrCss(rPr, withToggles = true) {
    const css = {};
    if (!rPr) return css;
    const font = this.fontFor(wChild(rPr, 'rFonts'));
    if (font) css['font-family'] = `"${font}", sans-serif`;
    const sz = wAttr(wChild(rPr, 'sz'), 'val');
    if (sz) css['font-size'] = `${Number(sz) / 2}pt`;
    const color = this.runColor(wChild(rPr, 'color'));
    if (color) css.color = color;
    // Character spacing (twentieths of a point) and text raised or lowered (half-points).
    const spacing = Number(wAttr(wChild(rPr, 'spacing'), 'val'));
    if (spacing) css['letter-spacing'] = `${spacing / 20}pt`;
    const position = Number(wAttr(wChild(rPr, 'position'), 'val'));
    if (position) css['vertical-align'] = `${position / 2}pt`;
    const hl = wAttr(wChild(rPr, 'highlight'), 'val');
    if (withToggles && hl && hl !== 'none') css['background-color'] = HIGHLIGHT[hl] || hl;
    const shd = this.shadingColor(wChild(rPr, 'shd'));
    if (!hl && shd) css['background-color'] = shd;
    const va = wAttr(wChild(rPr, 'vertAlign'), 'val');
    if (withToggles && va === 'superscript') Object.assign(css, { 'vertical-align': 'super', 'font-size': 'smaller' });
    if (withToggles && va === 'subscript') Object.assign(css, { 'vertical-align': 'sub', 'font-size': 'smaller' });
    if (onOff(wChild(rPr, 'caps'))) css['text-transform'] = 'uppercase';
    if (onOff(wChild(rPr, 'smallCaps'))) css['font-variant'] = 'small-caps';
    if (onOff(wChild(rPr, 'vanish'))) css.opacity = '0.5';
    if (withToggles) {
      const b = onOff(wChild(rPr, 'b'));
      if (b !== undefined) css['font-weight'] = b ? 'bold' : 'normal';
      const i = onOff(wChild(rPr, 'i'));
      if (i !== undefined) css['font-style'] = i ? 'italic' : 'normal';
      const deco = [];
      const u = wAttr(wChild(rPr, 'u'), 'val');
      if (wChild(rPr, 'u') && u !== 'none') deco.push('underline');
      if (onOff(wChild(rPr, 'strike')) || onOff(wChild(rPr, 'dstrike'))) deco.push('line-through');
      if (deco.length) css['text-decoration'] = deco.join(' ');
    }
    return css;
  }

  /** CSS declarations for paragraph properties (spacing, indentation, alignment). */
  pPrCss(pPr, rtl = onOff(wChild(pPr, 'bidi')) === true) {
    const css = {};
    if (!pPr) return css;
    // A right-to-left paragraph (Arabic, Hebrew, Persian, Urdu…).
    if (onOff(wChild(pPr, 'bidi')) === true) css.direction = 'rtl';
    const sp = wChild(pPr, 'spacing');
    if (sp) {
      const before = wAttr(sp, 'before');
      const after = wAttr(sp, 'after');
      if (before != null) css['margin-top'] = `${twipsToPt(before)}pt`;
      if (after != null) css['margin-bottom'] = `${twipsToPt(after)}pt`;
      const line = wAttr(sp, 'line');
      const rule = wAttr(sp, 'lineRule') || 'auto';
      if (line != null) {
        css['line-height'] = rule === 'auto' ? String((Number(line) / 240) * 1.17) : `${twipsToPt(line)}pt`;
      }
    }
    const ind = wChild(pPr, 'ind');
    if (ind) {
      const left = wAttr(ind, 'left') ?? wAttr(ind, 'start');
      const right = wAttr(ind, 'right') ?? wAttr(ind, 'end');
      if (left != null) css['padding-left'] = `${twipsToPt(left)}pt`;
      if (right != null) css['padding-right'] = `${twipsToPt(right)}pt`;
      const first = wAttr(ind, 'firstLine');
      const hanging = wAttr(ind, 'hanging');
      if (hanging != null) css['text-indent'] = `-${twipsToPt(hanging)}pt`;
      else if (first != null) css['text-indent'] = `${twipsToPt(first)}pt`;
    }
    const jc = wAttr(wChild(pPr, 'jc'), 'val');
    if (jc) css['text-align'] = jcToCss(jc, rtl);
    // Borders and shading. Word draws a border its "space" (points) away from the text.
    const pBdr = wChild(pPr, 'pBdr');
    for (const side of ['top', 'left', 'bottom', 'right']) {
      const b = wChild(pBdr, side) || (side === 'left' && wChild(pBdr, 'start')) || (side === 'right' && wChild(pBdr, 'end'));
      if (!b) continue;
      const border = borderCss(b);
      css[`border-${side}`] = border || 'none';
      if (!border) continue;
      const space = `${Number(wAttr(b, 'space')) || 0}pt`;
      if (side === 'top' || side === 'bottom') css[`padding-${side}`] = space;
      else if (!css[`padding-${side}`]) css[`padding-${side}`] = space;
    }
    const fill = this.shadingColor(wChild(pPr, 'shd'));
    if (fill) css['background-color'] = fill;
    return css;
  }

  /**
   * The colour of a run's <w:color>, as Word shows it: a theme colour (with its tint or shade) from
   * the document's theme, else its value. Null for automatic or none.
   */
  runColor(el) {
    if (!el) return null;
    const theme = this.themeColors[THEME_COLOR_NAMES[wAttr(el, 'themeColor')]];
    if (theme) return '#' + tintShade(theme, wAttr(el, 'themeTint'), wAttr(el, 'themeShade'));
    const val = wAttr(el, 'val');
    return val && /^[0-9a-f]{6}$/i.test(val) ? '#' + val : null;
  }

  /** The fill of a <w:shd> (paragraph or run shading), theme colours included, or null. */
  shadingColor(shd) {
    if (!shd || wAttr(shd, 'val') === 'nil') return null;
    const theme = this.themeColors[THEME_COLOR_NAMES[wAttr(shd, 'themeFill')]];
    if (theme) return '#' + tintShade(theme, wAttr(shd, 'themeFillTint'), wAttr(shd, 'themeFillShade'));
    const fill = wAttr(shd, 'fill');
    if (fill && /^[0-9a-f]{6}$/i.test(fill)) return '#' + fill;
    const color = wAttr(shd, 'color');
    if (wAttr(shd, 'val') === 'solid' && color && /^[0-9a-f]{6}$/i.test(color)) return '#' + color;
    return null;
  }

  /** A stylesheet for the editor page, scoped under `scope`. */
  css(scope) {
    const rules = [];
    const base = { ...this.pPrCss(this.defaultPPr), ...this.rPrCss(this.defaultRPr) };
    if (!base['font-family']) base['font-family'] = `"${this.fonts.minor}", sans-serif`;
    if (!base['font-size']) base['font-size'] = '10pt';
    const { 'margin-top': mt, 'margin-bottom': mb, 'line-height': lh, ...rest } = base;
    // Paragraph borders, shading and padding in the defaults are each paragraph's, not the page's.
    const pageCss = Object.fromEntries(Object.entries(rest).filter(([k]) => !/^(border|padding|background)/.test(k)));
    rules.push(`${scope} { ${decl(pageCss)} }`);
    rules.push(`${scope} :is(p,h1,h2,h3,h4,h5,h6) { margin-top: ${mt || '0'}; margin-bottom: ${mb || '0'}; ${lh ? `line-height: ${lh};` : ''} }`);
    const merged = (id) => {
      const out = {};
      for (const s of this.chain(id)) Object.assign(out, this.pPrCss(s.pPr), this.rPrCss(s.rPr));
      return out;
    };
    for (const s of this.byId.values()) {
      if (!s.id) continue;
      const sel = CSS.escape(s.id);
      if (s.type === 'paragraph') {
        rules.push(`${scope} [data-style="${sel}"] { ${decl(merged(s.id))} }`);
      } else if (s.type === 'character') {
        rules.push(`${scope} [data-rstyle="${sel}"] { ${decl(merged(s.id))} }`);
      } else if (s.type === 'table') {
        // A table style's paragraph settings (often "no space after") apply to the Normal-style
        // paragraphs in its cells, as Word shows them.
        const pCss = {};
        for (const t of this.chain(s.id)) Object.assign(pCss, this.pPrCss(t.pPr));
        if (Object.keys(pCss).length) {
          rules.push(`${scope} table[data-tstyle="${sel}"] td :is(p,h1,h2,h3,h4,h5,h6):not([data-style]) { ${decl(pCss)} }`);
        }
      }
    }
    if (this.defaultParagraph) {
      rules.push(`${scope} :is(p,h1,h2,h3,h4,h5,h6):not([data-style]) { ${decl(merged(this.defaultParagraph))} }`);
    }
    return rules.join('\n');
  }
}

/** Word's theme colour names (w:themeColor) -> the theme's colour slots. */
const THEME_COLOR_NAMES = {
  dark1: 'dk1', light1: 'lt1', dark2: 'dk2', light2: 'lt2', text1: 'dk1', background1: 'lt1', text2: 'dk2', background2: 'lt2',
  accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6',
  hyperlink: 'hlink', followedHyperlink: 'folHlink',
};

/** The colours of a theme's colour scheme: { dk1: 'RRGGBB', lt1, …, accent1…6, hlink, folHlink }. */
export function themeColors(themeDoc) {
  const out = {};
  const scheme = themeDoc?.getElementsByTagNameNS(NS.a, 'clrScheme')[0];
  for (const slot of scheme ? [...scheme.children] : []) {
    const c = slot.firstElementChild;
    const hex = c?.localName === 'srgbClr' ? c.getAttribute('val') : c?.localName === 'sysClr' ? c.getAttribute('lastClr') : null;
    if (hex && /^[0-9a-f]{6}$/i.test(hex)) out[slot.localName] = hex.toUpperCase();
  }
  return out;
}

/**
 * A theme colour with Word's tint or shade (hex bytes, "BF"): a shade darkens and a tint lightens
 * its luminance, as Word applies them (in HSL).
 */
export function tintShade(hex, tint, shade) {
  if (!tint && !shade) return hex;
  const [h, s, l] = rgbToHsl(hex);
  let lum = l;
  if (shade) lum = l * (parseInt(shade, 16) / 255);
  if (tint) lum = l * (parseInt(tint, 16) / 255) + (1 - parseInt(tint, 16) / 255);
  return hslToRgb(h, s, Math.min(1, Math.max(0, lum)));
}

function rgbToHsl(hex) {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb(h, s, l) {
  const f = (p, q, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const rgb = s === 0 ? [l, l, l] : [f(p, q, h + 1 / 3), f(p, q, h), f(p, q, h - 1 / 3)];
  return rgb.map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}

const BORDER_STYLES = {
  single: 'solid', thick: 'solid', double: 'double', dotted: 'dotted', dashed: 'dashed', dashSmallGap: 'dashed',
  dotDash: 'dashed', dotDotDash: 'dotted', triple: 'double', wave: 'solid', doubleWave: 'double',
  thinThickSmallGap: 'double', thickThinSmallGap: 'double', thinThickMediumGap: 'double', thickThinMediumGap: 'double',
  thinThickLargeGap: 'double', thickThinLargeGap: 'double', threeDEmboss: 'ridge', threeDEngrave: 'groove', outset: 'outset', inset: 'inset',
};

/** A Word border (<w:top w:val="single" w:sz="12" w:color="C55A11"/>, sz in eighths of a point) as CSS, or null for none. */
export function borderCss(b) {
  const val = wAttr(b, 'val');
  if (!val || val === 'nil' || val === 'none') return null;
  const style = BORDER_STYLES[val] || 'solid';
  const width = Math.max(style === 'double' ? 2.25 : 0.5, (Number(wAttr(b, 'sz')) || 4) / 8);
  const color = wAttr(b, 'color');
  return `${width}pt ${style} ${color && /^[0-9a-f]{6}$/i.test(color) ? '#' + color : 'currentColor'}`;
}

/**
 * CSS text-align for w:jc. In a right-to-left paragraph (w:bidi), Word's left and right are its start
 * and end, as LibreOffice and Word read them: the default, "left", is at the right.
 */
export function jcToCss(jc, rtl = false) {
  if (rtl) return { start: 'start', left: 'start', end: 'end', right: 'end', center: 'center', both: 'justify', distribute: 'justify' }[jc] || 'start';
  return { start: 'left', left: 'left', end: 'right', right: 'right', center: 'center', both: 'justify', distribute: 'justify' }[jc] || 'left';
}

export const HIGHLIGHT = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff',
  red: '#ff0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000',
  darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080',
  lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
};

export function decl(obj) {
  return Object.entries(obj).map(([k, v]) => `${k}: ${v};`).join(' ');
}
