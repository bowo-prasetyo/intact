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
      if (s.type === 'paragraph' && s.name.toLowerCase() === def.name) return s.id;
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
    const color = wAttr(wChild(rPr, 'color'), 'val');
    if (color && /^[0-9a-f]{6}$/i.test(color)) css.color = '#' + color;
    const hl = wAttr(wChild(rPr, 'highlight'), 'val');
    if (withToggles && hl && hl !== 'none') css['background-color'] = HIGHLIGHT[hl] || hl;
    const shd = wAttr(wChild(rPr, 'shd'), 'fill');
    if (!hl && shd && /^[0-9a-f]{6}$/i.test(shd)) css['background-color'] = '#' + shd;
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
  pPrCss(pPr) {
    const css = {};
    if (!pPr) return css;
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
    if (jc) css['text-align'] = jcToCss(jc);
    return css;
  }

  /** A stylesheet for the editor page, scoped under `scope`. */
  css(scope) {
    const rules = [];
    const base = { ...this.pPrCss(this.defaultPPr), ...this.rPrCss(this.defaultRPr) };
    if (!base['font-family']) base['font-family'] = `"${this.fonts.minor}", sans-serif`;
    if (!base['font-size']) base['font-size'] = '10pt';
    const { 'margin-top': mt, 'margin-bottom': mb, 'line-height': lh, ...pageCss } = base;
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

export function jcToCss(jc) {
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
