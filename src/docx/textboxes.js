// Text boxes and shapes with text, as Word writes them in a run: a DrawingML shape (<wps:wsp>, Word
// 2010 and later) with its text in <wps:txbx><w:txbxContent>, usually inside
// <mc:AlternateContent> with a VML copy (<w:pict><v:shape><v:textbox><w:txbxContent>) as the
// fallback for older readers; or, in older files, the VML shape alone. The text is ordinary
// WordprocessingML (paragraphs, tables), edited here like the body. Everything else about the shape
// (size, position, fill, outline, effects) is kept as it is.
import { NS } from './xml.js';
import { frameInfo } from './pictures.js';

const W = NS.w;
const A = NS.a;
const WP = NS.wp;
const V = NS.v;
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const EMU_PER_PX = 9525;
const EMU_PER_PT = 12700;

const kids = (el, ns, local) => (el ? [...el.children].filter((c) => c.namespaceURI === ns && (!local || c.localName === local)) : []);
const kid = (el, ns, local) => kids(el, ns, local)[0] || null;

function inside(node, stop, ns, local) {
  for (let p = node.parentNode; p && p !== stop; p = p.parentNode) if (p.namespaceURI === ns && p.localName === local) return true;
  return false;
}

/**
 * The text of a run's text box: its own <w:txbxContent> (not one nested in another text box),
 * split into the one shown and edited here (`main`) and copies of it for older readers
 * (`copies`, in <mc:Fallback>), which are written with the same text on save.
 */
export function textboxContents(r) {
  const own = [...r.getElementsByTagNameNS(W, 'txbxContent')].filter((t) => !inside(t, r, W, 'txbxContent'));
  return {
    main: own.filter((t) => !inside(t, r, MC, 'Fallback')),
    copies: own.filter((t) => inside(t, r, MC, 'Fallback')),
  };
}

/** The DrawingML frame (<wp:inline> / <wp:anchor>) of a run's shape: in the run, or in its <mc:Choice>. */
function frameOf(r) {
  let drawing = kid(r, W, 'drawing');
  if (!drawing) {
    for (const ac of kids(r, MC, 'AlternateContent')) {
      for (const choice of kids(ac, MC, 'Choice')) {
        drawing = kid(choice, W, 'drawing');
        if (drawing) break;
      }
      if (drawing) break;
    }
  }
  const frames = kids(drawing, WP).filter((c) => c.localName === 'inline' || c.localName === 'anchor');
  return frames.length === 1 ? frames[0] : null;
}

/** The legacy VML shape holding a run's text box (<w:pict><v:shape>…<v:textbox>), outside any <mc:Fallback>. */
function vmlShapeOf(r) {
  const pict = kid(r, W, 'pict');
  return pict ? kids(pict, V).find((s) => kid(s, V, 'textbox')) || null : null;
}

/**
 * What a run's text box looks like and where it is, or null if the run holds no text box that can
 * be edited here (groups of shapes, canvases and shapes with several texts stay as they are):
 * { wrap, cx, cy (EMU), h, v (as pictureInfo), fill, line: { color, width (px) }, insets (px,
 * [top, right, bottom, left]), anchor ('t', 'ctr', 'b'), autofit, vml }.
 */
export function textboxInfo(r, theme = {}) {
  if (r?.namespaceURI !== W || r.localName !== 'r') return null;
  const { main } = textboxContents(r);
  if (main.length !== 1) return null;
  const frame = frameOf(r);
  if (frame) {
    const data = kid(kid(frame, A, 'graphic'), A, 'graphicData');
    const wsp = kid(data, WPS, 'wsp');
    if (!wsp || !kid(wsp, WPS, 'txbx') || !main[0].parentNode || main[0].parentNode.parentNode !== wsp) return null;
    const spPr = kid(wsp, WPS, 'spPr');
    const style = kid(wsp, WPS, 'style');
    const bodyPr = kid(wsp, WPS, 'bodyPr');
    const fillRef = kid(style, A, 'fillRef');
    const lnRef = kid(style, A, 'lnRef');
    const ln = kid(spPr, A, 'ln');
    let fill = fillOf(spPr, theme);
    if (fill === undefined) fill = fillRef && fillRef.getAttribute('idx') !== '0' ? colorOf(fillRef, theme) : null;
    let line = null;
    const lineFill = ln ? fillOf(ln, theme) : undefined;
    const lineColor = lineFill === undefined ? (lnRef && lnRef.getAttribute('idx') !== '0' ? colorOf(lnRef, theme) : null) : lineFill;
    if (lineColor) line = { color: lineColor, width: Math.max(1, (Number(ln?.getAttribute('w')) || 9525) / EMU_PER_PX) };
    const ins = (attr, dflt) => (bodyPr?.hasAttribute(attr) ? Number(bodyPr.getAttribute(attr)) || 0 : dflt) / EMU_PER_PX;
    return {
      ...frameInfo(frame),
      fill,
      line,
      insets: [ins('tIns', 45720), ins('rIns', 91440), ins('bIns', 45720), ins('lIns', 91440)],
      anchor: bodyPr?.getAttribute('anchor') || 't',
      autofit: !!kid(bodyPr, A, 'spAutoFit'),
      geometry: kid(spPr, A, 'prstGeom')?.getAttribute('prst') || 'rect',
      vml: false,
    };
  }
  const shape = vmlShapeOf(r);
  if (!shape || kid(shape, V, 'textbox')?.firstElementChild !== main[0]) return null;
  return vmlInfo(shape);
}

/** A VML text box's layout and look, from its CSS-like style and attributes. */
function vmlInfo(shape) {
  const css = Object.fromEntries((shape.getAttribute('style') || '').split(';').map((d) => d.split(':').map((x) => x.trim())).filter((d) => d.length === 2));
  const emu = (v) => {
    const m = /^(-?[\d.]+)(pt|in|cm|mm|px)?$/.exec(v || '');
    if (!m) return 0;
    const n = Number(m[1]);
    return Math.round(n * { pt: EMU_PER_PT, in: 914400, cm: 360000, mm: 36000, px: EMU_PER_PX }[m[2] || 'px']);
  };
  const absolute = css.position === 'absolute';
  const wrapEl = [...shape.children].find((c) => c.localName === 'wrap');
  const wrapType = wrapEl?.getAttribute('type');
  const wrap = !absolute ? 'inline' : wrapType === 'topAndBottom' ? 'topAndBottom' : wrapType === 'square' || wrapType === 'tight' || wrapType === 'through'
    ? wrapType : Number(css['z-index']) < 0 ? 'behind' : 'front';
  const color = (v) => {
    const m = /^#([0-9a-f]{6}|[0-9a-f]{3})\b/i.exec(v || '');
    if (!m) return v === 'white' ? 'FFFFFF' : v === 'black' ? '000000' : null;
    return (m[1].length === 3 ? m[1].replace(/./g, (c) => c + c) : m[1]).toUpperCase();
  };
  const filled = shape.getAttribute('filled') !== 'f' && shape.getAttribute('filled') !== 'false';
  const stroked = shape.getAttribute('stroked') !== 'f' && shape.getAttribute('stroked') !== 'false';
  const inset = (kid(shape, V, 'textbox')?.getAttribute('inset') || '').split(',').map((x) => emu(x.trim()) / EMU_PER_PX);
  const dflt = [3.6, 7.2, 3.6, 7.2].map((pt) => (pt * EMU_PER_PT) / EMU_PER_PX);
  const [l, t, rr, b] = [0, 1, 2, 3].map((i) => (inset[i] || inset[i] === 0 ? inset[i] : null));
  return {
    wrap,
    cx: emu(css.width),
    cy: emu(css.height),
    descr: shape.getAttribute('alt') || '',
    h: { relativeFrom: css['mso-position-horizontal-relative'] || 'column', align: css['mso-position-horizontal'] && css['mso-position-horizontal'] !== 'absolute' ? css['mso-position-horizontal'] : null, offset: emu(css['margin-left']) },
    v: { relativeFrom: css['mso-position-vertical-relative'] || 'paragraph', align: css['mso-position-vertical'] && css['mso-position-vertical'] !== 'absolute' ? css['mso-position-vertical'] : null, offset: emu(css['margin-top']) },
    side: 'bothSides',
    fill: filled ? color(shape.getAttribute('fillcolor') || '#ffffff') : null,
    line: stroked ? { color: color(shape.getAttribute('strokecolor') || '#000000') || '000000', width: Math.max(1, emu(shape.getAttribute('strokeweight') || '0.75pt') / EMU_PER_PX) } : null,
    insets: [t ?? dflt[0], rr ?? dflt[1], b ?? dflt[2], l ?? dflt[3]],
    anchor: { middle: 'ctr', bottom: 'b' }[kid(shape, V, 'textbox')?.getAttribute('style')?.match(/v-text-anchor:\s*(\w+)/)?.[1]] || 't',
    autofit: /mso-fit-shape-to-text:\s*t/.test(kid(shape, V, 'textbox')?.getAttribute('style') || ''),
    geometry: shape.localName === 'roundrect' ? 'roundRect' : shape.localName === 'oval' ? 'ellipse' : 'rect',
    vml: true,
  };
}

/** The fill of a shape's properties (or its outline's): a colour 'RRGGBB', null for none, undefined if not said. */
function fillOf(el, theme) {
  if (!el) return undefined;
  if (kid(el, A, 'noFill')) return null;
  const solid = kid(el, A, 'solidFill');
  if (solid) return colorOf(solid, theme);
  const grad = kid(el, A, 'gradFill');
  if (grad) return colorOf(kid(kid(grad, A, 'gsLst'), A, 'gs'), theme); // its first colour
  return undefined;
}

const SCHEME = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2', phClr: 'accent1' };
const DEFAULT_THEME = {
  dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6', accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5',
  accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72',
};
const PRESET = { black: '000000', white: 'FFFFFF', gray: '808080', red: 'FF0000', blue: '0000FF', green: '008000', yellow: 'FFFF00' };

/** The colour (RRGGBB) inside a DrawingML element, with luminance changes applied; null if none. */
function colorOf(el, theme) {
  const c = el && [...el.children].find((k) => k.namespaceURI === A && /Clr$/.test(k.localName));
  if (!c) return null;
  let hex = null;
  if (c.localName === 'srgbClr') hex = c.getAttribute('val');
  else if (c.localName === 'schemeClr') {
    const slot = SCHEME[c.getAttribute('val')] || c.getAttribute('val');
    hex = theme[slot] || DEFAULT_THEME[slot];
  } else if (c.localName === 'prstClr') hex = PRESET[c.getAttribute('val')];
  else if (c.localName === 'sysClr') hex = c.getAttribute('lastClr');
  if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) return null;
  let [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const mod = (local) => {
    const m = kid(c, A, local);
    return m ? Number(m.getAttribute('val')) / 100000 : null;
  };
  const lumMod = mod('lumMod');
  const lumOff = mod('lumOff');
  const shade = mod('shade');
  const tint = mod('tint');
  if (lumMod != null || lumOff != null) {
    const max = Math.max(r, g, b) / 255;
    const min = Math.min(r, g, b) / 255;
    const l = (max + min) / 2;
    const nl = Math.min(1, Math.max(0, l * (lumMod ?? 1) + (lumOff ?? 0)));
    const scale = (v) => {
      const x = v / 255;
      // Keeps hue and saturation: moves the colour towards black or white as its lightness changes.
      return nl >= l ? 255 * (x + (1 - x) * (l >= 1 ? 0 : (nl - l) / (1 - l))) : 255 * x * (l ? nl / l : 0);
    };
    [r, g, b] = [r, g, b].map(scale);
  }
  if (shade != null) [r, g, b] = [r, g, b].map((v) => v * shade);
  if (tint != null) [r, g, b] = [r, g, b].map((v) => v * tint + 255 * (1 - tint));
  return [r, g, b].map((v) => Math.round(Math.min(255, Math.max(0, v))).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/**
 * Writes the edited text of a text box into a copy of its run: `blocks` (paragraphs and tables)
 * replace its shown text, and its copies for older readers get the same (see textboxContents).
 * Returns the copies' new contents, so revision ids can be given to them once known.
 */
export function setTextboxText(r, blocks) {
  const { main, copies } = textboxContents(r);
  const box = main[0];
  while (box.firstChild) box.removeChild(box.firstChild);
  for (const b of blocks) box.appendChild(b);
  if (!blocks.some((b) => b.namespaceURI === W && b.localName === 'p')) box.appendChild(box.ownerDocument.createElementNS(W, 'w:p'));
  // The last block of a text box is a paragraph, as in a table cell.
  const last = box.lastElementChild;
  if (last && last.localName !== 'p') box.appendChild(box.ownerDocument.createElementNS(W, 'w:p'));
  return { main: box, copies };
}

/** Puts the same text as `main` into each copy for older readers: copies of its blocks. */
export function syncCopies(main, copies) {
  const made = [];
  for (const copy of copies) {
    while (copy.firstChild) copy.removeChild(copy.firstChild);
    for (const b of main.childNodes) copy.appendChild(b.cloneNode(true));
    made.push(copy);
  }
  return made;
}

