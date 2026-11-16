// Pictures as Word lays them out (DrawingML in a run): in line with the text (<wp:inline>) or
// floating (<wp:anchor>) with a text wrapping (square, tight, through, top and bottom, behind or in
// front of the text) and a position, horizontal and vertical, either aligned (left, centre, right;
// top, middle, bottom) or at an offset from the column, margin, page or paragraph. Everything here
// changes a copy of a picture's run; the rest of its XML (effects, crop, ids) is kept.
import { NS } from './xml.js';

const WP = NS.wp;
const A = NS.a;
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
export const EMU_PER_PX = 9525;
export const EMU_PER_CM = 360000;

/** Text wrapping choices, as Word names them in Layout Options. */
export const WRAPS = [
  ['inline', 'In line with text'],
  ['square', 'Square'],
  ['tight', 'Tight'],
  ['through', 'Through'],
  ['topAndBottom', 'Top and bottom'],
  ['behind', 'Behind text'],
  ['front', 'In front of text'],
];

const child = (el, ns, local) => (el ? [...el.children].find((c) => c.namespaceURI === ns && c.localName === local) || null : null);

function drawingOf(r) {
  if (r?.namespaceURI === NS.w && r.localName === 'drawing') return r;
  return child(r, NS.w, 'drawing');
}

/** The <wp:inline> or <wp:anchor> of a picture's run, if it holds one picture that can be laid out here. */
function frameOf(r) {
  const drawing = drawingOf(r);
  if (!drawing) return null;
  const frames = [...drawing.children].filter((c) => c.namespaceURI === WP && (c.localName === 'inline' || c.localName === 'anchor'));
  if (frames.length !== 1) return null;
  const data = child(child(frames[0], A, 'graphic'), A, 'graphicData');
  if (data?.getAttribute('uri') !== PIC || !child(data, PIC, 'pic')) return null; // charts, shapes, groups: Word's
  return frames[0];
}

function position(el) {
  if (!el) return null;
  const align = child(el, WP, 'align');
  const offset = child(el, WP, 'posOffset');
  return {
    relativeFrom: el.getAttribute('relativeFrom') || '',
    align: align ? align.textContent.trim() : null,
    offset: offset ? Number(offset.textContent.trim()) || 0 : align ? null : 0,
  };
}

/**
 * What a picture's run says about its layout, or null if it isn't a picture laid out here:
 * { wrap, cx, cy (EMU), h, v ({ relativeFrom, align, offset }), descr, side (wrapText) }.
 */
export function pictureInfo(r) {
  const frame = frameOf(r);
  if (!frame) return null;
  const extent = child(frame, WP, 'extent');
  const docPr = child(frame, WP, 'docPr');
  const info = {
    wrap: 'inline',
    cx: Number(extent?.getAttribute('cx')) || 0,
    cy: Number(extent?.getAttribute('cy')) || 0,
    descr: docPr?.getAttribute('descr') || '',
    h: null,
    v: null,
  };
  if (frame.localName === 'anchor') {
    const wrapEl = [...frame.children].find((c) => c.namespaceURI === WP && c.localName.startsWith('wrap'));
    const kind = wrapEl?.localName;
    info.wrap = kind === 'wrapSquare' ? 'square' : kind === 'wrapTight' ? 'tight' : kind === 'wrapThrough' ? 'through'
      : kind === 'wrapTopAndBottom' ? 'topAndBottom' : frame.getAttribute('behindDoc') === '1' ? 'behind' : 'front';
    info.h = position(child(frame, WP, 'positionH')) || { relativeFrom: 'column', align: null, offset: 0 };
    info.v = position(child(frame, WP, 'positionV')) || { relativeFrom: 'paragraph', align: null, offset: 0 };
    info.side = wrapEl?.getAttribute('wrapText') || 'bothSides';
  }
  return info;
}

/** A picture's new size (EMU), on the frame and the picture's own shape. */
export function resizePicture(r, cx, cy) {
  const frame = frameOf(r);
  cx = Math.max(EMU_PER_PX, Math.round(cx));
  cy = Math.max(EMU_PER_PX, Math.round(cy));
  const extent = child(frame, WP, 'extent');
  extent.setAttribute('cx', String(cx));
  extent.setAttribute('cy', String(cy));
  const xfrm = child(child(child(child(child(frame, A, 'graphic'), A, 'graphicData'), PIC, 'pic'), PIC, 'spPr'), A, 'xfrm');
  const ext = child(xfrm, A, 'ext');
  if (ext) {
    ext.setAttribute('cx', String(cx));
    ext.setAttribute('cy', String(cy));
  }
}

/** Sets (or, when empty, removes) a picture's alt text: its description. */
export function setDescription(r, text) {
  const docPr = child(frameOf(r), WP, 'docPr');
  if (!docPr) return;
  if (text) docPr.setAttribute('descr', text);
  else docPr.removeAttribute('descr');
}

const make = (doc, local, attrs = {}) => {
  const el = doc.createElementNS(WP, `wp:${local}`);
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
  return el;
};

function wrapElement(doc, wrap) {
  if (wrap === 'square') return make(doc, 'wrapSquare', { wrapText: 'bothSides' });
  if (wrap === 'topAndBottom') return make(doc, 'wrapTopAndBottom');
  if (wrap === 'tight' || wrap === 'through') {
    // The outline text wraps to: the picture's rectangle, in Word's 21600 units.
    const el = make(doc, wrap === 'tight' ? 'wrapTight' : 'wrapThrough', { wrapText: 'bothSides' });
    const polygon = el.appendChild(make(doc, 'wrapPolygon', { edited: '0' }));
    polygon.appendChild(make(doc, 'start', { x: 0, y: 0 }));
    for (const [x, y] of [[0, 21600], [21600, 21600], [21600, 0], [0, 0]]) polygon.appendChild(make(doc, 'lineTo', { x, y }));
    return el;
  }
  return make(doc, 'wrapNone');
}

function positionElement(doc, local, { relativeFrom, align, offset }) {
  const el = make(doc, local, { relativeFrom });
  const value = align != null ? make(doc, 'align') : make(doc, 'posOffset');
  value.textContent = align != null ? align : String(Math.round(offset || 0));
  el.appendChild(value);
  return el;
}

/**
 * Changes a picture's text wrapping, as Word's Layout Options do: floating pictures become in line
 * (keeping size, alt text and the picture itself); in-line ones float where their paragraph starts.
 * `zOrder` gives a new floating picture its place in the stack (relativeHeight).
 */
export function setWrap(r, wrap, zOrder = 251659264) {
  const frame = frameOf(r);
  const doc = r.ownerDocument;
  const keep = (local) => child(frame, WP, local);
  if (wrap === 'inline') {
    if (frame.localName === 'inline') return;
    const inline = make(doc, 'inline', { distT: 0, distB: 0, distL: 0, distR: 0 });
    for (const local of ['extent', 'effectExtent', 'docPr', 'cNvGraphicFramePr']) if (keep(local)) inline.appendChild(keep(local));
    inline.appendChild(child(frame, A, 'graphic'));
    frame.replaceWith(inline);
    return;
  }
  let anchor = frame;
  if (frame.localName === 'inline') {
    anchor = make(doc, 'anchor', {
      distT: 0, distB: 0, distL: 114300, distR: 114300, simplePos: 0, relativeHeight: zOrder,
      behindDoc: 0, locked: 0, layoutInCell: 1, allowOverlap: 1,
    });
    anchor.appendChild(make(doc, 'simplePos', { x: 0, y: 0 }));
    anchor.appendChild(positionElement(doc, 'positionH', { relativeFrom: 'column', offset: 0 }));
    anchor.appendChild(positionElement(doc, 'positionV', { relativeFrom: 'paragraph', offset: 0 }));
    anchor.appendChild(keep('extent'));
    anchor.appendChild(keep('effectExtent') || make(doc, 'effectExtent', { l: 0, t: 0, r: 0, b: 0 }));
    anchor.appendChild(make(doc, 'wrapNone')); // replaced below
    for (const local of ['docPr', 'cNvGraphicFramePr']) if (keep(local)) anchor.appendChild(keep(local));
    anchor.appendChild(child(frame, A, 'graphic'));
    frame.replaceWith(anchor);
  }
  const old = [...anchor.children].find((c) => c.namespaceURI === WP && c.localName.startsWith('wrap'));
  const fresh = wrapElement(doc, wrap);
  if (old) old.replaceWith(fresh);
  else (child(anchor, WP, 'docPr') ? child(anchor, WP, 'docPr').before(fresh) : anchor.appendChild(fresh));
  anchor.setAttribute('behindDoc', wrap === 'behind' ? '1' : '0');
  if (wrap === 'topAndBottom') {
    anchor.setAttribute('distL', '0');
    anchor.setAttribute('distR', '0');
  } else if (anchor.getAttribute('distL') === '0' && anchor.getAttribute('distR') === '0' && wrap !== 'behind' && wrap !== 'front') {
    anchor.setAttribute('distL', '114300');
    anchor.setAttribute('distR', '114300');
  }
}

/**
 * Positions a floating picture: `h` and `v` are { relativeFrom, align } or { relativeFrom, offset }
 * (EMU), as Word's Position dialog sets them. A position left out stays as it is.
 */
export function setPosition(r, { h, v }) {
  const frame = frameOf(r);
  if (frame.localName !== 'anchor') return;
  const doc = r.ownerDocument;
  for (const [local, value] of [['positionH', h], ['positionV', v]]) {
    if (!value) continue;
    const fresh = positionElement(doc, local, value);
    const old = child(frame, WP, local);
    if (old) old.replaceWith(fresh);
    else if (local === 'positionH') child(frame, WP, 'simplePos').after(fresh);
    else child(frame, WP, 'positionH').after(fresh);
  }
}

/** True if a run holds a picture whose layout can be changed here. */
export const isLayoutPicture = (r) => !!frameOf(r);

// ---------------------------------------------------------------- crop, rotation, effects

const picOf = (r) => child(child(child(frameOf(r), A, 'graphic'), A, 'graphicData'), PIC, 'pic');
const blipFillOf = (r) => child(picOf(r), PIC, 'blipFill');
const spPrOf = (r) => {
  const pic = picOf(r);
  let spPr = child(pic, PIC, 'spPr');
  if (!spPr) {
    spPr = r.ownerDocument.createElementNS(PIC, 'pic:spPr');
    pic.appendChild(spPr);
  }
  return spPr;
};
const makeA = (doc, local, attrs = {}) => {
  const el = doc.createElementNS(A, `a:${local}`);
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
  return el;
};
/** Puts `el` into `parent` at its place in `order` (DrawingML is strict about order). */
function placeIn(parent, el, order) {
  const at = order.indexOf(el.localName);
  const before = [...parent.children].find((c) => order.indexOf(c.localName) > at || order.indexOf(c.localName) === -1 && c.localName === 'extLst');
  parent.insertBefore(el, before || null);
}
const SPPR_ORDER = ['xfrm', 'custGeom', 'prstGeom', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst'];
const EFFECT_ORDER = ['blur', 'fillOverlay', 'glow', 'innerShdw', 'outerShdw', 'prstShdw', 'reflection', 'softEdge'];
const pct = (v) => (Number(v) || 0) / 100000; // DrawingML percentages are in 1/1000 of a percent

// Office's default theme colours, for effects given as theme colours (accent1 and so on).
const THEME = {
  accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47',
  tx1: '000000', bg1: 'FFFFFF', dk1: '000000', lt1: 'FFFFFF', tx2: '44546A', bg2: 'E7E6E6', dk2: '44546A', lt2: 'E7E6E6',
};
const PRESET_COLORS = { black: '000000', white: 'FFFFFF', gray: '808080', red: 'FF0000', blue: '0000FF', green: '008000', yellow: 'FFFF00' };

/** The colour inside a DrawingML element (srgbClr, schemeClr, prstClr, sysClr), with its alpha: { color, alpha } or null. */
function colorIn(el) {
  const c = el && [...el.children].find((k) => k.namespaceURI === A && /Clr$/.test(k.localName));
  if (!c) return null;
  let color = '000000';
  if (c.localName === 'srgbClr') color = c.getAttribute('val') || color;
  else if (c.localName === 'schemeClr') color = THEME[c.getAttribute('val')] || color;
  else if (c.localName === 'prstClr') color = PRESET_COLORS[c.getAttribute('val')] || color;
  else if (c.localName === 'sysClr') color = c.getAttribute('lastClr') || color;
  const alpha = child(c, A, 'alpha');
  return { color: color.toUpperCase(), alpha: alpha ? pct(alpha.getAttribute('val')) : 1 };
}

const num = (el, attr, dflt = 0) => {
  const v = Number(el?.getAttribute(attr));
  return Number.isFinite(v) && el?.hasAttribute(attr) ? v : dflt;
};

const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
/** Where Intact keeps a picture's original when an artistic effect is drawn into it (Word ignores what it doesn't know in extLst). */
export const INTACT_PIC = 'urn:intact:picture';
const ORIGINAL_EXT = '{6E8D1F2A-7B3C-4D5E-9F60-A1B2C3D4E5F6}';

/**
 * Crop, rotation, shape, border, effects, 3-D and colour of a picture, with the values the file
 * gives (lengths in EMU, angles in degrees, alphas and percentages from 0 to 1):
 * { crop, rot, flipH, flipV, geometry, border: { color, width (pt), alpha } | null,
 *   shadow: { blur, dist, dir, color, alpha } | null, glow: { rad, color, alpha } | null,
 *   softEdge: { rad } | null, reflection: { blur, dist, stA, endA, endPos } | null,
 *   bevel: { prst, w, h } | null, camera, material, greyscale, brightness, contrast,
 *   artistic: { effect, original (relationship id) } | null }.
 */
export function pictureLook(r) {
  const src = child(blipFillOf(r), A, 'srcRect');
  const spPr = child(picOf(r), PIC, 'spPr');
  const xfrm = child(spPr, A, 'xfrm');
  const ln = child(spPr, A, 'ln');
  const effects = child(spPr, A, 'effectLst');
  const blip = child(blipFillOf(r), A, 'blip');
  const lum = child(blip, A, 'lum');
  const shadow = child(effects, A, 'outerShdw');
  const glow = child(effects, A, 'glow');
  const soft = child(effects, A, 'softEdge');
  const reflection = child(effects, A, 'reflection');
  const bevel = child(child(spPr, A, 'sp3d'), A, 'bevelT');
  const camera = child(child(spPr, A, 'scene3d'), A, 'camera');
  const lineColor = colorIn(child(ln, A, 'solidFill'));
  const original = blip && [...blip.getElementsByTagNameNS(INTACT_PIC, 'original')][0];
  return {
    crop: { l: pct(src?.getAttribute('l')), t: pct(src?.getAttribute('t')), r: pct(src?.getAttribute('r')), b: pct(src?.getAttribute('b')) },
    rot: num(xfrm, 'rot') / 60000,
    flipH: xfrm?.getAttribute('flipH') === '1',
    flipV: xfrm?.getAttribute('flipV') === '1',
    geometry: child(spPr, A, 'prstGeom')?.getAttribute('prst') || 'rect',
    border: ln && !child(ln, A, 'noFill') ? { color: lineColor?.color || '000000', alpha: lineColor?.alpha ?? 1, width: num(ln, 'w', 9525) / 12700 } : null,
    shadow: shadow ? { blur: num(shadow, 'blurRad'), dist: num(shadow, 'dist'), dir: num(shadow, 'dir') / 60000, ...(colorIn(shadow) || { color: '000000', alpha: 1 }) } : null,
    glow: glow ? { rad: num(glow, 'rad'), ...(colorIn(glow) || { color: 'FFC000', alpha: 1 }) } : null,
    softEdge: soft ? { rad: num(soft, 'rad') } : null,
    reflection: reflection ? {
      blur: num(reflection, 'blurRad'), dist: num(reflection, 'dist'),
      stA: pct(num(reflection, 'stA', 100000)), endA: pct(num(reflection, 'endA', 0)), endPos: pct(num(reflection, 'endPos', 100000)),
    } : null,
    bevel: bevel ? { prst: bevel.getAttribute('prst') || 'circle', w: num(bevel, 'w', 76200), h: num(bevel, 'h', 76200) } : null,
    camera: camera && camera.getAttribute('prst') !== 'orthographicFront' ? camera.getAttribute('prst') : null,
    // A rotation set by angles (Word's X, Y and Z rotation; a:rot lon, lat, rev) and a perspective (fov).
    rotation3d: camera && (child(camera, A, 'rot') || camera.hasAttribute('fov')) ? {
      x: num(child(camera, A, 'rot'), 'lon') / 60000, y: num(child(camera, A, 'rot'), 'lat') / 60000,
      z: num(child(camera, A, 'rot'), 'rev') / 60000, perspective: num(camera, 'fov') / 60000,
    } : null,
    light: lightOf(child(spPr, A, 'scene3d')),
    material: child(spPr, A, 'sp3d')?.getAttribute('prstMaterial') || null,
    greyscale: !!child(blip, A, 'grayscl'),
    brightness: pct(lum?.getAttribute('bright')),
    contrast: pct(lum?.getAttribute('contrast')),
    artistic: original ? {
      effect: original.getAttribute('effect'), original: original.getAttributeNS(NS.r, 'embed'),
      amount: original.hasAttribute('amount') ? Number(original.getAttribute('amount')) : null,
    } : null,
  };
}

/**
 * Crops a picture (fractions of the picture cut from each side, as Word's Crop does): what shows
 * keeps its scale, so the picture's frame grows or shrinks with the crop.
 */
export function setCrop(r, crop) {
  const blipFill = blipFillOf(r);
  const before = pictureLook(r).crop;
  const clamp = (v) => Math.max(-0.9, Math.min(0.95, Number(v) || 0));
  const c = { l: clamp(crop.l), t: clamp(crop.t), r: clamp(crop.r), b: clamp(crop.b) };
  if (c.l + c.r > 0.95) c.r = 0.95 - c.l;
  if (c.t + c.b > 0.95) c.b = 0.95 - c.t;
  const info = pictureInfo(r);
  const fullW = info.cx / (1 - before.l - before.r);
  const fullH = info.cy / (1 - before.t - before.b);
  resizePicture(r, fullW * (1 - c.l - c.r), fullH * (1 - c.t - c.b));
  let src = child(blipFill, A, 'srcRect');
  if (!c.l && !c.t && !c.r && !c.b) {
    if (src) blipFill.removeChild(src);
    return;
  }
  if (!src) {
    src = makeA(r.ownerDocument, 'srcRect');
    const blip = child(blipFill, A, 'blip');
    if (blip) blip.after(src);
    else blipFill.prepend(src);
  }
  for (const side of ['l', 't', 'r', 'b']) {
    const v = Math.round(c[side] * 100000);
    if (v) src.setAttribute(side, String(v));
    else src.removeAttribute(side);
  }
}

/** Rotates (degrees, clockwise) and flips a picture, as on its shape (a:xfrm). */
export function setRotation(r, { rot, flipH, flipV }) {
  const spPr = spPrOf(r);
  let xfrm = child(spPr, A, 'xfrm');
  if (!xfrm) {
    const info = pictureInfo(r);
    xfrm = makeA(r.ownerDocument, 'xfrm');
    xfrm.appendChild(makeA(r.ownerDocument, 'off', { x: 0, y: 0 }));
    xfrm.appendChild(makeA(r.ownerDocument, 'ext', { cx: info.cx, cy: info.cy }));
    spPr.prepend(xfrm);
  }
  const deg = ((Math.round(Number(rot) || 0) % 360) + 360) % 360;
  if (deg) xfrm.setAttribute('rot', String(deg * 60000));
  else xfrm.removeAttribute('rot');
  // In line with the text, a turned picture takes the room of its turned outline, as Word records
  // it (wp:effectExtent), so the text makes way for its corners.
  const frame = frameOf(r);
  if (frame.localName === 'inline') {
    const info = pictureInfo(r);
    const { dx, dy } = turnedRoom(info.cx, info.cy, deg);
    let extent = child(frame, WP, 'effectExtent');
    if (!extent) {
      extent = make(r.ownerDocument, 'effectExtent');
      child(frame, WP, 'extent').after(extent);
    }
    for (const [side, v] of [['l', dx], ['t', dy], ['r', dx], ['b', dy]]) extent.setAttribute(side, String(Math.round(v)));
  }
  for (const [name, on] of [['flipH', flipH], ['flipV', flipV]]) {
    if (on) xfrm.setAttribute(name, '1');
    else xfrm.removeAttribute(name);
  }
}

/**
 * A picture's effects, as Word's Picture Border, Effects and Colour menus set them: a border
 * (colour, width in points), an outer shadow, a glow, soft edges, a reflection; greyscale,
 * brightness and contrast. An effect left on keeps the values it has (only one turned on gets
 * Word's default); effects it doesn't set (blur, inner shadows…) are kept.
 */
export function setEffects(r, look) {
  const doc = r.ownerDocument;
  const spPr = spPrOf(r);
  const oldLn = child(spPr, A, 'ln');
  const wanted = look.border ? Math.round((Number(look.border.width) || 0.75) * 12700) : 0;
  const color = (look.border?.color || '000000').replace('#', '').toUpperCase();
  const same = oldLn && look.border && num(oldLn, 'w', 9525) === wanted && colorIn(child(oldLn, A, 'solidFill'))?.color === color;
  if (!same) {
    if (oldLn) spPr.removeChild(oldLn);
    if (look.border) {
      const ln = makeA(doc, 'ln', { w: wanted });
      ln.appendChild(makeA(doc, 'solidFill')).appendChild(makeA(doc, 'srgbClr', { val: color }));
      placeIn(spPr, ln, SPPR_ORDER);
    }
  }
  let effects = child(spPr, A, 'effectLst');
  if (!effects) effects = makeA(doc, 'effectLst');
  const add = (el) => placeIn(effects, el, EFFECT_ORDER);
  const toggle = (local, on, make) => {
    const old = child(effects, A, local);
    if (on && !old) add(make());
    else if (!on && old) effects.removeChild(old);
  };
  toggle('glow', look.glow, () => {
    const glow = makeA(doc, 'glow', { rad: 63500 });
    const c = typeof look.glow === 'string' ? look.glow : 'FFC000';
    glow.appendChild(makeA(doc, 'srgbClr', { val: c.replace('#', '').toUpperCase() })).appendChild(makeA(doc, 'alpha', { val: 60000 }));
    return glow;
  });
  toggle('outerShdw', look.shadow, () => {
    const shadow = makeA(doc, 'outerShdw', { blurRad: 50800, dist: 38100, dir: 2700000, algn: 'tl', rotWithShape: 0 });
    shadow.appendChild(makeA(doc, 'prstClr', { val: 'black' })).appendChild(makeA(doc, 'alpha', { val: 40000 }));
    return shadow;
  });
  toggle('reflection', look.reflection, () => makeA(doc, 'reflection', { blurRad: 6350, stA: 50000, endA: 300, endPos: 55000, dir: 5400000, sy: -100000, algn: 'bl', rotWithShape: 0 }));
  toggle('softEdge', look.softEdge, () => makeA(doc, 'softEdge', { rad: 63500 }));
  if (effects.children.length && !effects.parentNode) placeIn(spPr, effects, SPPR_ORDER);
  else if (!effects.children.length && effects.parentNode) spPr.removeChild(effects);
  const blip = child(blipFillOf(r), A, 'blip');
  if (!blip) return;
  for (const local of ['grayscl', 'lum']) {
    const old = child(blip, A, local);
    if (old) blip.removeChild(old);
  }
  const ext = child(blip, A, 'extLst');
  if (look.greyscale) blip.insertBefore(makeA(doc, 'grayscl'), ext);
  const bright = Math.round((Number(look.brightness) || 0) * 100000);
  const contrast = Math.round((Number(look.contrast) || 0) * 100000);
  if (bright || contrast) {
    const lum = makeA(doc, 'lum');
    if (bright) lum.setAttribute('bright', String(bright));
    if (contrast) lum.setAttribute('contrast', String(contrast));
    blip.insertBefore(lum, ext);
  }
}

// ---------------------------------------------------------------- shapes, 3-D and picture styles

/** Shapes a picture can be cut to, as Word's Crop to Shape offers the common ones. */
export const SHAPES = [
  ['rect', 'Rectangle'], ['roundRect', 'Rounded rectangle'], ['ellipse', 'Oval'],
  ['snip2DiagRect', 'Diagonal corners snipped'], ['round2DiagRect', 'Diagonal corners rounded'], ['snip1Rect', 'Single corner snipped'],
];

/** Bevels (the top edge of Word's 3-D Format) and 3-D rotations, by Word's names. */
export const BEVELS = [
  ['circle', 'Circle'], ['relaxedInset', 'Relaxed inset'], ['cross', 'Cross'], ['coolSlant', 'Cool slant'],
  ['angle', 'Angle'], ['softRound', 'Soft round'], ['convex', 'Convex'],
];
export const CAMERAS = [
  ['isometricOffAxis1Right', 'Off axis 1: right'], ['isometricOffAxis1Left', 'Off axis 1: left'],
  ['perspectiveRight', 'Perspective: right'], ['perspectiveLeft', 'Perspective: left'],
  ['perspectiveRelaxed', 'Perspective: relaxed'], ['isometricTopUp', 'Isometric: top up'],
];

/** Cuts a picture to a shape (a:prstGeom), as Crop to Shape does. */
export function setGeometry(r, prst) {
  const spPr = spPrOf(r);
  let geom = child(spPr, A, 'prstGeom');
  const custom = child(spPr, A, 'custGeom');
  if (custom) spPr.removeChild(custom);
  if (!geom) {
    geom = makeA(r.ownerDocument, 'prstGeom');
    geom.appendChild(makeA(r.ownerDocument, 'avLst'));
    placeIn(spPr, geom, SPPR_ORDER);
  }
  geom.setAttribute('prst', prst || 'rect');
}

/** A scene's light rig: { rig, dir, angle (degrees, Word's lighting Angle) }, or null. */
function lightOf(scene) {
  const rig = child(scene, A, 'lightRig');
  if (!rig) return null;
  return { rig: rig.getAttribute('rig') || 'threePt', dir: rig.getAttribute('dir') || 't', angle: num(child(rig, A, 'rot'), 'rev') / 60000 };
}

/** Word's materials and lighting (3-D Format), by Word's names. */
export const MATERIALS = [
  ['matte', 'Matte'], ['warmMatte', 'Warm Matte'], ['plastic', 'Plastic'], ['metal', 'Metal'], ['dkEdge', 'Dark Edge'],
  ['softEdge', 'Soft Edge'], ['flat', 'Flat'], ['legacyWireframe', 'Wireframe'], ['powder', 'Powder'],
  ['translucentPowder', 'Translucent Powder'], ['clear', 'Clear'], ['softmetal', 'Soft Metal'],
];
export const LIGHTS = [
  ['threePt', 'Three Point'], ['balanced', 'Balanced'], ['soft', 'Soft'], ['harsh', 'Harsh'], ['flood', 'Flood'],
  ['contrasting', 'Contrasting'], ['morning', 'Morning'], ['sunrise', 'Sunrise'], ['sunset', 'Sunset'],
  ['chilly', 'Chilly'], ['freezing', 'Freezing'], ['flat', 'Flat'], ['twoPt', 'Two Point'], ['glow', 'Glow'], ['brightRoom', 'Bright Room'],
];

/**
 * A picture's 3-D format and rotation (a:scene3d, a:sp3d), as Word's 3-D Format and 3-D Rotation
 * set them: a bevel on its top edge ({ prst, w, h } in EMU, Word's 6 pt × 2 pt unless given), a
 * material, a light rig ({ rig, dir, angle }), a camera preset, and a rotation by angles
 * ({ x, y, z, perspective } in degrees, written as a:rot and the camera's field of view).
 * Nothing set takes the 3-D format off.
 */
export function set3D(r, { bevel = null, camera = null, material = null, light = null, rotation = null } = {}) {
  const doc = r.ownerDocument;
  const spPr = spPrOf(r);
  for (const local of ['scene3d', 'sp3d']) {
    const old = child(spPr, A, local);
    if (old) spPr.removeChild(old);
  }
  const turned = rotation && (rotation.x || rotation.y || rotation.z || rotation.perspective);
  const lit = light && (light.rig !== 'threePt' || light.dir !== 't' || light.angle);
  if (!bevel && !camera && !material && !turned && !lit) return;
  const scene = makeA(doc, 'scene3d');
  const cam = scene.appendChild(makeA(doc, 'camera', { prst: camera || (rotation?.perspective ? 'perspectiveFront' : 'orthographicFront') }));
  if (rotation?.perspective) cam.setAttribute('fov', String(Math.round(rotation.perspective * 60000)));
  if (rotation && (rotation.x || rotation.y || rotation.z)) {
    const deg = (v) => String(Math.round((((Number(v) || 0) % 360) + 360) % 360 * 60000));
    cam.appendChild(makeA(doc, 'rot', { lat: deg(rotation.y), lon: deg(rotation.x), rev: deg(rotation.z) }));
  }
  const rig = scene.appendChild(makeA(doc, 'lightRig', { rig: light?.rig || 'threePt', dir: light?.dir || 't' }));
  if (light?.angle) rig.appendChild(makeA(doc, 'rot', { lat: 0, lon: 0, rev: Math.round((((light.angle % 360) + 360) % 360) * 60000) }));
  placeIn(spPr, scene, SPPR_ORDER);
  const sp3d = makeA(doc, 'sp3d', material ? { prstMaterial: material } : {});
  if (bevel) {
    const b = typeof bevel === 'string' ? { prst: bevel } : bevel;
    sp3d.appendChild(makeA(doc, 'bevelT', { w: Math.round(b.w ?? 76200), h: Math.round(b.h ?? 25400), prst: b.prst || 'circle' }));
  }
  placeIn(spPr, sp3d, SPPR_ORDER);
}

// Word's Picture Styles gallery: each is a shape, border, effects and 3-D format, written out as such
// (Word keeps no reference to the style). Lengths in EMU.
const WHITE_FRAME = { color: 'FFFFFF', width: 7 };
const SOFT_SHADOW = { blur: 50800, dist: 38100, dir: 45, color: '000000', alpha: 0.43 };
export const PICTURE_STYLES = [
  { name: 'Simple Frame, White', geometry: 'rect', border: WHITE_FRAME, shadow: SOFT_SHADOW },
  { name: 'Simple Frame, Black', geometry: 'rect', border: { color: '000000', width: 7 } },
  { name: 'Beveled Matte, White', geometry: 'rect', border: WHITE_FRAME, shadow: SOFT_SHADOW, bevel: { prst: 'circle', w: 25400, h: 19050 } },
  { name: 'Metal Frame', geometry: 'rect', border: { color: 'A6A6A6', width: 7 }, bevel: { prst: 'angle', w: 50800, h: 50800 } },
  { name: 'Drop Shadow Rectangle', geometry: 'rect', shadow: { blur: 127000, dist: 63500, dir: 45, color: '000000', alpha: 0.4 } },
  { name: 'Rounded Diagonal Corner, White', geometry: 'round2DiagRect', border: WHITE_FRAME, shadow: SOFT_SHADOW },
  { name: 'Snip Diagonal Corner, White', geometry: 'snip2DiagRect', border: WHITE_FRAME, shadow: SOFT_SHADOW },
  { name: 'Moderate Frame, Black', geometry: 'rect', border: { color: '000000', width: 4.5 }, shadow: SOFT_SHADOW },
  { name: 'Soft Edge Rectangle', geometry: 'rect', softEdge: { rad: 112500 } },
  { name: 'Soft Edge Oval', geometry: 'ellipse', softEdge: { rad: 112500 } },
  { name: 'Center Shadow Rectangle', geometry: 'rect', shadow: { blur: 190500, dist: 0, dir: 0, color: '000000', alpha: 0.4 } },
  { name: 'Reflected Rounded Rectangle', geometry: 'roundRect', reflection: { blur: 12700, dist: 0, stA: 0.38, endA: 0.003, endPos: 0.38 } },
  { name: 'Bevel Rectangle', geometry: 'rect', shadow: SOFT_SHADOW, bevel: { prst: 'circle', w: 63500, h: 25400 } },
  { name: 'Bevel Perspective', geometry: 'rect', border: WHITE_FRAME, bevel: { prst: 'circle', w: 25400, h: 19050 }, camera: 'perspectiveRelaxed' },
];

/**
 * Applies a picture style, as Word's gallery does: the shape, border, shadow, glow, soft edges,
 * reflection and 3-D format it gives replace the picture's (other effects are kept). `null` takes
 * them all off (Word's Reset Picture keeps size and crop too).
 */
export function applyPictureStyle(r, style) {
  const doc = r.ownerDocument;
  const st = style || {};
  setGeometry(r, st.geometry || 'rect');
  const spPr = spPrOf(r);
  const oldLn = child(spPr, A, 'ln');
  if (oldLn) spPr.removeChild(oldLn);
  if (st.border) {
    const ln = makeA(doc, 'ln', { w: Math.round(st.border.width * 12700) });
    ln.appendChild(makeA(doc, 'solidFill')).appendChild(makeA(doc, 'srgbClr', { val: st.border.color }));
    ln.appendChild(makeA(doc, 'miter', { lim: 800000 }));
    placeIn(spPr, ln, SPPR_ORDER);
  }
  let effects = child(spPr, A, 'effectLst');
  if (!effects) effects = makeA(doc, 'effectLst');
  for (const local of ['glow', 'outerShdw', 'reflection', 'softEdge']) {
    const old = child(effects, A, local);
    if (old) effects.removeChild(old);
  }
  const add = (el) => placeIn(effects, el, EFFECT_ORDER);
  const colour = (el, { color, alpha }) => {
    const c = el.appendChild(makeA(doc, 'srgbClr', { val: color }));
    if (alpha != null && alpha < 1) c.appendChild(makeA(doc, 'alpha', { val: Math.round(alpha * 100000) }));
    return el;
  };
  if (st.glow) add(colour(makeA(doc, 'glow', { rad: st.glow.rad }), st.glow));
  if (st.shadow) {
    add(colour(makeA(doc, 'outerShdw', {
      blurRad: st.shadow.blur, dist: st.shadow.dist, dir: Math.round(st.shadow.dir * 60000), algn: st.shadow.dist ? 'tl' : 'ctr', rotWithShape: 0,
    }), st.shadow));
  }
  if (st.reflection) {
    add(makeA(doc, 'reflection', {
      blurRad: st.reflection.blur, stA: Math.round(st.reflection.stA * 100000), endA: Math.round(st.reflection.endA * 100000),
      endPos: Math.round(st.reflection.endPos * 100000), dist: st.reflection.dist, dir: 5400000, sy: -100000, algn: 'bl', rotWithShape: 0,
    }));
  }
  if (st.softEdge) add(makeA(doc, 'softEdge', { rad: st.softEdge.rad }));
  if (effects.children.length && !effects.parentNode) placeIn(spPr, effects, SPPR_ORDER);
  else if (!effects.children.length && effects.parentNode) spPr.removeChild(effects);
  set3D(r, { bevel: st.bevel || null, camera: st.camera || null });
}

// ---------------------------------------------------------------- artistic effects

/** Word's artistic effects, by Word's names, all drawn here (see editor/artistic.js). */
export const ARTISTIC = [
  ['marker', 'Marker'], ['pencilGrayscale', 'Pencil Grayscale'], ['pencilSketch', 'Pencil Sketch'], ['lineDrawing', 'Line Drawing'],
  ['chalkSketch', 'Chalk Sketch'], ['paintStrokes', 'Paint Strokes'], ['paintBrush', 'Paint Brush'], ['glowDiffused', 'Glow Diffused'],
  ['blur', 'Blur'], ['lightScreen', 'Light Screen'], ['watercolorSponge', 'Watercolor Sponge'], ['filmGrain', 'Film Grain'],
  ['mosaicBubbles', 'Mosaic Bubbles'], ['glass', 'Glass'], ['cement', 'Cement'], ['texturizer', 'Texturizer'],
  ['crisscrossEtching', 'Crisscross Etching'], ['pastelsSmooth', 'Pastels Smooth'], ['plasticWrap', 'Plastic Wrap'],
  ['cutout', 'Cutout'], ['photocopy', 'Photocopy'], ['glowEdges', 'Glow Edges'],
];

/**
 * An artistic effect drawn into the picture: the picture shows `rendered` (a relationship id of the
 * drawn image), and the original is kept (`original`, with the effect's name) in an extension of the
 * picture, so the effect can be changed or taken off here. Word shows the drawn picture.
 */
export function setArtistic(r, { effect, rendered, original, amount = null }) {
  const doc = r.ownerDocument;
  const blip = child(blipFillOf(r), A, 'blip');
  clearArtisticMark(blip);
  blip.setAttributeNS(NS.r, 'r:embed', rendered);
  let extLst = child(blip, A, 'extLst');
  if (!extLst) extLst = blip.appendChild(makeA(doc, 'extLst'));
  const ext = extLst.appendChild(makeA(doc, 'ext', { uri: ORIGINAL_EXT }));
  const el = ext.appendChild(doc.createElementNS(INTACT_PIC, 'intact:original'));
  el.setAttributeNS(NS.r, 'r:embed', original);
  el.setAttribute('effect', effect);
  if (amount != null) el.setAttribute('amount', String(Math.round(amount)));
}

/** Takes an artistic effect off: the original picture shows again. */
export function clearArtistic(r) {
  const blip = child(blipFillOf(r), A, 'blip');
  const original = blip && [...blip.getElementsByTagNameNS(INTACT_PIC, 'original')][0];
  if (!original) return;
  blip.setAttributeNS(NS.r, 'r:embed', original.getAttributeNS(NS.r, 'embed'));
  clearArtisticMark(blip);
}

function clearArtisticMark(blip) {
  const extLst = child(blip, A, 'extLst');
  for (const ext of [...(extLst?.children || [])]) if (ext.getAttribute('uri') === ORIGINAL_EXT) extLst.removeChild(ext);
  if (extLst && !extLst.children.length) blip.removeChild(extLst);
}

/** The relationship id of the picture shown (a:blip/@r:embed). */
export function blipOf(r) {
  return child(blipFillOf(r), A, 'blip')?.getAttributeNS(NS.r, 'embed') || null;
}

// ---------------------------------------------------------------- wrapping to the outline

/** The outline text wraps to (Tight, Through), as points in Word's 21600 units: [[x, y]…], or null. */
export function wrapPolygon(r) {
  const frame = frameOf(r);
  const wrap = frame && [...frame.children].find((c) => c.namespaceURI === WP && (c.localName === 'wrapTight' || c.localName === 'wrapThrough'));
  const polygon = child(wrap, WP, 'wrapPolygon');
  if (!polygon) return null;
  return [...polygon.children].filter((c) => c.namespaceURI === WP).map((c) => [Number(c.getAttribute('x')) || 0, Number(c.getAttribute('y')) || 0]);
}

/** Sets the outline text wraps to (points in 21600 units, the first repeated at the end as Word writes it). */
export function setWrapPolygon(r, points) {
  const frame = frameOf(r);
  const wrap = [...frame.children].find((c) => c.namespaceURI === WP && (c.localName === 'wrapTight' || c.localName === 'wrapThrough'));
  if (!wrap || !points?.length) return;
  const doc = r.ownerDocument;
  const polygon = make(doc, 'wrapPolygon', { edited: '0' });
  const all = [...points, points[0]];
  all.forEach(([x, y], i) => polygon.appendChild(make(doc, i === 0 ? 'start' : 'lineTo', { x: Math.round(x), y: Math.round(y) })));
  const old = child(wrap, WP, 'wrapPolygon');
  if (old) old.replaceWith(polygon);
  else wrap.appendChild(polygon);
}

/** How much more room a picture of `w` × `h` needs on each side when turned `deg` degrees: { dx, dy }. */
export function turnedRoom(w, h, deg) {
  const a = (deg * Math.PI) / 180;
  const bw = Math.abs(w * Math.cos(a)) + Math.abs(h * Math.sin(a));
  const bh = Math.abs(w * Math.sin(a)) + Math.abs(h * Math.cos(a));
  return { dx: Math.max(0, (bw - w) / 2), dy: Math.max(0, (bh - h) / 2) };
}
