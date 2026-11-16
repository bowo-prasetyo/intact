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

/**
 * Crop, rotation and effects of a picture: { crop: { l, t, r, b } (fractions), rot (degrees),
 * flipH, flipV, border: { color, width (pt) } | null, shadow, glow, softEdge, reflection,
 * greyscale, brightness, contrast (−1 to 1) }.
 */
export function pictureLook(r) {
  const src = child(blipFillOf(r), A, 'srcRect');
  const spPr = child(picOf(r), PIC, 'spPr');
  const xfrm = child(spPr, A, 'xfrm');
  const ln = child(spPr, A, 'ln');
  const fill = child(ln, A, 'solidFill');
  const color = child(fill, A, 'srgbClr')?.getAttribute('val');
  const effects = child(spPr, A, 'effectLst');
  const blip = child(blipFillOf(r), A, 'blip');
  const lum = child(blip, A, 'lum');
  return {
    crop: { l: pct(src?.getAttribute('l')), t: pct(src?.getAttribute('t')), r: pct(src?.getAttribute('r')), b: pct(src?.getAttribute('b')) },
    rot: (Number(xfrm?.getAttribute('rot')) || 0) / 60000,
    flipH: xfrm?.getAttribute('flipH') === '1',
    flipV: xfrm?.getAttribute('flipV') === '1',
    border: ln && !child(ln, A, 'noFill') ? { color: color || '000000', width: (Number(ln.getAttribute('w')) || 9525) / 12700 } : null,
    shadow: !!child(effects, A, 'outerShdw'),
    glow: child(effects, A, 'glow') ? (child(child(effects, A, 'glow'), A, 'srgbClr')?.getAttribute('val') || 'FFC000') : null,
    softEdge: !!child(effects, A, 'softEdge'),
    reflection: !!child(effects, A, 'reflection'),
    greyscale: !!child(blip, A, 'grayscl'),
    brightness: pct(lum?.getAttribute('bright')),
    contrast: pct(lum?.getAttribute('contrast')),
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
  for (const [name, on] of [['flipH', flipH], ['flipV', flipV]]) {
    if (on) xfrm.setAttribute(name, '1');
    else xfrm.removeAttribute(name);
  }
}

/**
 * A picture's effects, as Word's Picture Border, Effects and Colour menus set them: a border
 * (colour, width in points), an outer shadow, a glow, soft edges, a reflection; greyscale,
 * brightness and contrast. Effects it doesn't set (blur, inner shadows, 3-D…) are kept.
 */
export function setEffects(r, look) {
  const doc = r.ownerDocument;
  const spPr = spPrOf(r);
  const oldLn = child(spPr, A, 'ln');
  if (oldLn) spPr.removeChild(oldLn);
  if (look.border) {
    const ln = makeA(doc, 'ln', { w: Math.round((Number(look.border.width) || 0.75) * 12700) });
    ln.appendChild(makeA(doc, 'solidFill')).appendChild(makeA(doc, 'srgbClr', { val: (look.border.color || '000000').replace('#', '').toUpperCase() }));
    placeIn(spPr, ln, SPPR_ORDER);
  }
  let effects = child(spPr, A, 'effectLst');
  if (!effects) effects = makeA(doc, 'effectLst');
  for (const local of ['glow', 'outerShdw', 'reflection', 'softEdge']) {
    const old = child(effects, A, local);
    if (old) effects.removeChild(old);
  }
  const add = (el) => placeIn(effects, el, EFFECT_ORDER);
  if (look.glow) {
    const glow = makeA(doc, 'glow', { rad: 63500 });
    glow.appendChild(makeA(doc, 'srgbClr', { val: String(look.glow).replace('#', '').toUpperCase() })).appendChild(makeA(doc, 'alpha', { val: 60000 }));
    add(glow);
  }
  if (look.shadow) {
    const shadow = makeA(doc, 'outerShdw', { blurRad: 50800, dist: 38100, dir: 2700000, algn: 'tl', rotWithShape: 0 });
    shadow.appendChild(makeA(doc, 'prstClr', { val: 'black' })).appendChild(makeA(doc, 'alpha', { val: 40000 }));
    add(shadow);
  }
  if (look.reflection) add(makeA(doc, 'reflection', { blurRad: 6350, stA: 50000, endA: 300, endPos: 55000, dir: 5400000, sy: -100000, algn: 'bl', rotWithShape: 0 }));
  if (look.softEdge) add(makeA(doc, 'softEdge', { rad: 63500 }));
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
