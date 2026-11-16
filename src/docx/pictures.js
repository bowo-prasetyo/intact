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
