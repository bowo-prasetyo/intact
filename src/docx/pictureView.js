// Drawing a picture on the page as Word shows it, from what its XML says (see pictures.js
// pictureLook): its crop, shape, border, shadow, glow, soft edges, reflection, bevel, 3-D rotation,
// rotation and colour, each with the file's own values. The layers, inside the picture's span
// (which places and turns it):
//   .pic-shadow  the shadow and glow (drop shadows of what is inside);
//     .pic-box   the picture cut to its shape and crop, with its border, soft edges and bevel;
//   .pic-reflect its reflection, below it.
// A bevel is lit per pixel, from its profile, the light rig and the material (see lighting.js);
// a 3-D rotation is a CSS 3-D transform with the file's angles and perspective.

import { lightingLayer, cameraTransform } from './lighting.js';

const EMU_PER_PX = 9525;
const px = (emu) => (emu || 0) / EMU_PER_PX;
const rgba = (hex, alpha = 1) => {
  const n = parseInt(hex || '000000', 16);
  return `rgb(${(n >> 16) & 255} ${(n >> 8) & 255} ${n & 255} / ${Math.round(alpha * 1000) / 10}%)`;
};

/** The CSS for a shape (Word's preset geometry, with its default adjustments): { radius, clip }. */
function shapeCss(prst, w, h) {
  const s = Math.min(w, h) * 0.16667;
  switch (prst) {
    case 'roundRect': return { radius: `${s}px` };
    case 'ellipse': return { radius: '50%' };
    case 'round2DiagRect': return { radius: `${s}px 0 ${s}px 0` };
    case 'snip2DiagRect': return { clip: `polygon(0 0, calc(100% - ${s}px) 0, 100% ${s}px, 100% 100%, ${s}px 100%, 0 calc(100% - ${s}px))` };
    case 'snip1Rect': return { clip: `polygon(0 0, calc(100% - ${s}px) 0, 100% ${s}px, 100% 100%, 0 100%)` };
    default: return {};
  }
}

/**
 * Brightness and contrast as Word applies them (a:lum): contrast stretches the tones about the
 * middle, brightness moves them all; greyscale by luminance. An SVG filter (shared by pictures with
 * the same settings); returns its CSS `url(#…)`, or '' for none.
 */
function colourFilter({ greyscale, brightness, contrast }) {
  if (!greyscale && !brightness && !contrast) return '';
  const slope = contrast >= 0 ? 1 / Math.max(0.01, 1 - contrast) : 1 + contrast;
  const intercept = 0.5 - 0.5 * slope + (brightness || 0);
  const id = `intact-colour-${greyscale ? 'g' : 'c'}-${Math.round(slope * 1000)}-${Math.round(intercept * 1000)}`;
  if (!document.getElementById(id)) {
    let svg = document.getElementById('intact-filters');
    if (!svg) {
      svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.id = 'intact-filters';
      svg.setAttribute('width', '0');
      svg.setAttribute('height', '0');
      svg.setAttribute('aria-hidden', 'true');
      svg.style.position = 'absolute';
      document.body.append(svg);
    }
    const NSV = 'http://www.w3.org/2000/svg';
    const filter = document.createElementNS(NSV, 'filter');
    filter.id = id;
    filter.setAttribute('color-interpolation-filters', 'sRGB');
    if (greyscale) {
      const m = document.createElementNS(NSV, 'feColorMatrix');
      m.setAttribute('type', 'saturate');
      m.setAttribute('values', '0');
      filter.append(m);
    }
    const t = document.createElementNS(NSV, 'feComponentTransfer');
    for (const c of ['R', 'G', 'B']) {
      const f = document.createElementNS(NSV, `feFunc${c}`);
      f.setAttribute('type', 'linear');
      f.setAttribute('slope', String(slope));
      f.setAttribute('intercept', String(intercept));
      t.append(f);
    }
    filter.append(t);
    svg.append(filter);
  }
  return `url(#${id})`;
}

/** How much more room a box of `w` × `h` needs on each side when turned `deg` degrees: { dx, dy }. */
export function turnedRoomPx(w, h, deg) {
  const a = (deg * Math.PI) / 180;
  const bw = Math.abs(w * Math.cos(a)) + Math.abs(h * Math.sin(a));
  const bh = Math.abs(w * Math.sin(a)) + Math.abs(h * Math.cos(a));
  return { dx: Math.max(0, (bw - w) / 2), dy: Math.max(0, (bh - h) / 2) };
}

/**
 * Draws a picture into its span (already placed by the caller): `w` × `h` px, `look` from
 * pictureLook, `url` the picture shown, `alt` its alt text, `inline` if in line with the text.
 */
export function drawPicture(span, { w, h, look, url, alt, inline }) {
  const img = span.querySelector('img');
  let shadow = span.querySelector(':scope > .pic-shadow');
  if (!shadow) {
    shadow = document.createElement('span');
    shadow.className = 'pic-shadow';
    const box = document.createElement('span');
    box.className = 'pic-box';
    if (img) box.append(img);
    shadow.append(box);
    span.prepend(shadow);
  }
  const box = shadow.querySelector('.pic-box');
  span.querySelector(':scope > .pic-reflect')?.remove();
  box.querySelector('.pic-bevel')?.remove();
  const st = span.style;
  if (!st.position) st.position = 'relative';
  st.width = `${w}px`;
  st.height = `${h}px`;
  // Turned (and in 3-D): the camera first, then the rotation and flips.
  const turns = [cameraTransform(look, w, h), look.rot && `rotate(${look.rot}deg)`, look.flipH && 'scaleX(-1)', look.flipV && 'scaleY(-1)']
    .filter(Boolean).join(' ');
  st.transform = turns;
  span.dataset.turn = turns;
  // In line with the text, a turned picture takes the room of its turned outline (as Word's effectExtent).
  if (inline) {
    const { dx, dy } = turnedRoomPx(w, h, look.rot || 0);
    st.margin = dx || dy ? `${dy}px ${dx}px` : '';
  }
  // Shadow and glow, from the file's distance, direction, blur, colour and transparency.
  const drops = [];
  if (look.shadow) {
    const a = (look.shadow.dir * Math.PI) / 180;
    const d = px(look.shadow.dist);
    drops.push(`drop-shadow(${(d * Math.cos(a)).toFixed(1)}px ${(d * Math.sin(a)).toFixed(1)}px ${(px(look.shadow.blur) / 2).toFixed(1)}px ${rgba(look.shadow.color, look.shadow.alpha)})`);
  }
  if (look.glow) {
    const r = Math.max(1, px(look.glow.rad) / 2);
    const c = rgba(look.glow.color, look.glow.alpha);
    drops.push(`drop-shadow(0 0 ${r.toFixed(1)}px ${c})`, `drop-shadow(0 0 ${r.toFixed(1)}px ${c})`);
  }
  Object.assign(shadow.style, { position: 'absolute', inset: '0', display: 'block', filter: drops.join(' ') });
  // The picture: cut to its shape and crop, with its border (on its edge, as Word draws it).
  const shape = shapeCss(look.geometry, w, h);
  const bs = box.style;
  Object.assign(bs, {
    position: 'absolute', inset: '0', display: 'block', overflow: 'hidden',
    borderRadius: shape.radius || '', clipPath: shape.clip || '',
    outline: look.border && !shape.clip ? `${Math.max(1, look.border.width * 1.333).toFixed(2)}px solid ${rgba(look.border.color, look.border.alpha)}` : '',
    outlineOffset: look.border ? `${(-look.border.width * 1.333 / 2).toFixed(2)}px` : '',
    maskImage: '', webkitMaskImage: '', maskComposite: '', webkitMaskComposite: '',
  });
  if (look.softEdge) {
    const r = Math.max(1, px(look.softEdge.rad));
    const fade = look.geometry === 'ellipse'
      ? `radial-gradient(closest-side, #000 calc(100% - ${r}px), transparent)`
      : `linear-gradient(to right, transparent, #000 ${r}px, #000 calc(100% - ${r}px), transparent), linear-gradient(to bottom, transparent, #000 ${r}px, #000 calc(100% - ${r}px), transparent)`;
    Object.assign(bs, { maskImage: fade, webkitMaskImage: fade });
    if (look.geometry !== 'ellipse') Object.assign(bs, { maskComposite: 'intersect', webkitMaskComposite: 'source-in' });
  }
  const c = look.crop;
  const fullW = w / (1 - c.l - c.r);
  const fullH = h / (1 - c.t - c.b);
  if (img) {
    if (url && img.getAttribute('src') !== url) img.src = url;
    img.alt = alt || '';
    Object.assign(img.style, {
      width: `${fullW}px`, height: `${fullH}px`, maxWidth: 'none', marginLeft: `${-c.l * fullW}px`, marginTop: `${-c.t * fullH}px`,
      filter: colourFilter(look),
    });
  }
  // A bevel, lit: the light and shade the light rig and material give it (see lighting.js).
  const lit = lightingLayer({ w, h, geometry: look.geometry, bevel: look.bevel, material: look.material, light: look.light });
  if (lit) {
    const layer = document.createElement('img');
    layer.className = 'pic-bevel';
    layer.alt = '';
    layer.draggable = false;
    layer.src = lit;
    Object.assign(layer.style, { position: 'absolute', inset: '0', width: '100%', height: '100%', borderRadius: 'inherit', pointerEvents: 'none', maxWidth: 'none', margin: '0', filter: '' });
    box.append(layer);
  }
  // A reflection below it, from the file's distance and start and end transparency and position.
  if (look.reflection) {
    const { stA, endA, endPos, dist, blur } = look.reflection;
    const reflect = document.createElement('span');
    reflect.className = 'pic-reflect';
    reflect.setAttribute('aria-hidden', 'true');
    const fade = `linear-gradient(to bottom, rgb(0 0 0 / ${stA * 100}%) 0%, rgb(0 0 0 / ${endA * 100}%) ${endPos * 100}%, transparent ${Math.min(100, endPos * 100 + 0.1)}%)`;
    Object.assign(reflect.style, {
      position: 'absolute', left: '0', top: `calc(100% + ${px(dist)}px)`, width: '100%', height: '100%', display: 'block',
      pointerEvents: 'none', maskImage: fade, webkitMaskImage: fade, filter: blur ? `blur(${(px(blur) / 2).toFixed(1)}px)` : '',
    });
    const copy = box.cloneNode(true);
    copy.style.transform = 'scaleY(-1)';
    reflect.append(copy);
    span.append(reflect);
  }
}


let markerIds = 0;
const SVG = 'http://www.w3.org/2000/svg';
const svgEl = (name, attrs = {}) => {
  const e = document.createElementNS(SVG, name);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, String(v));
  return e;
};

/**
 * Draws a shape without text (see pictures.js shapeOf): a line or arrow from its box's top left to
 * its bottom right (turned and flipped as Word records it, as pictures are), or a rectangle or an
 * ellipse, in an SVG over its box. The box keeps the shape's size (a straight line's is a few
 * pixels deep, so it can be seen and clicked); the drawing reaches past it by `pad`.
 */
export function drawShape(span, { w, h, look, shape, inline }) {
  const st = span.style;
  if (!st.position) st.position = 'relative';
  st.display = inline ? 'inline-block' : st.display;
  st.width = `${w}px`;
  st.height = `${h}px`;
  st.verticalAlign = inline ? 'middle' : '';
  const turns = [look?.rot && `rotate(${look.rot}deg)`, look?.flipH && 'scaleX(-1)', look?.flipV && 'scaleY(-1)'].filter(Boolean).join(' ');
  st.transform = turns;
  span.dataset.turn = turns;
  span.replaceChildren();
  if (!shape) return;
  const pad = 8;
  const lineW = shape.line?.width || 0;
  const color = shape.line ? `#${shape.line.color}` : 'none';
  const svg = svgEl('svg', { width: w + 2 * pad, height: h + 2 * pad, viewBox: `${-pad} ${-pad} ${w + 2 * pad} ${h + 2 * pad}`, 'aria-hidden': 'true' });
  svg.style.cssText = `position:absolute;left:${-pad}px;top:${-pad}px;overflow:visible;pointer-events:none`;
  if (shape.prst === 'line' || shape.prst === 'straightConnector1') {
    const defs = svgEl('defs');
    const marker = (type) => {
      const id = `shape-arrow-${++markerIds}`;
      const m = svgEl('marker', { id, viewBox: '0 0 10 10', refX: type === 'triangle' ? 8 : 5, refY: 5, markerWidth: 4, markerHeight: 4, orient: 'auto-start-reverse', markerUnits: 'strokeWidth' });
      m.append(type === 'oval' ? svgEl('circle', { cx: 5, cy: 5, r: 4, fill: color }) : svgEl('path', { d: type === 'arrow' ? 'M1 1L9 5L1 9' : 'M0 0L10 5L0 10z', fill: type === 'arrow' ? 'none' : color, stroke: color, 'stroke-width': type === 'arrow' ? 1.5 : 0 }));
      defs.append(m);
      return `url(#${id})`;
    };
    const ends = {};
    if (shape.tail) ends['marker-end'] = marker(shape.tail);
    if (shape.head) ends['marker-start'] = marker(shape.head);
    svg.append(defs);
    // A wide, unseen line under it, so a thin line is easy to click.
    const hit = svgEl('line', { x1: 0, y1: 0, x2: w, y2: h, stroke: 'transparent', 'stroke-width': Math.max(10, lineW + 6) });
    hit.style.pointerEvents = 'stroke';
    svg.append(hit, svgEl('line', { x1: 0, y1: 0, x2: w, y2: h, stroke: color, 'stroke-width': lineW, ...ends }));
  } else {
    const fill = shape.fill ? `#${shape.fill}` : 'none';
    const inset = lineW / 2;
    const body = shape.prst === 'ellipse'
      ? svgEl('ellipse', { cx: w / 2, cy: h / 2, rx: Math.max(0, w / 2 - inset), ry: Math.max(0, h / 2 - inset), fill, stroke: color, 'stroke-width': lineW })
      : svgEl('rect', { x: inset, y: inset, width: Math.max(0, w - lineW), height: Math.max(0, h - lineW), rx: shape.prst === 'roundRect' ? Math.min(w, h) / 6 : 0, fill, stroke: color, 'stroke-width': lineW });
    body.style.pointerEvents = 'visiblePainted';
    svg.append(body);
  }
  span.append(svg);
}
