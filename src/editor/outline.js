// The outline of a picture, for wrapping text tightly around it (Word's Tight and Through): the
// picture is drawn small, and each row's first and last visible (not transparent) pixel gives the
// outline's left and right edges, as points in Word's 21600 units.

const SAMPLE = 64; // the picture is sampled at most this many pixels across
const VISIBLE = 24; // alpha (0–255) above which a pixel counts as part of the picture

/**
 * The outline of what shows of `img` (cropped by `crop`, fractions from each side), as polygon
 * points [[x, y]…] in 21600 units, or null if it can't be read or fills its whole rectangle.
 */
export function pictureOutline(img, crop = { l: 0, t: 0, r: 0, b: 0 }) {
  if (!img?.naturalWidth) return null;
  const sx = Math.max(0, crop.l) * img.naturalWidth;
  const sy = Math.max(0, crop.t) * img.naturalHeight;
  const sw = img.naturalWidth * (1 - Math.max(0, crop.l) - Math.max(0, crop.r));
  const sh = img.naturalHeight * (1 - Math.max(0, crop.t) - Math.max(0, crop.b));
  if (sw <= 0 || sh <= 0) return null;
  const scale = Math.min(1, SAMPLE / Math.max(sw, sh));
  const w = Math.max(2, Math.round(sw * scale));
  const h = Math.max(2, Math.round(sh * scale));
  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const g = canvas.getContext('2d', { willReadFrequently: true });
  let data;
  try {
    g.drawImage(img, sx, sy, sw, sh, 0, 0, w, h);
    data = g.getImageData(0, 0, w, h).data;
  } catch {
    return null; // not readable here
  }
  const left = [];
  const right = [];
  let full = true;
  for (let y = 0; y < h; y++) {
    let a = -1;
    let b = -1;
    for (let x = 0; x < w; x++) {
      if (data[(y * w + x) * 4 + 3] > VISIBLE) {
        if (a < 0) a = x;
        b = x;
      }
    }
    if (a !== 0 || b !== w - 1) full = false;
    if (a < 0) continue; // an empty row: the outline passes it by
    left.push([a, y], [a, y + 1]);
    right.push([b + 1, y], [b + 1, y + 1]);
  }
  if (full || !left.length) return null;
  const points = [...left, ...right.reverse()];
  // Points on a straight line add nothing.
  const kept = points.filter((p, i) => {
    const prev = points[(i - 1 + points.length) % points.length];
    const next = points[(i + 1) % points.length];
    return (next[0] - prev[0]) * (p[1] - prev[1]) !== (next[1] - prev[1]) * (p[0] - prev[0]);
  });
  return kept.map(([x, y]) => [Math.round((x / w) * 21600), Math.round((y / h) * 21600)]);
}
