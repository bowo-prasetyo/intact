// Artistic effects, drawn in the browser as Word draws them into a picture: the picture is redrawn
// on a canvas with the effect (blur, a pencil sketch, a line drawing, glowing edges, a photocopy)
// and saved as a new picture; the original is kept beside it, so the effect can be taken off.

const MAX = 1600; // the drawn picture is at most this many pixels across

function load(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('The picture couldn’t be read'));
    img.src = url;
  });
}

/** Edge strength (0–255) at each pixel, from the luminance (Sobel). */
function edges(lum, w, h) {
  const out = new Float32Array(w * h);
  for (let y = 1; y < h - 1; y++) {
    for (let x = 1; x < w - 1; x++) {
      const i = y * w + x;
      const gx = -lum[i - w - 1] - 2 * lum[i - 1] - lum[i + w - 1] + lum[i - w + 1] + 2 * lum[i + 1] + lum[i + w + 1];
      const gy = -lum[i - w - 1] - 2 * lum[i - w] - lum[i - w + 1] + lum[i + w - 1] + 2 * lum[i + w] + lum[i + w + 1];
      out[i] = Math.min(255, Math.hypot(gx, gy) / 4);
    }
  }
  return out;
}

/** The picture at `url` drawn with artistic effect `effect` (see pictures.js ARTISTIC), as PNG bytes. */
export async function renderArtistic(url, effect) {
  const img = await load(url);
  const scale = Math.min(1, MAX / Math.max(img.naturalWidth, img.naturalHeight));
  const w = Math.max(1, Math.round(img.naturalWidth * scale));
  const h = Math.max(1, Math.round(img.naturalHeight * scale));
  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const g = canvas.getContext('2d', { willReadFrequently: true });
  if (effect === 'blur') {
    g.filter = `blur(${Math.max(2, Math.round(Math.max(w, h) / 90))}px)`;
    g.drawImage(img, 0, 0, w, h);
  } else {
    g.drawImage(img, 0, 0, w, h);
    const data = g.getImageData(0, 0, w, h);
    const p = data.data;
    const lum = new Float32Array(w * h);
    for (let i = 0; i < w * h; i++) lum[i] = 0.299 * p[i * 4] + 0.587 * p[i * 4 + 1] + 0.114 * p[i * 4 + 2];
    const e = edges(lum, w, h);
    for (let i = 0; i < w * h; i++) {
      const o = i * 4;
      let r;
      let gg;
      let b;
      if (effect === 'pencilGrayscale') {
        r = gg = b = Math.max(0, 255 - e[i] * 1.6 - (255 - lum[i]) * 0.15);
      } else if (effect === 'lineDrawing') {
        r = gg = b = e[i] > 40 ? 30 : 255;
      } else if (effect === 'glowEdges') {
        const k = Math.min(1, (e[i] * 2) / 255);
        r = p[o] * k + 255 * k * 0.3;
        gg = p[o + 1] * k + 255 * k * 0.3;
        b = p[o + 2] * k + 255 * k * 0.3;
      } else { // photocopy: dark tones and edges black, the rest white
        r = gg = b = lum[i] < 100 || e[i] > 70 ? 20 : 250;
      }
      p[o] = Math.min(255, r);
      p[o + 1] = Math.min(255, gg);
      p[o + 2] = Math.min(255, b);
    }
    g.putImageData(data, 0, 0);
  }
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
  return new Uint8Array(await blob.arrayBuffer());
}
