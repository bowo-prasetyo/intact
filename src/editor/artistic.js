// Word's artistic effects, drawn in the browser as Word draws them into a picture: the picture is
// redrawn on a canvas with the effect and saved as a new picture; the original is kept beside it, so
// the effect can be changed or taken off. Each effect has an amount (0–100), as Word's Artistic
// Effects Options have a main setting (a blur's radius, a pencil's size, a grain's size…); noise
// comes from a fixed seed, so an effect always draws the same.

const MAX = 1600; // the drawn picture is at most this many pixels across

function load(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('The picture couldn’t be read'));
    img.src = url;
  });
}

/** A seeded random number generator (mulberry32): the same numbers every time. */
function random(seed = 1234567) {
  let a = seed;
  return () => {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const clamp = (v) => (v < 0 ? 0 : v > 255 ? 255 : v);
const lumOf = (p, i) => 0.299 * p[i] + 0.587 * p[i + 1] + 0.114 * p[i + 2];

/** The picture drawn blurred by `r` px: its pixels. */
function blurred(img, w, h, r) {
  const c = document.createElement('canvas');
  c.width = w;
  c.height = h;
  const g = c.getContext('2d', { willReadFrequently: true });
  g.filter = r > 0 ? `blur(${r}px)` : 'none';
  g.drawImage(img, 0, 0, w, h);
  return g.getImageData(0, 0, w, h).data;
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

function lums(p, n) {
  const out = new Float32Array(n);
  for (let i = 0; i < n; i++) out[i] = lumOf(p, i * 4);
  return out;
}

const posterize = (v, levels) => Math.round((v / 255) * (levels - 1)) * (255 / (levels - 1));

/**
 * Each effect: (pixels out, the source pixels, a blur of them, size, amount 0–1, helpers) → fills
 * `out` (RGBA). `src` is the picture as it is.
 */
const EFFECTS = {
  blur: null, // drawn with the canvas's own blur (see renderArtistic)
  pencilGrayscale(out, src, ctx) {
    const { e, lum, n } = ctx;
    for (let i = 0; i < n; i++) out.set3(i, Math.max(0, 255 - e[i] * (1 + ctx.a * 2) - (255 - lum[i]) * 0.15));
  },
  pencilSketch(out, src, ctx) {
    const { e, n } = ctx;
    for (let i = 0; i < n; i++) {
      const k = Math.min(1, (e[i] * (1 + ctx.a * 2)) / 255);
      const o = i * 4;
      out.rgb(i, 255 - (255 - src[o]) * k, 255 - (255 - src[o + 1]) * k, 255 - (255 - src[o + 2]) * k);
    }
  },
  lineDrawing(out, src, ctx) {
    const t = 70 - ctx.a * 55;
    for (let i = 0; i < ctx.n; i++) out.set3(i, ctx.e[i] > t ? 30 : 255);
  },
  chalkSketch(out, src, ctx) {
    const { e, lum, n, rnd } = ctx;
    for (let i = 0; i < n; i++) {
      const chalk = Math.min(1, e[i] / 120 + (lum[i] / 255) * 0.5) * (0.7 + rnd() * 0.3 * (1 + ctx.a));
      out.set3(i, 55 + chalk * 200);
    }
  },
  marker(out, src, ctx) {
    const { soft, e, n } = ctx;
    const levels = Math.round(6 - ctx.a * 3);
    for (let i = 0; i < n; i++) {
      const o = i * 4;
      const dark = e[i] > 60 ? 0.45 : 1;
      out.rgb(i, posterize(soft[o], levels) * dark, posterize(soft[o + 1], levels) * dark, posterize(soft[o + 2], levels) * dark);
    }
  },
  paintStrokes(out, src, ctx) {
    const { w, h } = ctx;
    const len = Math.round(3 + ctx.a * 9);
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        let r = 0;
        let g = 0;
        let b = 0;
        let k = 0;
        for (let s = -len; s <= len; s++) { // a stroke along the diagonal
          const xx = Math.min(w - 1, Math.max(0, x + s));
          const yy = Math.min(h - 1, Math.max(0, y + s));
          const o = (yy * w + xx) * 4;
          r += src[o]; g += src[o + 1]; b += src[o + 2]; k++;
        }
        out.rgb(y * w + x, posterize(r / k, 8), posterize(g / k, 8), posterize(b / k, 8));
      }
    }
  },
  paintBrush(out, src, ctx) {
    const { w, h, soft, rnd } = ctx;
    const spread = 2 + ctx.a * 6;
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const xx = Math.min(w - 1, Math.max(0, Math.round(x + (rnd() - 0.5) * spread)));
        const yy = Math.min(h - 1, Math.max(0, Math.round(y + (rnd() - 0.5) * spread)));
        const o = (yy * w + xx) * 4;
        out.rgb(y * w + x, posterize(soft[o], 10), posterize(soft[o + 1], 10), posterize(soft[o + 2], 10));
      }
    }
  },
  glowDiffused(out, src, ctx) {
    const { soft, n } = ctx;
    for (let i = 0; i < n; i++) {
      const o = i * 4;
      const screen = (a, b) => 255 - ((255 - a) * (255 - b)) / 255; // the blur laid over in Screen
      const k = 0.4 + ctx.a * 0.5;
      out.rgb(i, src[o] * (1 - k) + screen(src[o], soft[o]) * k, src[o + 1] * (1 - k) + screen(src[o + 1], soft[o + 1]) * k, src[o + 2] * (1 - k) + screen(src[o + 2], soft[o + 2]) * k);
    }
  },
  lightScreen(out, src, ctx) {
    const { w, h } = ctx;
    const step = Math.round(6 - ctx.a * 3);
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const i = y * w + x;
        const o = i * 4;
        const line = x % step === 0 || y % step === 0 ? 60 : 0;
        out.rgb(i, src[o] + line, src[o + 1] + line, src[o + 2] + line);
      }
    }
  },
  watercolorSponge(out, src, ctx) {
    const { soft, e, n, rnd } = ctx;
    for (let i = 0; i < n; i++) {
      const o = i * 4;
      const t = (rnd() - 0.5) * 30 * (0.5 + ctx.a);
      const edge = e[i] > 50 ? 0.8 : 1;
      out.rgb(i, (posterize(soft[o], 7) + t) * edge, (posterize(soft[o + 1], 7) + t) * edge, (posterize(soft[o + 2], 7) + t) * edge);
    }
  },
  filmGrain(out, src, ctx) {
    const { n, rnd } = ctx;
    for (let i = 0; i < n; i++) {
      const o = i * 4;
      const grain = (rnd() + rnd() - 1) * (20 + ctx.a * 60);
      const l = lumOf(src, o);
      out.rgb(i, src[o] * 0.8 + l * 0.2 + grain, src[o + 1] * 0.8 + l * 0.2 + grain, src[o + 2] * 0.8 + l * 0.2 + grain);
    }
  },
  mosaicBubbles(out, src, ctx) {
    const { w, h } = ctx;
    const size = Math.max(4, Math.round(Math.max(w, h) / (60 - ctx.a * 40)));
    for (let cy = 0; cy < h; cy += size) {
      for (let cx = 0; cx < w; cx += size) {
        let r = 0;
        let g = 0;
        let b = 0;
        let k = 0;
        for (let y = cy; y < Math.min(h, cy + size); y++) {
          for (let x = cx; x < Math.min(w, cx + size); x++) {
            const o = (y * w + x) * 4;
            r += src[o]; g += src[o + 1]; b += src[o + 2]; k++;
          }
        }
        r /= k; g /= k; b /= k;
        for (let y = cy; y < Math.min(h, cy + size); y++) {
          for (let x = cx; x < Math.min(w, cx + size); x++) {
            const dx = (x - cx - size / 2) / (size / 2);
            const dy = (y - cy - size / 2) / (size / 2);
            const d = dx * dx + dy * dy;
            const shine = d < 1 ? Math.max(0, 1 - Math.hypot(dx + 0.35, dy + 0.35) * 2) * 90 : 0;
            const rim = d < 1 ? 1 - d * 0.25 : 0.7;
            out.rgb(y * w + x, r * rim + shine, g * rim + shine, b * rim + shine);
          }
        }
      }
    }
  },
  glass(out, src, ctx) {
    const { w, h, rnd } = ctx;
    const amp = 2 + ctx.a * 8;
    const field = new Float32Array((Math.ceil(w / 8) + 2) * (Math.ceil(h / 8) + 2) * 2).map(() => rnd() - 0.5);
    const fw = Math.ceil(w / 8) + 2;
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const f = (Math.floor(y / 8) * fw + Math.floor(x / 8)) * 2;
        const xx = Math.min(w - 1, Math.max(0, Math.round(x + field[f] * amp * 2)));
        const yy = Math.min(h - 1, Math.max(0, Math.round(y + field[f + 1] * amp * 2)));
        const o = (yy * w + xx) * 4;
        out.rgb(y * w + x, src[o], src[o + 1], src[o + 2]);
      }
    }
  },
  cement(out, src, ctx) {
    const { lum, w, n, rnd } = ctx;
    const noise = new Float32Array(n).map(() => rnd() * 255);
    for (let i = 0; i < n; i++) {
      const emboss = (noise[i] - (noise[i + w + 1] ?? noise[i])) * (0.3 + ctx.a * 0.5);
      out.set3(i, lum[i] * 0.75 + 40 + emboss);
    }
  },
  texturizer(out, src, ctx) {
    const { w, n, rnd } = ctx;
    const noise = new Float32Array(n).map(() => rnd() * 255);
    for (let i = 0; i < n; i++) {
      const o = i * 4;
      const emboss = (noise[i] - (noise[i + w + 1] ?? noise[i])) * (0.15 + ctx.a * 0.3);
      out.rgb(i, src[o] + emboss, src[o + 1] + emboss, src[o + 2] + emboss);
    }
  },
  crisscrossEtching(out, src, ctx) {
    const { w, h, lum } = ctx;
    const step = Math.round(7 - ctx.a * 3);
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const i = y * w + x;
        const l = lum[i];
        const one = (x + y) % step === 0 && l < 190;
        const two = (x - y + h * step) % step === 0 && l < 110;
        out.set3(i, one || two ? 40 : 250);
      }
    }
  },
  pastelsSmooth(out, src, ctx) {
    const { soft, n } = ctx;
    for (let i = 0; i < n; i++) {
      const o = i * 4;
      const l = lumOf(soft, o);
      const k = 0.3 + ctx.a * 0.3; // washed toward light, as pastels on paper
      out.rgb(i, posterize(soft[o] * (1 - k) + (l * 0.3 + 180) * k, 12), posterize(soft[o + 1] * (1 - k) + (l * 0.3 + 180) * k, 12), posterize(soft[o + 2] * (1 - k) + (l * 0.3 + 180) * k, 12));
    }
  },
  plasticWrap(out, src, ctx) {
    const { soft, w, n } = ctx;
    const sl = lums(soft, n);
    for (let i = 0; i < n; i++) {
      const o = i * 4;
      const ridge = Math.max(0, (sl[i] - (sl[i + w + 1] ?? sl[i])) * (2 + ctx.a * 4)); // light on the ridges of the wrap
      out.rgb(i, src[o] * 0.85 + ridge, src[o + 1] * 0.85 + ridge, src[o + 2] * 0.85 + ridge);
    }
  },
  cutout(out, src, ctx) {
    const { soft, n } = ctx;
    const levels = Math.round(5 - ctx.a * 2.5);
    for (let i = 0; i < n; i++) {
      const o = i * 4;
      out.rgb(i, posterize(soft[o], levels), posterize(soft[o + 1], levels), posterize(soft[o + 2], levels));
    }
  },
  photocopy(out, src, ctx) {
    const t = 80 + ctx.a * 50;
    for (let i = 0; i < ctx.n; i++) out.set3(i, ctx.lum[i] < t || ctx.e[i] > 70 ? 20 : 250);
  },
  glowEdges(out, src, ctx) {
    const { e, n } = ctx;
    for (let i = 0; i < n; i++) {
      const k = Math.min(1, (e[i] * (1.5 + ctx.a * 1.5)) / 255);
      const o = i * 4;
      out.rgb(i, src[o] * k + 255 * k * 0.3, src[o + 1] * k + 255 * k * 0.3, src[o + 2] * k + 255 * k * 0.3);
    }
  },
};

/**
 * The picture at `url` drawn with artistic effect `effect` (see pictures.js ARTISTIC) at `amount`
 * (0–100), as PNG bytes.
 */
export async function renderArtistic(url, effect, amount = 50) {
  const img = await load(url);
  const scale = Math.min(1, MAX / Math.max(img.naturalWidth, img.naturalHeight));
  const w = Math.max(1, Math.round(img.naturalWidth * scale));
  const h = Math.max(1, Math.round(img.naturalHeight * scale));
  const a = Math.max(0, Math.min(100, Number(amount) || 0)) / 100;
  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const g = canvas.getContext('2d', { willReadFrequently: true });
  if (effect === 'blur' || !EFFECTS[effect]) {
    g.filter = `blur(${Math.max(1, Math.round((Math.max(w, h) / 200) * (1 + a * 4)))}px)`;
    g.drawImage(img, 0, 0, w, h);
  } else {
    g.drawImage(img, 0, 0, w, h);
    const data = g.getImageData(0, 0, w, h);
    const src = new Uint8ClampedArray(data.data);
    const n = w * h;
    const lum = lums(src, n);
    const p = data.data;
    const out = {
      rgb(i, r, gg, b) { p[i * 4] = clamp(r); p[i * 4 + 1] = clamp(gg); p[i * 4 + 2] = clamp(b); },
      set3(i, v) { const c = clamp(v); p[i * 4] = c; p[i * 4 + 1] = c; p[i * 4 + 2] = c; },
    };
    const ctx = {
      w, h, n, a, lum, e: edges(lum, w, h), rnd: random(),
      soft: blurred(img, w, h, Math.max(1, Math.round(Math.max(w, h) / 300 * (1 + a * 2)))),
    };
    EFFECTS[effect](out, src, ctx);
    g.putImageData(data, 0, 0); // the alpha stays as it was
  }
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
  return new Uint8Array(await blob.arrayBuffer());
}

export const EFFECT_NAMES = Object.keys(EFFECTS);
