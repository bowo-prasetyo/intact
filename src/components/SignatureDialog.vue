<template>
  <div class="sig-backdrop" @mousedown.self="$emit('close')">
    <section ref="dialog" class="sig" tabindex="-1" role="dialog" aria-modal="true" aria-labelledby="sig-title" @keydown.esc.stop="$emit('close')">
      <header>
        <h2 id="sig-title">Draw your signature</h2>
        <button type="button" class="close" aria-label="Close" @click="$emit('close')">×</button>
      </header>
      <p class="hint">Sign with your finger, a pen or the mouse. It goes in at the caret, as a picture, or in place of the line you selected. It stays on this device: nothing is uploaded or kept.</p>
      <div class="pad-wrap">
        <canvas
          ref="canvas"
          class="pad"
          role="img"
          :aria-label="drawn ? 'Your signature' : 'Signature pad: draw here'"
          @pointerdown="start"
          @pointermove="move"
          @pointerup="end"
          @pointercancel="end"
          @lostpointercapture="end"
        />
        <span class="line" aria-hidden="true" />
        <span v-if="!drawn" class="placeholder" aria-hidden="true">Sign here</span>
      </div>
      <div class="row">
        <fieldset class="ink">
          <legend>Ink</legend>
          <label v-for="c in INKS" :key="c.value" :title="c.name">
            <input v-model="ink" type="radio" name="sig-ink" :value="c.value" />
            <span class="swatch" :style="{ background: c.value }" />{{ c.name }}
          </label>
        </fieldset>
        <button type="button" :disabled="!drawn" @click="clear">Clear</button>
      </div>
      <details class="typed">
        <summary>Or type your name</summary>
        <label for="sig-typed">Name, written in a handwriting font</label>
        <input id="sig-typed" v-model="typed" type="text" autocomplete="name" spellcheck="false" @input="drawTyped" />
      </details>
      <div class="actions">
        <button type="button" class="confirm" :disabled="!drawn" @click="insert">Insert signature</button>
        <button type="button" @click="$emit('close')">Cancel</button>
      </div>
    </section>
  </div>
</template>

<script setup>
// A signature drawn (or typed) here, for a form's "Signature: ____": the strokes are cropped and
// handed over as a PNG with a transparent background, to be inserted as an ordinary picture (alt
// text "Signature"). Everything happens in this page: the drawing isn't stored or sent anywhere.
import { ref, watch, onMounted } from 'vue';

const emit = defineEmits(['close', 'insert']);

const INKS = [
  { name: 'Black', value: '#1a1a1a' },
  { name: 'Blue', value: '#1f3a93' },
];
const WIDTH = 480; // the pad, in CSS pixels; a signature across it is inserted half as wide (about 6 cm)
const HEIGHT = 160;
const PEN = 2.6;

const canvas = ref(null);
const dialog = ref(null);
const ink = ref(INKS[0].value);
const drawn = ref(false);
const typed = ref('');

let ctx = null;
let scale = 1;
let strokes = []; // [{ ink, points: [[x, y], …] }], in CSS pixels, to draw again in another ink
let current = null;

function setup() {
  const el = canvas.value;
  scale = Math.max(2, Math.min(4, window.devicePixelRatio || 1)); // sharp when printed
  el.width = WIDTH * scale;
  el.height = HEIGHT * scale;
  ctx = el.getContext('2d');
  ctx.setTransform(scale, 0, 0, scale, 0, 0);
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
}

function point(e) {
  const box = canvas.value.getBoundingClientRect();
  return [((e.clientX - box.left) * WIDTH) / box.width, ((e.clientY - box.top) * HEIGHT) / box.height];
}

function start(e) {
  if (e.button > 0) return;
  e.preventDefault();
  canvas.value.setPointerCapture?.(e.pointerId);
  if (typed.value) {
    typed.value = ''; // drawing replaces a typed name
    strokes = [];
    redraw();
  }
  current = { ink: ink.value, points: [point(e)] };
  strokes.push(current);
  redraw();
}

function move(e) {
  if (!current) return;
  e.preventDefault();
  for (const ev of e.getCoalescedEvents?.() || [e]) current.points.push(point(ev));
  redraw();
}

function end() {
  if (!current) return;
  current = null;
  drawn.value = strokes.length > 0;
}

/** Draws a stroke through its points' midpoints, so it is smooth rather than jagged. */
function drawStroke(s) {
  const p = s.points;
  ctx.strokeStyle = s.ink;
  ctx.fillStyle = s.ink;
  ctx.lineWidth = PEN;
  if (p.length === 1) {
    ctx.beginPath();
    ctx.arc(p[0][0], p[0][1], PEN / 2, 0, Math.PI * 2);
    ctx.fill();
    return;
  }
  ctx.beginPath();
  ctx.moveTo(p[0][0], p[0][1]);
  for (let i = 1; i < p.length - 1; i++) {
    ctx.quadraticCurveTo(p[i][0], p[i][1], (p[i][0] + p[i + 1][0]) / 2, (p[i][1] + p[i + 1][1]) / 2);
  }
  ctx.lineTo(p[p.length - 1][0], p[p.length - 1][1]);
  ctx.stroke();
}

function redraw() {
  ctx.clearRect(0, 0, WIDTH, HEIGHT);
  if (typed.value) {
    ctx.fillStyle = ink.value;
    let size = 64;
    const font = (px) => `italic ${px}px 'Segoe Script', 'Bradley Hand', 'Brush Script MT', 'Lucida Handwriting', cursive`;
    ctx.font = font(size);
    while (size > 18 && ctx.measureText(typed.value).width > WIDTH - 40) ctx.font = font((size -= 4));
    ctx.textBaseline = 'alphabetic';
    ctx.fillText(typed.value, 20, HEIGHT * 0.68);
    return;
  }
  for (const s of strokes) drawStroke(s);
}

function drawTyped() {
  strokes = [];
  redraw();
  drawn.value = !!typed.value.trim();
}

function clear() {
  strokes = [];
  typed.value = '';
  redraw();
  drawn.value = false;
}

// A change of ink recolours what is drawn, as the choice is for the whole signature.
function recolour() {
  for (const s of strokes) s.ink = ink.value;
  redraw();
}

/** The drawn part of the pad, in canvas pixels, with a little room around it; null if empty. */
function inkBounds() {
  const { width, height } = canvas.value;
  const data = ctx.getImageData(0, 0, width, height).data;
  let x0 = width, y0 = height, x1 = -1, y1 = -1;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (data[(y * width + x) * 4 + 3] > 8) {
        if (x < x0) x0 = x;
        if (x > x1) x1 = x;
        if (y < y0) y0 = y;
        if (y > y1) y1 = y;
      }
    }
  }
  if (x1 < 0) return null;
  const pad = Math.round(4 * scale);
  x0 = Math.max(0, x0 - pad);
  y0 = Math.max(0, y0 - pad);
  x1 = Math.min(width - 1, x1 + pad);
  y1 = Math.min(height - 1, y1 + pad);
  return { x: x0, y: y0, w: x1 - x0 + 1, h: y1 - y0 + 1 };
}

function insert() {
  const box = inkBounds();
  if (!box) return;
  const out = document.createElement('canvas');
  out.width = box.w;
  out.height = box.h;
  out.getContext('2d').drawImage(canvas.value, box.x, box.y, box.w, box.h, 0, 0, box.w, box.h);
  out.toBlob((blob) => {
    if (!blob) return;
    const file = new File([blob], 'signature.png', { type: 'image/png' });
    // Half the pad's size: in pixels at 96 per inch, as pictures are inserted.
    emit('insert', { file, width: Math.max(1, Math.round(box.w / scale / 2)) });
  }, 'image/png');
}

watch(ink, recolour);

onMounted(() => {
  setup();
  redraw();
  dialog.value?.focus();
});
</script>

<style scoped>
.sig-backdrop {
  position: fixed; inset: 0; z-index: 20; display: flex; align-items: center; justify-content: center;
  padding: 16px; background: rgb(0 0 0 / 35%);
}
.sig:focus { outline: none; }
.sig {
  width: min(540px, 100%); max-height: calc(100vh - 32px); overflow: auto; padding: 0 20px 18px;
  background: var(--surface); color: var(--text); border-radius: 8px; box-shadow: 0 12px 40px rgb(0 0 0 / 30%);
  font-size: 13px;
}
header { display: flex; align-items: center; justify-content: space-between; padding: 16px 0 4px; }
h2 { margin: 0; font-size: 18px; }
.close { border: 0; background: none; color: inherit; font-size: 22px; line-height: 1; width: 32px; height: 32px; border-radius: 4px; cursor: pointer; }
.close:hover { background: var(--hover); }
.hint { margin: 0 0 10px; color: var(--muted); }
/* The pad is always white, like the page the signature goes on. */
.pad-wrap { position: relative; border: 1px solid var(--border); border-radius: 6px; background: #fff; overflow: hidden; }
.pad { display: block; width: 100%; aspect-ratio: 3 / 1; touch-action: none; cursor: crosshair; }
.line { position: absolute; left: 6%; right: 6%; top: 72%; border-top: 1px solid #b8bec4; pointer-events: none; }
.placeholder { position: absolute; left: 6%; top: 74%; color: #8a9096; font-size: 12px; pointer-events: none; }
.row { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin: 10px 0 4px; flex-wrap: wrap; }
.ink { display: flex; gap: 14px; align-items: center; border: 0; margin: 0; padding: 0; }
.ink legend { float: left; margin-right: 10px; color: var(--muted); }
.ink label { display: inline-flex; align-items: center; gap: 5px; cursor: pointer; }
.swatch { width: 12px; height: 12px; border-radius: 50%; display: inline-block; }
.typed { margin: 6px 0 2px; }
.typed summary { cursor: pointer; color: var(--accent); width: fit-content; }
.typed label { display: block; margin: 8px 0 4px; color: var(--muted); }
.typed input { width: 100%; box-sizing: border-box; padding: 6px 8px; font: inherit; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; }
.actions { display: flex; gap: 8px; margin-top: 14px; }
button:not(.close) {
  padding: 6px 12px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; font: inherit; cursor: pointer;
}
button:not(.close):hover:not(:disabled) { background: var(--hover); }
button:disabled { opacity: 0.5; cursor: default; }
.confirm { background: var(--brand) !important; border-color: var(--brand) !important; color: #fff !important; }
button:focus-visible, input:focus-visible, summary:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
</style>
