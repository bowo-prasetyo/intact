// Pictures laid out as Word lays them out: selected by a click, resized, cropped, rotated, wrapped and positioned, with Layout Options, picture styles, effects, 3-D and alt text.
// Split out of EditorView.vue: see src/views/editor/README.md. The editor's shared state comes in as
// `ctx`, read when it is used, never copied at setup.
import { reactive, ref, nextTick } from 'vue';
import { t } from '../../i18n.js';
import { WRAPS, PICTURE_STYLES, SHAPES, CAMERAS, ARTISTIC } from '../../docx/pictures.js';
import { CAMERA_ANGLES } from '../../docx/lighting.js';
import { EMU_PER } from '../../editor/units.js';
import { pictureOutline } from '../../editor/outline.js';
import { pictureZoom } from '../../editor/fitPictures.js';
import { renderArtistic } from '../../editor/artistic.js';

/**
 * `ctx`: { BLOCK_SELECTOR, changes, doc, docUnit, lockMode, markDirty, message, page, savedRange, sheetZoom, undoStack }, as the editor has them.
 * Returns { HANDLES, apply3D, applyAlt, applyArtistic, applyCameraPreset, applyCrop, applyEffects, applyPosition, applyRotation, applyShape, applySize, applyStyle, applyWrap, closePicturePanel, flip, layoutOptionsButton, picture, picturePanel, picturePanelEl, resetPictureSize, rotateBy, round, startCrop, startResize, startRotate, toggleCropping, togglePicturePanel, onPictureMousedown, pictureInSelection, selectPicture, showPicture, updatePictureFrame }.
 */
export function usePictureLayout(ctx) {
  const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
  const PX = 9525; // EMU per pixel
  const picture = reactive({ el: null, x: 0, y: 0, w: 0, h: 0, wrap: 'inline', cropping: false, shape: false });
  const picturePanel = reactive({
    unit: 'cm', // lengths in centimetres or inches (see editor/units.js)
    open: false, x: 0, y: 0, wrap: 'inline', hAlign: 'left', hOffset: 0, vAlign: 'offset', vOffset: 0, width: 0, height: 0, lock: true, alt: '',
    cropL: 0, cropT: 0, cropR: 0, cropB: 0, rot: 0, flipH: false, flipV: false,
    border: 'none', borderColor: '#000000', shadow: false, glow: false, softEdge: false, reflection: false, colour: 'normal', brightness: 0, contrast: 0,
    style: '', shape: 'rect', bevel: '', camera: '', artistic: '', drawing: false,
    bevelW: 6, bevelH: 2, material: '', lightRig: 'threePt', lightAngle: 0, rotX: 0, rotY: 0, rotZ: 0, perspective: 0, amount: 50,
  });
  const BORDERS = { thin: 0.75, medium: 1.5, thick: 3 };
  const picturePanelEl = ref(null);
  const layoutOptionsButton = ref(null);
  const round = (n, d = 2) => Math.round(n * 10 ** d) / 10 ** d;

  /** The picture a selection is on: exactly one picture and nothing else (as a click selects it). */
  function pictureInSelection(range) {
    if (!range || range.collapsed) return null;
    const frag = range.cloneContents();
    if (frag.textContent.trim() || frag.querySelectorAll('.picture').length !== 1) return null;
    return [...ctx.page.value.querySelectorAll('.picture')].find((p) => range.intersectsNode(p)) || null;
  }

  /** Selects a picture, as a click does in Word: the frame with its handles shows around it. */
  function selectPicture(span) {
    const range = document.createRange();
    range.selectNode(span);
    ctx.page.value.focus({ preventScroll: true });
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    ctx.savedRange = range.cloneRange();
    showPicture(span);
  }

  function showPicture(span) {
    if (span && ctx.lockMode.value) {
      // Resizing or moving a picture isn't tracked: under "tracked changes only" it can't be done either.
      picture.el = null;
      return;
    }
    if (picture.el !== span) {
      picturePanel.open = false;
      picture.cropping = false;
    }
    picture.el = span;
    picture.shape = span?.dataset.shape != null;
    updatePictureFrame();
    if (span) ctx.message.value = picture.shape ? 'Shape: drag it to move it, a handle to resize it, or open Layout Options to turn it' : 'Picture: drag a corner to resize, or open Layout Options for wrapping and position';
  }

  /** Puts the frame over the selected picture (in the sheet's coordinates). */
  function updatePictureFrame() {
    const span = picture.el;
    if (!span?.isConnected) {
      picture.el = null;
      picturePanel.open = false;
      return;
    }
    const sheet = ctx.page.value.closest('.docx-sheet').getBoundingClientRect();
    const r = span.getBoundingClientRect(); // the picture's frame (what shows of it, cropped)
    const z = ctx.sheetZoom.value; // the frame is drawn in the page's own pixels
    Object.assign(picture, { x: (r.left - sheet.left) / z, y: (r.top - sheet.top) / z, w: r.width / z, h: r.height / z, wrap: span.dataset.wrap || 'inline' });
  }

  /** Applies a change to the selected picture as one undoable step, keeping it selected. */
  function changePicture(change, done) {
    const span = picture.el;
    if (!span?.isConnected) return;
    ctx.undoStack.flush();
    change(span);
    ctx.undoStack.commit();
    ctx.markDirty();
    nextTick(() => {
      if (span.isConnected) selectPicture(span);
      fillPicturePanel();
      // A line, arrow or shape is changed as a picture is, and called what it is.
      if (done) ctx.message.value = picture.shape ? done.replace(/^Picture\b/, 'Shape') : done;
    });
  }

  /** Drag a handle: corners keep the proportions, sides stretch, as in Word. */
  function startResize(e, handle) {
    const span = picture.el;
    const info = span && ctx.doc.pictureLayout(span);
    if (!info) return;
    const look = ctx.doc.pictureLookOf(span);
    const start = { x: e.clientX, y: e.clientY, w: info.cx / PX, h: info.cy / PX };
    const zoom = pictureZoom(span) * ctx.sheetZoom.value; // shown smaller (or zoomed): the pointer moves in shown pixels
    const target = e.currentTarget;
    target.setPointerCapture(e.pointerId);
    let size = { w: start.w, h: start.h };
    const move = (ev) => {
      const dx = ((ev.clientX - start.x) / zoom) * (handle.includes('w') ? -1 : handle.includes('e') ? 1 : 0);
      const dy = ((ev.clientY - start.y) / zoom) * (handle.includes('n') ? -1 : handle.includes('s') ? 1 : 0);
      let w = start.w + dx;
      let h = start.h + dy;
      if (handle.length === 2) {
        const scale = Math.max(w / start.w, h / start.h);
        w = start.w * scale;
        h = start.h * scale;
      }
      size = { w: Math.max(8, w), h: Math.max(8, h) };
      ctx.doc.stylePicture(span, { ...info, cx: size.w * PX, cy: size.h * PX }, look); // a preview until it is let go
      updatePictureFrame();
    };
    const up = () => {
      target.removeEventListener('pointermove', move);
      target.removeEventListener('pointerup', up);
      target.removeEventListener('pointercancel', up);
      if (size.w === start.w && size.h === start.h) return;
      const unit = ctx.docUnit();
      changePicture((s) => ctx.doc.setPictureSize(s, size.w * PX, size.h * PX), `Picture resized to ${round(size.w * PX / EMU_PER[unit], 1)} × ${round(size.h * PX / EMU_PER[unit], 1)} ${unit}`);
    };
    target.addEventListener('pointermove', move);
    target.addEventListener('pointerup', up);
    target.addEventListener('pointercancel', up);
  }

  /**
   * A press on a picture selects it; dragging a selected floating picture moves it, its new place
   * measured from the left margin and from its paragraph (so it moves with the text, as by default
   * in Word).
   */
  function onPictureMousedown(e, span) {
    e.preventDefault();
    const wasSelected = picture.el === span;
    selectPicture(span);
    if (!wasSelected || (span.dataset.wrap || 'inline') === 'inline' || e.button !== 0) return;
    const start = { x: e.clientX, y: e.clientY };
    let moved = false;
    const move = (ev) => {
      const dx = ev.clientX - start.x;
      const dy = ev.clientY - start.y;
      if (!moved && Math.hypot(dx, dy) < 4) return;
      moved = true;
      const z = ctx.sheetZoom.value; // moved in shown pixels, placed in the page's own
      span.style.transform = `translate(${dx / z}px, ${dy / z}px) ${span.dataset.turn || ''}`;
      updatePictureFrame();
    };
    const up = (ev) => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      if (!moved) return;
      const r = span.getBoundingClientRect(); // where it was dropped
      span.style.transform = span.dataset.turn || '';
      const sheet = ctx.page.value.closest('.docx-sheet').getBoundingClientRect();
      const block = span.closest(ctx.BLOCK_SELECTOR) || span.parentElement;
      const z = ctx.sheetZoom.value;
      const left = (r.left - sheet.left) / z - ctx.doc.page.left;
      const top = (r.top - block.getBoundingClientRect().top) / z;
      changePicture((s) => ctx.doc.setPicturePosition(s, {
        h: { relativeFrom: 'margin', offset: Math.round(left) * PX },
        v: { relativeFrom: 'paragraph', offset: Math.round(top) * PX },
      }), 'Picture moved');
      void ev;
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  }

  /** Fills the Layout Options panel from the selected picture. */
  function fillPicturePanel() {
    if (!picture.el) return;
    const info = ctx.doc.pictureLayout(picture.el);
    if (!info) return;
    const paper = ctx.doc.page; // the page set up (not the editor's page element)
    const unit = ctx.docUnit();
    const PER_UNIT = EMU_PER[unit]; // EMU per unit shown: a centimetre or an inch
    const h = info.h || { relativeFrom: 'margin', align: 'left' };
    const v = info.v || { relativeFrom: 'paragraph', offset: 0 };
    let hAlign = 'word';
    let hOffset = 0;
    if (h.align && ['left', 'center', 'right'].includes(h.align) && ['margin', 'column'].includes(h.relativeFrom)) hAlign = h.align;
    else if (h.align == null && ['margin', 'column', 'page'].includes(h.relativeFrom)) {
      hAlign = 'offset';
      hOffset = round(((h.offset || 0) / PX - (h.relativeFrom === 'page' ? paper.left : 0)) * PX / PER_UNIT);
    }
    let vAlign = 'word';
    if (v.align && ['top', 'center', 'bottom'].includes(v.align) && v.relativeFrom === 'margin') vAlign = v.align;
    else if (v.align == null && ['paragraph', 'line'].includes(v.relativeFrom)) vAlign = 'offset';
    const look = ctx.doc.pictureLookOf(picture.el);
    const border = look.border ? (look.border.width <= 1 ? 'thin' : look.border.width <= 2 ? 'medium' : 'thick') : 'none';
    Object.assign(picturePanel, {
      unit, wrap: info.wrap, hAlign, hOffset, vAlign, vOffset: vAlign === 'offset' ? round((v.offset || 0) / PER_UNIT) : 0,
      width: round(info.cx / PER_UNIT), height: round(info.cy / PER_UNIT), alt: info.descr,
      cropL: round(look.crop.l * 100, 1), cropT: round(look.crop.t * 100, 1), cropR: round(look.crop.r * 100, 1), cropB: round(look.crop.b * 100, 1),
      rot: round(look.rot, 1), flipH: look.flipH, flipV: look.flipV,
      border, borderColor: `#${look.border?.color || '000000'}`.toLowerCase(), shadow: !!look.shadow, glow: !!look.glow, softEdge: !!look.softEdge,
      reflection: !!look.reflection, colour: look.greyscale ? 'greyscale' : 'normal', brightness: round(look.brightness * 100), contrast: round(look.contrast * 100),
      style: '', shape: look.geometry, bevel: look.bevel?.prst || '', camera: look.camera || '', artistic: look.artistic?.effect || '',
      amount: look.artistic?.amount ?? 50,
      bevelW: round((look.bevel?.w ?? 76200) / 12700, 1), bevelH: round((look.bevel?.h ?? 25400) / 12700, 1),
      material: look.material || '', lightRig: look.light?.rig || 'threePt', lightAngle: round(look.light?.angle || 0),
      ...(() => {
        const [x, y, z, fov] = look.rotation3d ? [look.rotation3d.x, look.rotation3d.y, look.rotation3d.z, look.rotation3d.perspective] : CAMERA_ANGLES[look.camera] || [0, 0, 0, 0];
        return { rotX: round(x, 1), rotY: round(y, 1), rotZ: round(z, 1), perspective: round(fov, 1) };
      })(),
    });
  }

  function togglePicturePanel() {
    if (picturePanel.open) return closePicturePanel();
    fillPicturePanel();
    const r = layoutOptionsButton.value.getBoundingClientRect();
    Object.assign(picturePanel, { open: true, x: Math.max(8, Math.min(r.right + 6, window.innerWidth - 310)), y: Math.max(8, r.top) });
    nextTick(() => {
      const el = picturePanelEl.value;
      if (!el) return;
      picturePanel.y = Math.max(8, Math.min(picturePanel.y, window.innerHeight - el.offsetHeight - 8)); // all of it on screen
      el.querySelector('input:checked, input')?.focus();
    });
  }

  function closePicturePanel() {
    picturePanel.open = false;
    if (picture.el?.isConnected) selectPicture(picture.el);
    else ctx.page.value.focus();
  }

  const WRAP_NAMES = Object.fromEntries(WRAPS);

  function applyWrap(wrap) {
    // Tight and Through: the text wraps to the picture's outline (what isn't transparent), as in Word.
    const outline = (wrap === 'tight' || wrap === 'through') && picture.el
      ? pictureOutline(picture.el.querySelector('img'), ctx.doc.pictureLookOf(picture.el).crop) : null;
    changePicture((s) => ctx.doc.setPictureWrap(s, wrap, outline), `Text wrapping: ${WRAP_NAMES[wrap]}${outline ? ', around its outline' : ''}`);
  }

  function applyPosition(which) {
    const p = picturePanel;
    const PER_UNIT = EMU_PER[p.unit];
    const h = which === 'h' && p.hAlign !== 'word'
      ? (p.hAlign === 'offset' ? { relativeFrom: 'margin', offset: (Number(p.hOffset) || 0) * PER_UNIT } : { relativeFrom: 'margin', align: p.hAlign })
      : null;
    const v = which === 'v' && p.vAlign !== 'word'
      ? (p.vAlign === 'offset' ? { relativeFrom: 'paragraph', offset: (Number(p.vOffset) || 0) * PER_UNIT } : { relativeFrom: 'margin', align: p.vAlign })
      : null;
    if (!h && !v) return;
    changePicture((s) => ctx.doc.setPicturePosition(s, { h, v }), 'Picture positioned');
  }

  function applySize(which) {
    const info = ctx.doc.pictureLayout(picture.el);
    if (!info) return;
    const PER_UNIT = EMU_PER[picturePanel.unit];
    let w = Math.max(0.1, Number(picturePanel.width) || 0.1) * PER_UNIT;
    let h = Math.max(0.1, Number(picturePanel.height) || 0.1) * PER_UNIT;
    if (picturePanel.lock) {
      if (which === 'width') h = (w * info.cy) / info.cx;
      else w = (h * info.cx) / info.cy;
    }
    changePicture((s) => ctx.doc.setPictureSize(s, w, h), `Picture resized to ${round(w / PER_UNIT, 1)} × ${round(h / PER_UNIT, 1)} ${picturePanel.unit}`);
  }

  /** Back to the picture's own size (at 96 pixels an inch), narrowed to fit between the margins. */
  function resetPictureSize() {
    const img = picture.el?.querySelector('img');
    if (!img?.naturalWidth) return;
    const max = ctx.doc.page.width - ctx.doc.page.left - ctx.doc.page.right;
    const scale = Math.min(1, max / img.naturalWidth);
    changePicture((s) => ctx.doc.setPictureSize(s, img.naturalWidth * scale * PX, img.naturalHeight * scale * PX), 'Picture back to its own size');
  }

  /** Crop with the handles, as Word's Crop button: dragging a handle cuts (or uncuts) that side. */
  function toggleCropping() {
    picture.cropping = !picture.cropping;
    ctx.message.value = picture.cropping ? 'Crop: drag the handles to cut the picture; Esc or Done cropping when finished' : 'Picture cropped';
  }

  function startCrop(e, handle) {
    const span = picture.el;
    const info = span && ctx.doc.pictureLayout(span);
    if (!info) return;
    const look = ctx.doc.pictureLookOf(span);
    const c0 = look.crop;
    const fullW = info.cx / PX / (1 - c0.l - c0.r);
    const fullH = info.cy / PX / (1 - c0.t - c0.b);
    const start = { x: e.clientX, y: e.clientY };
    const zoom = pictureZoom(span) * ctx.sheetZoom.value;
    const target = e.currentTarget;
    target.setPointerCapture(e.pointerId);
    let crop = c0;
    const move = (ev) => {
      const dx = (ev.clientX - start.x) / zoom / fullW;
      const dy = (ev.clientY - start.y) / zoom / fullH;
      crop = {
        l: handle.includes('w') ? c0.l + dx : c0.l,
        r: handle.includes('e') ? c0.r - dx : c0.r,
        t: handle.includes('n') ? c0.t + dy : c0.t,
        b: handle.includes('s') ? c0.b - dy : c0.b,
      };
      for (const side of ['l', 'r', 't', 'b']) crop[side] = Math.max(-0.5, Math.min(0.9, crop[side]));
      if (crop.l + crop.r > 0.95 || crop.t + crop.b > 0.95) return;
      const w = fullW * (1 - crop.l - crop.r);
      const h = fullH * (1 - crop.t - crop.b);
      ctx.doc.stylePicture(span, { ...info, cx: w * PX, cy: h * PX }, { ...look, crop });
      updatePictureFrame();
    };
    const up = () => {
      target.removeEventListener('pointermove', move);
      target.removeEventListener('pointerup', up);
      if (crop === c0) return;
      changePicture((s) => ctx.doc.setPictureCrop(s, crop), 'Picture cropped');
      nextTick(() => { picture.cropping = true; });
    };
    target.addEventListener('pointermove', move);
    target.addEventListener('pointerup', up);
  }

  /** Drag the round handle above a picture to rotate it (Shift: in steps of 15°), as in Word. */
  function startRotate(e) {
    const span = picture.el;
    const info = span && ctx.doc.pictureLayout(span);
    if (!info) return;
    const look = ctx.doc.pictureLookOf(span);
    const r = span.getBoundingClientRect();
    const center = { x: r.left + r.width / 2, y: r.top + r.height / 2 };
    const target = e.currentTarget;
    target.setPointerCapture(e.pointerId);
    let rot = look.rot;
    const move = (ev) => {
      let deg = (Math.atan2(ev.clientY - center.y, ev.clientX - center.x) * 180) / Math.PI + 90;
      if (ev.shiftKey) deg = Math.round(deg / 15) * 15;
      rot = ((Math.round(deg) % 360) + 360) % 360;
      ctx.doc.stylePicture(span, info, { ...look, rot });
      updatePictureFrame();
    };
    const up = () => {
      target.removeEventListener('pointermove', move);
      target.removeEventListener('pointerup', up);
      if (rot === look.rot) return;
      changePicture((s) => ctx.doc.setPictureRotation(s, { rot, flipH: look.flipH, flipV: look.flipV }), `Picture rotated to ${rot}°`);
    };
    target.addEventListener('pointermove', move);
    target.addEventListener('pointerup', up);
  }

  function applyCrop() {
    const p = picturePanel;
    const crop = { l: (Number(p.cropL) || 0) / 100, t: (Number(p.cropT) || 0) / 100, r: (Number(p.cropR) || 0) / 100, b: (Number(p.cropB) || 0) / 100 };
    changePicture((s) => ctx.doc.setPictureCrop(s, crop), 'Picture cropped');
  }

  function applyRotation() {
    const p = picturePanel;
    changePicture((s) => ctx.doc.setPictureRotation(s, { rot: Number(p.rot) || 0, flipH: p.flipH, flipV: p.flipV }), `Picture rotated to ${((Math.round(Number(p.rot) || 0) % 360) + 360) % 360}°`);
  }

  function rotateBy(deg) {
    picturePanel.rot = (((Number(picturePanel.rot) || 0) + deg) % 360 + 360) % 360;
    applyRotation();
  }

  function flip(axis) {
    picturePanel[axis] = !picturePanel[axis];
    const p = picturePanel;
    changePicture((s) => ctx.doc.setPictureRotation(s, { rot: Number(p.rot) || 0, flipH: p.flipH, flipV: p.flipV }), axis === 'flipH' ? 'Picture flipped horizontally' : 'Picture flipped vertically');
  }

  function applyEffects() {
    const p = picturePanel;
    const look = {
      border: p.border === 'none' ? null : { color: p.borderColor.replace('#', ''), width: BORDERS[p.border] },
      shadow: p.shadow, glow: p.glow ? 'FFC000' : null, softEdge: p.softEdge, reflection: p.reflection,
      greyscale: p.colour === 'greyscale', brightness: (Number(p.brightness) || 0) / 100, contrast: (Number(p.contrast) || 0) / 100,
    };
    changePicture((s) => ctx.doc.setPictureEffects(s, look), 'Picture effects changed');
  }

  /** A picture style from Word's gallery: its shape, border, effects and 3-D format replace the picture's. */
  function applyStyle() {
    const name = picturePanel.style;
    if (!name) return;
    const style = name === 'none' ? null : PICTURE_STYLES.find((st) => st.name === name);
    changePicture((s) => ctx.doc.setPictureStyle(s, style), style ? `Picture style: ${style.name}` : 'Picture style taken off');
  }

  function applyShape() {
    const label = SHAPES.find(([v]) => v === picturePanel.shape)?.[1] || picturePanel.shape;
    changePicture((s) => ctx.doc.setPictureGeometry(s, picturePanel.shape), `Picture shape: ${label}`);
  }

  /**
   * The 3-D format and rotation from the panel, as Word's 3-D Format and 3-D Rotation set them. The
   * angles (X, Y, Z rotation, perspective) are written when changed (`how` 'angles') or already set;
   * a preset alone stays a preset.
   */
  function apply3D(how) {
    const p = picturePanel;
    const look = ctx.doc.pictureLookOf(picture.el);
    const bevel = p.bevel ? { prst: p.bevel, w: Math.max(0, Number(p.bevelW) || 0) * 12700, h: Math.max(0, Number(p.bevelH) || 0) * 12700 } : null;
    const angles = how === 'angles' || look.rotation3d;
    const rotation = angles ? { x: Number(p.rotX) || 0, y: Number(p.rotY) || 0, z: Number(p.rotZ) || 0, perspective: Math.max(0, Number(p.perspective) || 0) } : null;
    changePicture((s) => ctx.doc.setPicture3D(s, {
      bevel, camera: p.camera || null, material: p.material || null, rotation,
      light: { rig: p.lightRig || 'threePt', dir: look.light?.dir || 't', angle: Number(p.lightAngle) || 0 },
    }), how === 'angles' ? '3-D rotation changed' : '3-D format changed');
  }

  /** A 3-D rotation preset: its angles replace any set by hand. */
  function applyCameraPreset() {
    const p = picturePanel;
    const look = ctx.doc.pictureLookOf(picture.el);
    const bevel = p.bevel ? { prst: p.bevel, w: Math.max(0, Number(p.bevelW) || 0) * 12700, h: Math.max(0, Number(p.bevelH) || 0) * 12700 } : null;
    const label = CAMERAS.find(([v]) => v === p.camera)?.[1];
    changePicture((s) => ctx.doc.setPicture3D(s, {
      bevel, camera: p.camera || null, material: p.material || null,
      light: { rig: p.lightRig || 'threePt', dir: look.light?.dir || 't', angle: Number(p.lightAngle) || 0 },
    }), label ? `3-D rotation: ${label}` : '3-D rotation taken off');
  }

  /** An artistic effect: the picture is drawn with it here (the original kept), or taken off. */
  async function applyArtistic() {
    const span = picture.el;
    const effect = picturePanel.artistic;
    if (!span) return;
    const label = ARTISTIC.find(([v]) => v === effect)?.[1];
    if (!effect) {
      changePicture((s) => ctx.doc.setPictureArtistic(s, null), 'Artistic effect taken off');
      return;
    }
    picturePanel.drawing = true;
    ctx.message.value = `Drawing ${label}…`;
    try {
      const amount = Number(picturePanel.amount) || 0;
      const png = await renderArtistic(ctx.doc.pictureOriginal(span).url, effect, amount);
      changePicture((s) => ctx.doc.setPictureArtistic(s, effect, png, amount), `Artistic effect: ${label}`);
    } catch (err) {
      ctx.message.value = err.message || 'The effect couldn’t be drawn';
    } finally {
      picturePanel.drawing = false;
    }
  }

  function applyAlt() {
    changePicture((s) => ctx.doc.setPictureDescription(s, picturePanel.alt.trim()), picturePanel.alt.trim() ? 'Alt text saved' : 'Alt text removed');
  }

  return { HANDLES, apply3D, applyAlt, applyArtistic, applyCameraPreset, applyCrop, applyEffects, applyPosition, applyRotation, applyShape, applySize, applyStyle, applyWrap, closePicturePanel, flip, layoutOptionsButton, picture, picturePanel, picturePanelEl, resetPictureSize, rotateBy, round, startCrop, startResize, startRotate, toggleCropping, togglePicturePanel, onPictureMousedown, pictureInSelection, selectPicture, showPicture, updatePictureFrame };
}
