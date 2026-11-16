// 3-D lighting of a picture's bevel, worked out per pixel as a 3-D renderer does: the bevel's
// profile (Word's presets: circle, angle, soft round…) gives the surface's height at each point near
// the edge, the height's slope gives the surface's normal, and the light rig (Word's: three point,
// harsh, sunset…; from its direction and angle) and the material (matte, plastic, metal…) give how
// bright each point is, by Lambert's and Phong's laws. Drawn as a layer over the picture: light
// where the bevel faces the lights, shade where it faces away, nothing on the flat face.

const EMU_PER_PX = 9525;
const MAX = 320; // the layer is worked out at most this many pixels across, then scaled
const cache = new Map();
const CONTRAST = 2.6; // how strongly light and shade show over the picture, as Word's bevels read

const clamp = (v, a = 0, b = 1) => Math.max(a, Math.min(b, v));
const smooth = (t) => t * t * (3 - 2 * t);

/** A bevel's profile: height (0–1) and its slope at `t`, the distance in from the edge (0–1 of the bevel's width). */
const PROFILES = {
  circle: (t) => [Math.sqrt(1 - (1 - t) ** 2), (1 - t) / Math.max(0.05, Math.sqrt(1 - (1 - t) ** 2))],
  relaxedInset: (t) => [1 - (1 - t) ** 2, 2 * (1 - t)],
  cross: (t) => (t < 0.5 ? [2 * t, 2] : [2 - 2 * t, -2]),
  coolSlant: (t) => [t, 1.4],
  angle: (t) => [t, 1],
  softRound: (t) => [smooth(t), 6 * t * (1 - t)],
  convex: (t) => [Math.sin((t * Math.PI) / 2), (Math.PI / 2) * Math.cos((t * Math.PI) / 2)],
  slope: (t) => [t, 1],
  divot: (t) => [Math.abs(Math.sin(t * Math.PI)) * 0.5 + t * 0.5, Math.cos(t * Math.PI) * 1.5 + 0.5],
  riblet: (t) => [t + Math.sin(t * Math.PI * 4) * 0.1, 1 + Math.cos(t * Math.PI * 4) * 1.2],
  hardEdge: (t) => (t < 0.25 ? [t * 4, 4] : [1, 0]),
  artDeco: (t) => [Math.floor(t * 4) / 4, (t * 4) % 1 < 0.25 ? 4 : 0],
};

/** Word's light rigs: lights as [azimuth offset from the rig's direction (deg), elevation (deg), strength], ambient light, tint. */
const RIGS = {
  threePt: { lights: [[0, 40, 0.8], [90, 50, 0.25], [-90, 50, 0.15]], ambient: 0.35 }, // a key light, fill lights at the sides
  balanced: { lights: [[0, 45, 0.55], [180, 45, 0.45]], ambient: 0.4 },
  soft: { lights: [[0, 55, 0.5]], ambient: 0.5 },
  harsh: { lights: [[0, 30, 1.05]], ambient: 0.15 },
  flood: { lights: [[0, 70, 0.8]], ambient: 0.35 },
  contrasting: { lights: [[0, 25, 1.1]], ambient: 0.1 },
  morning: { lights: [[0, 35, 0.8]], ambient: 0.35, tint: [255, 236, 210] },
  sunrise: { lights: [[0, 20, 0.85]], ambient: 0.3, tint: [255, 214, 170] },
  sunset: { lights: [[0, 15, 0.85]], ambient: 0.3, tint: [255, 190, 150] },
  chilly: { lights: [[0, 40, 0.75]], ambient: 0.35, tint: [214, 230, 255] },
  freezing: { lights: [[0, 40, 0.8]], ambient: 0.3, tint: [190, 215, 255] },
  twoPt: { lights: [[-60, 35, 0.6], [60, 35, 0.6]], ambient: 0.3 },
  glow: { lights: [[0, 60, 0.5]], ambient: 0.6, tint: [255, 255, 240] },
  brightRoom: { lights: [[0, 60, 0.45]], ambient: 0.7 },
};

/** Word's materials: diffuse and specular strength, shininess, and darkening toward the edge. */
const MATERIALS = {
  matte: { kd: 1, ks: 0, n: 1 },
  warmMatte: { kd: 1, ks: 0.05, n: 4, tint: [255, 240, 220] },
  plastic: { kd: 0.9, ks: 0.55, n: 24 },
  metal: { kd: 0.7, ks: 0.9, n: 40 },
  dkEdge: { kd: 1, ks: 0.2, n: 10, edge: 0.5 },
  softEdge: { kd: 0.8, ks: 0.15, n: 8, soft: true },
  powder: { kd: 0.9, ks: 0.05, n: 2 },
  translucentPowder: { kd: 0.6, ks: 0.05, n: 2 },
  clear: { kd: 0.5, ks: 0.7, n: 30 },
  softmetal: { kd: 0.8, ks: 0.6, n: 18 },
  legacyWireframe: { kd: 0.4, ks: 0, n: 1 },
};

/** Where light from a rig's direction comes from, as a screen angle (deg; 0 = from the right, 90 = from below). */
const DIRS = { r: 0, br: 45, b: 90, bl: 135, l: 180, tl: -135, t: -90, tr: -45 };

function unit([x, y, z]) {
  const l = Math.hypot(x, y, z) || 1;
  return [x / l, y / l, z / l];
}

/**
 * The lighting layer for a picture of `w` × `h` px with a bevel ({ prst, w, h } in EMU), on its shape
 * (`geometry`), lit by `light` ({ rig, dir, angle }) in `material`: an image address (PNG), or null
 * when nothing shows (no bevel, or flat lighting or material).
 */
export function lightingLayer({ w, h, geometry, bevel, material, light }) {
  if (!bevel || !w || !h) return null;
  const rigName = light?.rig || 'threePt';
  if (rigName === 'flat' || /^legacyFlat/.test(rigName) || material === 'flat') return null;
  const key = JSON.stringify([Math.round(w), Math.round(h), geometry, bevel, material, light]);
  if (cache.has(key)) return cache.get(key);
  const scale = Math.min(1, MAX / Math.max(w, h));
  const W = Math.max(2, Math.round(w * scale));
  const H = Math.max(2, Math.round(h * scale));
  const bw = Math.max(1, (bevel.w || 76200) / EMU_PER_PX) * scale; // the bevel's width, in layer pixels
  const depth = ((bevel.h || 25400) / EMU_PER_PX) * scale; // its height
  const profile = PROFILES[bevel.prst] || PROFILES.circle;
  const rig = RIGS[rigName] || RIGS.threePt;
  const mat = MATERIALS[material] || MATERIALS.matte;
  const base = (DIRS[light?.dir] ?? -90) + (light?.angle || 0);
  const lights = rig.lights.map(([az, el, k]) => {
    const a = ((base + az) * Math.PI) / 180;
    const e = (el * Math.PI) / 180;
    return { L: unit([Math.cos(a) * Math.cos(e), Math.sin(a) * Math.cos(e), Math.sin(e)]), k };
  });
  const shade = (N) => {
    let d = rig.ambient;
    let s = 0;
    for (const { L, k } of lights) {
      const nl = N[0] * L[0] + N[1] * L[1] + N[2] * L[2];
      d += mat.kd * k * Math.max(0, nl);
      if (mat.ks && nl > 0) {
        const rz = 2 * nl * N[2] - L[2]; // the reflected light, towards the viewer (0, 0, 1)
        s += mat.ks * k * Math.max(0, rz) ** mat.n;
      }
    }
    return { d, s };
  };
  const flat = shade([0, 0, 1]);
  const flatI = flat.d + flat.s;
  const tint = rig.tint || mat.tint || [255, 255, 255];
  const canvas = document.createElement('canvas');
  canvas.width = W;
  canvas.height = H;
  const g = canvas.getContext('2d');
  const img = g.createImageData(W, H);
  const p = img.data;
  const oval = geometry === 'ellipse';
  const a = W / 2;
  const b = H / 2;
  for (let y = 0; y < H; y++) {
    for (let x = 0; x < W; x++) {
      // Distance in from the edge, and the direction inwards.
      let dist;
      let gx;
      let gy;
      if (oval) {
        const nx = (x + 0.5 - a) / a;
        const ny = (y + 0.5 - b) / b;
        const r = Math.hypot(nx, ny) || 1e-6;
        dist = (1 - r) * Math.min(a, b);
        gx = -nx / r;
        gy = -ny / r;
      } else {
        const d = [x + 0.5, W - x - 0.5, y + 0.5, H - y - 0.5];
        const m = Math.min(...d);
        dist = m;
        [gx, gy] = [[1, 0], [-1, 0], [0, 1], [0, -1]][d.indexOf(m)];
      }
      if (dist <= 0 || dist >= bw) continue;
      const t = clamp(dist / bw);
      const [, slope] = profile(t);
      const k = (slope * depth) / bw;
      const N = unit([-gx * k, -gy * k, 1]);
      const { d, s } = shade(N);
      let delta = d + s - flatI;
      if (mat.edge) delta -= mat.edge * (1 - t) * 0.6;
      if (mat.soft) delta *= 0.6;
      const o = (y * W + x) * 4;
      if (delta > 0) {
        p[o] = tint[0];
        p[o + 1] = tint[1];
        p[o + 2] = tint[2];
        p[o + 3] = Math.round(clamp(delta * CONTRAST) * 235);
      } else {
        p[o] = 0;
        p[o + 1] = 0;
        p[o + 2] = 0;
        p[o + 3] = Math.round(clamp(-delta * CONTRAST) * 215);
      }
    }
  }
  g.putImageData(img, 0, 0);
  const url = canvas.toDataURL('image/png');
  cache.set(key, url);
  return url;
}

/**
 * Word's 3-D rotation presets as angles (X and Y rotation, Z rotation, perspective field of view;
 * degrees), as near as the presets look in Word.
 */
export const CAMERA_ANGLES = {
  orthographicFront: [0, 0, 0, 0],
  isometricTopUp: [45, 325, 0, 0],
  isometricTopDown: [315, 325, 0, 0],
  isometricBottomUp: [45, 35, 0, 0],
  isometricBottomDown: [315, 35, 0, 0],
  isometricLeftUp: [45, 0, 330, 0],
  isometricLeftDown: [45, 0, 30, 0],
  isometricRightUp: [315, 0, 30, 0],
  isometricRightDown: [315, 0, 330, 0],
  isometricOffAxis1Left: [64, 342, 0, 0],
  isometricOffAxis1Right: [296, 342, 0, 0],
  isometricOffAxis1Top: [314, 302, 324, 0],
  isometricOffAxis2Left: [26, 342, 0, 0],
  isometricOffAxis2Right: [334, 342, 0, 0],
  isometricOffAxis2Top: [46, 302, 36, 0],
  perspectiveFront: [0, 0, 0, 45],
  perspectiveLeft: [20, 0, 0, 45],
  perspectiveRight: [340, 0, 0, 45],
  perspectiveAbove: [0, 340, 0, 45],
  perspectiveBelow: [0, 20, 0, 45],
  perspectiveAboveLeftFacing: [21, 345, 0, 45],
  perspectiveAboveRightFacing: [339, 345, 0, 45],
  perspectiveContrastingLeftFacing: [44, 10, 356, 45],
  perspectiveContrastingRightFacing: [316, 10, 4, 45],
  perspectiveHeroicLeftFacing: [34, 8, 320, 80],
  perspectiveHeroicRightFacing: [326, 8, 40, 80],
  perspectiveRelaxed: [0, 309, 0, 45],
  perspectiveRelaxedModerately: [0, 324, 0, 45],
};

/**
 * The CSS 3-D transform for a picture's camera: its rotation by angles (Word's X, Y and Z rotation)
 * and perspective; from the file's own angles when given, else from its preset.
 */
export function cameraTransform({ camera, rotation3d }, w, h) {
  const preset = CAMERA_ANGLES[camera] || (camera ? [0, 0, 0, 20] : [0, 0, 0, 0]);
  const [x, y, z, fov] = rotation3d
    ? [rotation3d.x, rotation3d.y, rotation3d.z, rotation3d.perspective || (camera && CAMERA_ANGLES[camera] ? CAMERA_ANGLES[camera][3] : 0)]
    : preset;
  const signed = (v) => (((v % 360) + 540) % 360) - 180;
  const parts = [];
  if (fov) parts.push(`perspective(${Math.round(Math.max(w, h) / Math.tan(((fov / 2) * Math.PI) / 180))}px)`);
  if (signed(y)) parts.push(`rotateX(${-signed(y)}deg)`);
  if (signed(x)) parts.push(`rotateY(${signed(x)}deg)`);
  if (signed(z)) parts.push(`rotateZ(${signed(z)}deg)`);
  return parts.join(' ');
}
