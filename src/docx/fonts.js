// Fonts: the ones embedded in a document (word/fonts/*.odttf, listed in fontTable.xml) are shown
// with the document, as Word shows them; fonts it uses that are neither embedded nor on this device
// are named, so the reader knows the page shows another font in their place. Nothing is fetched:
// embedded fonts are read from the file, and whether a font is installed is measured locally.
import { NS, wChildren, wChild, wAttr } from './xml.js';

const W = NS.w;
const EMBEDS = { embedRegular: ['400', 'normal'], embedBold: ['700', 'normal'], embedItalic: ['400', 'italic'], embedBoldItalic: ['700', 'italic'] };

/**
 * Word obfuscates an embedded font (ECMA-376 Part 1, 17.8.1): its first 32 bytes are XORed with
 * the font key, a GUID, read as 16 bytes from its last hex pair to its first. Returns a copy.
 */
export function deobfuscate(bytes, fontKey) {
  const hex = (fontKey || '').replace(/[^0-9a-f]/gi, '');
  if (hex.length !== 32) return bytes;
  const key = [];
  for (let i = 30; i >= 0; i -= 2) key.push(parseInt(hex.slice(i, i + 2), 16));
  const out = new Uint8Array(bytes);
  for (let i = 0; i < 32 && i < out.length; i++) out[i] ^= key[i % 16];
  return out;
}

/**
 * The fonts embedded in a document, loaded for this page: [FontFace]. `part` is fontTable.xml as
 * loaded by loadPart; `read(path)` gives a file's bytes. A font that can't be read is skipped.
 */
export async function loadEmbeddedFonts(part, read) {
  if (!part || typeof FontFace === 'undefined') return [];
  const faces = [];
  for (const font of wChildren(part.xml.documentElement, 'font')) {
    const name = wAttr(font, 'name');
    if (!name) continue;
    for (const [local, [weight, style]] of Object.entries(EMBEDS)) {
      const embed = wChild(font, local);
      const rel = embed && part.rels.get(embed.getAttributeNS(NS.r, 'id'));
      if (!rel?.path) continue;
      try {
        const bytes = await read(rel.path);
        if (!bytes) continue;
        const data = /\.odttf$/i.test(rel.path) || wAttr(embed, 'fontKey') ? deobfuscate(bytes, wAttr(embed, 'fontKey')) : bytes;
        const face = new FontFace(name, data, { weight, style });
        await face.load();
        document.fonts.add(face);
        faces.push(face);
      } catch {
        // Not a font the browser can use (or a subset it rejects): the page shows another.
      }
    }
  }
  return faces;
}

const GENERIC = new Set(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui', 'math', 'emoji', 'inherit', 'initial']);
let canvas = null;

/** True if the font `name` is on this device (or loaded for the page), measured on a canvas. */
export function fontAvailable(name) {
  if (!name || GENERIC.has(name.toLowerCase())) return true;
  for (const face of document.fonts || []) if (face.family.replace(/^["']|["']$/g, '') === name && face.status === 'loaded') return true;
  canvas ||= document.createElement('canvas');
  const g = canvas.getContext('2d');
  if (!g) return true;
  const sample = 'mmmmmmmmmmlli1WQ@#&ÆgjpqyÐ';
  const quoted = `"${name.replace(/["\\]/g, '')}"`;
  for (const base of ['monospace', 'serif', 'sans-serif']) {
    g.font = `72px ${base}`;
    const without = g.measureText(sample).width;
    g.font = `72px ${quoted}, ${base}`;
    if (g.measureText(sample).width !== without) return true;
  }
  return false;
}

/** The first font named in a CSS font-family list, unquoted ("Playfair Display"). */
function firstFamily(list) {
  const first = (list || '').split(',')[0]?.trim() || '';
  return first.replace(/^["']|["']$/g, '');
}

/**
 * The fonts the text in `root` is shown in that aren't on this device: [name], sorted. Each is
 * what the document asks for; the browser shows another font in its place.
 */
export function missingFonts(root) {
  const families = new Set();
  const seen = new Set();
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let t = walker.nextNode(); t; t = walker.nextNode()) {
    const el = t.parentElement;
    if (!el || seen.has(el) || !t.data.trim()) continue;
    seen.add(el);
    const family = firstFamily(getComputedStyle(el).fontFamily);
    if (family) families.add(family);
  }
  return [...families].filter((f) => !fontAvailable(f)).sort((a, b) => a.localeCompare(b));
}
