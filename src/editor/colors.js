// Colours given by hex code ("More colours…"), for font colour and cell shading: a business's own
// brand colour, typed or picked once and then offered again, in any document, from this browser.

const KEY = 'intact.customColors';
const MAX = 10;

/** "#0E6467", "0e6467", "#0e6" or "0E6" as Word's RRGGBB ("0E6467"), or null if it isn't one. */
export function parseHex(text) {
  const t = String(text || '').trim().replace(/^#/, '');
  if (/^[0-9a-f]{6}$/i.test(t)) return t.toUpperCase();
  if (/^[0-9a-f]{3}$/i.test(t)) return t.replace(/./g, (c) => c + c).toUpperCase();
  return null;
}

/** The colours chosen by code in this browser, most recent first: ["0E6467", …]. */
export function rememberedColors() {
  try {
    const list = JSON.parse(localStorage.getItem(KEY) || '[]');
    return Array.isArray(list) ? list.map(parseHex).filter(Boolean).slice(0, MAX) : [];
  } catch {
    return [];
  }
}

/** Remembers a colour chosen by code (first in the list); returns the list. */
export function rememberColor(hex) {
  const list = [hex, ...rememberedColors().filter((c) => c !== hex)].slice(0, MAX);
  try { localStorage.setItem(KEY, JSON.stringify(list)); } catch { /* this session only */ }
  return list;
}
