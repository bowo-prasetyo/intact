// The interface's language: English, or Indonesian (Bahasa Indonesia). Chosen on the start page and
// remembered in this browser; else the browser's own language decides (Indonesian for "id", English
// for any other). Only the interface is translated: a document's own text, styles and fields are
// never changed by it.
//
// Strings are looked up by their English text, so a string not (yet) translated shows in English,
// as it did before: t('Page setup') is "Pengaturan halaman" in Indonesian. Placeholders in braces
// are filled in from `vars`: t('Section {n}', { n: 2 }). A language's strings are loaded when it is
// chosen (a file of the app's own, kept for offline use like the rest), so English costs nothing.
import { ref, shallowRef, watch } from 'vue';

export const LANGUAGES = [
  { code: 'en', name: 'English' },
  { code: 'id', name: 'Bahasa Indonesia' },
];
const LOADERS = { id: () => import('./locales/id.js') };
const dictionaries = shallowRef({}); // code -> its strings, once loaded
const KEY = 'intact.lang';

/** The language the browser asks for: 'id' for Indonesian, else 'en'. */
export function browserLanguage(language = (typeof navigator !== 'undefined' && navigator.language) || 'en') {
  return /^(id|in)\b/i.test(language) ? 'id' : 'en';
}

function rememberedLanguage() {
  try {
    const v = localStorage.getItem(KEY);
    return LANGUAGES.some((l) => l.code === v) ? v : null;
  } catch {
    return null;
  }
}

/** The interface's language now ('en' or 'id'), reactive. */
export const lang = ref(rememberedLanguage() || browserLanguage());

/** Chooses the interface's language, remembered in this browser. */
export function setLanguage(code) {
  if (!LANGUAGES.some((l) => l.code === code)) return;
  lang.value = code;
  try { localStorage.setItem(KEY, code); } catch { /* this session only */ }
}

// The page's language follows the interface's, for screen readers and the browser's own tools; the
// document's page says its own language (see EditorView).
if (typeof document !== 'undefined') {
  watch(lang, (code) => { document.documentElement.lang = code; }, { immediate: true });
}
watch(lang, async (code) => {
  if (!LOADERS[code] || dictionaries.value[code]) return;
  const strings = (await LOADERS[code]()).default;
  dictionaries.value = { ...dictionaries.value, [code]: strings };
}, { immediate: true });

/** `english` in the interface's language, with {name} placeholders filled in from `vars`. */
export function t(english, vars) {
  const dict = dictionaries.value[lang.value];
  const text = (dict && dict[english]) || english;
  return vars ? text.replace(/\{(\w+)\}/g, (m, k) => (k in vars ? String(vars[k]) : m)) : text;
}
