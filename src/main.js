import { createApp } from 'vue';
import App from './App.vue';
import router from './router.js';
import './style.css';

createApp(App).use(router).mount('#app');

// A version of the editor kept a content brief, writing goals and a daily word log in this browser,
// for good and for every document (intact.creator.*). They are gone from the editor, so they go from
// the browser too: on a shared or borrowed computer, nothing of them is left for the next person.
try {
  for (const key of Object.keys(localStorage)) if (key.startsWith('intact.creator.')) localStorage.removeItem(key);
} catch { /* storage unavailable: nothing was kept */ }

// A file dropped where nothing takes it would make the browser open or download it in place of the
// app, and an open document's edits with it. Places that take files (the start page, the editable
// page) say so with data-file-drop or by being editable; everywhere else a file can't be dropped.
// Text and other drags are left to the browser.
const takesFiles = (e) => e.target instanceof Element && e.target.closest('[data-file-drop], [contenteditable="true"]');
const draggingFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
window.addEventListener('dragover', (e) => {
  if (!draggingFiles(e) || takesFiles(e)) return;
  e.preventDefault();
  e.dataTransfer.dropEffect = 'none';
});
window.addEventListener('drop', (e) => {
  if (draggingFiles(e) && !e.defaultPrevented) e.preventDefault();
});

// Offline use, once opened online (production builds only; see src/sw-template.js). A browser that
// can't (no service workers, or not over HTTPS) works as before, online.
if (import.meta.env.PROD && 'serviceWorker' in navigator) {
  window.addEventListener('load', () => { navigator.serviceWorker.register('./sw.js').catch(() => {}); });
}
