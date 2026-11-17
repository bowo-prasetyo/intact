import { createApp } from 'vue';
import App from './App.vue';
import router from './router.js';
import './style.css';

createApp(App).use(router).mount('#app');

// Offline use, once opened online (production builds only; see src/sw-template.js). A browser that
// can't (no service workers, or not over HTTPS) works as before, online.
if (import.meta.env.PROD && 'serviceWorker' in navigator) {
  window.addEventListener('load', () => { navigator.serviceWorker.register('./sw.js').catch(() => {}); });
}
