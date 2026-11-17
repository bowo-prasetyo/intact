import { readFileSync, readdirSync, statSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join, relative } from 'node:path';
import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';

const pkg = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8'));

// The Content Security Policy of the built app: the "Private" promise, enforced by the browser.
// Nothing but the app's own files loads, and the page can't connect anywhere (connect-src 'none'):
// a document opened here can't be sent off, by the app or by anything injected into it. Pictures and
// fonts come from the document as blob: or data: URLs. Styles are inline (the document's formatting
// is applied as style attributes and a generated stylesheet), so 'unsafe-inline' is needed for
// styles, never for scripts. Self-hosters can send it as a header too, with frame-ancestors (which
// a <meta> can't set): see "Self-hosting" in the README.
export const CSP = [
  "default-src 'none'",
  "script-src 'self'",
  "style-src 'self' 'unsafe-inline'",
  "img-src 'self' blob: data:",
  "font-src 'self' blob: data:",
  "media-src 'self' blob: data:",
  "connect-src 'none'",
  "worker-src 'self'",
  "manifest-src 'self'",
  "base-uri 'none'",
  "form-action 'none'",
  "object-src 'none'",
  "frame-src 'none'",
].join('; ');

/** Every file under `dir`, as paths relative to it. */
function filesIn(dir) {
  const out = [];
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) out.push(...filesIn(path).map((p) => `${name}/${p}`));
    else out.push(relative(dir, path));
  }
  return out;
}

/**
 * Production only: puts the CSP in index.html, and writes sw.js, the service worker that keeps the
 * app's own files for offline use (see src/sw-template.js). The dev server needs inline scripts
 * and a websocket for hot reloading, so it runs without either.
 */
function hardening() {
  return {
    name: 'intact-hardening',
    apply: 'build',
    transformIndexHtml(html) {
      return html.replace('<meta charset="UTF-8" />', `<meta charset="UTF-8" />\n    <meta http-equiv="Content-Security-Policy" content="${CSP}" />`);
    },
    generateBundle(_, bundle) {
      const publicFiles = filesIn('public');
      const files = ['./', ...publicFiles, ...Object.keys(bundle)].filter((f) => f !== 'sw.js').sort();
      const version = createHash('sha256').update(`${pkg.version}\n${files.join('\n')}`).digest('hex').slice(0, 12);
      const template = readFileSync('src/sw-template.js', 'utf8');
      this.emitFile({
        type: 'asset',
        fileName: 'sw.js',
        source: template.replace('__VERSION__', version).replace('__FILES__', JSON.stringify(files.map((f) => (f === './' ? f : `./${f}`)))),
      });
    },
  };
}

// `base: './'` makes the build work from any GitHub Pages sub-path (https://<user>.github.io/<repo>/).
export default defineConfig({
  base: './',
  plugins: [vue(), hardening()],
  define: { __APP_VERSION__: JSON.stringify(pkg.version) },
});
