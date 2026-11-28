// scripts/check-headers.mjs, which checks a self-hosted Intact's response headers: against the
// built app served as README's Self-hosting recommends (every check passes), and served as a plain
// static server would serve it (it says what is missing, and fails).
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { test, expect } from '@playwright/test';

const DIST = path.resolve('dist');
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.webmanifest': 'application/manifest+json', '.svg': 'image/svg+xml' };

/** Serves dist/ on a free port, with `headers` added to every response; resolves to its URL. */
function serve(headers, { manifestType = true } = {}) {
  const server = http.createServer((req, res) => {
    const file = path.join(DIST, decodeURIComponent(new URL(req.url, 'http://x').pathname.replace(/^\/intact\//, '/')).replace(/\/$/, '/index.html'));
    if (!file.startsWith(DIST) || !fs.existsSync(file)) { res.writeHead(404); res.end(); return; }
    const ext = path.extname(file);
    res.writeHead(200, { ...(ext !== '.webmanifest' || manifestType ? { 'Content-Type': TYPES[ext] || 'application/octet-stream' } : {}), ...headers });
    res.end(fs.readFileSync(file));
  });
  return new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve({ url: `http://127.0.0.1:${server.address().port}/intact/`, close: () => server.close() })));
}

const run = async (url) => {
  try {
    const { stdout } = await promisify(execFile)('node', ['scripts/check-headers.mjs', url]);
    return { code: 0, out: stdout };
  } catch (e) {
    return { code: e.code, out: e.stdout };
  }
};

test.describe.configure({ mode: 'serial' });
test.skip(({ browserName }) => browserName !== 'chromium', 'a Node script: once is enough');

test('served as README recommends, every check passes', async () => {
  const csp = /<meta[^>]+http-equiv="Content-Security-Policy"[^>]+content="([^"]+)"/.exec(fs.readFileSync(path.join(DIST, 'index.html'), 'utf8'))[1];
  const site = await serve({ 'Content-Security-Policy': `${csp}; frame-ancestors 'none'`, 'Referrer-Policy': 'no-referrer', 'X-Content-Type-Options': 'nosniff' });
  try {
    const { code, out } = await run(site.url);
    expect(out).not.toContain('FAIL');
    expect(out).toContain('ok    frame-ancestors');
    expect(code).toBe(0);
  } finally { site.close(); }
});

test('served as a plain static server would, it says what is missing, and fails', async () => {
  const site = await serve({ 'Content-Security-Policy': "default-src 'none'; script-src 'self' 'unsafe-inline'" }, { manifestType: false });
  try {
    const { code, out } = await run(site.url);
    expect(code).toBe(1);
    expect(out).toMatch(/FAIL {2}The header is as strict as the page’s own policy – missing or wider: script-src, style-src/);
    expect(out).toContain('FAIL  frame-ancestors – missing');
    expect(out).toContain('FAIL  X-Content-Type-Options: nosniff – missing');
    expect(out).toContain('FAIL  manifest.webmanifest as application/manifest+json');
  } finally { site.close(); }
});
