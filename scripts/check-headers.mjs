#!/usr/bin/env node
// Checks the response headers of a self-hosted Intact, as README's Self-hosting recommends them:
//   node scripts/check-headers.mjs https://docs.example.org/intact/
// The page carries its Content Security Policy in a <meta> tag; a header can add what a <meta>
// can't (frame-ancestors, so no other site can frame it) and must be at least as strict. Each check
// prints ok or what is wrong; the exit code is 1 if any fails. It only reads: it sends nothing but
// the requests for the page and its manifest.
const base = process.argv[2];
if (!base) {
  console.error('Usage: node scripts/check-headers.mjs <the deployed app’s URL>');
  process.exit(2);
}

/** A policy as { directive: [sources] }. */
const parse = (policy) => Object.fromEntries((policy || '').split(';').map((d) => d.trim().split(/\s+/)).filter(([n]) => n).map(([n, ...v]) => [n.toLowerCase(), v]));

export async function checkHeaders(url) {
  const results = [];
  const check = (name, ok, detail) => results.push({ name, ok, detail });
  const res = await fetch(url, { redirect: 'follow' });
  const html = await res.text();
  const h = res.headers;
  check('The page loads', res.ok, `HTTP ${res.status}`);

  const tag = /<meta[^>]+http-equiv=["']Content-Security-Policy["'][^>]*>/i.exec(html)?.[0] || '';
  const raw = /content="([^"]*)"|content='([^']*)'/i.exec(tag);
  const meta = (raw?.[1] ?? raw?.[2])?.replace(/&#39;|&apos;/g, "'").replace(/&quot;/g, '"').replace(/&amp;/g, '&');
  const header = h.get('content-security-policy');
  check('Content-Security-Policy header', !!header, header ? 'sent' : 'missing: the <meta> policy applies, but nothing stops another site framing the app');
  if (header && meta) {
    const want = parse(meta);
    const got = parse(header);
    // Every directive of the page's own policy, with no source it doesn't allow.
    const weaker = Object.entries(want).filter(([d, sources]) => !got[d] || got[d].some((s) => !sources.includes(s))).map(([d]) => d);
    check('The header is as strict as the page’s own policy', !weaker.length, weaker.length ? `missing or wider: ${weaker.join(', ')}` : 'every directive');
  }
  const frames = header && parse(header)['frame-ancestors'];
  check('frame-ancestors', !!frames && frames.every((s) => ["'none'", "'self'"].includes(s)), frames ? frames.join(' ') : 'missing: another site could frame the app');
  check('X-Content-Type-Options: nosniff', h.get('x-content-type-options')?.toLowerCase() === 'nosniff', h.get('x-content-type-options') || 'missing');
  const referrer = h.get('referrer-policy');
  check('Referrer-Policy: no-referrer', referrer?.toLowerCase().split(/\s*,\s*/).includes('no-referrer'), referrer || 'missing (the page’s <meta name="referrer"> still applies)');

  const manifest = await fetch(new URL('manifest.webmanifest', res.url));
  const type = manifest.headers.get('content-type') || '';
  check('manifest.webmanifest as application/manifest+json', manifest.ok && /application\/manifest\+json/.test(type), manifest.ok ? type || 'no type' : `HTTP ${manifest.status}`);
  return results;
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const results = await checkHeaders(base);
  for (const r of results) console.log(`${r.ok ? 'ok  ' : 'FAIL'}  ${r.name}${r.detail ? ` – ${r.detail}` : ''}`);
  process.exit(results.every((r) => r.ok) ? 0 : 1);
}
