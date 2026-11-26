// Offline editorial helpers for the Content Creator workspace. These are deliberately
// deterministic checks and transformations: no document text leaves this device.
const WORD = /[\\p{L}\\p{N}]+(?:['’][\\p{L}\\p{N}]+)*/gu;
const SENTENCE = /[^.!?。！？]+[.!?。！？]?/gu;

export function wordsIn(text = '') {
  return (String(text).match(WORD) || []).length;
}

export function editorialChecks(text = '') {
  const source = String(text).replace(/\\s+/g, ' ').trim();
  const issues = [];
  const tokens = source.match(WORD) || [];
  const lower = tokens.map((w) => w.toLocaleLowerCase());
  for (let i = 1; i < lower.length; i++) {
    if (lower[i] === lower[i - 1] && !['very'].includes(lower[i])) {
      issues.push({ type: 'Repeated word', detail: `“${tokens[i]}” appears twice in a row.`, excerpt: tokens.slice(Math.max(0, i - 3), i + 3).join(' ') });
    }
  }
  const sentences = source.match(SENTENCE) || [];
  sentences.forEach((sentence) => {
    const count = wordsIn(sentence);
    if (count > 30) issues.push({ type: 'Long sentence', detail: `This sentence has ${count} words; consider splitting it.`, excerpt: sentence.trim() });
  });
  const paragraphs = String(text).split(/\\n\\s*\\n/).map((p) => p.trim()).filter(Boolean);
  paragraphs.forEach((paragraph) => {
    const count = wordsIn(paragraph);
    if (count > 150) issues.push({ type: 'Long paragraph', detail: `This paragraph has ${count} words; consider adding a break.`, excerpt: paragraph.slice(0, 180) + (paragraph.length > 180 ? '…' : '') });
  });
  const counts = new Map();
  lower.filter((w) => w.length > 4).forEach((w) => counts.set(w, (counts.get(w) || 0) + 1));
  [...counts].filter(([, count]) => count >= 6).sort((a, b) => b[1] - a[1]).slice(0, 5).forEach(([word, count]) => {
    issues.push({ type: 'Frequently used word', detail: `“${word}” appears ${count} times. Check whether repetition is intentional.`, excerpt: '' });
  });
  return { words: tokens.length, paragraphs: paragraphs.length, sentences: sentences.length, issues };
}

export function seoChecks({ text = '', title = '', description = '', keyword = '' } = {}) {
  const body = String(text);
  const normalized = (value) => String(value).toLocaleLowerCase();
  const key = normalized(keyword).trim();
  const firstWords = body.split(/\\s+/).slice(0, 100).join(' ');
  const headings = body.split('\\n').filter((line) => /^\\s{0,3}#{1,6}\\s+|^\\s*\\S.{0,80}$/.test(line));
  const links = /https?:\\/\\//i.test(body);
  const checks = [
    { label: 'SEO title length', ok: title.length >= 30 && title.length <= 60, detail: `${title.length} characters (recommended editorial range: 30–60).` },
    { label: 'Meta description length', ok: description.length >= 120 && description.length <= 160, detail: `${description.length} characters (recommended editorial range: 120–160).` },
    { label: 'Focus keyword provided', ok: !!key, detail: key ? `Focus keyword: “${keyword}”.` : 'Add a focus keyword to check its placement.' },
    { label: 'Keyword in title', ok: !!key && normalized(title).includes(key), detail: 'Include the exact focus keyword naturally in the title.' },
    { label: 'Keyword near the beginning', ok: !!key && normalized(firstWords).includes(key), detail: 'Use the focus keyword naturally in the opening 100 words.' },
    { label: 'Descriptive structure', ok: headings.length >= 2, detail: `${headings.length} possible headings detected in the supplied text.` },
    { label: 'Links included', ok: links, detail: 'Consider adding useful internal or authoritative external links.' },
  ];
  const occurrences = key ? (normalized(body).split(key).length - 1) : 0;
  return { checks, occurrences, note: 'These are editorial heuristics, not a search-ranking prediction. Search engines may use different lengths and signals.' };
}

function firstSentences(text, limit = 2) {
  return (String(text).trim().match(SENTENCE) || []).map((s) => s.trim()).filter(Boolean).slice(0, limit).join(' ');
}

export function repurpose(text = '', title = '') {
  const source = String(text).trim();
  const heading = String(title).trim() || 'A thought worth sharing';
  const opening = firstSentences(source, 2) || source.slice(0, 240);
  const paragraph = source.split(/\\n\\s*\\n/).map((p) => p.trim()).filter(Boolean)[0] || source;
  const compact = paragraph.replace(/\\s+/g, ' ').trim();
  const xPost = `${heading}: ${opening}`;
  return {
    linkedin: `${heading}\\n\\n${opening}\\n\\n${compact.length > opening.length ? compact.slice(opening.length).trim() : 'What is your perspective?'}\\n\\nWhat would you add to this conversation?`,
    x: xPost.length <= 280 ? xPost : `${xPost.slice(0, 277).trimEnd()}…`,
    instagram: `${heading}\\n\\n${opening}\\n\\nWhat do you think? Share your perspective below.\\n\\n#${heading.split(/\\s+/).filter(Boolean).slice(0, 3).join('')}`,
    newsletter: `${heading}\\n\\nIn this edition: ${opening}`,
  };
}
